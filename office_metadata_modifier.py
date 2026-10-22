#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Office 元数据修改工具：支持 .docx、.pptx、.xlsx（仅现代 OOXML 格式）。

仅依赖 Python 标准库。GUI 模式保留原有 Tkinter 操作风格；命令行模式适合批处理。
修改前自动创建同目录备份，使用临时文件与平台适配的原子替换，并在写后重新读取验证。
"""

from __future__ import annotations

import argparse
import csv
import ctypes
import datetime as _dt
import glob
import hashlib
import json
import os
import posixpath
import shutil
import sys
import tempfile
import tkinter as tk
import zipfile
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Any, Dict, Iterable, List, Optional, Tuple, Union
import xml.etree.ElementTree as ET


# ============================================================
# OOXML 命名空间与文件类型
# ============================================================

SUPPORTED_EXTENSIONS = {".docx", ".pptx", ".xlsx"}
# Bound pathological ZIP packages before streaming each member through CRC validation.
MAX_ZIP_ENTRIES = 100_000
MAX_ZIP_UNCOMPRESSED_SIZE = 2 * 1024 * 1024 * 1024  # 2 GiB total uncompressed size
ZIP_VALIDATION_CHUNK_SIZE = 4 * 1024 * 1024  # 4 MiB: fewer Python read iterations for large OOXML packages
PRESET_DIRECTORY = os.path.join(os.path.expanduser("~"), ".office_metadata_modifier")
PRESET_FILE = os.path.join(PRESET_DIRECTORY, "presets.json")
UTC_PLUS_8 = _dt.timezone(_dt.timedelta(hours=8))

CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
CORE_REL_TYPE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
APP_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
CORE_CONTENT_TYPE = "application/vnd.openxmlformats-package.core-properties+xml"
APP_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.extended-properties+xml"

# ElementTree 可使用任意合法 XML 前缀；为常见 Office 元数据注册稳定前缀。
for _prefix, _uri in (
    ("cp", CP_NS), ("dc", DC_NS), ("dcterms", DCTERMS_NS),
    ("xsi", XSI_NS), ("ep", APP_NS), ("vt", VT_NS),
    ("rel", PKG_REL_NS), ("ct", CT_NS),
):
    ET.register_namespace(_prefix, _uri)

CORE_FIELDS = {
    "author": (DC_NS, "creator"),
    "last_modified_by": (CP_NS, "lastModifiedBy"),
    "revision": (CP_NS, "revision"),
    "created": (DCTERMS_NS, "created"),
    "modified": (DCTERMS_NS, "modified"),
}
FIELD_LABELS = {
    "author": "作者",
    "last_modified_by": "最后一次保存者",
    "revision": "修订号",
    "created": "创建时间",
    "modified": "最后一次保存时间",
    "total_time_minutes": "总编辑时间（分钟）",
    "windows_creation_time": "Windows 文件系统创建时间",
}


# ============================================================
# 时间与值验证
# ============================================================

def parse_datetime_utc8(value: str) -> _dt.datetime:
    """严格解析原 GUI 使用的 YYYY-MM-DD HH:MM:SS 格式，并按 UTC+8 解释。"""
    if not isinstance(value, str):
        raise ValueError("时间必须是 YYYY-MM-DD HH:MM:SS 格式的文本")
    try:
        naive = _dt.datetime.strptime(value.strip(), "%Y-%m-%d %H:%M:%S")
    except ValueError as exc:
        raise ValueError(
            f"时间格式或日期值无效：{value!r}；要求 YYYY-MM-DD HH:MM:SS（UTC+8）"
        ) from exc
    # strptime 对某些平台/输入可能较宽松，再用 round-trip 强制固定宽度。
    canonical = naive.strftime("%Y-%m-%d %H:%M:%S")
    if canonical != value.strip():
        raise ValueError(
            f"时间格式不规范：{value!r}；要求 YYYY-MM-DD HH:MM:SS（例如 2026-10-20 02:23:18）"
        )
    return naive.replace(tzinfo=UTC_PLUS_8)


def datetime_to_core_text(value: Union[str, _dt.datetime]) -> str:
    """以带 UTC+8 偏移的 ISO-8601 写入；无时区 datetime 明确按 UTC+8 解释。"""
    dt = parse_datetime_utc8(value) if isinstance(value, str) else value
    if dt.tzinfo is None or dt.utcoffset() is None:
        dt = dt.replace(tzinfo=UTC_PLUS_8)
    return dt.astimezone(UTC_PLUS_8).isoformat(timespec="seconds")


def parse_nonnegative_int(value: Any, field_label: str) -> int:
    text = str(value).strip()
    # OOXML 数字字段应写入 ASCII 十进制文本；str.isdecimal() 也接受全角/阿拉伯-印地数字，
    # 某些 Office 解析器并不接受这些 Unicode 数字字符。
    if not text or not text.isascii() or not text.isdigit():
        raise ValueError(f"{field_label}必须是由 ASCII 数字组成的非负整数；收到：{value!r}")
    number = int(text, 10)
    if number < 0:
        raise ValueError(f"{field_label}不能为负数；收到：{value!r}")
    # OPC/Office 字段最终写入十进制整数文本，避免接收超大值导致应用异常。
    if number > 2_147_483_647:
        raise ValueError(f"{field_label}超出支持范围（0 至 2147483647）")
    return number


def normalize_updates(raw_updates: Dict[str, Any]) -> Dict[str, Any]:
    normalized: Dict[str, Any] = {}
    for key, value in raw_updates.items():
        if key not in FIELD_LABELS or key == "windows_creation_time":
            raise ValueError(f"不支持的元数据字段：{key}")
        if key in ("author", "last_modified_by"):
            if value is None:
                continue
            normalized[key] = str(value)
        elif key in ("created", "modified"):
            normalized[key] = datetime_to_core_text(value if isinstance(value, _dt.datetime) else str(value))
        elif key in ("revision", "total_time_minutes"):
            normalized[key] = parse_nonnegative_int(value, FIELD_LABELS[key])
    return normalized


def _qname(namespace: str, local_name: str) -> str:
    return f"{{{namespace}}}{local_name}"


def _xml_bytes(root: ET.Element) -> bytes:
    """Serialize an XML part as UTF-8 with a stable Office-friendly declaration."""
    payload = ET.tostring(root, encoding="utf-8", xml_declaration=False)
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + payload


# ============================================================
# ZIP / OOXML 包检查与元数据读取
# ============================================================

def _validate_office_file(filepath: str) -> None:
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"文件不存在：{filepath}")
    ext = os.path.splitext(filepath)[1].lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"不支持的文件类型：{ext or '(无扩展名)'}；仅支持 .docx、.pptx、.xlsx（不含旧版 .doc/.ppt/.xls 或宏启用格式）"
        )
    if not zipfile.is_zipfile(filepath):
        raise ValueError(f"文件不是有效的 Office Open XML ZIP 包：{filepath}")
    try:
        with zipfile.ZipFile(filepath, "r") as package:
            infos = package.infolist()
            if len(infos) > MAX_ZIP_ENTRIES:
                raise ValueError(
                    f"Office 文件包含 {len(infos)} 个 ZIP 部件，超过安全上限 {MAX_ZIP_ENTRIES}；已拒绝处理"
                )
            names = [info.filename for info in infos]
            if len(names) != len(set(names)):
                raise ValueError("Office 文件包含重复的 ZIP 部件名，为避免误改已拒绝处理")
            total_uncompressed = sum(max(0, info.file_size) for info in infos)
            if total_uncompressed > MAX_ZIP_UNCOMPRESSED_SIZE:
                size_gib = total_uncompressed / (1024 ** 3)
                raise ValueError(
                    f"Office 文件 ZIP 部件声明的解压后总大小约为 {size_gib:.2f} GiB，"
                    f"超过安全上限 2 GiB；可能是超大文件或异常压缩包，已拒绝处理"
                )
            if "[Content_Types].xml" not in names or "_rels/.rels" not in names:
                raise ValueError("缺少 [Content_Types].xml 或 _rels/.rels，不像有效的 OOXML Office 文件")
            # 再按块读取每个部件：这既触发 zipfile 的 CRC 检查，也对实际读出的
            # 解压字节数做累计限制，避免只依赖 central directory 的声明总大小。
            # 4 MiB 块可减少大型演示文稿中 Python 层的循环次数，峰值缓冲仍受控。
            # 注意：zipfile 自身也会遵循 ZIP 目录中声明的单部件 file_size；若威胁模型
            # 包含精心伪造 ZIP 头的对抗样本，应使用专门的隔离解析器，而不能把本检查
            # 视为对所有恶意 ZIP 的完整沙箱防护。
            actual_uncompressed = 0
            for info in infos:
                with package.open(info, "r") as member:
                    while True:
                        chunk = member.read(ZIP_VALIDATION_CHUNK_SIZE)
                        if not chunk:
                            break
                        actual_uncompressed += len(chunk)
                        if actual_uncompressed > MAX_ZIP_UNCOMPRESSED_SIZE:
                            raise ValueError(
                                "Office 文件实际读取的解压数据超过安全上限 2 GiB；已中止校验"
                            )
    except zipfile.BadZipFile as exc:
        raise ValueError(f"Office 文件 ZIP 结构或 CRC 校验失败：{exc}") from exc
    except NotImplementedError as exc:
        # NotImplementedError 是 RuntimeError 的子类，必须先捕获，避免被误报成加密文件。
        raise ValueError(
            f"Office 文件使用了当前 Python zipfile 不支持的 ZIP 压缩方法：{exc}"
        ) from exc
    except RuntimeError as exc:
        # zipfile 对加密条目通常抛出 RuntimeError；同类内部读取错误也转成用户可理解的诊断。
        raise ValueError(
            f"Office 文件包含无法读取的 ZIP 部件（可能已加密或需要密码）：{exc}"
        ) from exc


def _parse_xml(data: bytes, part_name: str) -> ET.Element:
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise ValueError(f"无法解析 Office 元数据 XML 部件 {part_name}：{exc}") from exc


def _resolve_part_target(target: str) -> str:
    if target.startswith("/"):
        resolved = posixpath.normpath(target.lstrip("/"))
    else:
        resolved = posixpath.normpath(target)
    if resolved in ("", ".", "..") or resolved.startswith("../"):
        raise ValueError(f"Office 包关系指向了无效部件路径：{target!r}")
    return resolved


def _find_relationship(root: ET.Element, relationship_type: str) -> Optional[str]:
    for rel in root.findall(_qname(PKG_REL_NS, "Relationship")):
        if rel.get("Type") == relationship_type:
            if (rel.get("TargetMode") or "").lower() == "external":
                raise ValueError(f"Office 元数据关系意外指向外部目标：{relationship_type}")
            target = rel.get("Target")
            if not target:
                raise ValueError(f"Office 元数据关系缺少 Target：{relationship_type}")
            return _resolve_part_target(target)
    return None


def _find_content_type_override(root: ET.Element, part_path: str) -> Optional[ET.Element]:
    expected = "/" + part_path.lstrip("/")
    for node in root.findall(_qname(CT_NS, "Override")):
        if node.get("PartName") == expected:
            return node
    return None


def _load_package_parts(package: zipfile.ZipFile) -> Tuple[Dict[str, Any], Dict[str, bytes]]:
    """读取 core/app 部件。包必须含有 _rels/.rels 与 [Content_Types].xml；属性部件可缺失。"""
    names = set(package.namelist())
    required_parts = {"_rels/.rels", "[Content_Types].xml"}
    missing_parts = sorted(required_parts - names)
    if missing_parts:
        raise ValueError(
            "Office 包缺少必要的包级部件：" + ", ".join(missing_parts)
        )
    rels_root = _parse_xml(package.read("_rels/.rels"), "_rels/.rels")
    content_types_root = _parse_xml(package.read("[Content_Types].xml"), "[Content_Types].xml")

    core_path = _find_relationship(rels_root, CORE_REL_TYPE)
    app_path = _find_relationship(rels_root, APP_REL_TYPE)
    if core_path is None and "docProps/core.xml" in names:
        core_path = "docProps/core.xml"
    if app_path is None and "docProps/app.xml" in names:
        app_path = "docProps/app.xml"
    core_path = core_path or "docProps/core.xml"
    app_path = app_path or "docProps/app.xml"
    if core_path == app_path:
        raise ValueError(
            "Office 包中 core-properties 与 extended-properties 关系指向同一部件，已拒绝处理"
        )

    parts: Dict[str, Any] = {
        "core_path": core_path,
        "app_path": app_path,
        "core_exists": core_path in names,
        "app_exists": app_path in names,
        "rels_root": rels_root,
        "content_types_root": content_types_root,
        "rels_bytes": package.read("_rels/.rels"),
        "content_types_bytes": package.read("[Content_Types].xml"),
    }
    xml_parts: Dict[str, bytes] = {}
    if core_path in names:
        xml_parts[core_path] = package.read(core_path)
    if app_path in names:
        xml_parts[app_path] = package.read(app_path)
    return parts, xml_parts


def _read_core_values(core_root: Optional[ET.Element]) -> Dict[str, Any]:
    values: Dict[str, Any] = {
        "author": None,
        "last_modified_by": None,
        "revision": None,
        "created": None,
        "modified": None,
    }
    if core_root is None:
        return values
    for key, (namespace, local) in CORE_FIELDS.items():
        node = core_root.find(_qname(namespace, local))
        values[key] = node.text if node is not None else None
    return values


def _read_total_time(app_root: Optional[ET.Element]) -> Any:
    if app_root is None:
        return None
    node = app_root.find(_qname(APP_NS, "TotalTime"))
    if node is None:
        # ElementTree 的标签使用“命名空间 URI + 本地名”，不包含 XML 前缀；
        # 因而不同前缀但同一命名空间已由上面的精确查找覆盖。这里只兼容无命名空间的旧式元素，
        # 不把其他扩展命名空间中的同名 TotalTime 误读成 Office 总编辑时间。
        for child in list(app_root):
            if child.tag == "TotalTime":
                node = child
                break
    return node.text if node is not None else None


def _read_core_and_app(filepath: str) -> Dict[str, Any]:
    """读取 Office 包的 core/app 属性。公开入口负责完整 ZIP 校验；此函数复用读取逻辑。"""
    with zipfile.ZipFile(filepath, "r") as package:
        parts, xml_parts = _load_package_parts(package)
    core_root = _parse_xml(xml_parts[parts["core_path"]], parts["core_path"]) if parts["core_exists"] else None
    app_root = _parse_xml(xml_parts[parts["app_path"]], parts["app_path"]) if parts["app_exists"] else None
    values = _read_core_values(core_root)
    total_time = _read_total_time(app_root)
    # 兼容旧版原脚本曾写入的非标准 cp:totalTime；标准 app.xml TotalTime 优先。
    if total_time is None and core_root is not None:
        legacy_total = core_root.find(_qname(CP_NS, "totalTime"))
        if legacy_total is not None:
            total_time = legacy_total.text
    values["total_time_minutes"] = total_time
    return values


def get_office_metadata(filepath: str) -> Dict[str, Any]:
    """读取六项 Office 内部属性；缺失的属性返回 None，不伪造默认值。"""
    _validate_office_file(filepath)
    values = _read_core_and_app(filepath)
    values["file_type"] = os.path.splitext(filepath)[1].lower()
    if os.name == "nt":
        try:
            values["windows_creation_time"] = _get_windows_creation_time_text(filepath)
        except OSError:
            values["windows_creation_time"] = None
    else:
        values["windows_creation_time"] = None
    return values


# ============================================================
# XML 属性写入（含缺失 part / element 的创建）
# ============================================================

def _set_xml_text(root: ET.Element, namespace: str, local_name: str, value: str) -> ET.Element:
    node = root.find(_qname(namespace, local_name))
    if node is None:
        node = ET.SubElement(root, _qname(namespace, local_name))
    node.text = value
    return node


def _ensure_relationship(root: ET.Element, relationship_type: str, target_path: str) -> bool:
    """确保包根关系存在；返回是否新增关系。"""
    for node in root.findall(_qname(PKG_REL_NS, "Relationship")):
        if node.get("Type") == relationship_type:
            expected = _resolve_part_target(node.get("Target", ""))
            if expected != target_path:
                raise ValueError(
                    f"已有关系 {relationship_type} 指向 {expected!r}，"
                    f"与期望的部件路径 {target_path!r} 不一致；已拒绝静默保留错误关系"
                )
            return False
    used_ids = {node.get("Id") for node in root.findall(_qname(PKG_REL_NS, "Relationship"))}
    index = 1
    while f"rId{index}" in used_ids:
        index += 1
    ET.SubElement(
        root,
        _qname(PKG_REL_NS, "Relationship"),
        {"Id": f"rId{index}", "Type": relationship_type, "Target": target_path},
    )
    return True


def _ensure_content_type(root: ET.Element, part_path: str, content_type: str) -> bool:
    node = _find_content_type_override(root, part_path)
    if node is not None:
        if node.get("ContentType") != content_type:
            node.set("ContentType", content_type)
            return True
        return False
    ET.SubElement(
        root,
        _qname(CT_NS, "Override"),
        {"PartName": "/" + part_path.lstrip("/"), "ContentType": content_type},
    )
    return True


def _write_package_metadata(filepath: str, output_path: str, updates: Dict[str, Any]) -> None:
    """修改目标 XML 部件；其它已存在部件的解压后内容逐字节复制。"""
    with zipfile.ZipFile(filepath, "r") as source:
        parts, xml_parts = _load_package_parts(source)
        names = set(source.namelist())

        core_path = parts["core_path"]
        app_path = parts["app_path"]
        rels_root = parts["rels_root"]
        content_types_root = parts["content_types_root"]
        modified_parts: Dict[str, bytes] = {}
        new_parts: Dict[str, bytes] = {}
        rels_changed = False
        content_types_changed = False

        core_keys = set(CORE_FIELDS).intersection(updates)
        if core_keys:
            if parts["core_exists"]:
                core_root = _parse_xml(xml_parts[core_path], core_path)
            else:
                core_root = ET.Element(_qname(CP_NS, "coreProperties"))
            for key in core_keys:
                namespace, local = CORE_FIELDS[key]
                node = _set_xml_text(core_root, namespace, local, str(updates[key]))
                if key in ("created", "modified") and node.get(_qname(XSI_NS, "type")) is None:
                    # dcterms 前缀已在模块初始化时注册；同时 node 本身属于 DCTERMS_NS，
                    # 因而序列化时一定会在 XML 中声明该前缀。不要把类型值与命名空间元素拆开重构。
                    node.set(_qname(XSI_NS, "type"), "dcterms:W3CDTF")
            modified_parts[core_path] = _xml_bytes(core_root)
            if not parts["core_exists"]:
                new_parts[core_path] = modified_parts[core_path]
            rels_changed = _ensure_relationship(rels_root, CORE_REL_TYPE, core_path) or rels_changed
            content_types_changed = _ensure_content_type(content_types_root, core_path, CORE_CONTENT_TYPE) or content_types_changed

        if "total_time_minutes" in updates:
            # 旧版本曾把总编辑时间写入非标准 cp:totalTime。若源文件仍含该元素，
            # 同步其值，避免同一文件出现 app.xml 与旧元素互相矛盾的两个数值。
            if parts["core_exists"]:
                existing_core_payload = modified_parts.get(core_path, xml_parts[core_path])
                legacy_core_root = _parse_xml(existing_core_payload, core_path)
                legacy_total = legacy_core_root.find(_qname(CP_NS, "totalTime"))
                if legacy_total is not None:
                    legacy_total.text = str(updates["total_time_minutes"])
                    modified_parts[core_path] = _xml_bytes(legacy_core_root)
                    rels_changed = _ensure_relationship(rels_root, CORE_REL_TYPE, core_path) or rels_changed
                    content_types_changed = _ensure_content_type(
                        content_types_root, core_path, CORE_CONTENT_TYPE
                    ) or content_types_changed
            if parts["app_exists"]:
                app_root = _parse_xml(xml_parts[app_path], app_path)
            else:
                app_root = ET.Element(_qname(APP_NS, "Properties"))
            # 若旧式 app.xml 还包含无命名空间的 TotalTime，也同步它，避免留下冲突值。
            legacy_app_total = app_root.find("TotalTime")
            if legacy_app_total is not None:
                legacy_app_total.text = str(updates["total_time_minutes"])
            _set_xml_text(app_root, APP_NS, "TotalTime", str(updates["total_time_minutes"]))
            modified_parts[app_path] = _xml_bytes(app_root)
            if not parts["app_exists"]:
                new_parts[app_path] = modified_parts[app_path]
            rels_changed = _ensure_relationship(rels_root, APP_REL_TYPE, app_path) or rels_changed
            content_types_changed = _ensure_content_type(content_types_root, app_path, APP_CONTENT_TYPE) or content_types_changed

        # 只有确实补建/修正关系或内容类型时才重写索引文件，避免无意义改动。
        if rels_changed:
            modified_parts["_rels/.rels"] = _xml_bytes(rels_root)
        if content_types_changed:
            modified_parts["[Content_Types].xml"] = _xml_bytes(content_types_root)

        with zipfile.ZipFile(output_path, "w") as dest:
            dest.comment = source.comment
            for info in source.infolist():
                payload = modified_parts.get(info.filename)
                if payload is None:
                    payload = source.read(info.filename)
                dest.writestr(info, payload)
            for part_name, payload in new_parts.items():
                if part_name not in names:
                    # 固定新建部件的 ZIP 时间戳，避免相同内容因当前时间不同而无意义地改变包哈希。
                    new_info = zipfile.ZipInfo(part_name, date_time=(1980, 1, 1, 0, 0, 0))
                    new_info.compress_type = zipfile.ZIP_DEFLATED
                    dest.writestr(new_info, payload)

    with zipfile.ZipFile(output_path, "r") as check:
        bad = check.testzip()
        if bad is not None:
            raise ValueError(f"写入后的 Office 包 ZIP 校验失败，损坏部件：{bad}")


# ============================================================
# Windows 文件系统创建时间 API
# ============================================================

if os.name == "nt":
    _kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    _FILE_READ_ATTRIBUTES = 0x0080
    _FILE_WRITE_ATTRIBUTES = 0x0100
    _FILE_SHARE_READ = 0x00000001
    _FILE_SHARE_WRITE = 0x00000002
    _FILE_SHARE_DELETE = 0x00000004
    _OPEN_EXISTING = 3
    _INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value

    class FILETIME(ctypes.Structure):
        _fields_ = [("dwLowDateTime", ctypes.c_uint32), ("dwHighDateTime", ctypes.c_uint32)]

    _kernel32.CreateFileW.argtypes = [
        ctypes.c_wchar_p, ctypes.c_uint32, ctypes.c_uint32,
        ctypes.c_void_p, ctypes.c_uint32, ctypes.c_uint32, ctypes.c_void_p,
    ]
    _kernel32.CreateFileW.restype = ctypes.c_void_p
    _kernel32.GetFileTime.argtypes = [
        ctypes.c_void_p, ctypes.POINTER(FILETIME), ctypes.POINTER(FILETIME), ctypes.POINTER(FILETIME),
    ]
    _kernel32.GetFileTime.restype = ctypes.c_int
    _kernel32.SetFileTime.argtypes = [
        ctypes.c_void_p, ctypes.POINTER(FILETIME), ctypes.POINTER(FILETIME), ctypes.POINTER(FILETIME),
    ]
    _kernel32.SetFileTime.restype = ctypes.c_int
    # ReplaceFileW 用于尽量保留被替换文件的安全描述符、创建时间及其他文件属性。
    _kernel32.ReplaceFileW.argtypes = [
        ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_wchar_p,
        ctypes.c_uint32, ctypes.c_void_p, ctypes.c_void_p,
    ]
    _kernel32.ReplaceFileW.restype = ctypes.c_int
    _kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
    _kernel32.CloseHandle.restype = ctypes.c_int


def _datetime_to_filetime(dt: _dt.datetime) -> "FILETIME":
    if os.name != "nt":
        raise RuntimeError("Windows FILETIME 仅能在 Windows 环境使用")
    utc_dt = dt.astimezone(_dt.timezone.utc)
    epoch = _dt.datetime(1601, 1, 1, tzinfo=_dt.timezone.utc)
    delta = utc_dt - epoch
    if delta.total_seconds() < 0:
        raise ValueError("Windows 文件系统时间不能早于 1601-01-01 UTC")
    ticks = delta.days * 864_000_000_000 + delta.seconds * 10_000_000 + delta.microseconds * 10
    return FILETIME(ticks & 0xFFFFFFFF, (ticks >> 32) & 0xFFFFFFFF)


def _filetime_to_ticks(value: "FILETIME") -> int:
    return (int(value.dwHighDateTime) << 32) | int(value.dwLowDateTime)


def _ticks_to_filetime(ticks: int) -> "FILETIME":
    return FILETIME(ticks & 0xFFFFFFFF, (ticks >> 32) & 0xFFFFFFFF)


def _open_windows_file(filepath: str, access: int) -> Any:
    handle = _kernel32.CreateFileW(
        os.path.abspath(filepath), access,
        _FILE_SHARE_READ | _FILE_SHARE_WRITE | _FILE_SHARE_DELETE,
        None, _OPEN_EXISTING, 0, None,
    )
    if handle == _INVALID_HANDLE_VALUE:
        raise ctypes.WinError(ctypes.get_last_error())
    return handle


def _capture_windows_filetimes(filepath: str) -> Tuple[int, int, int]:
    created, accessed, written = FILETIME(), FILETIME(), FILETIME()
    handle = _open_windows_file(filepath, _FILE_READ_ATTRIBUTES)
    try:
        if not _kernel32.GetFileTime(handle, ctypes.byref(created), ctypes.byref(accessed), ctypes.byref(written)):
            raise ctypes.WinError(ctypes.get_last_error())
    finally:
        _kernel32.CloseHandle(handle)
    return (_filetime_to_ticks(created), _filetime_to_ticks(accessed), _filetime_to_ticks(written))


def _set_windows_filetimes(filepath: str, values: Tuple[Optional[int], Optional[int], Optional[int]]) -> None:
    created = _ticks_to_filetime(values[0]) if values[0] is not None else None
    accessed = _ticks_to_filetime(values[1]) if values[1] is not None else None
    written = _ticks_to_filetime(values[2]) if values[2] is not None else None
    handle = _open_windows_file(filepath, _FILE_WRITE_ATTRIBUTES)
    try:
        if not _kernel32.SetFileTime(
            handle,
            ctypes.byref(created) if created is not None else None,
            ctypes.byref(accessed) if accessed is not None else None,
            ctypes.byref(written) if written is not None else None,
        ):
            raise ctypes.WinError(ctypes.get_last_error())
    finally:
        _kernel32.CloseHandle(handle)


def _set_windows_creation_time(filepath: str, dt: _dt.datetime) -> None:
    ticks = _filetime_to_ticks(_datetime_to_filetime(dt))
    _set_windows_filetimes(filepath, (ticks, None, None))


def _ticks_to_datetime_utc8(ticks: int) -> _dt.datetime:
    epoch = _dt.datetime(1601, 1, 1, tzinfo=_dt.timezone.utc)
    return (epoch + _dt.timedelta(microseconds=ticks // 10)).astimezone(UTC_PLUS_8)


def _get_windows_creation_time_text(filepath: str) -> Optional[str]:
    values = _capture_windows_filetimes(filepath)
    return _ticks_to_datetime_utc8(values[0]).strftime("%Y-%m-%d %H:%M:%S UTC+08:00")


def _capture_filesystem_times(filepath: str) -> Dict[str, Any]:
    if os.name == "nt":
        return {"platform": "windows", "filetimes": _capture_windows_filetimes(filepath)}
    info = os.stat(filepath)
    return {"platform": "posix", "atime_ns": info.st_atime_ns, "mtime_ns": info.st_mtime_ns}


def _restore_filesystem_times(filepath: str, snapshot: Dict[str, Any]) -> None:
    if snapshot["platform"] == "windows":
        _set_windows_filetimes(filepath, snapshot["filetimes"])
    else:
        os.utime(filepath, ns=(snapshot["atime_ns"], snapshot["mtime_ns"]))


def _atomic_replace_preserving_target_metadata(temporary_path: str, target_path: str) -> None:
    """以临时文件替换目标，同时尽可能保留目标文件的权限/安全属性。

    直接 os.replace 会让 POSIX 文件继承 mkstemp 的权限，且在 Windows 上可能
    丢失目标文件原有的 ACL。符号链接和硬链接不适合本工具的“替换文件”语义，
    为避免静默改变链接关系，在替换前明确拒绝。
    """
    if os.path.islink(target_path):
        raise ValueError(f"目标是符号链接；为避免替换链接本身而非预期文件，已拒绝修改：{target_path}")

    try:
        target_stat = os.stat(target_path, follow_symlinks=False)
    except OSError as exc:
        raise OSError(f"无法读取目标文件属性，未替换文件：{exc}") from exc
    if target_stat.st_nlink > 1:
        raise ValueError(
            f"目标文件具有 {target_stat.st_nlink} 个硬链接；原子替换会使其他硬链接仍指向旧内容，已拒绝修改：{target_path}"
        )

    if os.name == "nt":
        # ReplaceFileW 相比直接 os.replace 更适合在 Windows 上替换现有文件，
        # 可保留目标文件的安全描述符等属性。临时文件与目标在同一目录/卷。
        if not _kernel32.ReplaceFileW(target_path, temporary_path, None, 0, None, None):
            raise ctypes.WinError(ctypes.get_last_error())
        return

    # POSIX：先把所有者/组和 copystat 支持的模式位、时间、扩展属性复制到临时文件，
    # 然后再原子替换。所有者/组无法保留时在替换前失败，避免静默改变权限边界。
    try:
        temporary_stat = os.stat(temporary_path, follow_symlinks=False)
        if (temporary_stat.st_uid, temporary_stat.st_gid) != (target_stat.st_uid, target_stat.st_gid):
            os.chown(temporary_path, target_stat.st_uid, target_stat.st_gid)
        shutil.copystat(target_path, temporary_path, follow_symlinks=False)
    except OSError as exc:
        raise OSError(f"无法保留目标文件的所有者、权限或可复制属性；为避免丢失文件属性，未替换文件：{exc}") from exc
    os.replace(temporary_path, target_path)


def _apply_original_sync_file_creation_time(filepath: str, target_dt: _dt.datetime) -> None:
    if os.name != "nt":
        raise RuntimeError("原有同步模式需要 Windows API 设置文件系统创建时间；请在 Windows 上运行")
    _set_windows_creation_time(filepath, target_dt)


# ============================================================
# 备份、原子替换、回滚与写后验证
# ============================================================

def _create_backup(filepath: str) -> str:
    """以排他创建方式分配备份路径，避免复用现存文件/悬空符号链接。"""
    fd = None
    candidate = ""
    flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, "O_BINARY", 0)
    for index in range(1000):
        candidate = filepath + ".bak" if index == 0 else filepath + f".bak.{index}"
        try:
            fd = os.open(candidate, flags, 0o600)
            break
        except FileExistsError:
            continue
    else:
        raise RuntimeError(
            f"无法为文件分配唯一备份名：{filepath}（已尝试 1000 个候选；请清理或改名旧备份后重试）"
        )

    try:
        # 使用刚才排他创建得到的文件描述符写入，避免 shutil.copy2 在竞态下跟随
        # 被替换成符号链接的备份路径。写入后再复制可复制的时间、权限和扩展属性。
        with open(filepath, "rb") as source:
            with os.fdopen(fd, "wb") as destination:
                fd = None
                shutil.copyfileobj(source, destination, length=1024 * 1024)
                destination.flush()
                os.fsync(destination.fileno())
        shutil.copystat(filepath, candidate)
    except Exception:
        if fd is not None:
            try:
                os.close(fd)
            except OSError:
                pass
        try:
            os.remove(candidate)
        except OSError:
            pass
        raise
    return candidate


def _sha256_file(filepath: str) -> str:
    digest = hashlib.sha256()
    with open(filepath, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _backup_manifest_path(backup_path: str) -> str:
    return backup_path + ".manifest.json"


def _write_backup_manifest(backup_path: str, payload: Dict[str, Any]) -> str:
    """将撤销所需记录写到本地 JSON sidecar，不发送任何数据到网络。"""
    manifest_path = _backup_manifest_path(backup_path)
    directory = os.path.dirname(os.path.abspath(manifest_path))
    fd, temporary = tempfile.mkstemp(prefix=".office_metadata_manifest_", suffix=".json", dir=directory)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2, default=_json_default)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, manifest_path)
        return manifest_path
    finally:
        if os.path.exists(temporary):
            try:
                os.remove(temporary)
            except OSError:
                pass


class BackupManifestCorruptError(RuntimeError):
    """备份撤销清单存在，但损坏或无法读取；不能将它误认成旧版无清单备份。"""


def _read_backup_manifest(backup_path: str) -> Optional[Dict[str, Any]]:
    manifest_path = _backup_manifest_path(backup_path)
    if not os.path.lexists(manifest_path):
        return None
    if not os.path.isfile(manifest_path):
        raise BackupManifestCorruptError(
            f"备份撤销清单路径存在但不是普通文件：{manifest_path}；请检查后再恢复"
        )
    try:
        with open(manifest_path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise BackupManifestCorruptError(
            f"备份撤销清单存在但无法读取或解析：{manifest_path}；"
            "为避免把损坏清单误当成旧版备份，已停止恢复。请核对备份与清单"
        ) from exc
    if not isinstance(payload, dict):
        raise BackupManifestCorruptError(
            f"备份撤销清单格式无效（顶层必须为 JSON 对象）：{manifest_path}"
        )
    return payload


def find_backups(filepath: str) -> List[str]:
    """找到指定文件的备份，排除 JSON sidecar 文件。"""
    filepath = os.path.abspath(os.path.expanduser(filepath))
    candidates = []
    for candidate in glob.glob(glob.escape(filepath) + ".bak*"):
        if not os.path.isfile(candidate) or candidate.endswith(".manifest.json"):
            continue
        if not zipfile.is_zipfile(candidate):
            continue
        candidates.append(candidate)

    def rank(candidate: str) -> float:
        try:
            manifest = _read_backup_manifest(candidate)
        except BackupManifestCorruptError:
            # 仍显示备份供用户核查；恢复阶段会再次读取并明确拒绝损坏清单。
            manifest = None
        if manifest:
            try:
                return float(manifest.get("created_at_epoch", 0))
            except (TypeError, ValueError):
                pass
        try:
            return float(os.path.getctime(candidate))
        except OSError:
            return 0.0

    return sorted(candidates, key=rank, reverse=True)


class BackupRestoreConflictError(RuntimeError):
    """因无法确认备份仍对应当前文件而阻止恢复；GUI 可针对该错误询问是否强制覆盖。"""


def restore_backup(filepath: str, backup_path: Optional[str] = None, force: bool = False) -> Dict[str, Any]:
    """从本地备份恢复文件。若当前内容已变化，默认拒绝覆盖。"""
    filepath = os.path.abspath(os.path.expanduser(filepath))
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"待恢复的原文件不存在：{filepath}")
    if os.path.islink(filepath):
        raise ValueError(f"待恢复目标是符号链接；为避免替换链接本身，已拒绝恢复：{filepath}")
    if backup_path is None:
        candidates = find_backups(filepath)
        if not candidates:
            raise FileNotFoundError(f"未发现该文件的本地备份：{filepath}")
        backup_path = candidates[0]
    backup_path = os.path.abspath(os.path.expanduser(backup_path))
    target_dir = os.path.normcase(os.path.dirname(filepath))
    backup_dir = os.path.normcase(os.path.dirname(backup_path))
    target_backup_prefix = os.path.basename(filepath) + ".bak"
    backup_name = os.path.basename(backup_path)
    is_standard_backup_name = (
        backup_name == target_backup_prefix
        or (
            backup_name.startswith(target_backup_prefix + ".")
            and backup_name[len(target_backup_prefix) + 1:].isdigit()
        )
    )
    if target_dir != backup_dir or not is_standard_backup_name:
        raise ValueError("备份路径与原文件不匹配，已拒绝恢复")
    if not os.path.isfile(backup_path) or not zipfile.is_zipfile(backup_path):
        raise ValueError(f"备份不存在或不是有效的 Office ZIP 包：{backup_path}")

    manifest = _read_backup_manifest(backup_path)
    # 清单同时记录修改前/后的哈希。先验证备份本身，防止备份文件损坏或被意外替换后仍被恢复。
    backup_sha = _sha256_file(backup_path)
    if manifest is not None:
        expected_backup_hash = manifest.get("before_sha256")
        if not isinstance(expected_backup_hash, str) or len(expected_backup_hash) != 64:
            raise ValueError("备份清单缺少有效的 before_sha256，无法验证备份完整性；已拒绝恢复")
        if backup_sha.lower() != expected_backup_hash.lower():
            raise ValueError("备份文件与撤销校验清单中的哈希不一致，备份可能已损坏或被替换；已拒绝恢复")

    current_hash = _sha256_file(filepath)
    expected_after_hash = manifest.get("after_sha256") if manifest else None
    if manifest is None and not force:
        raise BackupRestoreConflictError(
            "这是一份旧版备份，没有撤销校验清单，无法确认当前文件是否在备份后又被修改。"
            "请核对后通过界面二次确认强制恢复。"
        )
    if manifest and not expected_after_hash and not force:
        raise BackupRestoreConflictError("备份清单缺少当前版本校验值；为防止误覆盖，已拒绝自动恢复。请二次确认强制恢复。")
    if expected_after_hash and current_hash != expected_after_hash and not force:
        raise BackupRestoreConflictError(
            "当前文件内容已不同于这份备份对应的修改后版本。为防止覆盖后续编辑，已拒绝自动恢复；"
            "请先核对文件，如确认要覆盖当前内容，再通过界面二次确认强制恢复。"
        )

    directory = os.path.dirname(filepath)
    fd, restore_temp = tempfile.mkstemp(
        prefix=".office_metadata_undo_", suffix=os.path.splitext(filepath)[1], dir=directory
    )
    os.close(fd)
    try:
        shutil.copy2(backup_path, restore_temp)
        _validate_office_file(restore_temp)
        # 在替换原文件前完成可执行的读取与哈希验证，避免验证失败时目标文件已被覆盖。
        restored_values = get_office_metadata(restore_temp)
        restore_temp_sha = _sha256_file(restore_temp)
        if restore_temp_sha != backup_sha:
            raise ValueError("备份复制到临时文件后哈希不一致；目标文件未被替换，已拒绝恢复")
        _atomic_replace_preserving_target_metadata(restore_temp, filepath)
        # 从这里开始不再对目标文件做内容读取/哈希验证；只恢复记录的文件系统时间。
        timestamp_restored = False
        timestamp_warning = None
        if manifest and isinstance(manifest.get("filesystem_snapshot"), dict):
            try:
                _restore_filesystem_times(filepath, manifest["filesystem_snapshot"])
                timestamp_restored = True
                snapshot = manifest["filesystem_snapshot"]
                if snapshot.get("platform") == "windows" and snapshot.get("filetimes"):
                    restored_values["windows_creation_time"] = _ticks_to_datetime_utc8(
                        int(snapshot["filetimes"][0])
                    ).strftime("%Y-%m-%d %H:%M:%S UTC+08:00")
            except Exception as exc:
                # 内容替换已经成功，不能再把整个操作报告为“未恢复”；把时间恢复失败
                # 作为单独警告返回，避免 GUI 将已恢复的文件错误标记为失败。
                timestamp_warning = (
                    f"备份内容已恢复，但文件系统时间恢复失败：{exc}。备份保留在：{backup_path}"
                )
        elif os.name == "nt":
            try:
                restored_values["windows_creation_time"] = _get_windows_creation_time_text(filepath)
            except OSError:
                restored_values["windows_creation_time"] = None
        return {
            "filepath": filepath,
            "backup": backup_path,
            "manifest": _backup_manifest_path(backup_path) if manifest else None,
            "metadata": restored_values,
            "filesystem_timestamp_restored": timestamp_restored,
            "timestamp_warning": timestamp_warning,
            "legacy_backup": manifest is None,
        }
    finally:
        if os.path.exists(restore_temp):
            try:
                os.remove(restore_temp)
            except OSError:
                pass


def _restore_from_backup(filepath: str, backup_path: str, fs_snapshot: Dict[str, Any]) -> None:
    fd, restore_temp = tempfile.mkstemp(prefix=".office_metadata_restore_", suffix=os.path.splitext(filepath)[1], dir=os.path.dirname(os.path.abspath(filepath)))
    os.close(fd)
    try:
        shutil.copy2(backup_path, restore_temp)
        _atomic_replace_preserving_target_metadata(restore_temp, filepath)
        _restore_filesystem_times(filepath, fs_snapshot)
    finally:
        if os.path.exists(restore_temp):
            try:
                os.remove(restore_temp)
            except OSError:
                pass


def modify_office_file(
    filepath: str,
    raw_updates: Optional[Dict[str, Any]] = None,
    sync_dt: Optional[_dt.datetime] = None,
) -> Dict[str, Any]:
    """修改单个 Office 文件，返回备份路径及前后元数据。

    sync_dt 非 None 时保留原程序的同步行为：Windows 文件创建时间、Office created/modified
    设置为同一 UTC+8 目标时间；raw_updates 中明确指定的字段在其后覆盖对应 Office 属性。
    """
    filepath = os.path.abspath(os.path.expanduser(filepath))
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"文件不存在：{filepath}")
    if os.path.islink(filepath):
        raise ValueError(f"目标是符号链接；为避免替换链接本身而非预期文件，已拒绝修改：{filepath}")

    # 先验证用户输入与运行环境，再打开/读取文件，避免无效输入导致 atime 被意外改变。
    effective: Dict[str, Any] = {}
    if sync_dt is not None:
        if os.name != "nt":
            raise RuntimeError("原有时间同步模式需要 Windows API 设置文件系统创建时间；请在 Windows 上运行")
        if sync_dt.tzinfo is None or sync_dt.utcoffset() is None:
            sync_dt = sync_dt.replace(tzinfo=UTC_PLUS_8)
        # 提前验证 Windows FILETIME 可表达范围，避免写入后才失败并回滚。
        _datetime_to_filetime(sync_dt)
        sync_text = datetime_to_core_text(sync_dt)
        effective["created"] = sync_text
        effective["modified"] = sync_text
    elif not raw_updates:
        raise ValueError("没有指定要修改的字段；请勾选字段或启用原有时间同步模式")
    effective.update(normalize_updates(raw_updates or {}))
    if not effective:
        raise ValueError("没有有效的元数据更新项")

    fs_snapshot = _capture_filesystem_times(filepath)
    try:
        _validate_office_file(filepath)
    except Exception:
        try:
            _restore_filesystem_times(filepath, fs_snapshot)
        except Exception:
            pass
        raise
    # 在动手前尽可能完整地读取当前状态，避免在损坏文件上直接创建备份/覆盖。
    try:
        source_sha_before = _sha256_file(filepath)
        before = get_office_metadata(filepath)
        backup_path = _create_backup(filepath)
        backup_sha = _sha256_file(backup_path)
        source_sha_after_backup = _sha256_file(filepath)
        if source_sha_before != backup_sha or source_sha_after_backup != backup_sha:
            raise RuntimeError(
                "读取元数据或创建备份期间，目标文件内容发生变化；为避免覆盖并非本次操作起点的内容，已拒绝修改。"
                f"已生成的备份保留在：{backup_path}"
            )
    except Exception:
        try:
            _restore_filesystem_times(filepath, fs_snapshot)
        except Exception:
            pass
        raise
    replaced = False
    temp_path: Optional[str] = None
    try:
        directory = os.path.dirname(filepath)
        fd, temp_path = tempfile.mkstemp(
            prefix=".office_metadata_",
            suffix=os.path.splitext(filepath)[1],
            dir=directory,
        )
        os.close(fd)
        _write_package_metadata(filepath, temp_path, effective)

        # 在替换原文件之前先对临时包做字段级验证。
        temp_values = _read_metadata_without_os_info(temp_path)
        _verify_values(effective, temp_values)

        # 乐观并发检查：若另一个程序在备份后改写了源文件，不要把那份新内容静默覆盖。
        # 这不是跨进程锁；仍可能存在最终哈希检查与原子替换之间极短的竞态窗口。
        if _sha256_file(filepath) != backup_sha:
            raise RuntimeError(
                "目标文件在处理期间发生变化；为避免覆盖其他程序刚保存的内容，本次修改已中止。"
                f"原始备份保留在：{backup_path}"
            )
        _atomic_replace_preserving_target_metadata(temp_path, filepath)
        temp_path = None
        replaced = True

        if sync_dt is not None:
            _apply_original_sync_file_creation_time(filepath, sync_dt)

        after = get_office_metadata(filepath)
        _verify_values(effective, after)
        # 在恢复文件系统时间前计算哈希；否则后续读取可能再次更新访问时间，抵消前面的恢复。
        after_sha256 = _sha256_file(filepath)
        if sync_dt is not None:
            expected_fs_time = sync_dt.astimezone(UTC_PLUS_8).strftime("%Y-%m-%d %H:%M:%S UTC+08:00")
            if after.get("windows_creation_time") != expected_fs_time:
                raise ValueError(
                    "Windows 文件系统创建时间写入后验证失败："
                    f"期望 {expected_fs_time!r}，实际 {after.get('windows_creation_time')!r}"
                )
        else:
            # 在所有读取验证完成后再恢复创建/访问/写入时间，避免验证读取再次改变访问时间。
            _restore_filesystem_times(filepath, fs_snapshot)
            if fs_snapshot.get("platform") == "windows":
                original_created = _ticks_to_datetime_utc8(fs_snapshot["filetimes"][0]).strftime("%Y-%m-%d %H:%M:%S UTC+08:00")
                after["windows_creation_time"] = original_created
        manifest_payload = {
            "schema_version": 1,
            "original_path": filepath,
            "backup_path": backup_path,
            "created_at": _dt.datetime.now(UTC_PLUS_8).isoformat(timespec="seconds"),
            "created_at_epoch": _dt.datetime.now().timestamp(),
            "before_sha256": backup_sha,
            "after_sha256": after_sha256,
            "filesystem_snapshot": fs_snapshot,
            "changed_fields": sorted(effective.keys()),
            "before_metadata": before,
            "after_metadata": after,
        }
        manifest_path = _write_backup_manifest(backup_path, manifest_payload)
        return {"filepath": filepath, "backup": backup_path, "manifest": manifest_path, "before": before, "after": after, "updates": effective}
    except Exception as exc:
        rollback_note = ""
        if replaced:
            try:
                _restore_from_backup(filepath, backup_path, fs_snapshot)
                rollback_note = "；已从备份恢复原文件"
            except Exception as rollback_exc:
                rollback_note = f"；警告：自动回滚失败（{rollback_exc}），请保留并手动恢复备份：{backup_path}"
        else:
            # 失败发生在原文件替换前时，原内容未被替换，但读取/备份可能改变 atime。
            try:
                _restore_filesystem_times(filepath, fs_snapshot)
            except Exception as timestamp_exc:
                rollback_note = f"；警告：原文件内容未替换，但文件系统时间恢复失败（{timestamp_exc}）"
        if isinstance(exc, OSError):
            # 原地增强 strerror 而不是重新构造 FileNotFoundError/PermissionError，
            # 以保留 errno、filename、filename2 等由 OSError 构造器解析出的属性。
            if rollback_note:
                if getattr(exc, "strerror", None) is not None:
                    exc.strerror = f"{exc.strerror}{rollback_note}"
                else:
                    exc.args = (f"{exc}{rollback_note}",)
            raise
        if isinstance(exc, (ValueError, RuntimeError)):
            raise type(exc)(f"{exc}{rollback_note}") from exc
        raise RuntimeError(f"修改失败：{exc}{rollback_note}") from exc
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


def _read_metadata_without_os_info(filepath: str) -> Dict[str, Any]:
    """读取元数据但不检查/读取文件系统时间，便于临时文件验证。"""
    return _read_core_and_app(filepath)


def _verify_values(expected: Dict[str, Any], actual: Dict[str, Any]) -> None:
    for key, expected_value in expected.items():
        actual_value = actual.get(key)
        if key in ("revision", "total_time_minutes"):
            # 应用可能把整数格式化成等值十进制文本；比较其数值而不是 Python 类型。
            try:
                matched = int(str(actual_value)) == int(str(expected_value))
            except (TypeError, ValueError):
                matched = False
        elif key in ("author", "last_modified_by"):
            # 空字符串表示清空属性；XML 中空元素重新读取时通常表现为 None。
            matched = str(actual_value or "") == str(expected_value or "")
        else:
            matched = str(actual_value) == str(expected_value)
        if not matched:
            raise ValueError(
                f"写入后验证失败：{FIELD_LABELS.get(key, key)}，期望 {expected_value!r}，实际 {actual_value!r}"
            )


def build_change_plan(filepath: str, raw_updates: Optional[Dict[str, Any]] = None,
                      sync_dt: Optional[_dt.datetime] = None) -> Dict[str, Any]:
    """只读生成单文件变更预演，完全不写文件、不创建备份。"""
    filepath = os.path.abspath(os.path.expanduser(filepath))
    effective: Dict[str, Any] = {}
    if sync_dt is not None:
        if os.name != "nt":
            raise RuntimeError("原有时间同步模式需要 Windows API 设置文件系统创建时间；请在 Windows 上运行")
        if sync_dt.tzinfo is None or sync_dt.utcoffset() is None:
            sync_dt = sync_dt.replace(tzinfo=UTC_PLUS_8)
        _datetime_to_filetime(sync_dt)
        sync_text = datetime_to_core_text(sync_dt)
        effective["created"] = sync_text
        effective["modified"] = sync_text
    elif not raw_updates:
        raise ValueError("没有指定要修改的字段")
    effective.update(normalize_updates(raw_updates or {}))
    if not effective:
        raise ValueError("没有有效的元数据更新项")

    # 先校验输入，再读取目标文件。预演会完整校验 OOXML ZIP 包，读取可能改变 atime；
    # 因预演承诺不写入文件，尽力恢复访问/修改时间（POSIX ctime 由系统管理，无法回退）。
    fs_snapshot = _capture_filesystem_times(filepath)
    try:
        before = get_office_metadata(filepath)
    finally:
        # 尽力恢复读取前的时间戳；若文件读取已经因损坏等原因失败，时间戳恢复错误
        # 不应覆盖更有诊断价值的原始异常。预演本身仍然不写入文件内容。
        try:
            _restore_filesystem_times(filepath, fs_snapshot)
        except Exception:
            pass
    rows = []
    for key, target in effective.items():
        rows.append({
            "filepath": filepath,
            "key": key,
            "field": FIELD_LABELS.get(key, key),
            "before": before.get(key),
            "after": target,
        })
    if sync_dt is not None and os.name == "nt":
        rows.insert(0, {
            "filepath": filepath,
            "key": "windows_creation_time",
            "field": FIELD_LABELS["windows_creation_time"],
            "before": before.get("windows_creation_time"),
            "after": sync_dt.astimezone(UTC_PLUS_8).strftime("%Y-%m-%d %H:%M:%S UTC+08:00"),
        })
    return {"filepath": filepath, "before": before, "effective": effective, "rows": rows}


def _load_local_presets() -> Dict[str, Any]:
    """预设只保存在当前用户本地配置目录，不连接网络。"""
    if not os.path.isfile(PRESET_FILE):
        return {}
    with open(PRESET_FILE, "r", encoding="utf-8") as handle:
        payload = json.load(handle)
    if not isinstance(payload, dict):
        raise ValueError(f"本机预设文件格式无效：{PRESET_FILE}")
    return payload


def _save_local_presets(presets: Dict[str, Any]) -> None:
    os.makedirs(PRESET_DIRECTORY, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=".presets_", suffix=".json", dir=PRESET_DIRECTORY)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(presets, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, PRESET_FILE)
    finally:
        if os.path.exists(temporary):
            try:
                os.remove(temporary)
            except OSError:
                pass


# ============================================================
# 文件发现、结果报告
# ============================================================

def collect_files(filepaths: Iterable[str], folder: Optional[str], recursive: bool = False) -> List[str]:
    output: List[str] = []
    seen = set()
    for candidate in filepaths:
        path = os.path.abspath(os.path.expanduser(candidate))
        if path not in seen:
            output.append(path)
            seen.add(path)
    if folder:
        folder_path = os.path.abspath(os.path.expanduser(folder))
        if not os.path.isdir(folder_path):
            raise NotADirectoryError(f"文件夹不存在：{folder_path}")
        if recursive:
            iterator = (
                os.path.join(root, name)
                for root, _dirs, names in os.walk(folder_path)
                for name in names
            )
        else:
            iterator = (os.path.join(folder_path, name) for name in os.listdir(folder_path))
        for path in iterator:
            if os.path.isfile(path) and os.path.splitext(path)[1].lower() in SUPPORTED_EXTENSIONS:
                path = os.path.abspath(path)
                if path not in seen:
                    output.append(path)
                    seen.add(path)
    return output


def _normalized_path_identity(path: str) -> str:
    return os.path.normcase(os.path.realpath(os.path.abspath(os.path.expanduser(path))))


def validate_report_destination(path: str, protected_paths: Iterable[str]) -> str:
    """拒绝把 CSV 报告写到输入 Office 文件、其任意既有备份或撤销清单上。"""
    report_path = os.path.abspath(os.path.expanduser(path))
    report_identity = _normalized_path_identity(report_path)
    protected_list = [os.path.abspath(os.path.expanduser(str(item))) for item in protected_paths if item]
    # 同时保护输入文件旁边所有既有 .bak* 工具产物，不只是本次新生成的备份；
    # 否则指定旧备份作为 --report 时，批处理后可能把它截断成 CSV。
    for original in list(protected_list):
        if os.path.splitext(original)[1].lower() in SUPPORTED_EXTENSIONS:
            protected_list.extend(glob.glob(glob.escape(original) + ".bak*"))
    unique_protected = list(dict.fromkeys(protected_list))
    for protected_path in unique_protected:
        protected_path = os.path.abspath(os.path.expanduser(str(protected_path)))
        same_path = report_identity == _normalized_path_identity(protected_path)
        if not same_path and os.path.exists(report_path) and os.path.exists(protected_path):
            try:
                same_path = os.path.samefile(report_path, protected_path)
            except OSError:
                same_path = False
        if same_path:
            raise ValueError(
                f"CSV 报告路径与受保护的 Office 文件/备份相同，已拒绝写入以防止数据覆盖：{protected_path}"
            )
    return report_path


def write_csv_report(path: str, rows: List[Dict[str, Any]]) -> None:
    protected_paths: List[str] = []
    for row in rows:
        if row.get("filepath"):
            protected_paths.append(str(row["filepath"]))
        if row.get("backup"):
            backup_path = str(row["backup"])
            protected_paths.extend((backup_path, _backup_manifest_path(backup_path)))
    report_path = validate_report_destination(path, protected_paths)
    parent = os.path.dirname(report_path)
    if parent and not os.path.isdir(parent):
        raise FileNotFoundError(f"报告所在文件夹不存在：{parent}")
    columns = ["filepath", "status", "backup", "before_json", "after_json", "message"]
    with open(report_path, "w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns)
        writer.writeheader()
        for row in rows:
            writer.writerow({column: row.get(column, "") for column in columns})


def _json_default(value: Any) -> Any:
    if isinstance(value, (_dt.datetime, _dt.date)):
        return value.isoformat()
    return str(value)


# ============================================================
# 命令行入口
# ============================================================

def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="读取或修改 .docx/.pptx/.xlsx 的 Office 元数据；不支持旧版 .doc/.ppt/.xls。",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--gui", action="store_true", help="启动图形界面")
    parser.add_argument("--file", action="append", default=[], help="指定文件；可重复指定")
    parser.add_argument("--folder", help="扫描文件夹中的受支持文件")
    parser.add_argument("--recursive", action="store_true", help="递归扫描子文件夹（默认仅扫描当前文件夹）")
    parser.add_argument("--read", action="store_true", help="只读取并显示当前元数据，不修改文件")
    parser.add_argument("--dry-run", action="store_true", help="只预演字段修改前后差异，不写入、不创建备份")
    parser.add_argument("--sync-time", help="沿用原有时间同步模式，格式 YYYY-MM-DD HH:MM:SS（UTC+8）")
    parser.add_argument("--author", help="作者；可与其他字段独立组合")
    parser.add_argument("--last-modified-by", dest="last_modified_by", help="最后一次保存者")
    parser.add_argument("--revision", help="修订号（非负整数）")
    parser.add_argument("--created", help="Office 内部创建时间，格式 YYYY-MM-DD HH:MM:SS（UTC+8）")
    parser.add_argument("--modified", help="Office 内部最后一次保存时间，格式 YYYY-MM-DD HH:MM:SS（UTC+8）")
    parser.add_argument("--total-time-minutes", dest="total_time_minutes", help="总编辑时间，非负整数分钟")
    parser.add_argument("--report", help="将批处理结果写入 CSV 文件")
    return parser


def run_cli(argv: List[str]) -> int:
    parser = _build_arg_parser()
    args = parser.parse_args(argv)
    modification_options = (
        args.sync_time is not None or args.author is not None or args.last_modified_by is not None
        or args.revision is not None or args.created is not None or args.modified is not None
        or args.total_time_minutes is not None
    )
    if args.gui:
        if args.file or args.folder or modification_options or args.read or args.dry_run or args.report:
            parser.error("--gui 不能与文件处理、读取、预演、修改或报告参数组合使用")
        launch_gui()
        return 0
    if args.read and (args.dry_run or modification_options):
        parser.error("--read 是只读模式，不能与 --dry-run 或任何修改字段参数组合使用")
    if not args.file and not args.folder:
        parser.error("请提供 --file 或 --folder；不带参数运行则自动启动 GUI")
    try:
        files = collect_files(args.file, args.folder, args.recursive)
    except Exception as exc:
        print(f"参数错误：{exc}", file=sys.stderr)
        return 2
    if not files:
        print("未找到可处理的 .docx、.pptx 或 .xlsx 文件。", file=sys.stderr)
        return 1
    if args.report:
        try:
            validate_report_destination(args.report, files)
        except (OSError, ValueError) as exc:
            print(f"报告路径错误：{exc}", file=sys.stderr)
            return 2

    raw_updates: Dict[str, Any] = {}
    for key in ("author", "last_modified_by", "revision", "created", "modified", "total_time_minutes"):
        value = getattr(args, key)
        if value is not None:
            raw_updates[key] = value
    if not args.read and args.sync_time is None and not raw_updates:
        parser.error("没有指定操作：使用 --read 读取，--dry-run 配合目标字段预演，或指定 --sync-time / 一个或多个元数据字段")
    if not args.read:
        try:
            normalize_updates(raw_updates)
        except ValueError as exc:
            print(f"参数错误：{exc}", file=sys.stderr)
            return 2

    sync_dt = None
    if args.sync_time is not None:
        try:
            sync_dt = parse_datetime_utc8(args.sync_time)
        except ValueError as exc:
            print(f"参数错误：{exc}", file=sys.stderr)
            return 2
        if os.name != "nt":
            print(
                "环境错误：--sync-time 会设置 Windows 文件系统创建时间，只能在 Windows 上执行。",
                file=sys.stderr,
            )
            return 2

    rows: List[Dict[str, Any]] = []
    success = 0
    failed = 0
    for filepath in files:
        try:
            if args.read:
                values = get_office_metadata(filepath)
                success += 1
                print(f"\n[{filepath}]\n" + json.dumps(values, ensure_ascii=False, indent=2, default=_json_default))
                rows.append({
                    "filepath": filepath, "status": "READ_OK", "backup": "",
                    "before_json": "", "after_json": json.dumps(values, ensure_ascii=False, default=_json_default), "message": "",
                })
            elif args.dry_run:
                plan = build_change_plan(filepath, raw_updates=raw_updates, sync_dt=sync_dt)
                success += 1
                print(f"\n预演：{filepath}")
                for difference in plan["rows"]:
                    before_text = "（未设置）" if difference.get("before") is None else str(difference.get("before"))
                    print(f"  {difference['field']}: {before_text}  ->  {difference.get('after')}")
                rows.append({
                    "filepath": filepath, "status": "DRY_RUN", "backup": "",
                    "before_json": json.dumps(plan["before"], ensure_ascii=False, default=_json_default),
                    "after_json": json.dumps(plan["effective"], ensure_ascii=False, default=_json_default),
                    "message": "只读预演；未写入文件、未创建备份",
                })
            else:
                result = modify_office_file(filepath, raw_updates=raw_updates, sync_dt=sync_dt)
                success += 1
                print(f"成功：{filepath}\n备份：{result['backup']}")
                rows.append({
                    "filepath": filepath, "status": "SUCCESS", "backup": result["backup"],
                    "before_json": json.dumps(result["before"], ensure_ascii=False, default=_json_default),
                    "after_json": json.dumps(result["after"], ensure_ascii=False, default=_json_default),
                    "message": "写入并重新读取验证通过",
                })
        except Exception as exc:
            failed += 1
            print(f"失败：{filepath}\n原因：{exc}", file=sys.stderr)
            rows.append({
                "filepath": filepath, "status": "FAILED", "backup": "",
                "before_json": "", "after_json": "", "message": str(exc),
            })
    report_failed = False
    if args.report:
        try:
            write_csv_report(args.report, rows)
            print(f"CSV 报告：{os.path.abspath(args.report)}")
        except Exception as exc:
            print(f"报告写入失败：{exc}", file=sys.stderr)
            report_failed = True
    if args.dry_run:
        print(f"\n预演完成：可预演 {success}，失败 {failed}，总计 {len(files)}；未修改文件")
    elif not args.read:
        print(f"\n处理完成：成功 {success}，失败 {failed}，总计 {len(files)}")
    else:
        print(f"\n读取完成：成功 {success}，失败 {failed}")
    return 0 if failed == 0 and not report_failed else 1


# ============================================================
# Tkinter GUI（基础 / 高级双层界面）
# ============================================================

class OfficeMetadataModifierApp:
    """面向日常使用的双标签页 GUI；核心读写逻辑与 CLI 保持共用。"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Office 元数据管理工具")
        self.root.geometry("1080x800")
        self.root.minsize(900, 720)
        self.target_dt = _dt.datetime(2026, 10, 20, 2, 23, 18, tzinfo=UTC_PLUS_8)
        self.selected_files: List[str] = []
        self.file_status: Dict[str, str] = {}
        self.log_lines: List[str] = []
        self.log_window: Optional[tk.Toplevel] = None
        self.log_text: Optional[tk.Text] = None

        # 基础功能：保留原脚本默认行为。
        self.sync_var = tk.BooleanVar(value=True)
        self.total_time_enabled = tk.BooleanVar(value=True)
        self.total_time_var = tk.StringVar(value="0")
        self.sync_time_var = tk.StringVar(value=self.target_dt.strftime("%Y-%m-%d %H:%M:%S"))

        # 高级功能：独立编辑各字段，默认不勾选，避免意外覆盖现有值。
        self.recursive_var = tk.BooleanVar(value=False)
        self.author_enabled = tk.BooleanVar(value=False)
        self.last_modified_by_enabled = tk.BooleanVar(value=False)
        self.revision_enabled = tk.BooleanVar(value=False)
        self.created_enabled = tk.BooleanVar(value=False)
        self.modified_enabled = tk.BooleanVar(value=False)
        self.author_var = tk.StringVar(value="")
        self.last_modified_by_var = tk.StringVar(value="")
        self.revision_var = tk.StringVar(value="")
        self.created_var = tk.StringVar(value="")
        self.modified_var = tk.StringVar(value="")
        self.summary_before_write_var = tk.BooleanVar(value=True)
        self.export_report_var = tk.BooleanVar(value=False)
        self.report_path_var = tk.StringVar(value="")

        self._configure_style()
        self._create_widgets()

    def _configure_style(self) -> None:
        style = ttk.Style(self.root)
        try:
            themes = ("vista", "xpnative", "clam") if os.name == "nt" else ("clam", "default")
            for theme in themes:
                try:
                    style.theme_use(theme)
                    break
                except tk.TclError:
                    continue
        except tk.TclError:
            pass
        style.configure("App.TFrame", background="#F3F6FA")
        style.configure("Header.TFrame", background="#17324D")
        style.configure("HeaderTitle.TLabel", background="#17324D", foreground="#FFFFFF", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("HeaderSub.TLabel", background="#17324D", foreground="#D5E2EF", font=("Microsoft YaHei UI", 9))
        style.configure("Section.TLabelframe", padding=(8, 4))
        style.configure("Section.TLabelframe.Label", font=("Microsoft YaHei UI", 10, "bold"), foreground="#243B53")
        style.configure("TLabel", font=("Microsoft YaHei UI", 9))
        style.configure("TButton", padding=(9, 5), font=("Microsoft YaHei UI", 9))
        style.configure("Primary.TButton", padding=(16, 8), font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Treeview", rowheight=25, font=("Microsoft YaHei UI", 9))
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9, "bold"))
        style.configure("TNotebook.Tab", padding=(16, 8), font=("Microsoft YaHei UI", 10, "bold"))

    def _create_widgets(self) -> None:
        outer = ttk.Frame(self.root, style="App.TFrame", padding=8)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer, style="Header.TFrame", padding=(16, 10))
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text="Office 元数据管理工具", style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="统一管理 Word / PowerPoint / Excel 文件属性 · 自动备份 · 写入后验证",
            style="HeaderSub.TLabel",
        ).pack(anchor="w", pady=(4, 0))
        ttk.Label(
            header,
            text="本地离线处理：不包含文件上传、联网请求、遥测或云端依赖",
            style="HeaderSub.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        files_frame = ttk.LabelFrame(outer, text="01  文件列表", style="Section.TLabelframe")
        files_frame.pack(fill="x", pady=(0, 9))
        button_row = ttk.Frame(files_frame)
        button_row.pack(fill="x", pady=(0, 7))
        ttk.Button(button_row, text="添加文件…", command=self.choose_single_file).pack(side="left", padx=(0, 6))
        ttk.Button(button_row, text="添加文件夹…", command=self.choose_folder).pack(side="left", padx=6)
        ttk.Button(button_row, text="移除选中", command=self.remove_selected_files).pack(side="left", padx=6)
        ttk.Button(button_row, text="清空列表", command=self.clear_files).pack(side="left", padx=6)
        ttk.Button(button_row, text="恢复最近备份", command=self.undo_selected_files).pack(side="right", padx=(6, 0))
        ttk.Button(button_row, text="读取元数据预览", command=self.preview_metadata).pack(side="right")

        tree_frame = ttk.Frame(files_frame)
        tree_frame.pack(fill="x", expand=False)
        columns = ("kind", "path", "status")
        self.file_tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings", height=2, selectmode="extended")
        self.file_tree.heading("#0", text="文件名")
        self.file_tree.heading("kind", text="类型")
        self.file_tree.heading("path", text="完整路径")
        self.file_tree.heading("status", text="状态")
        self.file_tree.column("#0", width=190, minwidth=120, stretch=False)
        self.file_tree.column("kind", width=90, minwidth=70, stretch=False, anchor="center")
        self.file_tree.column("path", width=570, minwidth=220, stretch=True)
        self.file_tree.column("status", width=100, minwidth=80, stretch=False, anchor="center")
        tree_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.file_tree.yview)
        tree_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.file_tree.xview)
        self.file_tree.configure(yscrollcommand=tree_y.set, xscrollcommand=tree_x.set)
        self.file_tree.grid(row=0, column=0, sticky="nsew")
        tree_y.grid(row=0, column=1, sticky="ns")
        tree_x.grid(row=1, column=0, sticky="ew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self.file_count_label = ttk.Label(files_frame, text="尚未添加文件")
        self.file_count_label.pack(anchor="w", pady=(5, 0))

        self.notebook = ttk.Notebook(outer)
        self.basic_tab = ttk.Frame(self.notebook, padding=12)
        self.advanced_tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(self.basic_tab, text="基础功能")
        self.notebook.add(self.advanced_tab, text="高级功能")
        self._build_basic_tab()
        self._build_advanced_tab()

        action_row = ttk.Frame(outer)
        action_row.pack(side="bottom", fill="x", pady=(2, 0))
        self.progress = ttk.Progressbar(action_row, mode="determinate", length=180)
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 12))
        ttk.Button(action_row, text="查看日志", command=self.open_log_window).pack(side="right", padx=(0, 8))
        ttk.Button(action_row, text="清空日志", command=self.clear_log).pack(side="right", padx=(0, 8))
        ttk.Button(action_row, text="开始处理", style="Primary.TButton", command=self.start_modification).pack(side="right")
        self.status_label = ttk.Label(outer, text="就绪 · 修改前自动备份；未勾选的独立字段不会被修改", anchor="w")
        self.status_label.pack(side="bottom", fill="x", pady=(7, 0))
        # 最后再 pack 可扩展的 Notebook，确保底部“开始处理 / 查看日志”和状态栏始终可见。
        self.notebook.pack(fill="both", expand=True, pady=(0, 9))

    def _build_basic_tab(self) -> None:
        content = ttk.Frame(self.basic_tab)
        content.pack(fill="x", expand=False)
        content.columnconfigure(0, weight=3)
        content.columnconfigure(1, weight=2)

        sync_frame = ttk.LabelFrame(content, text="时间同步", style="Section.TLabelframe")
        sync_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 8))
        ttk.Checkbutton(
            sync_frame,
            text="启用原有时间同步模式",
            variable=self.sync_var,
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 7))
        ttk.Label(sync_frame, text="目标时间（UTC+8）").grid(row=1, column=0, sticky="w")
        self.sync_time_entry = ttk.Entry(sync_frame, textvariable=self.sync_time_var, width=23)
        self.sync_time_entry.grid(row=1, column=1, sticky="ew", padx=(7, 7))
        ttk.Button(sync_frame, text="检查格式", command=self.check_time).grid(row=1, column=2, sticky="w")
        ttk.Button(sync_frame, text="设为当前时间", command=self.set_current_time).grid(row=2, column=1, sticky="w", pady=(7, 0))
        ttk.Label(
            sync_frame,
            text="同步 Windows 文件创建时间、Office 创建时间与最后保存时间。",
            foreground="#607080", wraplength=590, justify="left",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(7, 0))
        sync_frame.columnconfigure(1, weight=1)

        edit_frame = ttk.LabelFrame(content, text="总编辑时间", style="Section.TLabelframe")
        edit_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 0), pady=(0, 8))
        ttk.Checkbutton(edit_frame, text="同时设置总编辑时间", variable=self.total_time_enabled).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 7))
        ttk.Label(edit_frame, text="分钟").grid(row=1, column=0, sticky="w")
        ttk.Entry(edit_frame, textvariable=self.total_time_var, width=12).grid(row=1, column=1, sticky="w", padx=(7, 0))
        ttk.Label(
            edit_frame,
            text="非负整数。默认值为 0 分钟，以保留原脚本行为；取消勾选可保留文件原值。",
            foreground="#607080", wraplength=290, justify="left",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(7, 0))

        ttk.Label(
            content,
            text="使用流程：添加文件 → 检查时间 → 开始处理；只修改作者、修订号或单独时间，请切换到“高级功能”。",
            foreground="#425466", wraplength=960, justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(1, 0))

    def _build_advanced_tab(self) -> None:
        content = ttk.Frame(self.advanced_tab)
        content.pack(fill="x", expand=False)
        content.columnconfigure(0, weight=3)
        content.columnconfigure(1, weight=2)

        fields_frame = ttk.LabelFrame(content, text="独立元数据字段（仅修改勾选项）", style="Section.TLabelframe")
        fields_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 7))
        field_rows = [
            (self.author_enabled, "作者", self.author_var),
            (self.last_modified_by_enabled, "最后一次保存者", self.last_modified_by_var),
            (self.revision_enabled, "修订号", self.revision_var),
            (self.created_enabled, "Office 创建时间", self.created_var),
            (self.modified_enabled, "Office 最后保存时间", self.modified_var),
        ]
        for row_index, (enabled, label, variable) in enumerate(field_rows):
            ttk.Checkbutton(fields_frame, variable=enabled).grid(row=row_index, column=0, sticky="w", padx=(0, 3), pady=2)
            ttk.Label(fields_frame, text=label, width=19).grid(row=row_index, column=1, sticky="w", pady=2)
            ttk.Entry(fields_frame, textvariable=variable, width=28).grid(row=row_index, column=2, sticky="ew", pady=2)
        fields_frame.columnconfigure(2, weight=1)
        ttk.Label(
            fields_frame, text="时间格式：YYYY-MM-DD HH:MM:SS（UTC+8）；修订号为非负整数。",
            foreground="#607080", wraplength=580, justify="left",
        ).grid(row=len(field_rows), column=0, columnspan=3, sticky="w", pady=(6, 0))

        opts_frame = ttk.LabelFrame(content, text="批处理与安全选项", style="Section.TLabelframe")
        opts_frame.grid(row=0, column=1, sticky="nsew", pady=(0, 7))
        ttk.Checkbutton(opts_frame, text="扫描所选文件夹的子文件夹", variable=self.recursive_var).pack(anchor="w", pady=(0, 5))
        ttk.Checkbutton(opts_frame, text="写入前显示变更摘要", variable=self.summary_before_write_var).pack(anchor="w", pady=(0, 5))
        ttk.Checkbutton(opts_frame, text="导出本次处理 CSV 报告", variable=self.export_report_var, command=self._toggle_report_path).pack(anchor="w", pady=(0, 5))
        report_row = ttk.Frame(opts_frame)
        report_row.pack(fill="x", pady=(0, 7))
        self.report_entry = ttk.Entry(report_row, textvariable=self.report_path_var, state="disabled")
        self.report_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.report_browse_button = ttk.Button(report_row, text="路径…", command=self.choose_report_path, state="disabled")
        self.report_browse_button.pack(side="right")
        ttk.Label(
            opts_frame,
            text="自动备份 · 写后验证 · 失败回滚",
            foreground="#1B6B50", wraplength=300,
        ).pack(anchor="w", pady=(2, 7))

        preset_frame = ttk.LabelFrame(opts_frame, text="本机配置预设", style="Section.TLabelframe")
        preset_frame.pack(fill="x", pady=(0, 4))
        preset_row = ttk.Frame(preset_frame)
        preset_row.pack(fill="x", pady=(0, 5))
        self.preset_name_var = tk.StringVar(value="")
        self.preset_combo = ttk.Combobox(preset_row, textvariable=self.preset_name_var, state="readonly", width=23)
        self.preset_combo.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(preset_row, text="加载", command=self.load_selected_preset).pack(side="left", padx=(0, 4))
        preset_actions = ttk.Frame(preset_frame)
        preset_actions.pack(fill="x")
        ttk.Button(preset_actions, text="保存当前配置…", command=self.save_current_preset).pack(side="left", padx=(0, 5))
        ttk.Button(preset_actions, text="删除预设", command=self.delete_selected_preset).pack(side="left")
        ttk.Label(
            preset_frame,
            text="预设仅保存到本机用户目录；不保存文件列表或文件内容。",
            foreground="#607080", wraplength=300, justify="left",
        ).pack(anchor="w", pady=(5, 0))
        self.refresh_preset_names(show_errors=False)


    def _preset_config(self) -> Dict[str, Any]:
        """预设只记录输入项和开关，不记录文件路径、文件内容或报告路径。"""
        return {
            "sync_enabled": bool(self.sync_var.get()),
            "sync_time": self.sync_time_var.get(),
            "total_time_enabled": bool(self.total_time_enabled.get()),
            "total_time_minutes": self.total_time_var.get(),
            "recursive": bool(self.recursive_var.get()),
            "summary_before_write": bool(self.summary_before_write_var.get()),
            "fields": {
                "author": {"enabled": bool(self.author_enabled.get()), "value": self.author_var.get()},
                "last_modified_by": {"enabled": bool(self.last_modified_by_enabled.get()), "value": self.last_modified_by_var.get()},
                "revision": {"enabled": bool(self.revision_enabled.get()), "value": self.revision_var.get()},
                "created": {"enabled": bool(self.created_enabled.get()), "value": self.created_var.get()},
                "modified": {"enabled": bool(self.modified_enabled.get()), "value": self.modified_var.get()},
            },
        }

    def refresh_preset_names(self, show_errors: bool = True) -> None:
        try:
            presets = _load_local_presets()
            names = sorted(str(name) for name in presets.keys())
            self.preset_combo.configure(values=names)
            if self.preset_name_var.get() not in names:
                self.preset_name_var.set(names[0] if names else "")
        except Exception as exc:
            self.preset_combo.configure(values=[])
            if show_errors:
                messagebox.showerror("读取预设失败", f"无法读取本机预设文件：\n{exc}", parent=self.root)

    def save_current_preset(self) -> None:
        name = simpledialog.askstring("保存配置预设", "为当前设置输入一个名称：", parent=self.root)
        if name is None:
            return
        name = name.strip()
        if not name:
            messagebox.showwarning("预设名称为空", "请输入预设名称。", parent=self.root)
            return
        try:
            presets = _load_local_presets()
            if name in presets and not messagebox.askyesno("覆盖预设", f"预设“{name}”已存在。是否覆盖？", parent=self.root):
                return
            presets[name] = self._preset_config()
            _save_local_presets(presets)
            self.preset_name_var.set(name)
            self.refresh_preset_names(show_errors=False)
            self._append_log(f"已在本机保存配置预设：{name}（{PRESET_FILE}）")
            messagebox.showinfo("保存成功", f"预设“{name}”已保存到本机：\n{PRESET_FILE}", parent=self.root)
        except Exception as exc:
            messagebox.showerror("保存预设失败", str(exc), parent=self.root)

    def load_selected_preset(self) -> None:
        name = self.preset_name_var.get().strip()
        if not name:
            messagebox.showinfo("请选择预设", "请先从列表中选择一个预设。", parent=self.root)
            return
        try:
            presets = _load_local_presets()
            config = presets.get(name)
            if not isinstance(config, dict):
                raise ValueError(f"预设“{name}”不存在或格式无效")
            self.sync_var.set(bool(config.get("sync_enabled", True)))
            self.sync_time_var.set(str(config.get("sync_time", self.sync_time_var.get())))
            self.total_time_enabled.set(bool(config.get("total_time_enabled", True)))
            self.total_time_var.set(str(config.get("total_time_minutes", "0")))
            self.recursive_var.set(bool(config.get("recursive", False)))
            self.summary_before_write_var.set(bool(config.get("summary_before_write", True)))
            fields = config.get("fields", {})
            field_targets = {
                "author": (self.author_enabled, self.author_var),
                "last_modified_by": (self.last_modified_by_enabled, self.last_modified_by_var),
                "revision": (self.revision_enabled, self.revision_var),
                "created": (self.created_enabled, self.created_var),
                "modified": (self.modified_enabled, self.modified_var),
            }
            for key, (enabled_var, value_var) in field_targets.items():
                setting = fields.get(key, {}) if isinstance(fields, dict) else {}
                if not isinstance(setting, dict):
                    setting = {}
                enabled_var.set(bool(setting.get("enabled", False)))
                value_var.set(str(setting.get("value", "")))
            self.notebook.select(self.basic_tab if self.sync_var.get() else self.advanced_tab)
            self._append_log(f"已从本机加载配置预设：{name}")
            self.status_label.config(text=f"已加载预设：{name} · 文件列表未改变")
        except Exception as exc:
            messagebox.showerror("加载预设失败", str(exc), parent=self.root)

    def delete_selected_preset(self) -> None:
        name = self.preset_name_var.get().strip()
        if not name:
            messagebox.showinfo("请选择预设", "请先从列表中选择一个预设。", parent=self.root)
            return
        if not messagebox.askyesno("删除预设", f"确定删除本机配置预设“{name}”吗？\n不会删除任何 Office 文件。", parent=self.root):
            return
        try:
            presets = _load_local_presets()
            presets.pop(name, None)
            _save_local_presets(presets)
            self.preset_name_var.set("")
            self.refresh_preset_names(show_errors=False)
            self._append_log(f"已删除本机配置预设：{name}")
        except Exception as exc:
            messagebox.showerror("删除预设失败", str(exc), parent=self.root)

    def undo_selected_files(self) -> None:
        paths = self._selected_paths_in_tree()
        if not paths:
            messagebox.showwarning("请先选择文件", "在文件列表中选中要恢复的文件，然后点击“恢复最近备份”。", parent=self.root)
            return
        backup_map = {path: find_backups(path) for path in paths}
        found = {path: candidates[0] for path, candidates in backup_map.items() if candidates}
        missing = [path for path, candidates in backup_map.items() if not candidates]
        if not found:
            messagebox.showinfo("没有备份", "所选文件均没有找到可用的本地 .bak 备份。", parent=self.root)
            return
        preview_blocks = []
        corrupt_manifests: Dict[str, str] = {}
        for path, backup in found.items():
            try:
                manifest = _read_backup_manifest(backup)
            except BackupManifestCorruptError as exc:
                manifest = None
                corrupt_manifests[path] = str(exc)
                manifest_warning = (
                    "（撤销清单损坏；该备份将被跳过，不会进入强制恢复流程。"
                    f"原因：{exc}）"
                )
            else:
                manifest_warning = ""
            if manifest and manifest.get("created_at"):
                backup_time = str(manifest["created_at"]) + "（工具记录时间）"
            else:
                try:
                    backup_time = _dt.datetime.fromtimestamp(
                        os.path.getctime(backup), UTC_PLUS_8
                    ).strftime("%Y-%m-%d %H:%M:%S UTC+08:00") + "（旧版备份的文件系统时间，可能是估算值）"
                except OSError:
                    backup_time = "无法读取"
            backup_time += manifest_warning
            preview_blocks.append(
                f"原文件：{path}\n  备份：{backup}\n  备份时间：{backup_time}"
            )
        preview = "\n\n".join(preview_blocks)
        if missing:
            preview += "\n\n以下文件没有可用备份，将跳过：\n" + "\n".join(missing)
        if not messagebox.askyesno(
            "确认恢复最近备份",
            "将用对应的最近本地备份替换所选文件。恢复不会删除备份，但会覆盖当前文件内容。\n\n"
            + preview + "\n\n是否继续？",
            parent=self.root,
        ):
            return

        restored = 0
        failures = []
        skipped = []
        warnings = []
        for path, backup in found.items():
            if path in corrupt_manifests:
                reason = corrupt_manifests[path]
                skipped.append(f"{path}：撤销清单损坏，按安全策略跳过；不会进入强制恢复流程。{reason}")
                self.file_status[path] = "撤销清单损坏，已跳过"
                self._append_log(
                    f"已跳过恢复：{path}\n  备份：{backup}\n  原因：撤销清单损坏，无法验证备份完整性。"
                )
                continue
            try:
                try:
                    result = restore_backup(path, backup_path=backup, force=False)
                except BackupRestoreConflictError as exc:
                    if not messagebox.askyesno(
                        "恢复安全检查未通过",
                        f"文件：{path}\n\n{exc}\n\n若继续，将强制用备份覆盖当前文件。是否确认？",
                        parent=self.root,
                    ):
                        failures.append(f"{path}：用户取消强制恢复")
                        continue
                    result = restore_backup(path, backup_path=backup, force=True)
                restored += 1
                timestamp_warning = result.get("timestamp_warning")
                self.file_status[path] = "已恢复备份（时间恢复失败）" if timestamp_warning else "已恢复备份"
                note = "；旧版备份未保存文件系统时间记录" if result.get("legacy_backup") else ""
                if os.name != "nt":
                    note += "；POSIX 系统仅恢复访问/修改时间，inode ctime 由系统管理且无法回退"
                if timestamp_warning:
                    warnings.append(f"{path}：{timestamp_warning}")
                    note += f"；警告：{timestamp_warning}"
                self._append_log(f"已恢复：{path}\n  备份：{backup}{note}")
            except Exception as exc:
                failures.append(f"{path}：{exc}")
                self.file_status[path] = "恢复失败"
                self._append_log(f"恢复失败：{path}\n  原因：{exc}")
        self._refresh_file_tree()
        if failures or skipped or warnings:
            details = []
            if failures:
                details.append("恢复失败：\n" + "\n".join(failures))
            if skipped:
                details.append("因安全检查跳过：\n" + "\n".join(skipped))
            if warnings:
                details.append("已恢复内容，但有时间戳警告：\n" + "\n".join(warnings))
            messagebox.showwarning(
                "恢复完成（存在失败、跳过或警告）",
                f"成功恢复：{restored}\n失败：{len(failures)}\n安全跳过：{len(skipped)}\n警告：{len(warnings)}\n\n" + "\n\n".join(details),
                parent=self.root,
            )
        else:
            messagebox.showinfo("恢复完成", f"成功恢复 {restored} 个文件。\n原备份仍保留在本地。", parent=self.root)
        self.status_label.config(text=f"备份恢复完成 · 成功 {restored} / 失败 {len(failures)} / 安全跳过 {len(skipped)} / 警告 {len(warnings)}")

    def _append_log(self, text: str) -> None:
        line = text.rstrip()
        self.log_lines.append(line)
        if self.log_text is not None:
            try:
                if self.log_text.winfo_exists():
                    self.log_text.configure(state="normal")
                    self.log_text.insert(tk.END, line + "\n")
                    self.log_text.see(tk.END)
                    self.log_text.configure(state="disabled")
            except tk.TclError:
                self.log_text = None
                self.log_window = None

    def open_log_window(self) -> None:
        try:
            if self.log_window is not None and self.log_window.winfo_exists():
                self.log_window.deiconify()
                self.log_window.lift()
                self.log_window.focus_set()
                return
        except tk.TclError:
            self.log_window = None
            self.log_text = None
        window = tk.Toplevel(self.root)
        window.title("处理日志与元数据预览")
        window.geometry("850x520")
        window.minsize(620, 360)
        self.log_window = window
        frame = ttk.Frame(window, padding=10)
        frame.pack(fill="both", expand=True)
        text_widget = tk.Text(
            frame, wrap="word", font=("Cascadia Mono", 9),
            background="#FFFFFF", foreground="#253746", relief="flat", padx=8, pady=6,
        )
        scroll = ttk.Scrollbar(frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scroll.set)
        text_widget.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        text_widget.insert("1.0", "\n".join(self.log_lines) + ("\n" if self.log_lines else "暂无日志。添加文件后可读取或处理。\n"))
        text_widget.configure(state="disabled")
        self.log_text = text_widget
        def on_close() -> None:
            self.log_text = None
            self.log_window = None
            window.destroy()
        window.protocol("WM_DELETE_WINDOW", on_close)

    def clear_log(self) -> None:
        self.log_lines.clear()
        if self.log_text is not None:
            try:
                if self.log_text.winfo_exists():
                    self.log_text.configure(state="normal")
                    self.log_text.delete("1.0", tk.END)
                    self.log_text.configure(state="disabled")
            except tk.TclError:
                self.log_text = None
                self.log_window = None

    def check_time(self) -> None:
        try:
            parsed = parse_datetime_utc8(self.sync_time_var.get())
            self.target_dt = parsed
            messagebox.showinfo("时间格式正确", "时间格式有效，将按 UTC+8 解释。", parent=self.root)
        except ValueError as exc:
            messagebox.showerror("时间格式错误", str(exc), parent=self.root)

    def set_current_time(self) -> None:
        current = _dt.datetime.now(UTC_PLUS_8).replace(microsecond=0)
        self.target_dt = current
        self.sync_time_var.set(current.strftime("%Y-%m-%d %H:%M:%S"))

    def choose_single_file(self) -> None:
        filenames = filedialog.askopenfilenames(
            parent=self.root,
            title="选择一个或多个 Office 文件",
            filetypes=[
                ("支持的 Office 文件", "*.docx *.pptx *.xlsx"),
                ("Word 文档", "*.docx"), ("PowerPoint", "*.pptx"),
                ("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*"),
            ],
        )
        if filenames:
            self._add_files(list(filenames))

    def choose_folder(self) -> None:
        folder = filedialog.askdirectory(parent=self.root, title="选择包含 Office 文件的文件夹")
        if not folder:
            return
        try:
            found = collect_files([], folder, self.recursive_var.get())
        except Exception as exc:
            messagebox.showerror("选择文件夹失败", str(exc), parent=self.root)
            return
        if not found:
            messagebox.showinfo("未找到文件", "该文件夹中没有发现受支持的 .docx、.pptx 或 .xlsx 文件。", parent=self.root)
            return
        self._add_files(found)
        self._append_log(f"已从文件夹添加 {len(found)} 个文件：{folder}")

    def _add_files(self, paths: Iterable[str]) -> None:
        seen = set(self.selected_files)
        added = 0
        for candidate in paths:
            path = os.path.abspath(os.path.expanduser(candidate))
            if path not in seen and os.path.splitext(path)[1].lower() in SUPPORTED_EXTENSIONS:
                self.selected_files.append(path)
                self.file_status[path] = "待处理"
                seen.add(path)
                added += 1
        self._refresh_file_tree()
        self.status_label.config(text=f"已添加 {added} 个文件；当前共 {len(self.selected_files)} 个")

    def _refresh_file_tree(self) -> None:
        for iid in self.file_tree.get_children():
            self.file_tree.delete(iid)
        for index, path in enumerate(self.selected_files):
            ext = os.path.splitext(path)[1].lower()
            kind = {".docx": "Word", ".pptx": "PowerPoint", ".xlsx": "Excel"}.get(ext, ext.upper())
            name = os.path.basename(path)
            # 使用稳定的内部编号，避免特殊路径字符影响 Treeview 项目标识。
            iid = f"file_{index}"
            self.file_tree.insert("", "end", iid=iid, text=name, values=(kind, path, self.file_status.get(path, "待处理")))
        self.file_count_label.config(text=f"共 {len(self.selected_files)} 个文件 · 支持 DOCX / PPTX / XLSX")

    def _selected_paths_in_tree(self) -> List[str]:
        paths: List[str] = []
        for iid in self.file_tree.selection():
            values = self.file_tree.item(iid, "values")
            if len(values) >= 2:
                paths.append(str(values[1]))
        return paths

    def remove_selected_files(self) -> None:
        remove = set(self._selected_paths_in_tree())
        if not remove:
            messagebox.showinfo("提示", "请先在文件列表中选中要移除的文件。", parent=self.root)
            return
        self.selected_files = [path for path in self.selected_files if path not in remove]
        for path in remove:
            self.file_status.pop(path, None)
        self._refresh_file_tree()
        self.status_label.config(text=f"已移除 {len(remove)} 个文件；剩余 {len(self.selected_files)} 个")

    def clear_files(self) -> None:
        self.selected_files.clear()
        self.file_status.clear()
        self._refresh_file_tree()
        self.status_label.config(text="文件列表已清空")

    def _toggle_report_path(self) -> None:
        state = "normal" if self.export_report_var.get() else "disabled"
        self.report_entry.configure(state=state)
        self.report_browse_button.configure(state=state)

    def choose_report_path(self) -> None:
        path = filedialog.asksaveasfilename(
            parent=self.root, title="保存批处理报告", defaultextension=".csv",
            filetypes=[("CSV 报告", "*.csv")],
            initialfile="office_metadata_report.csv",
        )
        if path:
            self.report_path_var.set(path)

    def preview_metadata(self) -> None:
        if not self.selected_files:
            messagebox.showwarning("提示", "请先添加文件或文件夹。", parent=self.root)
            return
        self.clear_log()
        for index, path in enumerate(self.selected_files, 1):
            self._append_log(f"[{index}/{len(self.selected_files)}] {path}")
            try:
                values = get_office_metadata(path)
                for key in ("author", "last_modified_by", "revision", "created", "modified", "total_time_minutes", "windows_creation_time"):
                    label = FIELD_LABELS[key]
                    self._append_log(f"  {label}：{values.get(key) if values.get(key) is not None else '（未设置）'}")
                self.file_status[path] = "读取成功"
            except Exception as exc:
                self._append_log(f"  读取失败：{exc}")
                self.file_status[path] = "读取失败"
            self._append_log("")
            self._refresh_file_tree()
            self.root.update_idletasks()
        self.status_label.config(text=f"元数据预览完成：{len(self.selected_files)} 个文件")

    def _collect_updates_from_gui(self) -> Dict[str, Any]:
        updates: Dict[str, Any] = {}
        checkboxes = [
            (self.author_enabled, "author", self.author_var),
            (self.last_modified_by_enabled, "last_modified_by", self.last_modified_by_var),
            (self.revision_enabled, "revision", self.revision_var),
            (self.created_enabled, "created", self.created_var),
            (self.modified_enabled, "modified", self.modified_var),
        ]
        for enabled, key, variable in checkboxes:
            if enabled.get():
                updates[key] = variable.get()
        if self.total_time_enabled.get():
            updates["total_time_minutes"] = self.total_time_var.get()
        # 验证输入，但返回原始 GUI 值。下游的 build_change_plan / modify_office_file
        # 会各自规范化一次；如果提前返回 ISO 时间字符串，下游会把它误判为非法输入并二次解析失败。
        normalize_updates(updates)
        return updates

    def _show_change_preview(self, updates: Dict[str, Any], sync_dt: Optional[_dt.datetime]) -> bool:
        """逐字段展示写入前后值；预演失败时禁用确认，避免盲目批量修改。"""
        plans = []
        errors = []
        for index, path in enumerate(self.selected_files, 1):
            try:
                plan = build_change_plan(path, raw_updates=updates, sync_dt=sync_dt)
                plans.extend(plan["rows"])
                self._append_log(f"预演成功：{path}（{len(plan['rows'])} 个字段变化）")
            except Exception as exc:
                errors.append((path, str(exc)))
                self._append_log(f"预演失败：{path}：{exc}")

        window = tk.Toplevel(self.root)
        window.title("修改预演 · 字段前后对比")
        window.geometry("1040x570")
        window.minsize(820, 420)
        window.transient(self.root)
        body = ttk.Frame(window, padding=12)
        body.pack(fill="both", expand=True)
        if errors:
            caption = f"预演遇到 {len(errors)} 个错误。为避免不完整预览，本次操作不会继续；请修复或移除失败文件后重试。"
        else:
            caption = f"共 {len(self.selected_files)} 个文件，预览 {len(plans)} 项字段变化。确认后才会写入；此窗口本身不会修改文件。"
        ttk.Label(body, text=caption, wraplength=990, justify="left").pack(anchor="w", pady=(0, 8))
        table_frame = ttk.Frame(body)
        table_frame.pack(fill="both", expand=True)
        columns = ("file", "field", "before", "after")
        table = ttk.Treeview(table_frame, columns=columns, show="headings", height=16)
        table.heading("file", text="文件")
        table.heading("field", text="字段")
        table.heading("before", text="原始值")
        table.heading("after", text="目标值")
        table.column("file", width=290, minwidth=180, stretch=True)
        table.column("field", width=160, minwidth=130, stretch=False)
        table.column("before", width=270, minwidth=180, stretch=True)
        table.column("after", width=270, minwidth=180, stretch=True)
        ybar = ttk.Scrollbar(table_frame, orient="vertical", command=table.yview)
        xbar = ttk.Scrollbar(table_frame, orient="horizontal", command=table.xview)
        table.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        table.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        for row in plans:
            table.insert("", "end", values=(
                row["filepath"], row["field"],
                "（未设置）" if row.get("before") is None else str(row.get("before")),
                "（清空）" if row.get("after") == "" and row["key"] in ("author", "last_modified_by") else str(row.get("after")),
            ))
        for path, error in errors:
            table.insert("", "end", values=(path, "预演失败", error, "请修复文件或从列表移除"))

        decision = {"confirmed": False}
        button_row = ttk.Frame(body)
        button_row.pack(fill="x", pady=(9, 0))
        def close_preview() -> None:
            try:
                window.grab_release()
            except tk.TclError:
                pass
            window.destroy()
        def confirm_preview() -> None:
            decision["confirmed"] = True
            close_preview()
        ttk.Button(button_row, text="取消", command=close_preview).pack(side="right", padx=(6, 0))
        ttk.Button(
            button_row, text="确认并执行", style="Primary.TButton",
            command=confirm_preview, state=("disabled" if errors else "normal"),
        ).pack(side="right")
        window.protocol("WM_DELETE_WINDOW", close_preview)
        window.grab_set()
        self.root.wait_window(window)
        return decision["confirmed"]

    def start_modification(self) -> None:
        if not self.selected_files:
            messagebox.showwarning("提示", "请先添加文件或文件夹。", parent=self.root)
            return
        sync_dt = None
        if self.sync_var.get():
            try:
                sync_dt = parse_datetime_utc8(self.sync_time_var.get())
            except ValueError as exc:
                messagebox.showerror("时间格式错误", str(exc), parent=self.root)
                self.notebook.select(self.basic_tab)
                return
        try:
            updates = self._collect_updates_from_gui()
        except ValueError as exc:
            error_text = str(exc)
            messagebox.showerror("字段值错误", error_text, parent=self.root)
            # “总编辑时间”位于基础页；其他独立字段位于高级页。
            target_tab = self.basic_tab if FIELD_LABELS["total_time_minutes"] in error_text else self.advanced_tab
            self.notebook.select(target_tab)
            return
        if not sync_dt and not updates:
            messagebox.showwarning("没有修改项", "请启用时间同步，或至少勾选一个独立字段/总编辑时间。", parent=self.root)
            return
        if sync_dt and os.name != "nt":
            messagebox.showerror("环境不支持", "原有时间同步模式需要 Windows API 设置文件系统创建时间，请在 Windows 上运行。", parent=self.root)
            return
        if self.export_report_var.get() and not self.report_path_var.get().strip():
            self.choose_report_path()
            if not self.report_path_var.get().strip():
                return
        if self.export_report_var.get():
            try:
                validate_report_destination(self.report_path_var.get().strip(), self.selected_files)
            except (OSError, ValueError) as exc:
                messagebox.showerror("报告路径错误", str(exc), parent=self.root)
                return

        fields = [FIELD_LABELS[key] for key in updates]
        description = (
            f"文件数量：{len(self.selected_files)}\n"
            + (f"时间同步：{self.sync_time_var.get()}（UTC+8）\n" if sync_dt else "时间同步：关闭\n")
            + (f"独立字段：{'、'.join(fields)}\n" if fields else "独立字段：无\n")
            + "安全措施：本机预演、自动备份、写后验证，失败时尝试回滚。\n\n是否继续？"
        )
        if self.summary_before_write_var.get():
            if not self._show_change_preview(updates, sync_dt):
                return
        elif not messagebox.askyesno("确认修改", description, parent=self.root):
            return

        self.clear_log()
        self.progress.configure(maximum=max(1, len(self.selected_files)), value=0)
        self.status_label.config(text="正在处理…")
        self.root.update_idletasks()
        successes = 0
        failures: List[str] = []
        report_rows: List[Dict[str, Any]] = []
        for index, path in enumerate(self.selected_files, 1):
            self._append_log(f"[{index}/{len(self.selected_files)}] 处理：{path}")
            try:
                result = modify_office_file(path, updates, sync_dt)
                successes += 1
                self.file_status[path] = "处理成功"
                self._append_log(f"  成功；备份：{result['backup']}")
                for key in result["updates"]:
                    self._append_log(f"  {FIELD_LABELS.get(key, key)}：{result['after'].get(key)}")
                report_rows.append({
                    "filepath": path, "status": "SUCCESS", "backup": result["backup"],
                    "before_json": json.dumps(result["before"], ensure_ascii=False, default=_json_default),
                    "after_json": json.dumps(result["after"], ensure_ascii=False, default=_json_default),
                    "message": "写入并重新读取验证通过",
                })
            except Exception as exc:
                failures.append(f"{path}：{exc}")
                self.file_status[path] = "处理失败"
                self._append_log(f"  失败：{exc}")
                report_rows.append({
                    "filepath": path, "status": "FAILED", "backup": "",
                    "before_json": "", "after_json": "", "message": str(exc),
                })
            self.progress.configure(value=index)
            self._refresh_file_tree()
            self.root.update_idletasks()
            self._append_log("")

        report_error = None
        report_note = ""
        if self.export_report_var.get():
            try:
                write_csv_report(self.report_path_var.get().strip(), report_rows)
                report_note = f"\nCSV 报告：{self.report_path_var.get().strip()}"
                self._append_log(report_note.strip())
            except Exception as exc:
                report_error = f"CSV 报告写入失败：{exc}"
                report_note = f"\n{report_error}"
                self._append_log(report_error)
        summary = f"处理完成！\n成功：{successes}\n文件处理失败：{len(failures)}\n总计：{len(self.selected_files)}{report_note}"
        if failures or report_error:
            details = []
            if failures:
                details.append("文件处理失败详情：\n" + "\n".join(failures))
            if report_error:
                details.append(report_error)
            summary += "\n\n" + "\n\n".join(details)
            messagebox.showwarning("处理完成（存在警告或失败）", summary, parent=self.root)
        else:
            messagebox.showinfo("处理成功", summary, parent=self.root)
        report_status = " / 报告失败" if report_error else ""
        self.status_label.config(text=f"处理完成 · 成功 {successes} / 文件失败 {len(failures)}{report_status}")


def launch_gui() -> None:
    root = tk.Tk()
    OfficeMetadataModifierApp(root)
    root.mainloop()


# ============================================================
# 控制台编码兼容
# ============================================================

def configure_console_error_handling() -> None:
    """避免旧版 Windows 控制台遇到不可编码字符时直接抛出 UnicodeEncodeError。

    不强行把控制台编码改为 UTF-8，以免破坏仍使用 GBK/系统代码页的终端；
    对当前代码页无法表示的字符使用 \\u/\\U 转义，文件本身仍统一使用 UTF-8。
    """
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if stream is None or not hasattr(stream, "reconfigure"):
            continue
        try:
            stream.reconfigure(errors="backslashreplace")
        except (OSError, ValueError):
            # 例如被某些宿主替换的流不支持 reconfigure；不影响 GUI 和文件读写。
            pass


# ============================================================
# 程序入口
# ============================================================

if __name__ == "__main__":
    configure_console_error_handling()
    if len(sys.argv) == 1:
        try:
            launch_gui()
        except tk.TclError as exc:
            print(f"无法启动 GUI（可能缺少图形桌面或 DISPLAY）：{exc}", file=sys.stderr)
            print("请在有图形桌面的环境中启动，或使用命令行参数运行，例如 --help。", file=sys.stderr)
            raise SystemExit(1) from exc
    else:
        raise SystemExit(run_cli(sys.argv[1:]))
