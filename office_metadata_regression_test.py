#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""便携式边界回归测试。将本脚本、office_metadata_modifier.py 与验证样例 ZIP 放在同一目录运行。
只使用 Python 标准库；测试只在临时目录中操作副本，不会修改用户的实际文档。
"""
from __future__ import annotations
import hashlib
import io
import importlib.util
import os
import shutil
import subprocess
import sys
import tempfile
import types
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).resolve().parent
SCRIPT = HERE / "office_metadata_modifier.py"
SAMPLES_ZIP = HERE / "office_metadata_validation_samples.zip"
if not SCRIPT.is_file() or not SAMPLES_ZIP.is_file():
    raise SystemExit("请将本脚本、office_metadata_modifier.py 和 office_metadata_validation_samples.zip 放在同一目录。")

spec = importlib.util.spec_from_file_location("office_metadata_under_test", SCRIPT)
module = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = module
assert spec.loader is not None
spec.loader.exec_module(module)

passed = 0
def check(condition: bool, name: str) -> None:
    global passed
    if not condition:
        raise AssertionError(name)
    passed += 1
    print("PASS", name)

with tempfile.TemporaryDirectory(prefix="office_metadata_regression_") as temp_dir:
    root = Path(temp_dir)
    samples_root = root / "samples"
    samples_root.mkdir()
    with zipfile.ZipFile(SAMPLES_ZIP) as archive:
        archive.extractall(samples_root)
    sample_dir = samples_root / "office_metadata_validation_samples"

    # OOXML write/read/undo across Word, PowerPoint and Excel samples.
    for ext in ("docx", "pptx", "xlsx"):
        target = root / f"roundtrip.{ext}"
        shutil.copy2(sample_dir / f"verify.{ext}", target)
        before_hash = hashlib.sha256(target.read_bytes()).hexdigest()
        result = module.modify_office_file(str(target), {
            "author": "回归测试作者😀", "revision": "17",
            "modified": "2026-10-22 12:34:56", "total_time_minutes": "89",
        })
        values = module.get_office_metadata(str(target))
        check(values["author"] == "回归测试作者😀", f"{ext} author round-trip")
        check(int(values["revision"]) == 17 and int(values["total_time_minutes"]) == 89, f"{ext} integer fields round-trip")
        check(values["modified"] == "2026-10-22T12:34:56+08:00", f"{ext} modified time round-trip")
        with zipfile.ZipFile(target) as package:
            xml = package.read("docProps/core.xml")
            check(xml.startswith(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'), f"{ext} XML declaration")
            ET.fromstring(xml)
        module.restore_backup(str(target), result["backup"], force=True)
        check(hashlib.sha256(target.read_bytes()).hexdigest() == before_hash, f"{ext} exact-byte undo")

    # A glob-special source filename must match backups literally.
    special = root / "file[1].docx"
    shutil.copy2(sample_dir / "verify.docx", special)
    special_backup = Path(str(special) + ".bak")
    shutil.copy2(sample_dir / "verify.docx", special_backup)
    check(str(special_backup) in module.find_backups(str(special)), "glob-special backup discovery")
    try:
        module.validate_report_destination(str(special_backup), [str(special)])
    except ValueError:
        check(True, "CSV report protected from glob-special backup overwrite")
    else:
        check(False, "CSV report protected from glob-special backup overwrite")

    # Same core/app relationship target is rejected.
    malformed = root / "samepart.docx"
    with zipfile.ZipFile(sample_dir / "verify.docx") as source, zipfile.ZipFile(malformed, "w") as dest:
        for info in source.infolist():
            payload = source.read(info.filename)
            if info.filename == "_rels/.rels":
                payload = payload.replace(b'Target="docProps/app.xml"', b'Target="docProps/core.xml"')
            dest.writestr(info, payload)
    try:
        module.get_office_metadata(str(malformed))
    except ValueError as exc:
        check("指向同一部件" in str(exc), "same core/app part rejected")
    else:
        check(False, "same core/app part rejected")

    rel_root = ET.Element(module._qname(module.PKG_REL_NS, "Relationships"))
    ET.SubElement(rel_root, module._qname(module.PKG_REL_NS, "Relationship"), {
        "Id": "rId1", "Type": module.CORE_REL_TYPE, "Target": "docProps/other.xml"
    })
    try:
        module._ensure_relationship(rel_root, module.CORE_REL_TYPE, "docProps/core.xml")
    except ValueError:
        check(True, "relationship mismatch fails loudly")
    else:
        check(False, "relationship mismatch fails loudly")

    # Corrupt sidecars are not treated as old backups without sidecars.
    corrupt_target = root / "corrupt_manifest.docx"
    shutil.copy2(sample_dir / "verify.docx", corrupt_target)
    result = module.modify_office_file(str(corrupt_target), {"author": "test"})
    Path(result["manifest"]).write_text("{not-json", encoding="utf-8")
    try:
        module.restore_backup(str(corrupt_target), result["backup"], force=True)
    except module.BackupManifestCorruptError:
        check(True, "corrupt manifest explicitly refused even with force")
    else:
        check(False, "corrupt manifest explicitly refused even with force")

    # Backup naming attempts have a fixed upper bound.
    allocation_source = root / "allocation.docx"
    shutil.copy2(sample_dir / "verify.docx", allocation_source)
    for index in range(1000):
        Path(str(allocation_source) + (".bak" if index == 0 else f".bak.{index}")).touch()
    try:
        module._create_backup(str(allocation_source))
    except RuntimeError as exc:
        check("1000" in str(exc), "backup allocation retries are bounded")
    else:
        check(False, "backup allocation retries are bounded")

    # ZIP size limits are checked before CRC/decompression; mock package metadata only.
    fake_path = root / "fake.docx"
    fake_path.write_bytes(b"placeholder")
    real_is_zipfile = module.zipfile.is_zipfile
    real_zipfile = module.zipfile.ZipFile
    class FakeInfo:
        def __init__(self, filename, file_size=0):
            self.filename, self.file_size = filename, file_size
    class FakePackage:
        testzip_called = False
        def __enter__(self): return self
        def __exit__(self, *args): return False
        def infolist(self):
            return [FakeInfo("[Content_Types].xml"), FakeInfo("_rels/.rels"), FakeInfo("payload.bin", module.MAX_ZIP_UNCOMPRESSED_SIZE + 1)]
        def testzip(self):
            self.testzip_called = True
            raise AssertionError("oversized package must be rejected before testzip")
    fake = FakePackage()
    try:
        module.zipfile.is_zipfile = lambda path: True
        module.zipfile.ZipFile = lambda *args, **kwargs: fake
        try:
            module._validate_office_file(str(fake_path))
        except ValueError as exc:
            check("超过安全上限 2 GiB" in str(exc), "ZIP uncompressed-size limit")
        else:
            check(False, "ZIP uncompressed-size limit")
        check(not fake.testzip_called, "oversized ZIP rejected before decompression")
    finally:
        module.zipfile.is_zipfile = real_is_zipfile
        module.zipfile.ZipFile = real_zipfile

    # GUI collection must keep the original date string so downstream normalization is idempotent.
    class FakeVar:
        def __init__(self, value): self.value = value
        def get(self): return self.value
    class GuiStub:
        author_enabled = FakeVar(False)
        author_var = FakeVar("")
        last_modified_by_enabled = FakeVar(False)
        last_modified_by_var = FakeVar("")
        revision_enabled = FakeVar(False)
        revision_var = FakeVar("1")
        created_enabled = FakeVar(True)
        created_var = FakeVar("2026-10-20 02:23:18")
        modified_enabled = FakeVar(True)
        modified_var = FakeVar("2026-10-21 03:24:25")
        total_time_enabled = FakeVar(False)
        total_time_var = FakeVar("0")
    gui_updates = module.OfficeMetadataModifierApp._collect_updates_from_gui(GuiStub())
    check(gui_updates["created"] == "2026-10-20 02:23:18", "GUI preserves raw created timestamp")
    check(gui_updates["modified"] == "2026-10-21 03:24:25", "GUI preserves raw modified timestamp")
    normalized_gui = module.normalize_updates(gui_updates)
    check(normalized_gui["created"] == "2026-10-20T02:23:18+08:00", "GUI timestamp can be normalized downstream")
    gui_target = root / "gui_timestamp.docx"
    shutil.copy2(sample_dir / "verify.docx", gui_target)
    gui_plan = module.build_change_plan(str(gui_target), gui_updates)
    planned_gui_values = {row["key"]: row["after"] for row in gui_plan["rows"]}
    check(planned_gui_values["created"] == "2026-10-20T02:23:18+08:00", "GUI created timestamp preview regression")
    check(planned_gui_values["modified"] == "2026-10-21T03:24:25+08:00", "GUI modified timestamp preview regression")
    module.modify_office_file(str(gui_target), gui_updates)
    gui_readback = module.get_office_metadata(str(gui_target))
    check(gui_readback["created"] == "2026-10-20T02:23:18+08:00", "GUI created timestamp write/read regression")
    check(gui_readback["modified"] == "2026-10-21T03:24:25+08:00", "GUI modified timestamp write/read regression")

    # Exercise the complete start_modification orchestration with checkbox-like GUI values,
    # a real preview plan, confirmed preview, and a real write/readback against a temporary copy.
    gui_flow_target = root / "gui_flow_timestamp.docx"
    shutil.copy2(sample_dir / "verify.docx", gui_flow_target)
    class FlowGuiStub:
        pass
    flow_gui = object.__new__(module.OfficeMetadataModifierApp)
    flow_gui.selected_files = [str(gui_flow_target)]
    flow_gui.sync_var = FakeVar(False)
    flow_gui.sync_time_var = FakeVar("2026-10-20 02:23:18")
    flow_gui.author_enabled = FakeVar(False)
    flow_gui.author_var = FakeVar("")
    flow_gui.last_modified_by_enabled = FakeVar(False)
    flow_gui.last_modified_by_var = FakeVar("")
    flow_gui.revision_enabled = FakeVar(False)
    flow_gui.revision_var = FakeVar("1")
    flow_gui.created_enabled = FakeVar(True)
    flow_gui.created_var = FakeVar("2026-10-20 02:23:18")
    flow_gui.modified_enabled = FakeVar(True)
    flow_gui.modified_var = FakeVar("2026-10-21 03:24:25")
    flow_gui.total_time_enabled = FakeVar(False)
    flow_gui.total_time_var = FakeVar("0")
    flow_gui.summary_before_write_var = FakeVar(True)
    flow_gui.export_report_var = FakeVar(False)
    flow_gui.report_path_var = FakeVar("")
    flow_gui.root = types.SimpleNamespace(update_idletasks=lambda: None)
    flow_gui.progress = types.SimpleNamespace(configure=lambda **kwargs: None)
    flow_gui.status_label = types.SimpleNamespace(config=lambda **kwargs: None)
    flow_gui.notebook = types.SimpleNamespace(select=lambda tab: None)
    flow_gui.basic_tab = object()
    flow_gui.advanced_tab = object()
    flow_gui.file_status = {}
    flow_gui.clear_log = lambda: None
    flow_gui._append_log = lambda message: None
    flow_gui._refresh_file_tree = lambda: None
    preview_state = {"called": False, "confirmed": False}
    def gui_flow_preview(updates, sync_dt):
        preview_state["called"] = True
        plan = module.build_change_plan(str(gui_flow_target), updates, sync_dt)
        values_by_key = {row["key"]: row["after"] for row in plan["rows"]}
        check(values_by_key["created"] == "2026-10-20T02:23:18+08:00", "GUI full-flow preview includes created timestamp")
        check(values_by_key["modified"] == "2026-10-21T03:24:25+08:00", "GUI full-flow preview includes modified timestamp")
        preview_state["confirmed"] = True
        return True
    flow_gui._show_change_preview = gui_flow_preview
    original_showinfo = module.messagebox.showinfo
    original_showwarning = module.messagebox.showwarning
    flow_messages = []
    try:
        module.messagebox.showinfo = lambda *args, **kwargs: flow_messages.append(("info", args))
        module.messagebox.showwarning = lambda *args, **kwargs: flow_messages.append(("warning", args))
        module.OfficeMetadataModifierApp.start_modification(flow_gui)
    finally:
        module.messagebox.showinfo = original_showinfo
        module.messagebox.showwarning = original_showwarning
    flow_readback = module.get_office_metadata(str(gui_flow_target))
    check(preview_state["called"] and preview_state["confirmed"], "GUI start_modification ran preview and confirmation path")
    check(flow_gui.file_status[str(gui_flow_target)] == "处理成功", "GUI full-flow timestamp modification succeeded")
    check(flow_readback["created"] == "2026-10-20T02:23:18+08:00" and flow_readback["modified"] == "2026-10-21T03:24:25+08:00", "GUI full-flow timestamps persisted and read back")

    # A filesystem-time restoration error must not hide the original preview/read error.
    orig_capture = module._capture_filesystem_times
    orig_read_metadata = module.get_office_metadata
    orig_restore_times = module._restore_filesystem_times
    try:
        module._capture_filesystem_times = lambda path: {"platform": "posix", "atime_ns": 0, "mtime_ns": 0}
        def broken_read(path): raise ValueError("original damaged Office package error")
        def broken_restore(path, snapshot): raise OSError("secondary timestamp restore error")
        module.get_office_metadata = broken_read
        module._restore_filesystem_times = broken_restore
        try:
            module.build_change_plan(str(root / "does_not_need_to_exist.docx"), {"author": "test"})
        except ValueError as exc:
            check("original damaged Office package error" in str(exc), "preview preserves primary exception")
        else:
            check(False, "preview preserves primary exception")
    finally:
        module._capture_filesystem_times = orig_capture
        module.get_office_metadata = orig_read_metadata
        module._restore_filesystem_times = orig_restore_times

    # Exception augmentation must not discard the filename/errno of an OSError.
    os_error_target = root / "os_error_metadata.docx"
    shutil.copy2(sample_dir / "verify.docx", os_error_target)
    orig_write_package = module._write_package_metadata
    orig_restore_times = module._restore_filesystem_times
    def injected_file_not_found(source, output, updates):
        raise FileNotFoundError(2, "synthetic missing part", str(source))
    def injected_timestamp_restore_error(path, snapshot):
        raise OSError("synthetic timestamp restore error")
    try:
        module._write_package_metadata = injected_file_not_found
        module._restore_filesystem_times = injected_timestamp_restore_error
        try:
            module.modify_office_file(str(os_error_target), {"author": "test"})
        except FileNotFoundError as exc:
            check(exc.errno == 2 and exc.filename == str(os_error_target), "OSError errno/filename preserved")
            check("时间恢复失败" in str(exc), "OSError rollback warning remains visible")
        else:
            check(False, "OSError errno/filename preserved")
    finally:
        module._write_package_metadata = orig_write_package
        module._restore_filesystem_times = orig_restore_times

    # Missing app.xml must be emitted with a fixed ZIP timestamp for repeatable output.
    minimal_source = root / "minimal.docx"
    with zipfile.ZipFile(minimal_source, "w") as zf:
        zf.writestr("[Content_Types].xml", f'<Types xmlns="{module.CT_NS}"/>')
        zf.writestr("_rels/.rels", f'<Relationships xmlns="{module.PKG_REL_NS}"/>')
    output_a, output_b = root / "minimal_a.docx", root / "minimal_b.docx"
    module._write_package_metadata(str(minimal_source), str(output_a), {"total_time_minutes": 7})
    module._write_package_metadata(str(minimal_source), str(output_b), {"total_time_minutes": 7})
    with zipfile.ZipFile(output_a) as zf:
        new_info = zf.getinfo("docProps/app.xml")
        check(new_info.date_time == (1980, 1, 1, 0, 0, 0), "new OOXML part has deterministic ZIP timestamp")
    check(hashlib.sha256(output_a.read_bytes()).digest() == hashlib.sha256(output_b.read_bytes()).digest(), "repeatable OOXML package bytes")

    # Streaming guard checks bytes read, even if mocked directory metadata understates them.
    fake_stream_path = root / "fake_stream.docx"
    fake_stream_path.write_bytes(b"placeholder")
    real_is_zipfile = module.zipfile.is_zipfile
    real_zipfile = module.zipfile.ZipFile
    real_max = module.MAX_ZIP_UNCOMPRESSED_SIZE
    class SmallAdvertisedInfo:
        def __init__(self, filename): self.filename, self.file_size = filename, 1
    class StreamingFakePackage:
        def __enter__(self): return self
        def __exit__(self, *args): return False
        def infolist(self):
            return [SmallAdvertisedInfo("[Content_Types].xml"), SmallAdvertisedInfo("_rels/.rels"), SmallAdvertisedInfo("payload.bin")]
        def open(self, info, mode):
            return io.BytesIO(b"X" * 11 if info.filename == "payload.bin" else b"x")
    try:
        module.zipfile.is_zipfile = lambda path: True
        module.zipfile.ZipFile = lambda *args, **kwargs: StreamingFakePackage()
        module.MAX_ZIP_UNCOMPRESSED_SIZE = 10
        try:
            module._validate_office_file(str(fake_stream_path))
        except ValueError as exc:
            check("实际读取的解压数据超过" in str(exc), "streamed uncompressed size guard")
        else:
            check(False, "streamed uncompressed size guard")
    finally:
        module.zipfile.is_zipfile = real_is_zipfile
        module.zipfile.ZipFile = real_zipfile
        module.MAX_ZIP_UNCOMPRESSED_SIZE = real_max

    # Encrypted ZIP members raise RuntimeError in zipfile; expose a focused user-facing error.
    encrypted_path = root / "encrypted.docx"
    encrypted_path.write_bytes(b"placeholder")
    real_is_zipfile = module.zipfile.is_zipfile
    real_zipfile = module.zipfile.ZipFile
    class EncryptedFakePackage:
        def __enter__(self): return self
        def __exit__(self, *args): return False
        def infolist(self):
            return [SmallAdvertisedInfo("[Content_Types].xml"), SmallAdvertisedInfo("_rels/.rels")]
        def open(self, info, mode):
            raise RuntimeError("File is encrypted, password required")
    try:
        module.zipfile.is_zipfile = lambda path: True
        module.zipfile.ZipFile = lambda *args, **kwargs: EncryptedFakePackage()
        try:
            module._validate_office_file(str(encrypted_path))
        except ValueError as exc:
            check("可能已加密或需要密码" in str(exc), "encrypted ZIP converted to actionable error")
        else:
            check(False, "encrypted ZIP converted to actionable error")
    finally:
        module.zipfile.is_zipfile = real_is_zipfile
        module.zipfile.ZipFile = real_zipfile

    # A compression method unsupported by the running Python must not be misreported as encryption.
    unsupported_path = root / "unsupported_compression.docx"
    unsupported_path.write_bytes(b"placeholder")
    real_is_zipfile = module.zipfile.is_zipfile
    real_zipfile = module.zipfile.ZipFile
    class UnsupportedCompressionPackage(EncryptedFakePackage):
        def open(self, info, mode):
            raise NotImplementedError("That compression method is not supported")
    try:
        module.zipfile.is_zipfile = lambda path: True
        module.zipfile.ZipFile = lambda *args, **kwargs: UnsupportedCompressionPackage()
        try:
            module._validate_office_file(str(unsupported_path))
        except ValueError as exc:
            check("不支持的 ZIP 压缩方法" in str(exc), "unsupported ZIP compression gets targeted diagnostic")
            check("加密" not in str(exc), "unsupported ZIP compression is not misreported as encryption")
        else:
            check(False, "unsupported ZIP compression gets targeted diagnostic")
    finally:
        module.zipfile.is_zipfile = real_is_zipfile
        module.zipfile.ZipFile = real_zipfile

    # Defensive helper use must produce a clear ValueError instead of a KeyError when required parts are missing.
    missing_root_part = root / "missing_root_part.docx"
    with zipfile.ZipFile(missing_root_part, "w") as zf:
        zf.writestr("[Content_Types].xml", f'<Types xmlns="{module.CT_NS}"/>')
    try:
        with zipfile.ZipFile(missing_root_part, "r") as package:
            module._load_package_parts(package)
    except ValueError as exc:
        check("缺少必要的包级部件" in str(exc) and "_rels/.rels" in str(exc), "OOXML package-part loader validates its preconditions")
    else:
        check(False, "OOXML package-part loader validates its preconditions")

    # In a headless POSIX environment, GUI startup should fail with a concise message, not a traceback.
    if os.name == "posix":
        headless_env = os.environ.copy()
        headless_env.pop("DISPLAY", None)
        headless_env.pop("WAYLAND_DISPLAY", None)
        no_gui = subprocess.run([sys.executable, str(SCRIPT)], env=headless_env, capture_output=True, text=True, encoding="utf-8", errors="replace")
        check(no_gui.returncode == 1 and "无法启动 GUI" in no_gui.stderr and "TclError" not in no_gui.stderr, "headless GUI startup handled")

    # Damaged-manifest backups are reported as security skips, not ordinary restore failures.
    gui_skip_target = root / "gui_skip.docx"
    shutil.copy2(sample_dir / "verify.docx", gui_skip_target)
    shutil.copy2(sample_dir / "verify.docx", Path(str(gui_skip_target) + ".bak"))
    Path(str(gui_skip_target) + ".bak.manifest.json").write_text("{damaged", encoding="utf-8")
    gui_messages = []
    gui_logs = []
    gui_status = {}
    gui_stub = object.__new__(module.OfficeMetadataModifierApp)
    gui_stub.root = object()
    gui_stub.file_status = {}
    gui_stub._selected_paths_in_tree = lambda: [str(gui_skip_target)]
    gui_stub._append_log = lambda message: gui_logs.append(message)
    gui_stub._refresh_file_tree = lambda: None
    gui_stub.status_label = types.SimpleNamespace(config=lambda **kwargs: gui_status.update(kwargs))
    old_askyesno = module.messagebox.askyesno
    old_showwarning = module.messagebox.showwarning
    old_showinfo = module.messagebox.showinfo
    try:
        module.messagebox.askyesno = lambda *args, **kwargs: True
        module.messagebox.showwarning = lambda *args, **kwargs: gui_messages.append(args)
        module.messagebox.showinfo = lambda *args, **kwargs: gui_messages.append(args)
        module.OfficeMetadataModifierApp.undo_selected_files(gui_stub)
    finally:
        module.messagebox.askyesno = old_askyesno
        module.messagebox.showwarning = old_showwarning
        module.messagebox.showinfo = old_showinfo
    gui_summary = str(gui_messages[-1][1]) if gui_messages else ""
    check("失败：0" in gui_summary and "安全跳过：1" in gui_summary, "GUI restore summary separates safe skip from failures")
    check("安全跳过 1" in gui_status.get("text", ""), "GUI status bar counts safe skips")

    # Damaged manifests are explicitly excluded from the GUI force-restore path.
    source_code = SCRIPT.read_text(encoding="utf-8")
    check("该备份将被跳过，不会进入强制恢复流程" in source_code and "if path in corrupt_manifests" in source_code, "GUI skips corrupt-manifest backups clearly")

    cli = subprocess.run(
        [sys.executable, str(SCRIPT), "--file", str(sample_dir / "verify.docx"), "--read"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    check(cli.returncode == 0 and "读取完成：成功 1，失败 0" in cli.stdout, "CLI read success count")

print(f"\n回归测试完成：{passed} 项通过。测试只处理临时副本。")
