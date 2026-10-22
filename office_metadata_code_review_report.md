# Office 元数据修改工具：代码审查报告

审查对象：`office_metadata_modifier.py`（Tkinter GUI、CLI、OOXML XML/ZIP 读写、备份/撤销、CSV 报告、时间处理）。

## 问题清单（按严重程度）

| 严重程度 | 位置 | 类型 | 问题与原因 | 修改结果 |
|---|---|---|---|---|
| 高 | `validate_report_destination()` 第 1150 行附近；`write_csv_report()` 第 1176 行附近；`run_cli()` 第 1228 行附近；GUI `start_modification()` 第 2063 行附近 | 文件破坏/路径校验 | 报告路径若与输入 Office 文件相同，旧代码会在处理结束后用 `open(..., "w")` 截断原文件；若指向旧 `.bak` 或撤销清单，也可能破坏恢复材料 | 增加规范化路径与 `samefile` 校验，保护输入文件、现有 `.bak*` 备份及清单；CLI/GUI 先检查输入文件冲突，实际写报告时再检查 |
| 高 | `_atomic_replace_preserving_target_metadata()` 第 563 行附近；`modify_office_file()` 第 853 行附近 | 文件系统元数据/安全 | 直接 `os.replace` 会使用临时文件替换原文件，可能改变 POSIX 权限、所有者/组及可复制属性；Windows 下可能丢失原文件 ACL；符号链接/硬链接可能出现与预期不一致的替换语义 | Windows 使用 `ReplaceFileW`；POSIX 替换前尽量复制所有者/组、模式、时间与 `copystat` 支持的属性。拒绝符号链接与硬链接，避免静默破坏链接关系 |
| 高 | `_create_backup()` 第 611 行附近 | 文件系统/竞态 | 旧备份名分配采用“先检查再复制”，有路径竞争窗口；悬空符号链接可能被误当成不存在的目标 | 改为 `O_CREAT | O_EXCL` 排他创建，使用已打开文件描述符写入，避免复用现存路径 |
| 中 | `modify_office_file()` 第 898、930 行附近 | 并发/一致性 | 另一个程序在本工具读取、备份或写入临时文件期间修改源文件时，旧逻辑可能覆盖未纳入本次备份的内容 | 备份前后以及最终替换前比较 SHA-256；检测到变化就停止并保留备份 |
| 中 | `build_change_plan()` 第 1037 行附近 | 边界条件/文件时间 | 只读预演会读取并校验整个 ZIP 包，可能改变文件访问时间 | 预演前保存文件时间，结束后尽力恢复访问/修改时间；POSIX ctime 由系统管理，无法保证回退 |
| 中 | `restore_backup()` 第 724 行附近；GUI `undo_selected_files()` 第 1702 行附近 | 异常处理/状态准确性 | 文件内容替换成功后若恢复时间戳失败，旧实现抛出异常，界面会把内容已恢复的文件错误标记为恢复失败 | 返回 `timestamp_warning`；GUI 把它显示为时间戳警告，而非把内容恢复本身计为失败 |
| 中 | `_read_total_time()` 第 262 行附近；`_write_package_metadata()` 第 384 行附近 | 数据一致性/解析 | 读取逻辑可能将其他命名空间的同名 `TotalTime` 误认作 Office 字段；旧版本遗留的 `cp:totalTime` 可能在标准 `app.xml/TotalTime` 更新后保留旧值 | 精确匹配标准命名空间；保留无命名空间旧式读取兼容；写入时同步已有旧版 `cp:totalTime` 和无命名空间旧式 `TotalTime` |
| 中 | `run_cli()` 第 1228 行附近；GUI 批处理汇总第 2127 行附近 | 参数验证/统计 | `--read` 与修改参数组合可能导致用户误以为字段已处理；报告写入失败还会污染文件失败数，使失败数与总文件数不一致 | 拒绝冲突参数组合；将报告输出错误与 Office 文件处理失败分开统计，报告故障仍返回非零退出码 |

## 已执行验证

- `python -m py_compile office_metadata_modifier.py`：通过。
- `python office_metadata_modifier.py --help`：通过。
- 对有效 `.docx`、`.pptx`、`.xlsx` 样例执行字段写入、回读验证、自动备份和撤销；撤销后原始 SHA-256 与写入前一致：通过。
- POSIX 环境下替换后权限位、UID、GID保持：通过。
- 报告路径指向输入 Office 文件、已有备份时均会拒绝；被保护目标的 SHA-256 保持不变：通过。
- 悬空符号链接占用首选备份路径时，备份器选择下一序号，没有写入链接目标：通过。
- `--dry-run` 前后 atime/mtime 一致：通过。
- 符号链接/硬链接目标被拒绝且原内容/链接保持：通过。
- 模拟处理期间其他程序改写源文件，工具检测变化且没有覆盖该版本：通过。
- 模拟恢复内容成功但时间戳恢复失败，工具返回明确警告：通过。
- 旧版 `cp:totalTime` 与标准 `app.xml/TotalTime` 同步：通过。
- CLI 模式冲突、报告输出失败统计：通过。
- Xvfb 下 Tkinter GUI 启动，两个标签页存在：通过。
- AST 静态导入审计未发现 HTTP、上传、遥测或常见网络请求依赖：通过。

## 尚未验证 / 疑似风险

1. 当前执行环境不是 Windows 桌面环境，未实际运行 `ReplaceFileW`、Windows ACL、Windows 创建时间 API，也未在 Microsoft Office GUI 中实测；这些必须在目标 Windows 机器上用测试副本验证。
2. SHA-256 检查是乐观并发保护，不是操作系统级排他锁。最终检查与替换之间仍存在极短竞态；若文件正被 Word、同步客户端或其他程序写入，建议先关闭相关程序。
3. POSIX 不同文件系统对 ACL、扩展属性的支持差异较大；`shutil.copystat` 只复制当前平台支持的属性，不保证复制所有特定文件系统安全标签。
4. 预演在 POSIX 上可尽力恢复 atime/mtime，但 ctime 无法由普通程序回退。
5. 本次样例和异常注入测试不能代替用户真实文件、真实权限条件和 Windows Office 本体验证。

## 第二轮边界审查（2026-10-22）

以下修复针对另一份逐项代码审查意见，按严重程度补充。代码已再次修改，旧的行号仅可用于定位函数名，具体行号以下载的最新版为准。

| 严重程度 | 位置 | 类型 | 问题与原因 | 修改结果 / 验证 |
|---|---|---|---|---|
| 高 | `find_backups()`、`validate_report_destination()` | 路径解析/误匹配 | 原始路径包含 `[`、`]`、`*`、`?` 时，`glob` 会将其解释为模式字符，可能漏找真实备份或误保护/误匹配其他文件 | 对输入文件路径使用 `glob.escape()`，只让末尾 `.bak*` 保持为通配模式；`report[1].docx` 测试通过 |
| 高 | `_load_package_parts()`、`_write_package_metadata()` | XML 关系/静默数据丢失 | 恶意或损坏的关系文件可使 core-properties 与 extended-properties 指向同一部件，第二次序列化可能覆盖第一次修改 | 加载部件时发现路径相同即拒绝；构造同目标关系的 DOCX 测试通过 |
| 中 | `_ensure_relationship()` | 错误处理 | 已有关系若指向与调用方期望不同的部件，原函数静默返回 `False`，让调用方误以为关系状态可接受 | 改为抛出 `ValueError`，直接单元场景测试通过 |
| 中 | `_read_backup_manifest()`、`restore_backup()`、GUI 撤销预览 | 备份完整性/状态提示 | 清单存在但损坏时原函数返回 `None`，与真正没有清单的旧版备份混为一谈 | 新增 `BackupManifestCorruptError`；损坏清单显示明确警告，恢复阶段拒绝把它当旧版清单缺失处理。强制恢复也不能绕过备份清单损坏检查；需先核实备份并单独移走损坏清单，之后才能按旧版流程二次确认 |
| 中 | `_create_backup()` | 资源/活锁 | 备份名分配的 `while True` 在大量文件名冲突或持续抢占时没有上限 | 限制为 1000 个候选名，耗尽后抛出可读错误；占满候选名测试通过 |
| 低 | `_xml_bytes()` | XML 兼容 | XML 头未显式带 `standalone="yes"`；一般并非必然错误，但可统一序列化形式 | 固定输出 UTF-8、`standalone="yes"` 声明；DOCX/PPTX/XLSX 均通过重新打开测试 |
| 低（原修复后来发现是回归） | `_collect_updates_from_gui()` | 数据规范化/一致性 | 初版修复让该函数直接返回 `normalize_updates(updates)`，但时间字段第一次规范化后变为带偏移的 ISO 字符串，下游再次调用严格格式解析器时会失败；具体修正见下方“第三轮复审” | 最终版改为仅调用 `normalize_updates(updates)` 做验证，再返回原始输入，让预演与写入各自规范化一次 |
| 低 | `start_modification()` | GUI 可用性 | 总编辑时间输入错误时仍切换到高级页，而该字段实际在基础页 | 根据错误信息所含字段标签选择基础/高级标签页；语法和 GUI smoke test 通过，真实弹窗交互仍需 Windows 人工检查 |
| 低 | `run_cli()` 的 `--read` 分支 | 统计一致性 | 读取模式依赖 `len(files)-failed` 算成功数，虽当前结果正确，但与其他操作的计数方式不同 | 读取成功时显式递增 `success`，最终统一输出该计数；CLI 单文件回读输出测试通过 |
| 中 | `_validate_office_file()` | 资源耗尽/拒绝服务 | `ZipFile.testzip()` 会解压检查所有部件，恶意 ZIP 可声明极大解压体积或海量部件 | 新增部件数上限 100,000、声明解压后总大小上限 2 GiB，并先做名称去重和大小检查，再进行 CRC 解压校验；模拟超限包证明会在 `testzip()` 前拒绝 |
| 低 | GUI `undo_selected_files()`、使用指南 | 语义准确性 | POSIX 上文件内容与 atime/mtime 可恢复，但 inode ctime 无法精确回退，界面原提示容易被理解为所有时间戳均恢复 | 在 POSIX 撤销日志明确指出 ctime 无法回退，并更新使用指南 |
| 低 | `OfficeMetadataModifierApp._build_advanced_tab()` | 冗余初始化/维护性 | 作者等 `StringVar` 在 `__init__` 和高级页构建中重复定义 | 删除高级页中重复初始化，统一由 `__init__` 创建；GUI smoke test 通过 |

### 审查中确认无须单独修改的项目

- **`xsi:type="dcterms:W3CDTF"` 的 `dcterms` 前缀**：当前模块在启动时已注册 `dcterms` 命名空间；创建或更新 `dcterms:created` / `dcterms:modified` 时，序列化器会同时输出该命名空间声明。实测“只更新 modified”后 XML 仍包含合法的 `dcterms:modified` 前缀，且三种第三方 Office 文件库均可重新打开。因此，当前代码中这不是独立的可复现缺陷；没有手工塞入 `xmlns:dcterms`，以避免 ElementTree 重复声明。
- **`force=True` 仍然校验备份文件哈希**：这是正确的安全约束，保留不改。force 只允许跳过“当前文件是否被后续改动”的冲突检查，不应跳过“备份是否完整/是否被替换”的检查。
- **CLI 增加 `--restore`、允许跳过预演失败文件、GUI 后台线程、改用 `%APPDATA%` 存放预设**：这些属于功能扩展或 UX 设计，不是当前必然错误。为遵守“只修复错误、不改变外部接口与原有意图”的要求，本轮未实施。

### 本轮执行的验证

- `python -m py_compile`：通过。
- 原有三类文件回归 `run_tests.py`：通过。
- 本轮 `round3_regression.py`：通配符路径、core/app 同路径、关系目标冲突、损坏清单、1000 次备份名上限、XML 声明、ZIP 解压上限、CLI 读取成功计数全部通过。
- 三类 Office 库 smoke test：分别修改 `modified` 与清空作者字段，确认 `standalone="yes"` 和 `dcterms:modified` 输出，使用 `python-docx`、`python-pptx`、`openpyxl` 重新打开，并通过备份恢复校验；全部通过。
- Xvfb GUI smoke test：双标签页、规范化字段值采集通过。

### 仍未验证

本轮环境不是 Windows 桌面系统，依然未验证 Windows `ReplaceFileW`、实际 ACL/创建时间语义、正在被 Office 占用或只读文件的真实交互，以及 Microsoft Office 桌面应用本体。ZIP 2 GiB / 100,000 部件是新增的安全边界，超出上限的合法大型 Office 包也会被拒绝；如业务确需处理更大文件，应在了解资源风险后再调整常量并测试。


## 第三轮复审与修复（2026-10-22）

本轮依据修订版复审报告重新审查并直接修改 `office_metadata_modifier.py`，新增测试已并入便携式 `office_metadata_regression_test.py`。

| 严重程度 | 位置 | 问题 | 修复与验证 |
|---|---|---|---|
| 高 | `OfficeMetadataModifierApp._collect_updates_from_gui()` | 先前直接返回规范化后的 ISO 日期字符串；`build_change_plan()` / `modify_office_file()` 又执行严格规范化，导致 GUI 勾选创建时间或最后保存时间时失败 | 现在仅用 `normalize_updates(updates)` 验证字段，然后返回原始 GUI 值；回归测试覆盖时间字段采集、预演以及实际写入后回读，均通过 |
| 中 | `build_change_plan()` 的文件时间恢复 `finally` | 文件读取异常与时间恢复异常同时发生时，后者会掩盖原始损坏/读取异常 | 时间恢复异常被局部捕获，保留主要读取异常；通过异常注入测试 |
| 中 | 程序入口 / `launch_gui()` | 无图形环境时直接输出 Tk `TclError` 堆栈 | 入口捕获 `tk.TclError`，输出中文原因和 `--help` 提示并返回退出码 1；无 `DISPLAY` 的 POSIX 子进程测试通过 |
| 低 | GUI `undo_selected_files()` | 损坏撤销清单无法走强制恢复流程，但预览表述容易让用户误以为可以二次强制覆盖 | 预览中明确该备份会跳过且不会进入强制恢复流程；执行循环主动跳过并记录原因。清单完整性校验仍不可被 `force=True` 绕过 |
| 低 | `_write_package_metadata()` 新建 ZIP 部件 | `writestr(name, payload)` 会使用当前时间戳，导致重复生成的包字节可能不同 | 对新增部件使用固定 ZIP 时间戳 `1980-01-01 00:00:00`；从不含 `app.xml` 的最小样例重复生成两次，整包 SHA-256 一致 |
| 低 | `_write_package_metadata()` 日期类型属性 | `dcterms` 声明依赖 ElementTree 对同一根节点命名空间的序列化行为 | 增加注释记录前缀注册与对应 DCTERMS 元素之间的契约；现有输出测试确认声明合法，不手工追加 `xmlns`，避免命名空间重复声明 |
| 低 | `_validate_office_file()` | 之前只依赖 ZIP central directory 声明的大小上限 | 在声明大小预筛之后按块读取各 ZIP 部件并累计读取字节，同时触发 CRC 校验；测试用模拟 ZIP 流确认流式累计超限时中止。该方案能强化一般异常包防护，但 Python `zipfile` 仍遵循目录中声明的单部件大小，**不是用于解析任意对抗性 ZIP 的完整沙箱** |
| 低 | `_format_field_value()` | 无调用点的未使用辅助函数 | 删除未使用函数；不影响对外接口 |

### 本轮验证结果

- `python -m py_compile office_metadata_modifier.py office_metadata_regression_test.py`：通过。
- `python office_metadata_regression_test.py`：**37 项通过**，覆盖 DOCX/PPTX/XLSX 写入回读和精确撤销、特殊字符路径、关系异常、损坏清单、备份名上限、ZIP 声明/流式大小限制、GUI 时间字段采集、预演、真实样例写入回读、异常优先级、固定 ZIP 时间戳、重复输出字节稳定、无显示环境启动提示及 CLI 读取成功统计。测试只操作临时副本。
- `python office_metadata_modifier.py --help`：通过。
- 已生成最终版工具包并进行 ZIP 完整性检查。

### 本轮保留的设计边界

- 遇到撤销清单损坏时仍拒绝恢复，不提供绕过完整性检查的“强制恢复”选项；GUI 现在会清晰告知并跳过。
- 未新增 CLI `--restore`、跳过失败预演的开关、后台线程或 `%APPDATA%` 配置迁移。这些属于功能/UX 变化，当前需求是修错且保持接口。
- 当前环境不是 Windows 桌面环境，仍需用户在 Windows 本机验证 `ReplaceFileW`、ACL、文件系统创建时间、占用/只读文件行为及 Microsoft Office 本体。


## 第四轮复审与增补修复（2026-10-22）

针对第三轮复审报告中未阻塞发布的项目，继续修改 `office_metadata_modifier.py` 并扩展 `office_metadata_regression_test.py`。

| 严重程度 | 位置 | 问题 | 修改与验证 |
|---|---|---|---|
| 低 | `_validate_office_file()` | 加密 ZIP 部件在 `package.open()` 期间会抛 `RuntimeError`，诊断信息不够直接 | 捕获 ZIP 读取期间的 `RuntimeError`，转为包含“可能已加密或需要密码”的 `ValueError`；模拟加密条目测试通过 |
| 低 | `OfficeMetadataModifierApp.undo_selected_files()` | 损坏撤销清单的备份被跳过，但被计入普通失败，容易让用户误判 | 新增 `skipped` 统计；弹窗、状态栏分别列出失败数、安全跳过数和时间戳警告。通过真实方法调用加 mock 消息框的集成测试 |
| 低 | `_validate_office_file()` | 大型 Office 包逐条读取时 Python 层循环次数较多 | 将读取块从 1 MiB 调整为 4 MiB，保留逐条 CRC 与实际解压累计限制。尚未跑多 GB 文件性能基准，故不宣称实测加速 |
| 低 | `modify_office_file()` 异常回报 | 重建 `FileNotFoundError` / `PermissionError` 可能丢失 `errno`、`filename` 等属性 | 原地更新 `OSError.strerror` 以附加时间恢复警告，然后重新抛出原异常实例；异常注入验证属性保留且警告可读 |

### 最新回归

- `python -m py_compile office_metadata_modifier.py office_metadata_regression_test.py`：通过。
- `python office_metadata_regression_test.py`：**42 项通过**。新增覆盖加密 ZIP 错误转换、`OSError` 属性保留、损坏清单安全跳过计数。
- CLI `--help`：通过。
- 从更新后的离线 ZIP 中核对主程序、回归测试、样例和报告版本；ZIP `testzip()`：通过。

### 明确不改的事项

- 两个元数据读取入口保留各自职责；减少少量重复代码不值得扩大本轮改动范围。
- CLI `--restore`、批处理跳过失败预演、GUI 后台线程与 `%APPDATA%` 迁移仍属于功能/UX 增强，不是必须修复的 bug，本轮未改。
- 4 MiB 分块尚无大型文件基准数据；这只是降低 Python 读取循环开销的取舍，并非经基准测试证明的性能提升。
- 最新测试仍在 Linux/POSIX 环境完成，未替代 Windows 桌面 Office 本体验收。

## 第五轮复审与修复（2026-10-22）

根据第四轮复审报告，本轮只处理边界诊断和重复逻辑，不增加 CLI 新接口、不改变现有 GUI/CLI 的功能意图。

| 严重程度 | 位置 | 问题 | 修复与验证 |
|---|---|---|---|
| 低 | `_validate_office_file()` | `zipfile` 遇到当前 Python 不支持的 ZIP 压缩方法会抛 `NotImplementedError`。它是 `RuntimeError` 的子类，如果放在通用 `RuntimeError` 之后会被误报为加密/需要密码 | 新增更早的 `except NotImplementedError`，提示“当前 Python zipfile 不支持的 ZIP 压缩方法”；模拟不支持压缩方法测试通过，并确认不会误报为加密 |
| 低 | `_load_package_parts()` | 此内部函数隐式依赖调用者先验证包级部件存在；被直接或未来被其他调用路径使用时，缺少 `_rels/.rels` / `[Content_Types].xml` 会以 `KeyError` 形式暴露 | 函数内显式检查两个必要部件，缺失时抛出可读 `ValueError`；直接调用内部函数的缺失部件测试通过 |
| 低 | `get_office_metadata()`、`_read_metadata_without_os_info()` | core/app XML 解析和旧版 `cp:totalTime` 回退逻辑重复，维护时容易一处更新、另一处遗漏 | 提取私有 `_read_core_and_app()`，两个入口共用 XML 读取；`get_office_metadata()` 仍先做完整 ZIP 校验并独立添加 `file_type`、Windows 创建时间；临时输出验证入口仍不重复做全包预校验。三类文件读写/撤销回归通过 |
| 建议测试 | GUI `start_modification()` | 需要锁定“创建时间/最后保存时间复选框输入 → 预演 → 确认 → 写入 → 回读”的核心回归路径 | 扩展回归测试：调用真实 `start_modification()` 编排和真实 `modify_office_file()`，对预演窗口使用替身回调以模拟确认；写入后重新读取两项时间并与期望值比较，测试通过。此项不等价于 Windows 桌面上人工点击真实 Tk 窗口 |

### 本轮验证

- `python -m py_compile office_metadata_modifier.py office_metadata_regression_test.py`：通过。
- `python office_metadata_modifier.py --help`：通过。
- `python office_metadata_regression_test.py`：**50 项通过**。在此前 42 项基础上，新增不支持压缩方法的定向错误/不误报验证、包级部件缺失的防御检查，以及 `start_modification()` 的预演—确认—真实写入与回读验证。
- 三种 Office 文件写入、回读、精确撤销与输出字节稳定测试继续通过。
- 本轮回归测试只处理临时目录中的样例副本。

### 仍未验证的边界

当前仍为 Linux/POSIX 容器环境。真实 Windows 桌面交互、`ReplaceFileW`、Windows ACL、Windows 文件系统创建时间，以及 Microsoft Word/PowerPoint/Excel 本体验证仍需在目标 Windows 电脑上完成。GUI 编排测试中的预演确认使用测试替身，不应描述成已人工点击真实窗口。
