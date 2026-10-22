# Office 元数据修改工具使用说明

## 支持范围

- 文件类型：`.docx`、`.pptx`、`.xlsx`。
- 支持读取/写入：作者、最后一次保存者、修订号、创建时间、最后一次保存时间、总编辑时间（分钟）。
- 原有同步模式：在 Windows 上同时设置文件系统创建时间、Office 内部创建时间和最后保存时间。
- 修改前自动在同目录生成不覆盖已有备份的 `.bak` 文件；写入使用临时文件和原子替换，并进行写后读取验证。
- 仅使用 Python 标准库；无需安装 `python-docx`、`python-pptx` 或 `openpyxl` 才能运行工具。

## 图形界面

双击运行或在命令行执行：

```powershell
python office_metadata_modifier.py
```

原有同步模式默认勾选，默认时间延续原脚本的 `2026-10-20 02:23:18`（UTC+8）。总编辑时间默认勾选为 0 分钟，以保留原脚本的默认行为。

界面分为两个标签页：

- **基础功能**：原有时间同步、设置总编辑时间、使用当前 UTC+8 时间、检查日期格式。
- **高级功能**：分别编辑作者、最后保存者、修订号、Office 创建时间和最后保存时间；设置递归扫描、写入前变更摘要和 CSV 报告；管理本机配置预设。

文件列表支持一次添加多个文件、追加文件夹、移除选中项和清空列表，并显示文件类型与处理状态。点击“读取元数据预览”只读取当前值；默认开启的“写入前显示变更摘要”会打开字段对比窗口，展示每个文件的原值和目标值，预演通过并由用户确认后才会写入。详细日志可通过“查看日志”打开。

选中一个或多个文件后，可以使用“恢复最近备份”撤销修改。新版备份旁会保存本机 JSON 撤销清单，记录修改前文件系统时间与修改后文件哈希；如果当前文件在修改后又发生了变化，工具会拒绝直接覆盖并要求二次确认。旧版本备份没有此清单，恢复时也会要求二次确认，且无法保证还原原始 Windows 创建时间。

配置预设只保存在本机用户配置目录 `~/.office_metadata_modifier/presets.json`（Windows 下通常是 `%USERPROFILE%\.office_metadata_modifier\presets.json`），只保存字段开关和输入值，不保存文件列表或文件正文。预设不会上传。

**隐私/联网说明：**本工具不包含 HTTP 请求、文件上传、遥测或云端接口，运行和读写元数据不需要互联网连接；Office 文件、备份、报告、撤销清单及预设都由本机 Python 进程写入本地文件系统。若文件本身放在 OneDrive、网盘同步目录或网络驱动器，外部同步客户端仍可能自行同步该目录，这不由本工具触发或控制。

只想改某一项时，请取消“原有时间同步模式”，并只勾选需要修改的字段。未勾选的独立字段不会被修改。

## 命令行示例

只读取现有属性：

```powershell
python office_metadata_modifier.py --file "C:\\Users\\liuar\\Desktop\\harness\\example.docx" --read
```

单独更新作者和修订号：

```powershell
python office_metadata_modifier.py --file "C:\\Users\\liuar\\Desktop\\harness\\example.docx" --author "测试作者" --revision 5
```

批量更新某个文件夹中的 Word、PPT 和 Excel 文件，并导出 CSV 报告：

```powershell
python office_metadata_modifier.py --folder "C:\\Users\\liuar\\Desktop\\harness" --author "测试作者" --last-modified-by "保存者" --revision 5 --created "2026-10-21 12:34:56" --modified "2026-10-22 14:01:02" --total-time-minutes 90 --report "C:\\Users\\liuar\\Desktop\\harness\\metadata_report.csv"
```

沿用原有“生成时间同步”功能：

```powershell
python office_metadata_modifier.py --file "C:\\Users\\liuar\\Desktop\\harness\\example.docx" --sync-time "2026-10-20 02:23:18" --total-time-minutes 0
```

递归处理子文件夹时增加 `--recursive`。单独修改 Office 创建/最后保存时间时，使用 `--created` 或 `--modified`，它们不会顺带改 Windows 文件系统创建时间。

## 校验规则与注意事项

- 时间输入格式固定为 `YYYY-MM-DD HH:MM:SS`，按 UTC+8 解释，不接受附加时区文本。
- 修订号与总编辑时间均为 0 至 2147483647 的整数；总编辑时间单位固定为分钟。
- `--read` 是只读操作，不能与写入参数同时使用。
- 对属性缺失的文件会创建对应 XML 属性；若整个 `docProps/core.xml` 或 `docProps/app.xml` 部件缺失，在需要写入时也会创建该部件及必要关系。
- 本工具针对现代 OOXML 文件，不支持旧版 `.doc`、`.ppt`、`.xls`，也不处理宏启用的 `.docm`、`.pptm`、`.xlsm`。
- Windows 文件系统创建时间的同步功能依赖 Windows API；在非 Windows 系统上会明确拒绝执行，而不会假装成功。
- 撤销时优先选择工具生成的最新本地备份；新版本备份通过 SHA-256 防止误覆盖修改后的文件，用户显式确认后才允许强制恢复。若撤销清单文件存在但损坏/不可解析，工具会明确拒绝将其当作“旧版无清单备份”；先核实备份确实可信，再将损坏的 `.manifest.json` 清单单独移走后，才可按旧版备份流程进行二次确认恢复。
- ZIP 包校验前会检查部件数和声明的解压后总大小：最多 100,000 个部件、总解压大小不超过 2 GiB。超过限制的超大型 Office 包会被拒绝处理，以降低 ZIP 炸弹和异常包造成资源耗尽的风险。 校验过程中还会按块读取部件并累计读取到的解压字节。该措施是纵深防护，不是针对任意恶意 ZIP 的完整安全沙箱；Python `zipfile` 仍依赖 ZIP 目录中的单部件大小字段。
- 备份旁的 JSON 撤销清单包含原始文件路径、修改前后元数据和哈希；这些信息可能含有用户名、作者名或本地目录结构。请像保护原始文档一样妥善保管备份与清单。
- 备份会尽可能复制原文件的时间和权限属性；在某些 POSIX 文件系统上，`copystat` 可能连同 setuid/setgid 等特殊权限位一起复制。请仅将备份保存在可信目录中。
- 应用程序本体验证仍需在目标 Windows 电脑的 Microsoft Word、PowerPoint、Excel 中打开样例或实际文档检查。当前测试环境未运行 Windows 桌面 Office，因此报告不会把 OOXML/第三方 Python 库验证表述为 Office 本体验证。


## 编码兼容说明（GBK / ANSI / UTF-8）

- **Python 源码：统一使用 UTF-8。** 文件开头声明了 `# -*- coding: utf-8 -*-`。不要把 `.py` 另存为 GBK/ANSI；若编辑器提示编码，请选择“以 UTF-8 保存”。把源码改存为 GBK/ANSI 可能导致语法错误或中文字符串损坏。
- **Office 内部元数据：使用 OOXML/XML 的 UTF-8 序列化。** 作者等字段中的中文和其他 Unicode 字符会写入 XML，不依赖 Windows 当前 ANSI 代码页。
- **本机配置预设：JSON 使用 UTF-8 读写。** 不需要根据系统代码页转换。
- **CSV 报告：采用 UTF-8 with BOM（`utf-8-sig`）。** 这种形式通常更方便 Windows Excel 直接识别中文；其他 CSV 工具也可以选择 UTF-8 导入。
- **图形界面：** Tkinter 字段使用 Unicode。若个别字显示为方框，通常是系统字体缺字或字体回退问题，不一定是文件编码错误。
- **命令行：** 程序不会强行把终端改为 UTF-8，以免破坏仍在使用 GBK/系统代码页的终端；同时启用安全输出策略，遇到当前代码页无法表示的字符时使用 `\u` / `\U` 转义，避免因 `UnicodeEncodeError` 中断。多数中文在 GBK 中可正常显示，但终端本身配置错误仍可能导致显示乱码。

`ANSI` 不是唯一固定编码，它通常表示当前 Windows 系统的活动代码页；简体中文 Windows 上常见的是 Windows-936/GBK。建议保留源码 UTF-8，不要为了适配控制台而改写 Office XML 或源码编码。若在旧版 CMD 中连中文路径/参数都显示异常，可先执行 `chcp 65001` 再运行；使用现代 PowerShell 或 Windows Terminal 通常更省事。


## 本轮代码审查补充（2026-10-22）

- CSV 报告路径不能与待处理 Office 文件、其现有 `.bak*` 备份或撤销清单重合；命令行和图形界面会在开始处理前检查输入文件冲突，写报告时还会再检查一次。
- 为避免意外破坏文件链接关系，修改和恢复操作拒绝符号链接或硬链接文件。请先将其复制为普通独立文件，再处理副本。
- 文件替换时会尽可能保留原文件的权限/安全属性：Windows 使用 `ReplaceFileW`，POSIX 复制所有者/组及 `copystat` 支持的属性。若 POSIX 所有者或组无法保留，工具会在替换前中止。
- 修改前后增加 SHA-256 检查，尽量发现 Word、同步客户端或其他程序在处理期间改写目标文件的情况。它不是操作系统级锁，不能彻底消除最后校验与替换之间的并发竞态；处理时建议关闭正在编辑目标文件的 Office 程序。
- 对旧版遗留的非标准 `cp:totalTime`，更新总编辑时间时会同步其数值，防止与标准 `docProps/app.xml/TotalTime` 不一致。
- 文件名包含 `[`、`]`、`*`、`?` 等通配符字符时，备份发现和报告路径校验会将原始路径按字面值处理，不会误匹配其他备份。
- 如果 Office 包的 core-properties 和 extended-properties 关系指向同一部件，或已有关系指向与当前操作期望不同的部件，工具会显式拒绝，而不是静默覆盖一部分属性。
- XML 属性部件统一输出为 UTF-8，并带 `standalone="yes"` 声明。
- POSIX 环境下撤销成功后只能恢复访问时间和修改时间；inode 的 ctime 由文件系统管理，无法精确回退。Windows 上创建时间是否正确恢复仍需在目标 Windows 系统验证。

## 便携式回归测试（可选）

完整离线工具包附带 `office_metadata_regression_test.py` 和验证样例压缩包。将文件解压到同一目录后运行：

```powershell
python office_metadata_regression_test.py
```

测试会在系统临时目录中解压样例并处理副本，检查三种文件的元数据回读与撤销、带通配符的路径、GUI 时间字段采集/预演/写入、预演异常保留、损坏清单处理与安全跳过统计、关系目标校验、备份命名上限、XML 声明、新部件 ZIP 时间戳稳定性、ZIP 体积限制与加密条目错误提示、OSError 属性保留、无图形环境启动提示和 CLI 读取统计。它不会修改你的实际文档。该脚本仅依赖 Python 标准库；它不能替代 Windows Office GUI 的最终检查。


## 第四轮代码审查说明

- ZIP 校验期间遇到加密或无法读取的部件时，会提供可能需要密码的友好错误提示，而不暴露原始 RuntimeError 作为唯一说明。
- 撤销清单损坏的备份会安全跳过，统计窗口将“安全跳过”与“恢复失败”分开显示。
- ZIP 部件校验以 4 MiB 块分段读取以减少 Python 循环次数；校验仍会读取所有部件并检查实际读出的累计字节。该设置未经过多 GB 文件性能基准验证。
- 最新便携式回归套件包含 42 项检查。当前测试环境为 Linux/POSIX；Windows `ReplaceFileW`、Windows ACL、Windows 文件系统创建时间和 Microsoft Office 本体仍需在目标 Windows 电脑上验证。

## ZIP 压缩方法兼容性补充

本工具依赖 Python 标准库 `zipfile` 读取 Office OOXML 包。若包内使用当前 Python 版本不支持的 ZIP 压缩方法，工具会明确提示“不支持的 ZIP 压缩方法”并拒绝处理；若部件需要密码或属于加密 ZIP，则显示可能已加密/需要密码的提示。工具不会联网安装额外依赖，也不会尝试绕过加密。

第五轮审查后，便携式回归测试共有 50 项。它包含 `start_modification()` 处理编排测试，但预演窗口的确认动作在该测试中由回调替身模拟；这不是 Windows 桌面人工交互或 Office 本体验收。
