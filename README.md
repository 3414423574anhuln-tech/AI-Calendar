# AI-Calendar · 日历提醒

**简体中文** | [English](README.en.md)

“日历提醒”是一款面向 Windows 10/11 x64 的中文桌面日程软件。核心日历、提醒、普通导入导出均可离线使用；可选的 AI 导入需要联网及用户自己的供应商 API Key。程序不依赖浏览器或 Microsoft Office，数据库默认保存在当前用户目录。

本项目采用 [MIT 许可证](LICENSE) 开源。第三方依赖仍遵循各自许可证。

## 功能

- 新建、查看、修改、删除日程；主页面每条日程均有独立删除按钮，并支持勾选、多选、全选和事务式批量删除
- 日程时间与提醒时间分别设置，统一显示为 `yyyy/MM/dd HH:mm`；旧数据库自动将提醒时间迁移为原日程时间
- SQLite 本地持久化、时间索引、重复检查和批量导入事务
- 每 25 秒检查一次到达提醒时间的日程，通过托盘气泡和置顶桌面弹窗提醒；点击主窗口关闭按钮后继续在后台驻留
- 导入预览可编辑、可取消勾选、可删除识别结果，并显示识别/重复/错误状态
- 重复项支持跳过、仍然导入、替换原日程
- 当前用户范围的开机自动启动开关
- 开机自启后把程序未运行期间到期的全部提醒合并成一个汇总弹窗
- 所有详细异常写入本地日志，普通界面仅显示中文错误摘要
- 独立 API 页面可配置 DeepSeek、Kimi、OpenAI、Anthropic（Claude）及各自 2—4 个模型；密钥由 Windows 凭据管理器加密保存
- AI 导入先生成规范 `.xlsx`，再进入可编辑预览；模型分批读取全部已有日程并拦截时间与事项高度重合的结果

## 技术栈

- C#、.NET 8、WPF、MVVM（CommunityToolkit.Mvvm）
- SQLite（Microsoft.Data.Sqlite）
- ClosedXML、ExcelDataReader（Excel）
- DocumentFormat.OpenXml（Word、PowerPoint）
- PDFsharp、PdfPig（PDF）
- Windows.Data.Pdf、Windows.Media.Ocr（扫描 PDF 与图片 OCR）
- System.Drawing / WPF 桌面能力（PNG、JPG、系统托盘）
- `HttpClient`（模型 API）、Windows Credential Manager（API Key）

## 开发环境与命令

项目实际使用 .NET SDK `8.0.423`。如果系统 `dotnet` 未包含 SDK，可使用微软官方 `dotnet-install.ps1` 安装到隔离目录，再显式调用其中的 `dotnet.exe`。

```powershell
dotnet restore CalendarReminder.sln
dotnet build CalendarReminder.sln -c Release
dotnet test CalendarReminder.sln -c Release
dotnet publish src/CalendarReminder.App/CalendarReminder.App.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:DebugType=None -p:DebugSymbols=false -o dist
```

## 数据位置

- 数据库：`%LocalAppData%\CalendarReminder\calendar.db`
- 日志：`%LocalAppData%\CalendarReminder\logs\`
- API 非敏感配置：`%LocalAppData%\CalendarReminder\api-settings.json`
- AI 生成的规范 Excel：`%LocalAppData%\CalendarReminder\ai-imports\`

API Key 不写入配置文件，而是以 `CalendarReminder.ApiKey.<供应商>` 项保存在当前用户的 Windows 凭据管理器中。

首次运行会自动创建目录、数据库、数据表和索引。

## 导入与导出

支持导入：`.xls`、`.xlsx`、`.docx`、`.pdf`、`.pptx`、`.png`、`.jpg`、`.jpeg`。

支持导出：`.xlsx`、`.docx`、`.pdf`、`.pptx`、`.png`、`.jpg`。图片内容较多时按 `文件名_01`、`文件名_02` 的方式分页。

Excel 优先识别“日期 / 日期时间 / 时间 / 安排 / 内容 / 事项”等表头；Word 优先读取表格，同时读取段落；PowerPoint 读取文本框和表格；PDF 先提取文本，文本不足时渲染页面后 OCR；图片始终使用 Windows 本地 OCR。所有识别结果必须经过预览确认后才会写入数据库。

导入记录已识别到明确年月日、但没有识别到具体小时和分钟时，自动把日程时间设为当天 `07:30`，提醒时间设为前一天 `07:30`，并在导入预览中显示实际提醒时间。缺少年份且没有唯一年份上下文的月份或月份周次，以及其他无法确定绝对日期的记录仍标记为“需要确认”，不会擅自补全具体日期。

精度不足日期的默认规则（时间均为 `07:30`）：仅有年份时取该年 `01/01`，提醒取上一年 `12/31`；仅有月份、月初或上旬时取该月 1 日；中旬取 13 日；下旬取 19 日；月底、月末或月尾取 25 日；月份规则的提醒均为默认日期前一天。月份本身不含年份时，仅当文件中存在唯一、明确的年份上下文才会套用；完全没有年份或出现多个不同年份时仍标为“需要确认”。支持“八月”等中文月份及常见同义表达。

日期范围按闭区间展开，起止日期均生成日程。例如 `2026年7月20日至22日` 生成 20、21、22 日三条记录；每条安排文字完全相同，并添加规范化的 `日期范围：2026/07/20—2026/07/22`，所有记录的提醒时间统一为范围左端前一天 `07:30`。支持同月、跨月和跨年范围。

“某月第 N 周”按周一至周日计算：包含该月第一个星期日的周为第一周，第二周是其后的七天，以此类推；因此第一周可能包含上个月末的日期。识别后把完整七天作为闭区间，按上述范围规则生成七条日程。

Excel 导入支持“年份标题 + 月份行 + 周次行 + 横向计划单元格”的项目计划矩阵。程序会合并客户、终端客户、品名、项目、内容、状态、问题点和计划文字；日期范围和月份周次会展开为逐日记录，同一单元格包含多个日期表达时分别解析；仅有横轴周序号、没有足够日期语义的记录仍要求确认。

## AI 导入

API 页面提供 DeepSeek、Kimi（月之暗面）、OpenAI 和 Anthropic（Claude）。每个供应商提供 2—4 个当前模型选项；“Claude Code”是开发工具而非 API 供应商，因此界面使用其实际供应商 Anthropic。模型名称会随供应商发布和下线而变化，若供应商返回“模型不存在”，需要升级软件中的模型目录。

AI 导入在本地先提取 Excel、Word、PDF、PowerPoint 或图片中的文字，再发送给所选模型整理成结构化日程。候选结果会与全部已有日程分批比较；被模型判定为时间和事项均高度重合的记录标为“高度重合”、默认取消勾选，并由数据库事务层强制跳过。整理完成后生成规范 Excel，再打开原有导入预览，用户确认前不会修改数据库。

每次发送前会显示供应商、模型、已有日程数量和数据范围并要求确认。文件文字以及已有日程的日程时间、提醒时间和安排描述会离开本机，可能产生 API 费用，并受对应供应商隐私政策约束。普通本地导入不发送任何数据。

## OCR 与扫描 PDF 限制

- OCR 依赖 Windows 已安装的语言包，优先简体中文，并兼容用户语言列表中的英文和数字；没有可用语言包时会给出明确提示。
- 扫描 PDF 依赖 Windows.Data.Pdf 成功渲染页面；加密、损坏、极端超大或 Windows 解码器不支持的 PDF 可能无法导入。
- OCR 识别本身可能出错或缺失，因此结果不会绕过预览直接写入数据库；不完整或含糊的日期会标记为“需要确认”。
- 密码保护的 Office/PDF 文件不在支持范围内。
- AI 能改善复杂表格的语义整理，但不能保证正确；模糊日期仍需在预览中人工确认。大量已有日程会拆成多次语义比对请求，耗时和费用随数量增加。

## 提醒与退出

程序运行时约每 25 秒查询已到提醒时间且尚未提醒的日程。每条日程可以分别设置“日程日期及时间”和“提醒日期及时间”，提醒时间允许早于或晚于日程时间。到时会显示托盘气泡和置顶桌面弹窗；用户确认弹窗后写入 `ReminderTriggered = true` 防止重复提醒，多条同时到期时会依次显示。点击主窗口右上角关闭按钮只会隐藏到托盘并继续检查；右键托盘图标可重新打开或选择“退出”，只有“退出”会结束程序。

程序在每次成功完成提醒查询后保存检查点。启用开机自动启动时，如果下次启动发现检查点之后、本次启动之前有尚未提醒的日程，会只显示一个“关机期间的日程提醒”汇总窗口，列出全部日程时间、提醒时间和完整安排；确认后一次性标记这些项目，避免逐条弹窗。异常断电时最多按最近一次约 25 秒检查点回溯，因此也能覆盖强制关机前尚未来得及弹出的提醒。

## 最终程序

`dist\日历提醒.exe`

这是 Windows x64 自包含单文件，不要求用户另外安装 .NET Runtime。双击即可运行。

可直接发送给其他 Windows 10/11 x64 用户的压缩包：`outputs\日历提醒_1.4.2_win-x64_便携版.zip`。对方解压后双击其中的 `日历提醒.exe` 即可。

## 应用图标

版本 1.4.2 使用用户提供的蓝色日历、时钟与黄色提醒铃图片作为应用图标。项目保留原始 PNG，并生成包含 16、24、32、48、64、128、256 像素七种尺寸的 Windows ICO；该图标已统一应用到 EXE 文件、主窗口、编辑/导入/导出/提醒窗口、任务栏和系统托盘。
