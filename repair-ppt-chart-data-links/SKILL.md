---
name: repair-ppt-chart-data-links
description: Use when a .pptx chart reports that a linked file is unavailable, Edit Data cannot open the original Excel workbook, chart values remain visible only in cache, or a repaired chart must be shared to another local Windows computer.
---

# PPT图表数据断链更新助手

作者：XLoffice-Fyl

Copyright 2026 XLoffice-Fyl

License: Apache-2.0（见同目录 `LICENSE`）

## 用途

从常规 PowerPoint 图表的 OOXML 缓存恢复图表实际引用的数据，重建 Excel，并用“Excel 原生图表载体 + Office 原生粘贴 + OOXML 数据关系移植”建立可编辑的真实外部链接。默认本机重连；用户明确需要跨电脑传输时，额外生成可迁移数据包。

完整流程要求 Windows 桌面版 Excel、PowerPoint。不得联网补数据，不得覆盖原 PPT，不得把 XML 关系正确单独当成修复成功。

## 固定边界

- 只处理 `.pptx` 中的常规 PowerPoint 图表。
- PivotChart、ChartEx、图片、在线数据源、链接表格和其他 OLE 对象只报告，不转换。
- 先扫描、再向用户展示目标图表清单；只有用户明确确认的图表才能恢复或重连。
- 按旧外部工作簿完整路径归并；同文件名但旧目录不同的工作簿不得自动合并。
- A 类缓存可以自动恢复；B 类可生成草稿 Excel，但必须暂停重连；C/U/N 类不得自动恢复。
- 空白点必须按缓存索引定位。连续前置/尾部空白可保留为图表语义；无索引缺失或同单元格冲突必须阻断。
- 组合图的 OOXML 系列顺序不等于 Office `SeriesCollection` 顺序。行为探针必须按系列名称唯一匹配；系列名称重复或无法唯一定位时停止，不得用数字序号猜测。
- 页面标题、正文、脚注和年份只生成候选，取得确认后才能修改。
- 不使用 `Chart.SetSourceData`，不删除原图表后替换为普通链接 Excel 对象，不静默降级。
- OneDrive、SharePoint、UNC/网络路径和映射网络驱动器不支持。

## 运行环境

使用当前环境中可用的 Python 3.10 或更高版本。若平台提供经过验证的捆绑 Python，优先使用。

先验证：

```powershell
py -3 -c "import lxml, openpyxl"
```

若没有 `py` 启动器，再分别尝试 `python3` 或 `python`。不要使用 Microsoft Store 的 Python 别名；必须实际执行导入探针，不能只检查可执行文件路径。没有含 `lxml`、`openpyxl` 的可靠运行时就暂停。

## 标准流程

### 1. 只读扫描

```powershell
& $python scripts/chart_recovery.py inspect --pptx input.pptx --work-dir work
```

读取 `chart-inventory.json`，向用户列出文件页码、页脚页码、图表 ID、形状名、旧路径、缓存等级和问题。依据 [supported-chart-types.md](references/supported-chart-types.md) 分类。

### 2. 取得确认并恢复 Excel

用户确认后，用明确的图表 ID 写入授权边界：

```powershell
& $python scripts/chart_recovery.py confirm-selection --manifest work/chart-inventory.json --chart-id slide-13-chart-1 --chart-id slide-17-chart-1
```

然后恢复：

先确定最终交付目录，再恢复；这样 Office 从一开始就绑定最终绝对路径，避免验收后再搬文件或重做链接：

```powershell
& $python scripts/chart_recovery.py recover --manifest work/chart-inventory.json --output-dir output
```

不传 `--output-dir` 时保留原有默认行为。传入后，Excel 写入 `output/recovered-workbooks`，修复版 PPT 和报告也直接写入 `output`。

逐图表复核：公式范围、缓存索引、系列、分类、值、工作表和单元格一一对应。数据映射结构见 [manifest-schema.md](references/manifest-schema.md)。

若 `relink_allowed` 为 false，交付草稿 Excel 和阻断项，停止。不得猜测缺失值。

### 3. 可选应用用户更新值

更新文件必须含 `updates_confirmed: true`，每项记录图表、系列、期间、旧值、新值、单位、目标单元格和来源说明：

```powershell
& $python scripts/chart_recovery.py apply-updates --manifest work/repair-manifest.json --updates updates.json
```

旧值、单位、范围或加总关系验证失败时停止。

### 4. 原生重连

先完成所有无需 PowerPoint 的准备。重连前检测 `POWERPNT.EXE`：

- PowerPoint 未运行：继续。
- PowerPoint 正在运行：立即暂停并明确提示“PowerPoint正在运行，请保存并关闭所有PowerPoint窗口后继续”，返回状态码 20。
- 不得保存、关闭或附着用户现有 PowerPoint 会话。

用户关闭后执行：

```powershell
& $python scripts/chart_recovery.py relink --manifest work/repair-manifest.json
```

重连引擎必须按以下顺序工作：

1. 为每个系列保留分类、数值、缓存索引、主次轴和组合图结构。
2. 在恢复工作簿中建立一次性 Excel 原生图表载体。
3. 由桌面 Excel 复制，PowerPoint 用 `PasteSourceFormatting` 生成真实链接图表。
4. 从临时链接图表提取 Office 生成的公式和外部数据关系。
5. 只把链接关系移植回原图表部件，保留原图表位置、大小、类型、轴、样式、颜色和层级。
6. 删除临时粘贴对象和载体工作表；最终 Excel 不得残留辅助图表工作表。

性能要求：同一批次只启动 1 个隔离 Excel 进程和 1 个隔离 PowerPoint 进程；所有恢复工作簿各打开一次，所有已确认图表在同一 PPT 会话内依次原生粘贴，整批只保存一次。粘贴后先完成 OOXML 关系移植和保存，再用 1 次只读 PowerPoint 会话从最终 PPT 捕获权威运行时系列名称并写回 manifest；临时粘贴对象中的名称只能作为诊断信息。名称缺失、重复或无法唯一映射时立即阻断，不得在后续验收中按数字序号猜测。

每次中断都应清理载体工作表；重复运行不得产生重复图表、重复链接或重复工作表。

### 5. 严格 Office 验收

```powershell
& $python scripts/chart_recovery.py verify --manifest work/repair-manifest.json --office-required
```

验证器对每张图表在三个隔离 PPT 探针会话中执行：

1. `ChartData.IsLinked` 为真；
2. `ActivateChartDataWindow` 成功并显示 Excel；
3. `ChartData.Activate` 后 `Workbook.FullName` 精确等于恢复工作簿路径；
4. 修改一个真实引用单元格并保存；
5. 按系列名称在 COM `SeriesCollection` 中唯一定位后，验证 `LinkFormat.Update`、`Chart.Refresh` 使对应系列值变化；
6. 保存、关闭、fresh PowerPoint 重开后仍为新值；
7. 用字节备份恢复 Excel，交付 PPT 不被 mutation probe 改写；
8. Office 保存后再次核对 ZIP、关系、公式、缓存、图表数量和未授权部件。

路径核对与可见“编辑数据”窗口合并在第一次 `Primary` 会话内完成；关闭该会话后再改工作簿，由 fresh `MutationUpdate` 会话更新并保存探针，最后由 fresh read-only 会话重开复核。不要把编辑数据与改值更新硬塞进同一会话：Office 会切换工作簿实例，可能出现空路径、未登记 Excel 进程或更新不收敛。所有图表通过后，只启动 1 次 PowerPoint，按目标页码去重批量导出 PNG。验收 JSON 必须记录 `elapsed_seconds`、`office_launch_records` 和 `visual_exports`，便于判断瓶颈与核查页面。

完整标准见 [verification-standard.md](references/verification-standard.md)。任一图表失败，整批不得标记为完成。

## 用户办公共存

- 扫描、缓存恢复、Excel OOXML 写入和结构验证可在用户继续办公时运行。
- Excel 自动化记录运行前 PID，并记录每个实际创建的 Office PID；只允许清理已登记 PID，或当前工作脚本直接派生且不属于运行前基线的 Office 子进程，不得按“运行后新增进程”批量清理。
- 如果 `ChartData.Activate` 或“编辑数据”复用了运行前已有的 Excel PID，立即停止；不得关闭该工作簿或该 Excel 进程。
- PowerPoint 无法可靠多实例隔离，所以原生粘贴和 Office 验收阶段必须要求用户关闭所有 PowerPoint。
- 将关闭提示推迟到所有非 PowerPoint 准备完成后，并在报告中记录该独占阶段耗时。
- Office 测试期间可能短暂出现 Excel 数据窗口；不得改写用户已有工作簿。

## 可迁移数据包（可选）

只有用户需要跨电脑传输，且本机 Office 硬验收已经全部通过时执行：

```powershell
& $python scripts/chart_recovery.py portable-package --manifest work/repair-manifest.json --output-dir output
```

数据包包含可迁移母版、`data` 文件夹、本机重绑启动器、包清单和使用说明。目标电脑不需要 AI 助手或本 Skill；需要 Windows 桌面版 Excel、PowerPoint。

目标电脑上双击启动器后：

1. 校验母版、清单和工作簿；
2. 按数据包当前绝对路径生成“本机链接版”，不覆盖母版；
3. 把授权图表的关系目标重绑到 `data` 中的 Excel；
4. 用隔离 Excel 克隆和 PPT 探针执行编辑数据、路径、改值、保存重开测试；
5. 真实数据工作簿不参与 mutation probe；
6. 通过后打开本机链接版。

整个数据包移动或改名后，重新运行启动器。不得声称单个 PPT 在没有工作簿和重绑步骤时可跨电脑保持链接。

## 停止条件

- 输入不是 `.pptx`、ZIP 损坏、加密、含宏或 Office 不可用。
- 缓存不存在、点数/索引不明、同单元格冲突、公式无法映射、用户旧值不匹配。
- PowerPoint 正在运行且用户尚未关闭。
- Office 粘贴失败、图表类型/系列/轴不匹配、辅助工作表未清理。
- 编辑数据失败、工作簿路径不匹配、改值不回读、保存重开失败、乱码或未授权部件变化。
- 可迁移包位于不支持的云盘或网络路径。

## 默认交付

- 修复版 PPT 副本；
- 重建 Excel；
- 数据恢复与严格验证报告。

盘点 JSON、配方、粘贴结果、探针 PPT、载体工作表和过程日志只保留在工作目录，不作为默认用户交付。用户选择跨电脑模式时，额外交付可迁移数据包。
