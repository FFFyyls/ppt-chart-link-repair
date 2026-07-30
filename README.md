# PPT 图表数据断链修复 Skill

从 PowerPoint 图表仍然保留的 OOXML 缓存中恢复数据，重建 Excel 工作簿，并通过桌面版 Excel、PowerPoint 恢复和验证可编辑的真实外部数据链接。

> 本项目完全在使用者自己的 Windows 电脑上运行，不会把 PPT、Excel 或恢复数据上传给项目作者或第三方服务。

## 适用场景

- PPT 图表仍能显示，但提示原始 Excel 不可用。
- “编辑数据”无法打开原工作簿。
- 需要从图表缓存重建 Excel，并让图表重新指向该工作簿。
- 需要把修复结果连同数据工作簿迁移到另一台本地 Windows 电脑。

## 不适用或暂不支持

- PivotChart、ChartEx、图片化图表、链接表格和其他 OLE 对象。
- 缓存数据已经缺失、索引冲突或无法可靠映射的图表。
- 加密、损坏、含宏的演示文稿。
- OneDrive、SharePoint、UNC、网络盘和映射网络驱动器。
- macOS、Linux、PowerPoint Online；完整流程依赖 Windows 桌面版 Microsoft Excel 和 PowerPoint。

完整兼容范围见 [supported-chart-types.md](repair-ppt-chart-data-links/references/supported-chart-types.md)。

## 环境

- Windows 10 或 Windows 11
- Python 3.10 或更高版本
- Microsoft Excel 桌面版
- Microsoft PowerPoint 桌面版
- `lxml`、`openpyxl`

安装 Python 依赖：

```powershell
py -3 -m pip install -r requirements.txt
```

## 安装 Skill

克隆仓库：

```powershell
git clone https://github.com/FFFyyls/ppt-chart-link-repair.git
```

将 `repair-ppt-chart-data-links` 整个目录复制到所用 Agent 支持的 Skills 目录。通用 Agent Skills 目录示例：

```powershell
Copy-Item `
  -LiteralPath ".\ppt-chart-link-repair\repair-ppt-chart-data-links" `
  -Destination "$env:USERPROFILE\.agents\skills\repair-ppt-chart-data-links" `
  -Recurse
```

不同 Agent 可能使用不同的 Skills 目录，请以相应产品说明为准。不要只复制 `SKILL.md`，脚本和参考文件同样是运行所必需的。

## 使用示例

```text
使用 repair-ppt-chart-data-links Skill，
扫描 C:\Documents\example.pptx 中的图表断链情况。
只生成盘点结果，不修改文件。
```

扫描后，Skill 会列出可恢复图表并要求确认。只有明确确认的图表才会恢复和重连。

确认后可继续：

```text
修复我确认的图表，保留原 PPT，
重建 Excel，并完成 Office 严格验证。
```

## 工作流程

1. 只读扫描 PPTX，识别图表、旧链接、缓存等级和阻断项。
2. 等待使用者确认目标图表。
3. 从完整缓存恢复 Excel 工作簿。
4. 使用 Office 原生图表和粘贴行为建立真实外部链接。
5. 验证“编辑数据”、工作簿路径、修改刷新、保存和重新打开。
6. 输出修复版 PPT 副本、重建 Excel 和验证报告。

原 PPT 不会被覆盖；不能可靠恢复的数据不会被猜测或联网补全。

## 已验证结果

- 当前自动化测试：76项。
- 参考回归样本：22张图表全部通过严格 Office 验证。
- 参考环境总耗时约19分30秒。

实际耗时取决于图表数量、Office 启动速度、文件复杂度和电脑性能；上述数据不是固定时限承诺。

## 本地文件安全

PPTX 是 ZIP/XML 容器。即使处理完全发生在本地，来自邮件、互联网或其他人员的恶意文件仍可能消耗大量内存或磁盘资源。本项目会在读取前检查 ZIP 成员数量、展开大小、压缩比、路径和 XML DTD/实体声明。

脚本只允许清理本次任务登记的 Office 进程，不按“新出现的 Excel/PowerPoint 进程”批量终止。便携包先写入临时目录，全部成功后再替换由本工具创建的旧输出；已有的非本工具目录不会被删除。

安全问题报告方式见 [SECURITY.md](SECURITY.md)。请勿在公开 Issue 上传含有客户或个人数据的 PPT、Excel、日志或恢复报告。

## 许可证

Copyright 2026 XLoffice-Fyl

本项目采用 [Apache License 2.0](LICENSE)。允许使用、修改、分发和商业使用；使用者和再分发者应遵守许可证中的版权、许可证、NOTICE 和修改标识要求。
