# API 速查 · ProgID / 对象模型 / 文字表格演示差异

> wps-live-edit 参考分册。坑的细节见 pits-word.md / pits-sheet.md / pits-ppt.md。

## ProgID 对照

| 组件 | WPS 正式版 | WPS 抢先版 | MS Office |
|---|---|---|---|
| 文字 | `wps.Application` | `kwps.Application` | `Word.Application` |
| 表格 | `et.Application` | `ket.Application` | `Excel.Application` |
| 演示 | `wpp.Application` | `kwpp.Application` | `PowerPoint.Application` |

**最关键的一条：在很多机器上，`Word.Application` 已经被 WPS 注册并指向同一个 WPS 实现。** 判断真实身份看这几个属性：

```python
word = win32.Dispatch("Word.Application")
word.Name      # 这里可能会骗你，返回 "Microsoft Word"
word.Version   # 12.0  ← 可疑，正常 Word 不是这个版本号
word.Path      # C:\Users\<用户>\AppData\Local\Kingsoft\WPS Office\<版本>\office6  ← 真相在这里
```

`Path` 指向 WPS 安装目录，就说明你手上的其实是 WPS 兼容层，必须按 WPS 的脾气来写（见 references/pits-word.md 坑 1）。

## 文字 / 表格 / 演示 API 对照

| 概念 | 文字 (wps/kwps) | 表格 (et/ket) | 演示 (wpp/kwpp) |
|---|---|---|---|
| 文档集合 | `app.Documents` | `app.Workbooks` | `app.Presentations` |
| 单个文档 | `doc = app.Documents.Item(i)` | `wb = app.Workbooks.Item(i)` | `p = app.Presentations.Item(i)` |
| 文档名 | `doc.Name` | `wb.Name` | `p.Name` |
| 内部页/表 | — | `wb.Sheets.Item(i)` | `p.Slides.Item(i)` |
| 页内对象 | — | `sh.Cells(r, c)` | `s.Shapes.Item(i)` / `s.Shapes("名字")` |
| 读内容 | `doc.Range(a, b).Text` | `sh.Cells(r, c).Value` 或 `sh.Range(c1, c2).Value` | `s.Shapes(i).TextFrame.TextRange.Text` |
| 写内容 | `doc.Range(a, b).Text = 新文本` | `sh.Cells(r, c).Value = 值` 或 `sh.Range(...).Value = tuple` | `sh.TextFrame.TextRange.Text = 新文本` |
| 写公式 | — | `sh.Cells(r, c).Formula = "=SUM(...)"` | — |
| 保存 | `doc.Save()` | `wb.Save()` | `p.Save()` |
| 强制重算 | — | `app.Calculate()` | — |
| 文档数 | `app.Documents.Count` | `app.Workbooks.Count` | `app.Presentations.Count` |
| 全局开关 | `app.DisplayAlerts` | `app.DisplayAlerts` | **`app.DisplayAlerts`**（`p.DisplayAlerts` 不存在） |

## Word/Excel vs PPT 差异速查表

| 维度 | 文字 / 表格 | 演示（PPT） |
|---|---|---|
| 段落分隔 | `\n` / 段落对象 | **只能用 `\r`**；`\n` 不换行 |
| 行距 / 段距 | `ParagraphFormat` 可用 | **绝对禁用** `TextRange.ParagraphFormat`（references/pits-ppt.md 坑 45），只能调 `Font.Size` + 形状高度 |
| 全局开关 | `app.DisplayAlerts` / `ScreenUpdating` / `Calculation` | 只有 `app.DisplayAlerts`；`p.DisplayAlerts`、`p.ScreenUpdating` **不存在**（references/pits-ppt.md 坑 46） |
| 出框判定 | 单元格宽度、`ColumnWidth` | `BoundHeight` 只在没碰过 ParagraphFormat 时可信，否则只能导 PNG 看（references/pits-ppt.md 坑 48） |
| 预览核验 | 回读单元格值、扫错误值 | `Slide.Export(路径, "PNG", 1600, 900)` + 肉眼（references/pits-ppt.md 坑 49） |
| 形状分组 | — | `GroupItems(j)` 成员的 `Left/Top` 是**页面绝对坐标**（references/pits-ppt.md 坑 47） |
| 文字引用 | Range 稳定 | 写完 `Text` 必须**重取** `sh.TextFrame.TextRange` |
