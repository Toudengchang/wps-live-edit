---
name: wps-live-edit
description: 通过 win32com 挂载用户已经打开的 WPS 文字、表格或演示实例读写内容，支持 WPS 文字（.docx/.doc）、WPS 表格（.xlsx/.xls）与 WPS 演示（.pptx/.ppt）。用于文档已被 WPS 打开/独占锁定、不能关闭也不能另开副本、只改局部内容且必须保住其余格式的场景。提供两种操作模式：显性模式（抢焦点，Activate 并把目标工作簿/sheet/单元格/形状选中，让用户看得见操作过程，用于调试和演示）与勿扰模式（后台静默，不 Activate、不碰 Selection、不抢焦点）。执行前必须先确认用哪种模式，用户没有明确指定时应先询问。核心是挂载式编辑：不新起进程、绝不 Quit、只动目标文档、写入回读再 Save。演示（PPT）里绝对不要碰 TextRange.ParagraphFormat 的行距/段距（会摧毁渲染且零报错），段落分隔只能用 \r，改完必须 Export 导 PNG 核验。
---

# WPS 在用手改文档（挂载式后台编辑）

> ⚠️ **本技能有两种操作模式：显性操作 / 勿扰操作。执行前必须先确认用哪种；用户没有明确指定时，先询问，不要自行默认。** 详见下方「操作模式」。

## 什么时候用它

**该用：**
- 目标文档已经在 WPS 里开着，你不能关它，也不该另存一份副本
- 只改局部内容（几个字、几段话、几个单元格、几页幻灯片里的几个文本框），其余内容和格式必须原样保留
- 文件被 WPS 独占锁定，直接读写文件（python-docx、zipfile、openpyxl、python-pptx）会失败或被顶回
- **勿扰模式**：用户正在用电脑忙别的，你的操作**不能**把 WPS 顶到前台、不能抢焦点、不能弹窗
- **显性模式**：用户要盯着看你的操作过程，需要像人一样把目标表/sheet/单元格/形状选中、把进度显示出来

**不该用：**
- 无人值守的批量处理 —— 那种场景应该自己开实例（`Dispatch`），做完再 `Quit`，与本技能路线相反
- 目标文档根本没打开 —— 直接改文件更简单可靠

## 操作模式：显性 / 勿扰（执行前必须先确认）

本技能有**两种操作模式**，动作方式相反。**开始前必须确认用户要哪一种；用户没有明确指定时，先询问，不要自行默认。**

| | 勿扰模式（原行为） | 显性模式（新增） |
|---|---|---|
| 焦点 | 不抢焦点，后台静默 | **抢焦点**，把目标窗口顶到前台 |
| 窗口 | 不动 `Visible` / `WindowState` | `app.Visible = True`，必要时置前/最大化 |
| 激活 | 不 `Activate` | `wb.Activate()` / `doc.Activate()` / `sh.Activate()` |
| 选中 | 不碰 `Selection` | 操作前 `Range.Select()` 选中将改的区域 |
| 用途 | 用户在用电脑、怕被打断 | **调试 / 演示，用户要看得见操作进度** |
| 演示效果 | 只见结果 | 像人一样一步步"点"到目标表/sheet/单元格 |

**两种模式共用的铁律不变**：挂载不新建、**绝不 `Quit()`**、只动目标文档、写完回读再 `Save()`、用完恢复 `DisplayAlerts`。

### 显性模式怎么写：让操作"被看见"

像人一样：**先激活目标工作簿 → 激活目标工作表 → 选中将操作的单元格/区域 → 再写入**，每一步用户都能在屏幕上看到。

```python
app.Visible = True                 # 确保可见
wb.Activate()                      # 目标工作簿顶到前台
sh.Activate()                      # 激活要操作的工作表
sh.Range(sh.Cells(r1, c1), sh.Cells(r2, c2)).Select()   # 选中将改的区域，用户可见
# ……执行写入 / 写公式……
app.Calculate()
wb.Save()                          # 仍然只保存目标文档
# 绝不 app.Quit()
```

文字同理：`doc.Activate()` 后，把要改的 `Range` 选中（`rng.Select()`）再写入。

### 勿扰模式怎么写：后台静默（= 原来的写法）

不 `Activate`、不改 `Visible`、不碰 `Selection`，直接对目标对象读写：

```python
sh.Range(sh.Cells(3, 1), sh.Cells(3 + len(rows) - 1, len(rows[0]))).Value = rows
app.Calculate()
wb.Save()
```

**注意**：显性模式下 `Selection` 会被脚本改变，用户当前光标/选区会被移动——这是"要被看见"的预期代价，但要提前跟用户说清楚；无论哪种模式，都**不要 `Quit()`**、不要关用户其他文档。

## 五条铁律

1. **挂载，不新建**：用 `win32com.client.GetActiveObject(...)` 挂到用户**已经在运行的** WPS 上。不要用 `Dispatch(...)`，那会另起一个进程、另开一份文件，还可能与用户正在编辑的副本冲突。
2. **绝不 `Quit()`**：WPS 是单实例多文档，`Quit()` 会**关掉用户全部已打开文档**。只对目标文档调 `Save()`。
3. **按模式决定是否抢焦点**：**勿扰模式绝不碰** `Activate()` / `Visible` / `Selection`（前两个会把窗口顶到前台，`Selection` 会抢走用户的光标和选区）；**显性模式才用**这三个，并严格按"激活工作簿 → 激活 sheet → 选中目标区域"的顺序显式展示（见「操作模式」）。两种模式下读写本身都优先走对象直取：文字 `Document.Range(a, b).Text`，表格 `Cells(r, c).Value` 或 `Range.Value`，演示 `Shape.TextFrame.TextRange.Text`（改完要重取 TextRange）。
4. **只动目标文档**：不要 `SaveAs`、不要 `Close()` 别人的文档、不要改 `ActiveDocument`。
5. **善后（强制）**：任何 `app` 级开关（`DisplayAlerts` / `ScreenUpdating` / `Interactive` / `Calculation` / `CutCopyMode`）都必须**先备份原值再修改**，并在 `finally` 里恢复；任务结束前做一次状态体检。详见下方「强制规则 A/B/C」。

## 挂载与 ProgID（速记）

| 组件 | WPS 正式版 | WPS 抢先版 | MS Office |
|---|---|---|---|
| 文字 | `wps.Application` | `kwps.Application` | `Word.Application` |
| 表格 | `et.Application` | `ket.Application` | `Excel.Application` |
| 演示 | `wpp.Application` | `kwpp.Application` | `PowerPoint.Application` |

**最关键的一条：很多机器上，`Word.Application` 已经被 WPS 注册并指向同一个 WPS 实现。** `word.Name` 会骗你（返回 "Microsoft Word"），`word.Path` 指向 WPS 安装目录才是真相——完整鉴别方法、三组件 API 对照表和 Word/Excel vs PPT 差异速查表见 [references/api-matrix.md](references/api-matrix.md)。

## 强制规则：WPS 全局状态必须还原

> 违反后果：用户会觉得 WPS「不受控制」——关文件不再问是否保存、点单元格没反应、公式不重算、复制虚线一直挂着。这是**对用户工作习惯的破坏**，按最高优先级处理。

### 强制规则 A：脚本结束必须完整还原 WPS 状态

| 开关 | 被改坏后的表现 | 正确值 |
|---|---|---|
| `DisplayAlerts` | 关闭文件不再提示「是否保存」 | `True` |
| `ScreenUpdating` | 点单元格没反应、界面不重绘 | `True` |
| `Interactive` | 用户完全无法操作 | `True` |
| `Calculation` | 公式不自动重算，数据看着是旧的 | `-4105`（xlCalculationAutomatic） |
| `CutCopyMode` | 单元格一直虚线闪烁、复制状态残留 | `False` |

> **演示（PPT）只有 `DisplayAlerts` 一个开关**，且在 `app` 上（`app.DisplayAlerts = 0` 关、`1` 开）。`p.DisplayAlerts`、`p.ScreenUpdating`、`app.Calculation`、`app.CutCopyMode` 在 PPT 上都不存在，照抄表格的还原代码会一路抛异常（见 [references/pits-ppt.md](references/pits-ppt.md) 坑 46）。

**推荐写法（`finally` 兜底）：**

```python
alerts_bak = app.DisplayAlerts      # 先备份
screen_bak = app.ScreenUpdating
calc_bak   = app.Calculation
try:
    app.DisplayAlerts = False
    app.ScreenUpdating = True
    # ... 全部操作 ...
finally:
    try: app.DisplayAlerts = alerts_bak
    except Exception: pass
    try: app.ScreenUpdating = screen_bak
    except Exception: pass
    try: app.Calculation = calc_bak
    except Exception: pass
    try: app.CutCopyMode = False
    except Exception: pass
    try: app.Interactive = True
    except Exception: pass
```

**收尾体检（不通过就修）：**

```python
assert app.DisplayAlerts is True
assert app.ScreenUpdating is True
assert app.Interactive is True
assert app.Calculation == -4105
assert app.CutCopyMode == 0
```

### 强制规则 B：备份必须写在修改之前

```python
# 错：先改后备份，备份到的是改后的值，恢复也恢复不回去
app.DisplayAlerts = False
alerts_bak = app.DisplayAlerts

# 对：先备份，再改
alerts_bak = app.DisplayAlerts
app.DisplayAlerts = False
```

判断口诀：**看到「`xxx = False` 在前、`bak = xxx` 在后」，就一定拿不到原值。** 多脚本连续操作时，每个脚本都独立备份+恢复，不能指望上一个脚本恢复好了。

### 强制规则 C：多脚本任务，每个脚本都要独立善后

- 每个脚本自带 `try/finally` 恢复
- 任务最后单独跑一次「状态还原脚本」，把 5 个开关全部拨回默认
- 还原脚本**只读写开关，不改任何单元格、不保存任何文件**

可直接用 `scripts/wps_com_utils.py` 的 `restore_app_state(app)` 和 `check_app_state(app)`。

## 标准流程 — 文字

```python
import pythoncom
import win32com.client as win32

pythoncom.CoInitialize()
app = win32.GetActiveObject("kwps.Application")   # 挂载用户已开实例
alerts_bak = app.DisplayAlerts

# 1. 列出已打开文档，确认目标在前、谁在前台
for i in range(app.Documents.Count):
    d = app.Documents.Item(i + 1)
    print(d.Name, "| 已保存:", d.Saved)
active = app.ActiveDocument.Name   # 用户当前正在操作的文档

# 2. 按名字找到目标文档
doc = None
for i in range(app.Documents.Count):
    if app.Documents.Item(i + 1).Name == 目标文件名:
        doc = app.Documents.Item(i + 1)
        break

# 3. 后台改内容（只用 Range，不 Activate）
app.DisplayAlerts = False
for pos in 命中位置列表从后往前:
    doc.Range(pos, pos + len(旧文本)).Text = 新文本

# 4. 只保存目标文档，然后恢复设置
doc.Save()
app.DisplayAlerts = alerts_bak
```

改内容的具体写法（WPS 的 `Find.Execute` 会假成功，必须用字符定位法）见 [references/pits-word.md](references/pits-word.md) 坑 1，或直接用 `scripts/wps_live.py`。

## 标准流程 — 表格

```python
import pythoncom
import win32com.client as win32

pythoncom.CoInitialize()
app = win32.GetActiveObject("ket.Application")   # 挂载用户已开 WPS 表格实例
alerts_bak = app.DisplayAlerts

# 1. 列出已打开工作簿
for i in range(app.Workbooks.Count):
    wb = app.Workbooks.Item(i + 1)
    print(f"Item {i+1}: {wb.Name} | Sheets: {wb.Sheets.Count} | 已保存: {wb.Saved}")

# 2. 按序号取工作簿和 sheet（COM 读出的 sheet 名可能是乱码，优先用序号）
wb = app.Workbooks.Item(目标序号)
sh = wb.Sheets.Item(sheet序号)   # 不要用 Sheets("汇总")，可能取不到

# 3. 批量读取区域
data = sh.Range(sh.Cells(1, 1), sh.Cells(61, 10)).Value   # 返回 tuple of tuples

# 4. 批量写入（必须用 tuple，不能用 list；逐格写会超时）
rows = ((r1c1, r1c2, ...), (r2c1, r2c2, ...), ...)
sh.Range(sh.Cells(3, 1), sh.Cells(3 + len(rows) - 1, 列数)).Value = rows

# 5. 写公式（行号用 f-string 动态拼）
sh.Cells(r, 5).Formula = f"=COUNTIFS(清单!J:J,D{r})"
sh.Cells(r, 7).Formula = f"=XLOOKUP(B{r},索引!$A:$A,索引!$C:$C)"

# 6. 删除整行（不留空行）
sh.Rows(行号).Delete(Shift=-4161)   # xlUp，下方行上移

# 7. 强制重算公式 + 保存
app.Calculate()
wb.Save()
app.DisplayAlerts = alerts_bak
# 绝不调用 app.Quit()
```

表格的坑最多（坑 5~44），动手前按需读 [references/pits-sheet.md](references/pits-sheet.md)；大批量操作先看坑 44。

## 标准流程 — 演示（PPT）

PPT 没有"按文件名精确匹配"的入口（`Presentations` 只按序号取），所以：**先用 `app.ActivePresentation`，找不到目标就遍历 `Presentations` 按 `Name` 比对。**

```python
import os
import pythoncom
import win32com.client as win32

pythoncom.CoInitialize()
app = win32.GetActiveObject("kwpp.Application")   # 抢先版；正式版 wpp.Application
old_al = app.DisplayAlerts                          # 注意：PPT 的开关在 app 上，不在 p 上
app.DisplayAlerts = 0                               # ppAlertsNone
try:
    # 1. 定位演示文稿（PPT 没有按名字取的口子，只能遍历比对 Name）
    p = app.ActivePresentation
    if p is None or p.Name != 目标文件名:
        p = None
        for i in range(1, app.Presentations.Count + 1):
            cand = app.Presentations.Item(i)
            if cand.Name == 目标文件名:
                p = cand
                break
    assert p is not None, "目标演示文稿没在 WPS 里打开"

    # 2. 取页、取形状（形状优先按 Name 精确取，不要靠序号猜）
    s = p.Slides.Item(页号)
    sh = s.Shapes("标题条")          # 或 for i in range(1, s.Shapes.Count+1): 逐个比对 Name

    # 3. 只改文字：先赋值，再重取 TextRange（旧引用会失效）
    sh.TextFrame.TextRange.Text = "新文字"       # 段落分隔只能用 \r
    r = sh.TextFrame.TextRange                   # ← 必须重取
    r.Font.Size = 20                             # 字号用 Font.Size，不要用 ParagraphFormat

    # 4. 收尾三件套：回读校验 → Save → 导 PNG 肉眼核验
    assert sh.TextFrame.TextRange.Text.startswith("新文字")
    p.Save()
    png = os.path.join(临时目录, "_preview.png")
    s.Export(png, "PNG", 1600, 900)              # 导出有渲染延迟，见坑 49
    open(os.path.join(临时目录, "_log.txt"), "w", encoding="utf-8").write("ok")
finally:
    try:
        app.DisplayAlerts = old_al
    except Exception:
        pass
# 绝不调用 app.Quit()
```

> ⚠️ **PPT 与表格最大的不同：没有"写进去就一定对"的自动校验。** 表格能回读单元格值、能扫 `#` 错误；PPT 只能靠 `Slide.Export` 导 PNG 用眼睛看。所以 PPT 任务**必须**把「导 PNG 核验」写进收尾步骤，不能省。
>
> PPT 的坑全在 [references/pits-ppt.md](references/pits-ppt.md)（坑 45~50），其中**坑 45（绝不碰 `TextRange.ParagraphFormat`，会零报错摧毁整页文本渲染）是全局最高危规则**；与文字/表格的差异速查表见 [references/api-matrix.md](references/api-matrix.md)。

## 坑库索引（细节按需读对应分册）

- **文字 → [references/pits-word.md](references/pits-word.md)**（坑 1~4）：Find.Execute 假成功与字符定位法 · 整段赋值污染格式 · os.replace 跨盘失败 · 旧版 .doc 字节校验
- **表格 → [references/pits-sheet.md](references/pits-sheet.md)**（坑 5~44 + 建表经验）：逐格读写都慢、批量 Range（5/26/44）· ClearContents 连公式清/留空行（6/7）· Range.Sort 不生效改 Sort 对象（8）· sheet 名乱码用序号（9）· 写公式后 Calculate（10）· 数据在 WPS 内部流转 + 手法 1 四硬限制「筛含表头、抄不含表头、同会话 Copy、回读落点」（11）· 贴前校对表头（12）· 公式行数跟数据走（13）· 动态数组固化覆盖溢出区（14）· 特多区域 SpecialCells 删行静默失败（15）· UsedRange 黏旧值（16）· 大批量切手动计算（17）· 簿序号漂移（18）· 跳过 ~$ 锁文件（19）· 筛选态只清可见行（20）· Worksheet.Copy 吞参数（21）· Interior.Color 是 BGR 且要回读（22）· 去外链删前缀变 #REF!（23）· 写入回读再 Save（24）· 读也走挂载 COM（25）· 双值文本 #VALUE!（27）· NumberFormat 整区单串（28）· 跨源数字归一化（29）· 名字含 \xa0 要容错（30）· 跨表公式用真实 sheet 名（31）· 省 token 批量取数（32）· 绝对引用不随粘贴偏移（33）· WrapText 撑行高（34）· PowerShell $ 插值（35）· RPC_E_CALL_REJECTED 重试（36）· 扫 .Text 找错误值（37）· 找簿判空（38）· 业务口径先找现成台账（39）· 跨脚本粘贴贴空（40）· 行数看 Area 跨度（41）· 复现按真实文件形状（42）· 簿名/sheet 顺序不可靠（43）
- **演示 → [references/pits-ppt.md](references/pits-ppt.md)**（坑 45~50）：绝不碰 ParagraphFormat（45）· 不支持属性清单（46）· GroupItems 绝对坐标（47）· BoundHeight 边界（48）· Export 渲染延迟（49）· VerticalAnchor 改 Height（50）· 纯净重建法

## 自检（仅勿扰模式）：怎么证明"没打扰用户"

操作前后各采一次前台窗口标题，比对是否变化：

```python
import win32gui
win32gui.GetWindowText(win32gui.GetForegroundWindow())
```

再对比用户文档的内容指纹、字数、`Saved` 状态、`Selection.Start`：

- 内容指纹和字数不变 → 没改到用户的文档
- `Saved` 不变 → 没触发保存、没清掉用户的未保存标记
- 前台窗口标题不变 → 没抢焦点

`Selection.Start` 只作参考：字数没变时光标的位移一般是用户自己在动，不要轻易归因到自己的脚本。

## 任务收尾清单（每次挂载式编辑都要过）

- [ ] 目标区域回读校验通过（无 `#` 错误值，见 [references/pits-sheet.md](references/pits-sheet.md) 坑 24 / 37）
- [ ] 大批量读写已改为**整块批量**（不逐格 `Cells`，见坑 44）
- [ ] 修改期间 `Calculation` 已设手动，完成后统一 `Calculate()` 再还原
- [ ] 关键数据勾稽一致（分项合计 = 总计行）
- [ ] **（PPT）全程未碰 `TextRange.ParagraphFormat`，段落分隔只用 `\r`**（见 [references/pits-ppt.md](references/pits-ppt.md) 坑 45）
- [ ] **（PPT）已 `Slide.Export` 导 PNG 并肉眼核验排版**（见坑 48 / 49）
- [ ] 已 `wb.Save()` / `p.Save()` 且 `Saved == True`
- [ ] 未调用 `app.Quit()`
- [ ] 未关闭、未保存其他工作簿
- [ ] `restore_app_state(app)` 已执行
- [ ] `check_app_state(app)` 返回空列表
- [ ] 向用户说明：改了哪个文件、哪些区域、备份在哪、有什么待确认项

## 配套脚本

`scripts/wps_live.py` 是文字文档的命令行封装：

```
python wps_live.py list                                  # 列出已打开文档，标出谁在前台
python wps_live.py info   "<文档名>"                      # 页数/字数/字体摘要
python wps_live.py dump   "<文档名>" [--text-only]        # 导出全部段落（含表格）
python wps_live.py find   "<文档名>" "<关键词>"            # 定位关键词
python wps_live.py replace "<文档名>" "<旧>" "<新>" [--save]   # 后台替换，默认只预演不保存
```

脚本默认**不保存**，确认没问题再加 `--save`。运行前确保系统 Python 已装 `pywin32`。

表格操作目前没有命令行封装，按「标准流程 — 表格」中的代码模式直接在 Python 脚本中调用 COM。

演示（PPT）同样没有命令行封装，按「标准流程 — 演示（PPT）」的骨架写脚本；动手前先看 [references/pits-ppt.md](references/pits-ppt.md) 坑 45~50 和差异速查表。

`scripts/wps_com_utils.py` 是表格挂载脚本的公共函数库，写脚本时优先复用（见强制规则 C、坑 29 / 36 / 37 / 38 / 40 / 41 / 44）：

```python
from wps_com_utils import (rtry, wb_by, scan_errors, restore_app_state,
                           check_app_state, norm, filter_move)

rtry(lambda: app.Workbooks.Count)        # COM 调用套重试，防 RPC_E_CALL_REJECTED
w = wb_by(app, "台账")                    # 按名找工作簿（容错 \xa0 / 全角空格），返回 None 要判空
bad = scan_errors(sh, 36, 1, 63, 20)     # 扫区域错误值，返回 [(r, c, "#N/A"), ...]
restore_app_state(app)                   # 5 个 app 级开关全部拨回默认
check_app_state(app)                     # 返回 [(属性, 当前值, 期望值)]，空列表 = 正常
norm(v)                                  # 跨源比对前归一化：数字 round(,6) 转 str，空值统一 ""
n = filter_move(app, sh, tgt, fld=6, val="华东", r_end=301,
                r_start=2, paste_cell="A3", col_last="I", expect_n=60)
# ↑ 筛选后整块搬运：A1 起筛（含表头）、A2 起抄（不含表头）、同会话 Copy、
#   按 Area 跨度算真实行数、粘贴后回读落点自检；任一条不过直接抛错，不会静默错位
```

## 已知代价

改动期间（尤其 `Save()`）WPS 进程会被 COM 调用占用，用户敲键盘可能有极短的顿感。这是 WPS 单进程 COM 的固有代价，改的内容越多、文档越大，占用越久。把操作切成小步、只保存必要的文档，可以把影响压到最小。
