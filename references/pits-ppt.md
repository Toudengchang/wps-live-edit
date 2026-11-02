# 演示（wpp/kwpp）的坑 · 坑 45~50

> wps-live-edit 坑库 · 演示分册。与文字/表格共用的铁律、双模式和收尾规则见根目录 SKILL.md。
>
> 这一节的坑比表格更凶：**PPT 出错时几乎不报错**，文字烂成一团也照样 `Save()` 成功。所以 PPT 的每一步都必须靠「导 PNG 看」来验证。

## 坑 45（最严重）：碰 `TextRange.ParagraphFormat` 会**摧毁整个 PPT 的文本渲染**

**现象**：设过一次行距/段距之后，文本框里的文字**不再换行、不再分段**，横向冲出页面，`Shape.Height` / `BoundHeight` 返回 `1.2 / 3.6 / 6.1 / 17.4` 这种垃圾值。**全程零报错**，`Save()` 也成功，只能靠 `Slide.Export` 导 PNG 才看得出已经烂了。

**有毒代码（都踩过）：**

```python
r.ParagraphFormat.LineSpacing = 1.2      # AttributeError：PPT 没有 LineSpacing
r.ParagraphFormat.LineRuleWithin = -1    # ← 真正的元凶，设了它整页文字就废
r.ParagraphFormat.SpaceWithin = 1.2      # 同类，段距不要碰
```

**排查过程（这些尝试全都无效，别再重复）：**

| 尝试 | 结果 |
|---|---|
| 改 `Shape.Width` / 反复微调 | 无效，Width 抖动改变不了 |
| 改 `TextFrame.AutoSize` 各值 | 无效 |
| 改 `TextFrame.WordWrap` 各值 | 无效 |
| 新建 `AddTextbox` 再写 | **新文本框也烂** → 说明是应用级/文档级状态被污染，不是单个形状的问题 |
| 重开文件再 Save | 无效，污染已落盘 |

最后靠**纯净对照实验**定位：新建一份干净 PPT，只逐个试属性，只有 `LineRuleWithin` 一设就复现 → 确认元凶。

**铁律：绝不碰 `TextRange.ParagraphFormat`。** 行距/段距需求改用下面两条替代：

1. **字号 + 形状高度**：`r.Font.Size = 20`，靠 `sh.Height` 撑出行间距视觉效果
2. **空段落占位**：需要行间空隙就多写一个 `\r`，不要用段距参数

**另外两条硬要求：**
- 段落分隔**只能用 `\r`**（`\n` 在 PPT 里不换行，直接显示成一个怪符号）
- 写完 `sh.TextFrame.TextRange.Text = ...` 之后**必须重新取** `r = sh.TextFrame.TextRange`，旧引用指向的内容已经失效

一旦已经踩了，单个文本框**修不回来**，只能走下面的「纯净重建法」。

## 坑 46：PPT 不支持的对象/属性清单

| 写法 | 实际结果 |
|---|---|
| `p.DisplayAlerts = 0` | 不存在，开关在 `app` 上 → `app.DisplayAlerts = 0` |
| `p.ScreenUpdating = False` | 不存在（PPT 没有这个开关），别照抄表格的写法 |
| `r.ParagraphFormat.LineSpacing = ...` | 不可设（见坑 45） |
| `r.Words()` | 报 `-2147352567`，PPT 的 Words 集合不支持这样取 |
| `r.Paragraphs().Count` / `r.Paragraphs(j)` | ✅ 可用，数段落/取段落用这个 |

## 坑 47：`GroupItems(j)` 成员的 `Left/Top` 是**页面绝对坐标**

组合（Group）里的成员，坐标不是相对组合左上角的，而是**相对整页**。所以：

- 想判断某个成员是谁，用绝对坐标比较：`it.Top < sh.Top + 10` → 说明它是组合顶部的标题条
- 改成员文字必须走 `it.TextFrame.TextRange`；**组合本身没有 `TextFrame`**，对 `sh`（组合）读 `.TextFrame` 会失败

```python
for j in range(1, sh.GroupItems.Count + 1):
    it = sh.GroupItems.Item(j)
    if it.Top < sh.Top + 10:                 # 绝对坐标 → 顶部那条就是标题
        it.TextFrame.TextRange.Text = 新标题
```

## 坑 48：`BoundHeight` 体检法有边界——只在没碰过 `ParagraphFormat` 时可信

`sh.TextFrame.TextRange.BoundHeight` 本来是判断"文字有没有溢出文本框"的好工具，但**一旦这个文档被 `ParagraphFormat` 污染过，BoundHeight 就开始返回垃圾值**（1.2 / 3.6 / 6.1 / 17.4），拿它做判断会得出完全错误的结论。

**规则：**
- 干净文档 → 可以用 `BoundHeight` 与 `sh.Height` 比较做自动体检
- 碰过 ParagraphFormat（或来源不明）→ 只能 `Slide.Export` 导 PNG 肉眼验，不要相信任何数值

## 坑 49：`Slide.Export` 有渲染延迟，刚 Save 完可能导到上一版画面

`s.Export(path, "PNG", 1600, 900)` 执行成功、文件也生成了，但画面可能是**保存前的旧版本**。

**铁律：图里看到位置异常，先回读坐标，坐标正常就重导一次，别急着改代码。**

```python
print(sh.Left, sh.Top, sh.Width, sh.Height)   # 先确认坐标是对的
s.Export(png, "PNG", 1600, 900)               # 再重导一遍
```

很多"看起来错位"的情况，改代码是白改——图是旧的。

## 坑 50：切 `VerticalAnchor` 会顺带改 `Shape.Height`

设 `tf.VerticalAnchor = 3`（垂直居中）之后，`sh.Height` 会从 158 变成 88 之类，**形状自己缩了**，文字就被裁掉。

**做法：改完锚点必须把高度重设回去，并断言校验：**

```python
tf.VerticalAnchor = 3
sh.Height = 126.8
assert abs(sh.Height - 126.8) < 0.5, "高度被锚点改了"
```

## 纯净重建法：被污染的文本框只能删了重造

单个文本框一旦被 `ParagraphFormat` 污染，**改属性、改宽度、重设文本都救不回来**。唯一可靠办法是删掉重建：

```python
# 1. 先记下原样（后面要一模一样造回来）
L, T, W, H, name = sh.Left, sh.Top, sh.Width, sh.Height, sh.Name
text = sh.TextFrame.TextRange.Text
size = sh.TextFrame.TextRange.Font.Size
bold = sh.TextFrame.TextRange.Font.Bold

# 2. 删掉旧的
sh.Delete()

# 3. 造一个同位置同尺寸的新文本框（1 = msoTextOrientationHorizontal）
new = s.Shapes.AddTextbox(1, L, T, W, H)
new.Name = name

# 4. 重取 TextRange，只设 Text + Font —— 一个 ParagraphFormat 属性都不要碰
new.TextFrame.TextRange.Text = text          # 段落分隔只用 \r
r = new.TextFrame.TextRange                  # ← 重取
r.Font.Size = size
r.Font.Bold = bold
```

> 口诀：**记录 → 删除 → 重建 → 重取 → 只设 Text 和 Font。**
