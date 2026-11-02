# 文字（wps/kwps）的坑 · 坑 1~4

> wps-live-edit 坑库 · 文字分册。与表格/演示共用的铁律、双模式和收尾规则见根目录 SKILL.md。

## 坑 1：WPS 的 `Find.Execute(Replace=2)` 会**假成功**

它会返回 `True`，但文档一个字都没改，而且**不报错**。用 `cell.Range.Find`、`Content.Find` 都一样。这不是偶尔失败，是稳定复现的。

**不要用它做替换。** 改用字符定位法——先读出范围内文本，算出子串偏移，再直接对 `Range` 赋值：

```python
def replace_in_range(doc, rng, old, new):
    txt = rng.Text
    positions = []
    j = txt.find(old)
    while j >= 0:
        positions.append(rng.Start + j)
        j = txt.find(old, j + 1)
    for pos in sorted(positions, reverse=True):      # 从后往前，避免偏移错乱
        doc.Range(pos, pos + len(old)).Text = new
    return len(positions)
```

把文档的段落（**含表格单元格内的段落**）逐个过一遍就行。这样既绕开了 WPS 的 Find bug，又能精确保住段落格式。

## 坑 2：整段赋值会把格式污染到新文本

`rng.Text = 整段新文本` 时，新文本会继承该 Range 起始处的字符属性，实测会凭空长出**青色高亮底纹**（`HighlightColorIndex = 9999999`）和**红色字体**（`Font.Color = 255`）。

**做法：**
- 优先用「子串替换」（只替换命中的那几个字），不要整段重写。子串替换实测不会污染，新文本继承原位置的格式，正好。
- 万一还是脏了，事后统一清理：

```python
for i in range(1, doc.Tables(1).Range.Cells.Count + 1):
    c = doc.Tables(1).Range.Cells(i)
    c.Range.HighlightColorIndex = 0     # 清高亮
    c.Range.Font.Color = -16777216      # wdColorAutomatic，恢复"自动"色
```

注意 `-16777216` 是「自动色」，`0` 是「纯黑」，视觉一样但属性不同。原文档是自动色就还原成自动色，别改成本色。

## 坑 3：`os.replace()` 跨盘会失败

把临时文件搬到另一个盘（例如 C: → E:）会报 `[WinError 17] 系统无法将文件移到不同的磁盘驱动器`。跨盘一律用 `shutil.copy()`（要保留元数据用 `shutil.copy2`）。

## 坑 4：旧版 .doc（二进制格式）的字节校验

WPS 保存 `.doc` 后想确认改动真的落盘，用 UTF-16LE 直接搜关键词：

```python
data = open(path, "rb").read()
print(旧文本.encode("utf-16-le") in data, 新文本.encode("utf-16-le") in data)
```

比重新用 COM 读一遍更硬。
