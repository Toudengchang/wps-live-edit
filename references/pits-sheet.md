# 表格（et/ket）的坑 · 坑 5~44

> wps-live-edit 坑库 · 表格分册（坑号按历史编号保留，与各处交叉引用一致）。与文字/演示共用的铁律、双模式和收尾规则见根目录 SKILL.md。

## 坑 5：逐格写入会超时

`sh.Cells(r, c).Value = v` 逐格写几十行 × 几列（数百次 COM 调用）就会明显变慢，量一大直接超时。COM 每次跨进程调用都有固定开销，单元格级别调用累积后非常慢。

**必须用 `Range.Value` 批量赋值，且数据类型必须是 `tuple`（不能用 `list`）：**

```python
# 正确：一次 Range 调用写完
data = tuple(tuple(row) for row in rows_list)
sh.Range(sh.Cells(3, 1), sh.Cells(3 + len(data) - 1, len(data[0]))).Value = data
```

## 坑 6：`ClearContents` 会清掉公式

`sh.Range(...).ClearContents()` 清区域数据时，公式也一起没了。如果区域中有 COUNTIFS、XLOOKUP 等公式，清完后必须重新写入。

**两种避法：**
- 清区域时避开公式列，只清值列
- 清完后重新写回所有公式（用 `Cells(r, c).Formula = "=..."`）

## 坑 7：`ClearContents` 留空行影响公式和筛选

删除数据时只清内容不删行，会留下空行。空行会导致 COUNTIFS 统计错误、筛选出现空白项、XLOOKUP 命中空值。

**删行必须用 `EntireRow.Delete()`，不要用 `ClearContents`：**

```python
sh.Rows(63).Delete(Shift=-4161)   # xlUp，下方行上移
```

批量删多行时，**从后往前删**（行号大的先删），否则删一行后后面行号全变。

## 坑 8：`Range.Sort(OrderCustom=N)` 不生效，必须改用 `Sort` 对象

WPS 表格对 `Range.Sort(OrderCustom=N)` 的实现是**残缺的**：传任何值都不生效，也不报错。实测遍历 `OrderCustom=1~18`，排序结果完全一样（退化成固定顺序，自定义序列完全丢失）。

**正解：走 `Worksheet.Sort` 对象 + `SortFields`——这正是用户手工点"数据 → 排序"走的同一条代码路径：**

```python
so = sh.Sort
so.SortFields.Clear()
so.SortFields.Add(Key=sh.Cells(3, 关键字列), SortOn=0, Order=1,
                  CustomOrder="紧急,高,中,低")   # 自定义序列：逗号分隔字符串，或序列索引
so.SortFields.Add(Key=sh.Cells(3, 第二列), SortOn=0, Order=1)   # 第二级
so.SetRange(sh.Range(sh.Cells(3, 1), sh.Cells(末行, 末列)))
so.Header = 2      # 2=xlNo（排序区域不含表头）；区域含表头时传 1
so.Apply()
```

`CustomOrder` 传自定义序列字符串（逗号分隔），也可以传自定义序列的索引，两种都生效。

**第二级排序的陷阱**：如果第二关键字是"有值 / 空值"型字段，**空单元格在排序中永远排最后**，升序、降序、CustomOrder 都改不动它。此时：
- 业务上不要求该列先后顺序 → 直接排上就行，无需额外处理
- 确实要求空值排在前 → 只能临时加辅助列 `=IF(E3="",2,1)` 把空值转成数值再排，**用完立即删除**；且加列前必须先确认那一列确实是空的，别覆盖其他任务的数据

**不要为了排序往数据表加辅助列**：其他任务的最后一列很可能有值，加列会直接覆盖数据。

## 坑 9：sheet 名 COM 读出乱码

`wb.Sheets("汇总")` 按名字取 sheet 可能取不到——COM 读出的 sheet 名在某些环境下是乱码（编码问题），和用户在 WPS 界面看到的不一致。

**用序号访问，不要用名字：**

```python
sh = wb.Sheets.Item(1)   # 按 1-based 序号取，不依赖名字
# 如需确认 sheet 身份，读 A1 值或 UsedRange 内容来核对
```

## 坑 10：写完公式不自动算

写入 XLOOKUP、COUNTIFS 等公式后，WPS 可能不立即重算，读到的 `.Value` 是缓存的旧值。

**写完公式后调用 `app.Calculate()` 强制全表重算：**

```python
sh.Cells(r, 5).Formula = f"=COUNTIFS(清单!J:J,D{r})"
app.Calculate()   # 强制重算，之后读 .Value 才是新值
```

## 坑 11：逐格读写都慢，让数据在 WPS 内部流转（不进 Python）

不只是逐格**写**慢（坑 5），逐格**读**同样慢。数据量小时还好，数据量一大，把内容灌进 Python 内存再处理会非常慢。

**原则：Python 只发指令，数据全程留在 WPS 里。** 四种常用手法：

```python
# 手法1 剪贴板复制粘贴（筛选后搬运，数据不经 Python）
#   ⚠️ 必须「筛含表头、抄不含表头」，且 Copy 与 Paste 同会话，详见下方「手法 1 的四个硬限制」
src.Range("A1:I%d" % 末行).AutoFilter(Field=6, Criteria1="华东")   # 筛选范围必须含表头（A1 起）
vis = src.Range("A2:I%d" % 末行).SpecialCells(12)                 # 12 = xlCellTypeVisible，从数据首行起
vis.Copy()                                                      # 必须在同一脚本内 Copy
tgt.Range("A3").PasteSpecial(Paste=-4163)                        # -4163 = xlPasteValues，只贴值
app.CutCopyMode = False                                         # 立刻清虚线框
src.AutoFilterMode = False                                      # 用完取消筛选

# 手法2 批量写公式（不读数据，WPS 自动按相对引用调整行号）
sh.Range(sh.Cells(3, 7), sh.Cells(63, 7)).Formula = "=XLOOKUP(B3,索引!$A:$A,索引!$C:$C)"

# 手法3 整块左右/上下搬运
sh.Range("D1:H24").Copy()
sh.Range("K1:O24").PasteSpecial(Paste=-4163)

# 手法4 排序交给 SortFields（见坑 8），不要自己算顺序重写
```

#### 手法 1 的四个硬限制：筛含表头，抄不含表头

> 真实事故复盘产出：用旧版写法（`A2` 起筛 + `SpecialCells("A2:..")` 之外的变体）搬 60 行名单，结果 61 行里只有 46 个唯一值、15 行重复、14 行缺失，且清单数据区首行变成了表头文字。
>
> 排除"源端/目标端带筛选状态"（4 组受控对照全部 100% 正确，已证伪）后，定位到下面两个**组合生效**的静默错位条件。

```python
# sh = 源表, tgt = 目标表, fld = 筛选字段号(从1起), val = 筛选值, n = 应搬运行数
lrow = sh.Cells(sh.Rows.Count, 1).End(-4162).Row            # 源表末行

sh.Range("A1:I%d" % lrow).AutoFilter(Field=fld, Criteria1=val)  # ① 必须 A1 起（含表头）
vis = sh.Range("A2:I%d" % lrow).SpecialCells(12)                # ② 必须 A2 起（不含表头）
vis.Copy()                                                      # ③ 必须同一脚本内 Copy
tgt.Range("A3").PasteSpecial(Paste=-4103)                       # xlPasteAll
app.CutCopyMode = False                                         # ④ 立刻清虚线框

# ⑤ 必做：回读落点自检，全过才算成功
assert tgt.Cells(3, 1).Value not in ("序号", "姓名", "日期"), "落点被表头占据 → 已下移 1 行"
assert tgt.Cells(3, 2).Value == 首个应到值
assert tgt.Cells(2 + n, 2).Value == 末个应到值
assert tgt.Cells(3 + n, 2).Value in (None, ""), "末行下一行非空 → 溢出"
```

**铁律 ①：`AutoFilter` 范围必须含表头行（A1 起）**

AutoFilter 把 Range 的**第一行**当列标题行：永久可见、**永远不参与筛选**。若从 `A2` 起筛，第 2 行就成了列标题，无论它是什么值都会**无条件显示、无条件进入 SpecialCells 结果**。

实测（把第 2 行临时改成"华南"，筛"华东"）：

| 写法 | `Rows(2).Hidden` | `Area1` | 结论 |
|---|---|---|---|
| `Range("A2:I301").AutoFilter(6,"华东")` | `False` | `$A$2:$I$2` | 该行没被筛掉，被当成结果一起带走 ❌ |
| `Range("A1:I301").AutoFilter(6,"华东")` | `True` | — | 第 2 行被正确筛掉 ✅ |

**铁律 ②：`SpecialCells` 范围必须从数据首行起（不含表头）**

表头行因为"永久可见"会被 `SpecialCells(12)` 收入，与紧邻的首个可见数据行**合并成双行 Area**（如 `$A$1:$I$2`）。复制内容因此多 1 行 → 粘到 A3 后表头占了 A3、真实数据从 A4 起 → **整体下移 1 行，末尾溢出 1 行**。

实测对照（同 60 行、同目标表、粘贴起点 A3）：

| 写法 | `Area1` | 粘贴后 `A3` | 粘贴后末行 | 结论 |
|---|---|---|---|---|
| `SpecialCells("A1:I301")`（含表头） | `$A$1:$I$2` | `'序号'` | **A63** | 下移 1 行 ❌ |
| `SpecialCells("A2:I301")`（不含表头） | `$A$2:$I$2` | `'名002'` | **A62** | 落点正确 ✅ |

**注意**：两种写法的 `Areas.Count` 都是 60，**只看 Area 个数判断不出错位**，必须看 `Area1` 的地址跨度（见坑 41）。

**铁律 ③：`Copy()` 与 `PasteSpecial()` 必须在同一脚本内**

见坑 40：剪贴板与发出 Copy 的脚本进程绑定，跨脚本粘贴会静默贴空。

**铁律 ④：粘贴完必须回读落点，不能只看"没报错"**

判读经验：
- `A3` 出现表头文字 = 已下移 1 行
- 末行比预期多 1 = 复制内容多 1 行
- 末行下一行非空 = 溢出，目标表行数不够

**口诀：筛含表头，抄不含表头。A1 起筛，A2 起抄。**

可直接用 `scripts/wps_com_utils.py` 的 `filter_move(...)`，它把上面四条都包好了。

## 坑 12：贴数据前先校对表头

上游数据（别人发的表、系统导出）列结构可能变。**贴之前逐列对比源表和目标表的表头**，对不上就停下报告，不要闷头贴——错列比不贴更糟。

## 坑 13：清空重填时，公式列的行数要跟着数据走

数据变少时（如 61 行变 50 行），如果只清值列、公式列留着旧行数，会留下多余公式。**做法：整区清空（含公式列）→ 贴数据 → 按实际数据行数精确重写公式 → 末尾多余行整行 `Delete()`。**

```python
# 找到数据最后一行，删除其后的所有残留行
for r in range(used_last, data_last, -1):
    sh.Rows(r).Delete(Shift=-4161)
```

## 坑 14：动态数组（LET/FILTER）固化必须覆盖整个溢出区

LET、FILTER、UNIQUE、HSTACK 等动态数组函数，`.Formula` **只在左上角锚点格有值**，溢出到下面/右边的单元格读 `.Formula` 是空的。直接改锚点格会报"不能更改数组的某一部分"。

**正确做法**：`Range.Copy()` + `PasteSpecial(-4163)`（贴值）必须覆盖**整个溢出区域**，从锚点格到最后一行/最后一列，不能只贴锚点那一格。

```python
# 错误：只贴锚点
sh.Cells(2, 65).Copy(); sh.Cells(2, 65).PasteSpecial(Paste=-4163)   # 报错

# 正确：覆盖整个溢出区
last_row = sh.Cells(sh.Rows.Count, 65).End(-4162).Row
rng = sh.Range(sh.Cells(1, 65), sh.Cells(last_row, 66))
rng.Copy(); rng.PasteSpecial(Paste=-4163)
app.CutCopyMode = False
```

## 坑 15：`SpecialCells(12).EntireRow.Delete` 超多区域时**静默失败**

筛选后可见区域超过几千个时（例如 7220 个非连续可见行区域），
`sh.Range(sh.Rows(2), sh.Rows(last)).SpecialCells(12).EntireRow.Delete(Shift=-4161)`
会**不报错、也不删任何行**，行数纹丝不动。

**首选做法（= 用户手工法）**：筛选状态下直接对**连续行区域**删，WPS 会只删可见行，
等价于手工「选中第 2 行整行 → Ctrl+Shift+↓ → 删除」：

```python
sh.Rows("2:%d" % last).Delete(Shift=-4161)
```

**备用做法**：自己按标志列切出连续同值块，从下往上每 250 块 `app.Union(...)` 合成多区域 Range，
再一次性 `.Delete(Shift=-4161)`，每批用 `End(xlUp)` 复核最后行。

## 坑 16：`UsedRange.Rows.Count` 删行后会黏住旧值

删行后用 `UsedRange.Rows.Count` 核对行数不准——它可能还停留在删之前的数字，不会自动收缩。
**核行数必须用** `sh.Cells(sh.Rows.Count, col).End(-4162).Row`（xlUp）。

## 坑 17：大批量操作前先把计算模式切手动

删几万行、贴大量值、固化动态数组时，WPS 会边操作边重算，巨慢还容易卡死。
做法：

```python
calc_bak = app.Calculation
app.Calculation = -4135   # xlCalculationManual
try:
    ...   # 批量操作
finally:
    app.Calculation = calc_bak
```

## 坑 18：工作簿序号会漂移，每次操作前重新核对

操作过程中如果新建/关闭了别的工作簿，`Workbooks.Item(n)` 的序号会变，可能指向错的文件。
每次批量操作前，按工作簿名或 A1 特征值重新确认目标工作簿的序号。

## 坑 19：扫描文件时跳过 `~$` 锁文件

WPS 打开的文件，同目录会有 `~$文件名.xlsx` 的锁文件。按修改时间排序时它可能排最前面，
用 zipfile/openpyxl 去读会 Permission denied。**候选文件过滤必须加 `not f.startswith("~$")`**。

## 坑 20：筛选状态下整片 `ClearContents` 只清可见行

工作表处于筛选状态时，`sh.Range("M2:N53").ClearContents()` **只作用在可见行**上，被筛掉（隐藏）的行内容原样保留——看着"清了"，其实没清干净，后面统计就会多算。

**避法：先取消筛选再清（`sh.AutoFilterMode = False`），或改为逐格补清。**

清多个不连续区域可以一次写完，等价于手工 Ctrl 多选 + Del：

```python
sh.Range("M2:N53,R7:S22,T18,U52").ClearContents
```

## 坑 21：`Worksheet.Copy` 建副本会被 WPS 吞掉参数

用 WPS 原生"创建副本"搬表（表头、格式、公式会一起带过来）走 `Worksheet.Copy`。
但 WPS 对 `Copy` 的参数实现残缺：**传进去的位置/前后参数会被吞掉**，副本可能被丢进一个新建的"工作簿N"里，而不是你指定的位置。

**避法：换传参方式，或复制后显式把副本移到目标工作簿/位置再改名，搬完核对副本实际落在哪个工作簿。**

## 坑 22：单元格底色（`Interior.Color`）是 BGR 序，写完要回读校验

标色用 `Range.Interior.Color = <整数>`。**这个整数是 BGR 顺序**（不是 RGB），别把 BGR 值和视觉色的 RGB 十六进制混写：

```python
sh.Cells(r, c).Interior.Color = 3368703   # 0x3366FF（BGR 序）→ 视觉色 #FF6633（橙）
```

- 换算：视觉色 RGB → BGR 十进制 = R + G×256 + B×65536；写注释时别把 BGR 值标成 RGB 十六进制
- 同一工作簿里同一种语义色只用一个固定值，别两套混用

写完**必须回读 `Interior.Color` 比对**，值一致才算标色成功，不能凭"命令没报错"就认定成了。

## 坑 23：去外链——直接删 `[文件名]` 前缀会变 `#REF!`

跨工作簿粘贴/引用后，WPS 会把公式写成 `[某数据源.xlsx]Sheet1!A1` 这种**真外链**（取的是源工作簿的数）。
想把它变回本簿引用，**不能简单把 `[文件名]` 替换为空**——那样只会得到 `#REF!`，一批本来还能取数的公式直接废掉。

**做法：只有确认本簿存在同名 sheet 时才替换前缀；否则一律先停下跟用户确认怎么处置，不要闷头替换。**
确认安全后，可用内置查找替换一次性抹掉所有 `[文件名]` 前缀（实测一次能处理数百处）。

## 坑 24：任何写入都要回读校验后再 `Save()`

写值、写公式、标色之后，都**回读**确认真成了再 `Save()`：
- 公式回读 `.Formula`（公式串），不是 `.Value`（那是算出来的值）
- 颜色回读 `.Interior.Color`
- 值回读 `.Value` / `.Value2`

不要凭"命令没报错"就认为改成了。

## 坑 25：读也走挂载 COM，不要用 openpyxl 直读磁盘

读取同样挂载已开的 WPS 实例、走 COM 实时读（`sh.Range(...).Value`），**不要让 openpyxl 直接去读磁盘文件**——文件正被 WPS 独占时可能失败，也可能读到手改前的旧快照。

要读哪份，就让用户在 WPS 里开着哪份；**读完不要关用户的文件**，留给用户自己核对。

## 坑 26：逐格写会触发全表重算（不只是慢）

大表（成百上千公式）里逐格写，每写一格都可能触发一次全表重算——清 24 格 = 重算 24 遍，会非常慢。

**避法：改成一整次 `Range` 范围操作；真正的大批量先临时切手动计算（见坑 17），改完再切回并 `Calculate()` 一次。**

## 坑 27：业务数据里的"双值文本"套公式会 `#VALUE!`

有些表某格塞的是"双值文本"（一格写两个百分数，如 `12.93%/12.37%`）。这种格子**套进算术公式会直接 `#VALUE!`**。

**做法：需要参与计算的列，先和用户确认数值口径，把需要的数值单独取出、填成纯数字，不要拿双值文本格去算。**

## 坑 28：`Range.NumberFormat` 对多格区域只返回**单个字符串**，不能批量比对

`.Value` / `.Formula` 读多格区域返回的是 tuple of tuples，但 **`.NumberFormat` 对整片区域只返回一个字符串**（区域格式一致时是那个格式串，不一致时是 `None` / 混合），**不是逐格 tuple**。

拿它去跟逐格的 openpyxl 格式列表比对，会**大面积误报"格式不同"**（实测一次误报 1088 处）。

**做法：格式只能逐格读** `sh.Cells(r, c).NumberFormat` 再比对，不要指望一次 Range 调用拿到整片格式。

## 坑 29：COM 读数字统一成 double，跨源比对前必须归一化

WPS COM 读出来的数字统一是 **double**（单元格写 `25`，读出来是 `25.0`）；而 openpyxl 读同样内容可能是 **int**（`25`）。两者直接比会产生**大面积假差异**（"值源=25 / 主表=25.0"这种噪音），这是跨源核对"满屏不同"的常见根因。

**做法：跨源比对前先归一化**——

```python
def norm(v):
    if v is None or v == "":
        return ""
    if isinstance(v, (int, float)):
        return str(round(float(v), 6))
    return str(v).strip()
```

数字统一 `round(x, 6)` 再转 `str`（或都转 float），空值（`None` / `""` / NaN）统一成同一种表示。

## 坑 30：按名找簿/找表要容错（名字可能带 `\xa0`、全角空格）

WPS 里工作簿 / sheet 名常夹带**不换行空格 `\xa0`**、全角空格或首尾空白，界面看着一样，但**全等匹配会静默找不到**。

**做法：用 `in` / `startswith` 这类包含匹配，或先 `strip()` + 把 `\xa0` / 全角空格规整后再比。**
坑 18 讲的是"序号会漂移"，这里讲的是"名字匹配不上"，两类问题都要防。

## 坑 31：写跨表公式前先 dump 真实 sheet 名（含空格要加单引号）

坑 9 说"名字乱码时用序号访问"，那说的是**通过 COM 定位 sheet**。但**写跨表公式**（如 `=SUM(BU!S10:S13)`）时必须用**真实 sheet 名**——名字里有空格或特殊字符还要加英文单引号：`=SUM('某 表'!S10:S13)`。

**做法：写公式前先把目标 sheet 的真实名字 dump 出来核对再用**，别凭记忆或界面显示的名字硬写。这是和坑 9 互补的另一类场景。

## 坑 32（省 token 的读法）：批量取主体 → Python 内过滤比对 → 只回读结论

整表 `dump` 进上下文极耗 token。**推荐：用一次 `Range.Value` 把主体数据批量取到 Python，在 Python 里过滤/比对，只把最终结论写进一个小文件再读回来**，不要让整张表的内容进上下文。

## 坑 33：整块 `Copy + PasteSpecial` 时，绝对引用不会随粘贴偏移

**现象：** 把做好的表格块（如 R5:R32）整块复制到 R36:R63，`=E7/D7` 这类相对引用会自动变成 `=E38/D38`，但 `=RANK(F7,$F$7:$F$21)` 里的 `$F$7:$F$21` **纹丝不动**，结果全表 `#N/A`，COM 读 `.Value` 是整数 `-2146826246`。

**根因：** 复制粘贴只调整相对引用，`$` 锁定的绝对引用永远不变。块内一旦有「引用本块内区域的绝对引用」（排名、占比、条件统计），整块复制必然错位。

**正解：** 粘贴后逐列检查含 `$` 的公式，按行号偏移手工重写：

```python
# 块1 在 R7:R21，块2 在 R38:R52，偏移 +31
for r in range(38, 53):
    sh.Cells(r, 7).Formula = '=RANK(F{0},$F$38:$F$52)'.format(r)

# 多区域联合引用同理
for r in range(55, 62):
    sh.Cells(r, 7).Formula = '=RANK(F{0},($F$38:$F$44,$F$46:$F$52,$F$55:$F$61))'.format(r)
```

**验收：** 粘贴后立刻扫目标区 `.Text`，以 `#` 开头的都要查（见坑 37）。

**补充：** `-2146826246` = VBA 域错误码，等价 `#N/A`。看到这个数字先查公式引用区域，不要以为是数据问题。

## 坑 34：设 `WrapText = True` 会把行高自动撑大

设 `WrapText=True` 时 WPS 会按内容长度自动调行高，并覆盖原有行高（实测 77 → 409.5，整张表被撑得很难看）。

**正解：先设 `WrapText`，再显式设回 `RowHeight`**（顺序反了会被再次覆盖）：

```python
sh.Cells(63, 3).WrapText = True
sh.Cells(63, 13).WrapText = True
sh.Rows(63).RowHeight = sh.Rows(32).RowHeight   # 取源块的行高
```

## 坑 35：PowerShell 内联 `python -c` 时，公式里的 `$` 会被插值

```powershell
# 错：$F 被 PowerShell 当变量，报 InvalidVariableReferenceWithDrive，或公式被吃掉
python -c "sh.Cells(r,7).Formula='=RANK(F%d,$F$38:$F$52)'%r"
```

**正解：凡是要写含 `$` 的 Excel 公式，一律写成 `.py` 文件再执行**，不要用 `python -c` 内联。`\$` 是 bash 的转义写法，PowerShell 不认。

## 坑 36：`RPC_E_CALL_REJECTED (-2147418111)` 必须加重试包装

**现象：** 长时间连续操作时，某次读 sheet / 读单元格突然报 `com_error: (-2147418111, '被呼叫方拒绝接收呼叫。')`，脚本中断，前面的写入还只在内存里。

**根因：** WPS 正在处理上一条指令或用户正在操作，COM 通道临时忙，**是瞬时的**，隔几百毫秒重试即可。

**正解：所有 COM 调用套一层重试，不要裸调。**

```python
def rtry(fn, n=15, w=0.7):
    for i in range(n):
        try:
            return fn()
        except Exception:
            if i == n - 1:
                raise
            time.sleep(w)
```

重试对**读**绝对安全；对**写**要注意幂等性（`Cells.Value = x` 幂等，`Rows.Delete()` 不幂等，删行的重试要放外层先判断是否已删）。可直接用 `scripts/wps_com_utils.py` 的 `rtry`。

## 坑 37：批量找错误值，扫 `.Text` 而不是 `.Value`

`.Value` 读到的是 `-2146826246` 这类整数，看不出是什么错；**`.Text` 会返回 `#N/A`、`#VALUE!`、`#DIV/0!`、`#REF!`**，判断以 `#` 开头即可。

```python
# 整块批量读（一次 COM 往返），判断放在 python 里 —— 见坑 44
rng = sh.Range(sh.Cells(r1, c1), sh.Cells(r2, c2))
vals = rtry(lambda: rng.Text)          # 显示文本：错误值是 '#N/A' 这类字符串
if not isinstance(vals, tuple):
    vals = ((vals,),)
bad = []
for i, row in enumerate(vals):
    if not isinstance(row, tuple):
        row = (row,)
    for j, t in enumerate(row):
        if isinstance(t, str) and t.startswith('#'):
            bad.append((r1 + i, c1 + j, t))
```

> ⚠️ 逐格扫描（`Cells(r,c).Text` 循环）只适合几百格的小区域。区域一上万格就会拖到几分钟，见坑 44：一律改成整块批量读。

| `.Value` 整数 | 含义 |
|---|---|
| `-2146826246` | `#N/A` |
| `-2146826281` | `#VALUE!` |
| `-2146826289` | `#DIV/0!` |
| `-2146826273` | `#REF!` |
| `-2146826265` | `#NAME?` |

## 坑 38：按名找工作簿返回 `None` 必须先判空

用户中途关掉某个工作簿是常态，`wb_by()` 找不到会返回 `None`，下一行 `.Sheets` 就 `AttributeError: 'NoneType' object has no attribute 'Sheets'`。

**正解：判空 + 明确报错 + 准备降级路径；动手前先列出当前打开清单。**

```python
w = rtry(lambda: wb_by(app, '台账'))
if w is None:
    print('目标工作簿未打开，改用其他数据源')
    # 降级：从别的 sheet 取数

for i in range(1, app.Workbooks.Count + 1):
    w = app.Workbooks(i)
    print(i, w.Name, 'Saved=', w.Saved)
```

## 坑 39：业务口径不要自己按明细重算，先找现成台账 sheet

**现象：** 按明细表的「负责人」字段 + 明细行自己重算某项考核指标，得到的人数、达标率与用户历史口径全对不上。

**根因：** 明细里的「负责人/经办人」字段混着大量非考核对象（协作岗、外包岗、管理岗等）；考核名单是**单独维护**的，自己按明细聚合等于悄悄换了口径。

**正解：先扫一遍所有 sheet 名和行数，找现成的考核/通报表。**

```python
for i in range(1, wb.Sheets.Count + 1):
    print(i, wb.Sheets(i).Name, wb.Sheets(i).UsedRange.Rows.Count)
```

实战中正是这样发现工作簿里已有算好的通报表（行数 = 表头行数 + 应考核人数），直接读数，与历史口径完全吻合。

**通用原则：** 涉及「考核 / 通报 / 排名」这类有固定口径的指标，工作簿里大概率已有算好的表；**先找台账/索引/通报 sheet，找不到再问用户**，不要默认按明细字段聚合。

**附带：** 遇到「分母为 0」这类口径细节（如 0/0 算不算达标），不同业务有不同约定，先问用户口径，不要自行假设。

## 坑 40：`CutCopyMode` 虚线框「看起来还在，实际已失效」——跨脚本粘贴会静默贴空

**现象：** 脚本 A 里 `Copy()` 后没有清虚线框（屏幕上虚线框还在、`CutCopyMode=1`）；脚本 B（新进程 / 新 COM 会话）里直接 `PasteSpecial` → **什么都没贴进去、不报错、不抛异常**。

**实测：**
- 脚本 B 首次运行：目标列全 `None`，末个非空行停在表头行（`A3='序号'`）
- 同脚本内补一次 `vis.Copy()` 后立刻正常：`B3='名002'`、`B4='名007'`、末行正确到 A62

**原因：** 剪贴板内容与发出 `Copy` 的脚本进程绑定，脚本退出即失效；**但界面上的虚线框不会自动消失**，视觉上骗人，很容易误判"剪贴板还在"。

**对策：**
1. 粘贴前必须在**同一脚本内**重新 `Copy()`，绝不能指望上一个脚本留下的剪贴板
2. 粘贴完成后立即 `app.CutCopyMode = False`
3. "没生效"的信号：`CutCopyMode` 还是 1，但目标区一个字都没多

## 坑 41：判断「复制了几行」不能看 `Areas.Count`

`SpecialCells(12).Areas.Count = 60` 不代表 60 行——第一个 Area 可能是双行，真实复制行数 = 各 Area 跨度之和。

实测：`Areas.Count = 60`，但 `Area1 = $A$1:$I$2`（2 行）+ 其余 59 个单行 = **62 行**。

**对策：**

```python
total = 0
for i in range(1, vis.Areas.Count + 1):
    total += vis.Areas.Item(i).Rows.Count
print(vis.Areas.Item(1).Address)     # 一定要打 Area1 完整地址，不要只数个数
assert total == 预期行数, "复制行数不符，停下查，不要继续粘贴"
```

## 坑 42（方法论）：复现 / 演示前必须按**真实文件形状**构造数据

**教训：** 做一次筛选搬运事故的复现时，我构造的演示簿是 `A1 表头 / A2 空行 / A3 起数据`，而真实源表是 `A1 表头 / A2 就是数据行、无空行`。因为演示簿多了一个空行，从 A2 起筛时 A2 成列标题、与首个数据行合并成双行 Area，于是得出"**空行导致下移**"的错误结论，被用户连问两次才纠正。

**规则：做任何复现 / 演示 / 受控实验前，先只读核对真实文件——**

1. 表头在第几行
2. 数据从第几行起
3. 中间有没有空行 / 合并单元格 / 隐藏行 / 隐藏列
4. 末行行号、末列列号
5. 是否有辅助统计块（如源表右侧还有备注与统计列）

**为什么必须守：** 正确结论（表头行被 `SpecialCells` 收入）和错误结论（空行被收入）**外在表现完全相同**——都是 `$A$1:$I$2` 这类双行 Area、都是下移 1 行。只有回到真实形状才分得清。

**推论纪律：** 没实测过的解释不要写进结论；已被数据推翻的结论必须显式撤回，不能留在技能里。

## 坑 43：工作簿名、sheet 顺序都不可靠

| 现象 | 实测 | 对策 |
|---|---|---|
| `Workbook.Name = "xxx"` 赋值 | **不真生效**：赋值后立即读回是新名，跨脚本读回又变回默认名"工作簿N" | 定位工作簿**按 sheet 名判断**，不要按簿名 |
| `Sheets.Add()` 的位置 | 新 sheet 被插到**活动 sheet 之前**，`Item(1)` / `Item(2)` 顺序会变 | 不按序号假定顺序，按名字取 |
| `finally` 里读 `wb.Name` | 可能抛 `com_error`，导致后面的还原逻辑**全部跳过**（曾造成测试簿残留 + `DisplayAlerts=False` 没还原） | 先把簿名存进局部变量，还原逻辑放在最外层 finally 且不依赖任何 COM 读 |
| `et.Application` | 本机报"无效的类字符串" | WPS 抢先版表格用 `ket.Application` |

## 坑 44：COM 调用次数本身就是性能瓶颈——禁止逐格读写，批量读 + 手动计算

**现象：** 一次例行更新脚本跑了好几分钟，业务逻辑其实很简单（贴一块数据、修几条公式、排一次序）。复盘确认，时间几乎全花在「读单元格」和「重算公式」两件事上。

**根因一：逐格 COM 读取 = 十几万次跨进程 RPC**

```python
# 反面写法：一个格一次调用
for r in range(r1, r2 + 1):
    for c in range(c1, c2 + 1):
        t = sh.Cells(r, c).Text          # ← 每格一次 COM 往返
```

本次量级：一张 1788 行 × 64 列的主表 ≈ 11.4 万次、一张 1612 × 17 的表 ≈ 2.7 万次，五张表合计约 **14.7 万次**。每次 `Cells().Text` 都是一趟 WPS RPC，哪怕按 1ms 算也要两分半，实测更慢。

**根因二：`app.Calculate()` 全量重算**

公式列大量用**整列引用**（`XLOOKUP` 扫其他 sheet 的整列、或跨工作簿外链），行数一上来每次重算就是百万级扫描；再叠加多 sheet 的 COUNTIFS 联动，单次 `Calculate` 就很沉。更要命的是**删行、贴数、改公式期间若保持自动计算，会被反复触发**。

**正解 (a)：整块批量读写，判断放到 Python 里**

```python
# 一次调用把整块取回
rng = sh.Range(sh.Cells(r1, c1), sh.Cells(r2, c2))
vals = rtry(lambda: rng.Text)          # 显示文本，错误值为 '#N/A' 这类字符串
if not isinstance(vals, tuple):
    vals = ((vals,),)
bad = []
for i, row in enumerate(vals):
    if not isinstance(row, tuple):
        row = (row,)
    for j, t in enumerate(row):
        if isinstance(t, str) and t.startswith('#'):
            bad.append((r1 + i, c1 + j, t))
```

同一数量级的数据，逐格读与整块读相差**几百倍**；本次扫描改成整块读后，主表从数分钟降到秒级。

**正解 (b)：写入期间切手动计算，全部改完再算一次**

```python
calc_bak = app.Calculation
app.Calculation = -4135                # xlCalculationManual
try:
    # 删行 / 贴数 / 改公式 / 排序 … 全部在这里做
    pass
finally:
    app.Calculation = calc_bak         # 恢复自动
app.Calculate()                        # 最后统一算一次
```

**动手前先估「预计 COM 调用次数」：** 几百次以内随便写；上万次必须批量；十万次以上必然卡几分钟，别写。

**注意：** 批量 `.Value` 拿的是计算值，错误显示为整数错误码（如 `-2146826246`）；批量 `.Text` 拿的是显示文本，错误是 `#N/A` 这种字符串，判错更方便。若某块区域批量 `.Text` 返回的不是二维数组，按行列自适应展开，或退回 `.Value` + 坑 37 的错误码表。

**本类任务的重算压力来源（写脚本前先看，决定要不要切手动计算）：**

- 整列引用的 `XLOOKUP`（本表 → 其他 sheet 的整列，或跨工作簿外链）
- 整列引用的 `COUNTIFS`（如 `AR:AR` × `BL:BL`）
- `LET` / `RIGHT` / `XLOOKUP` 组合列
- 多张表公式互相联动时，任一表变动都会触发全链重算

## 建表经验

1. **块复制是最高效的建表方式**：做好第一块后 `Copy + PasteSpecial(-4104)` 到目标位置，格式、合并单元格、行高、数字格式、公式结构全部继承，只需改「数值 + 表头文字 + 绝对引用区域」三类（绝对引用见坑 33）。
2. **`PasteSpecial(Paste=-4104)` = xlPasteAll**，保留公式/格式/边框；`-4163` 只贴值会丢公式，建表不要用。
3. **`Worksheet.Sort` 对象排序**（坑 8 延伸）：`SortFields.Add(...)` + `SetRange` + `Header=2` + `Apply()` 在 WPS 稳定可用，位置参数比命名参数可靠。
4. **改表头文字用字符串替换**（如 `'上月' → '本月'`），比逐字敲更容易保持一致；但前提是先读到完整原值，避免截断。
5. **写完数据必须 `app.Calculate()` 再读**，否则 RANK / COUNTIFS 这类依赖其他区域的公式读到的是旧值。
6. **动用户正在编辑的文件前先备份**：`shutil.copy2` 一份 `xxx_备份YYYYMMDD.xlsx`，并告诉用户备份路径。
