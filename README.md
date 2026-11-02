# wps-live-edit

> 让 AI 挂载你**正在用**的 WPS：不关源文件、不打断手头工作、像人一样一步步改 Word / Excel / PPT。

`Windows` · `WPS Office`（兼容 MS Word/Excel/PowerPoint）· `pywin32` · `Agent Skill`

---

## 为什么写这个

想必有许多朋友和我一样，工作中主要用 WPS，尤其经常用 WPS 表格处理数据。把这类活交给 AI 时，常见的情况是：

- AI 往往**重新造一个新文件**出来，或者要求你**必须关闭源文件**才能修改；
- 你**看不到 AI 正在改什么、改到哪一步了**；
- 反复开关文件，对**老旧机子和超大文件**都十分不友好；
- 长流程里没办法让 AI **主动停下来，让人接管**。

这个 skill 就是从近期的实际工作里一点点踩坑整理出来的：让 AI 修改 Word / Excel / PPT 时**不关闭源文件、不打断你的其他工作**，并且**像人一样**操作修改文档——可以像带实习生那样，让 AI 走一步、你看一步。

希望对你有帮助。**欢迎把你遇到的其他踩坑点提 issue**，实测验证后会合并进对应分册。

## 它是怎么工作的

- **挂载，不新建**：用 `GetActiveObject` 挂到你已经打开的 WPS 实例上，不另起进程、不另开副本、**绝不 `Quit()`、绝不关你的文档**；
- **只动目标文档**：写入后回读校验再 `Save()`，其余内容和格式原样保留；
- **两种操作模式**，执行前会先问你用哪种：

| | 勿扰模式（后台静默） | 显性模式（看得见） |
|---|---|---|
| 焦点 | 不抢焦点、不碰你的选区 | 把目标窗口顶到前台 |
| 动作 | 直接对目标对象读写 | 像人一样：激活工作簿 → 激活 sheet → 选中目标区域 → 写入 |
| 适用 | 你还要用电脑干别的 | 调试 / 演示，你要盯着它操作 |

- **善后强制还原**：`DisplayAlerts`、`ScreenUpdating`、`Calculation` 等 app 级开关先备份、后修改、`finally` 恢复，结束还有一次状态体检，不会把你的 WPS 改得"不受控制"。

## 使用案例

**例 1 · 显性模式（我看着你改）**

> 帮我把新数据贴到通报模板里，新数据我已经下载好并打开了。使用 wps-live-edit 技能，**显性模式**，我看着你操作；操作完先别保存，等我自己确认后再保存。

AI 会像人一样依次把目标工作簿顶到前台、激活对应 sheet、选中将改的区域再写入，每一步你都看得见；最后停在"未保存"状态等你确认。

**例 2 · 勿扰模式（别打扰我）**

> 把我框选出来的单元格按表头和列名信息补充完整，关键数据在 Sheet1 的 D–F 列。使用 wps-live-edit 技能，**勿扰模式**后台静默操作，别抢我的焦点，我还要干别的活。

AI 在后台直写目标区域，不 `Activate`、不碰你的选区；前后各采一次前台窗口标题自检，证明没打扰你。

## 踩坑库（这个 skill 的核心价值）

50 条全部来自真实使用中稳定复现的坑，其中不少是 WPS 特有行为，在 MS Office 的文档里根本查不到。随便举几个：

- **WPS 的 `Find.Execute(Replace=2)` 会假成功**：返回 `True`，但文档一个字都没改，且零报错（坑 1）；
- **PPT 里碰 `TextRange.ParagraphFormat` 会摧毁整页文本渲染**：文字不再换行、横向冲出页面，全程零报错、照样保存成功（坑 45，全局最高危红线）；
- **筛选 + `SpecialCells` 搬运会静默错位**：粘贴后整体下移 1 行、末尾溢出，而且 `Areas.Count` 完全看不出异常（坑 11 / 41）；
- **跨脚本粘贴静默贴空**：界面上虚线框还在，其实剪贴板早失效了（坑 40）。

坑按组件分册，按需加载：

| 分册 | 覆盖 | 代表内容 |
|---|---|---|
| [references/pits-word.md](wps-live-edit/references/pits-word.md) | 坑 1~4 | Find 假成功、整段赋值污染格式 |
| [references/pits-sheet.md](wps-live-edit/references/pits-sheet.md) | 坑 5~44 | 批量读写、筛选搬运四硬限制、RPC 重试、错误值扫描 |
| [references/pits-ppt.md](wps-live-edit/references/pits-ppt.md) | 坑 45~50 | ParagraphFormat 红线、纯净重建法、导 PNG 核验 |
| [references/api-matrix.md](wps-live-edit/references/api-matrix.md) | 速查 | ProgID 对照、三组件 API 对照、差异速查表 |

## 安装

1. 环境要求：Windows + 已安装 WPS Office（装了 MS Office 也能用，ProgID 会自动降级），Python 3 且已装 `pywin32`：

   ```
   pip install pywin32
   ```

2. 把 `wps-live-edit/` 整个目录放进你的 agent 技能目录，例如 Claude Code（Windows）：

   ```
   %USERPROFILE%\.claude\skills\wps-live-edit
   ```

   其他支持 Agent Skill 规范的工具同理（目录结构见下）。

3. 使用时先在 WPS 里打开目标文档，然后直接对 AI 说你的需求，并指明用哪种模式（见上方案例）。

## 目录结构

```
wps-live-edit/
├── SKILL.md                # 主技能文件：何时用 / 双模式 / 五条铁律 / 强制规则 / 标准流程
├── references/             # 坑库与速查（按需加载）
│   ├── pits-word.md        # 坑 1~4（文字）
│   ├── pits-sheet.md       # 坑 5~44（表格）+ 建表经验
│   ├── pits-ppt.md         # 坑 45~50（演示）+ 纯净重建法
│   └── api-matrix.md       # ProgID / API / 差异速查
├── scripts/
│   ├── wps_live.py         # 文字组件命令行封装（list / info / dump / find / replace）
│   └── wps_com_utils.py    # 表格公共函数库（重试、找簿、错误扫描、状态还原、筛选搬运）
├── icon.png
└── VERSION.txt
```

## 不适合的场景

- **无人值守的批量处理**——那种场景应该自己 `Dispatch` 开实例、做完 `Quit`，和本技能的路线正好相反；
- **目标文档根本没打开**——直接改文件更简单可靠，不必挂载。

## 贡献

遇到新的踩坑点欢迎提 issue，最好附上：WPS 版本号、最小复现代码、现象描述（有截图更好）。实测确认后会合并进对应分册并注明来源。PR 同样欢迎。

---
如果你也在让 AI 帮你干 WPS 的活，希望这个 skill 能让你少踩一遍我踩过的坑。
