# -*- coding: utf-8 -*-
"""wps_com_utils.py —— 挂载式 WPS COM 编辑的通用工具函数。

配合 wps-live-edit 技能使用。所有函数都不新建 WPS 实例、不 Quit。

  rtry(fn, n=15, w=0.7)            COM 调用重试，专治 RPC_E_CALL_REJECTED(-2147418111)
  wb_by(app, name)                 按名字包含匹配找工作簿，容错 \\xa0 / 全角空格；找不到返回 None
  scan_errors(sh, r1, c1, r2, c2)  扫区域内以 # 开头的错误值（整块批量读 .Text，勿逐格）
  restore_app_state(app)           收尾：把 5 个全局开关拨回默认
  check_app_state(app)             收尾体检：返回异常项列表，空列表 = 全部正常
  norm(v)                          跨源比对归一化（COM 的 double vs openpyxl 的 int、空值统一）
  filter_move(app, sh, tgt, ...)   筛选后整块搬运：筛含表头、抄不含表头、同会话 Copy、粘贴后自检
"""
import time


def rtry(fn, n=15, w=0.7):
    """COM 调用重试包装，专治 RPC_E_CALL_REJECTED。

    重试对读操作绝对安全；写操作需自行保证幂等
    （Cells.Value = x 幂等；Rows.Delete() 不幂等，删行重试要放外层先判断是否已删）。
    """
    for i in range(n):
        try:
            return fn()
        except Exception:
            if i == n - 1:
                raise
            time.sleep(w)


def _clean_name(s):
    """把不换行空格 \\xa0、全角空格规整成普通空格并去首尾空白。"""
    s = str(s)
    return s.replace("\xa0", " ").replace("\u3000", " ").strip()


def wb_by(app, name):
    """按名字包含匹配找工作簿（容错 \\xa0、全角空格、首尾空白）。

    找到返回 workbook 对象，找不到返回 None —— 调用方必须先判空。
    """
    target = _clean_name(name)
    for i in range(1, app.Workbooks.Count + 1):
        try:
            w = app.Workbooks(i)
            if target in _clean_name(w.Name):
                return w
        except Exception:
            continue
    return None


def scan_errors(sh, r1, c1, r2, c2):
    """扫描区域内所有错误值，返回 [(row, col, '#N/A'), ...]。

    整块批量读（一次 COM 往返），判断放在 python 里 —— 逐格 Cells().Text
    在万格量级会拖到分钟级，见 references/pits-sheet.md 坑 44。
    读 .Text（字符串 #N/A / #VALUE! / #DIV/0! / #REF!），
    不要读 .Value（那是 -2146826246 这类整数，看不出是什么错）。
    """
    rng = sh.Range(sh.Cells(r1, c1), sh.Cells(r2, c2))
    vals = rtry(lambda: rng.Text)
    if not isinstance(vals, tuple):
        vals = ((vals,),)
    bad = []
    for i, row in enumerate(vals):
        if not isinstance(row, tuple):
            row = (row,)
        for j, t in enumerate(row):
            if isinstance(t, str) and t.startswith("#"):
                bad.append((r1 + i, c1 + j, t))
    return bad


_APP_DEFAULTS = (
    ("DisplayAlerts", True),
    ("ScreenUpdating", True),
    ("Interactive", True),
    ("Calculation", -4105),     # xlCalculationAutomatic
    ("CutCopyMode", False),
)


def restore_app_state(app):
    """收尾：把 WPS 全局开关全部拨回默认。

    只读写开关，不碰任何单元格、不保存任何文件。
    """
    for attr, val in _APP_DEFAULTS:
        try:
            setattr(app, attr, val)
        except Exception:
            pass


def check_app_state(app):
    """收尾体检：返回 [(开关, 当前值, 期望值), ...]，空列表表示全部正常。"""
    bad = []
    for attr, val in _APP_DEFAULTS:
        try:
            cur = getattr(app, attr)
        except Exception:
            continue
        if attr == "CutCopyMode":
            if cur not in (0, False, None):
                bad.append((attr, cur, val))
        elif attr == "Calculation":
            if cur != val:
                bad.append((attr, cur, val))
        else:
            if cur not in (True, 1):
                bad.append((attr, cur, val))
    return bad


def norm(v):
    """跨源比对归一化：COM 数字统一 double（25→25.0），openpyxl 可能是 int。

    数字统一 round(x, 6) 再转 str；None / "" / NaN 统一成 ""。
    """
    if v is None or v == "":
        return ""
    if isinstance(v, (int, float)):
        f = float(v)
        if f != f:          # NaN
            return ""
        return str(round(f, 6))
    return str(v).strip()


def filter_move(app, sh, tgt, fld, val, r_end, r_start=2, paste_cell="A3",
                col_last="I", expect_n=None, header_words=None):
    """筛选后把一块数据从源表搬到目标表（数据全程在 WPS 内部流转，不经 Python）。

    硬限制（真实事故复盘沉淀，详见 references/pits-sheet.md 坑 11「手法 1 的四个硬限制」）：
      ① AutoFilter 范围必须含表头（A1 起），否则数据首行会被当列标题、无条件进入结果
      ② SpecialCells 范围必须从数据首行起（不含表头），否则表头与首个可见行合并成双行
         Area，粘贴后整体下移 1 行、末尾溢出
      ③ Copy 与 PasteSpecial 必须在本函数（同一 COM 会话）内完成
      ④ 粘贴后回读落点自检，判据不过直接抛错，不会静默错位

    参数：
      app          挂载的 ket.Application
      sh           源表
      tgt          目标表
      fld          筛选字段号（从 1 起，对应 Range 的第 fld 列）
      val          筛选值
      r_end        源表数据末行
      r_start      源表数据首行，默认 2（即第 1 行是表头）
      paste_cell   粘贴起点，默认 "A3"
      col_last     源表数据末列字母，默认 "I"
      expect_n     预期搬运行数，给了就强校验，不符直接抛错不粘贴
      header_words 可选的表头词元组/列表；不给时默认用「源表表头首格」判定落点
                   是否被表头占据（更通用，推荐留空）

    返回：实际粘贴的数据行数。
    """
    rng_filter = "A1:%s%d" % (col_last, r_end)          # ① 含表头
    rng_value = "A%d:%s%d" % (r_start, col_last, r_end)  # ② 不含表头

    try:
        sh.AutoFilterMode = False                        # 清掉可能残留的旧筛选
    except Exception:
        pass
    sh.Range(rng_filter).AutoFilter(Field=fld, Criteria1=val)

    vis = sh.Range(rng_value).SpecialCells(12)           # 12 = xlCellTypeVisible

    area1 = vis.Areas.Item(1).Address
    if area1.startswith("$A$1:"):
        raise RuntimeError("Area1 含表头行（%s）：SpecialCells 范围写错了" % area1)

    total = 0                                            # ③ 真实行数 = 各 Area 跨度之和
    for i in range(1, vis.Areas.Count + 1):
        total += vis.Areas.Item(i).Rows.Count
    if total <= 0:
        raise RuntimeError("筛选结果为空：Field=%r Criteria1=%r" % (fld, val))
    if expect_n is not None and total != expect_n:
        raise RuntimeError("复制行数 %d != 预期 %d（Area1=%s），先查清再粘贴"
                           % (total, expect_n, area1))

    vis.Copy()                                           # 同一会话内 Copy
    tgt.Range(paste_cell).PasteSpecial(Paste=-4103)      # -4103 = xlPasteAll
    app.CutCopyMode = False                              # ④ 立刻清虚线框

    row0 = int("".join(ch for ch in paste_cell if ch.isdigit()) or "1")

    # ④ 落点自检：首格是否被表头占据、末尾是否溢出
    first = tgt.Cells(row0, 1).Value
    try:
        src_hdr = norm(sh.Cells(r_start - 1, 1).Value)   # 源表表头首格（r_start-1 行）
    except Exception:
        src_hdr = ""
    words = tuple(header_words or ())
    if isinstance(first, str):
        f = first.strip()
        if (src_hdr and f == src_hdr) or (words and f in words):
            raise RuntimeError("落点被表头占据（%s = %r）→ 已下移 1 行"
                               % (paste_cell, first))
    if tgt.Cells(row0 + total, 1).Value not in (None, ""):
        raise RuntimeError("粘贴溢出：第 %d 行非空，目标表行数不够" % (row0 + total))
    return total
