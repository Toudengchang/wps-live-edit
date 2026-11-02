# -*- coding: utf-8 -*-
"""wps_live.py —— 挂载已打开的 WPS 文字实例，在后台读写文档。

设计原则（务必保持）：
  * 只用 GetActiveObject 挂载，绝不 Dispatch 新建实例
  * 两种操作模式，执行前先确认用哪种（见 SKILL.md「操作模式」）：
      - 勿扰模式（默认，本脚本当前实现）：不 Activate、不改 Visible、不碰 Selection，后台静默
      - 显性模式（另说，调试/演示用）：app.Visible=True → doc.Activate() → 目标 Range.Select() 再写入
  * 但无论哪种模式都绝不调用 Quit，只对目标文档 Save，不动用户的其他文档
  * 5 个全局开关（DisplayAlerts/ScreenUpdating/Interactive/Calculation/CutCopyMode）先备份后改，用完恢复并体检

用法：
  python wps_live.py list
  python wps_live.py info    "<文档名>"
  python wps_live.py dump    "<文档名>" [--limit N]
  python wps_live.py find    "<文档名>" "<关键词>"
  python wps_live.py replace "<文档名>" "<旧文本>" "<新文本>" [--save] [--cell-only]
"""
import argparse
import io
import os
import sys
import time

import pythoncom
import win32com.client as win32

if sys.stdout.encoding is None or sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PROGIDS = ("kwps.Application", "wps.Application", "Word.Application")


# ---------------------------------------------------------------- 基础

def attach():
    """挂到用户已经在运行的 WPS/Word 上。不新建实例。"""
    pythoncom.CoInitialize()
    last = None
    for pid in PROGIDS:
        try:
            app = win32.GetActiveObject(pid)
            return app, pid
        except Exception as e:          # noqa: BLE001
            last = e
    raise SystemExit(
        "没有找到正在运行的 WPS 文字实例（试过 %s）。\n"
        "请先在 WPS 里打开目标文档，再运行本脚本。\n最后一个错误：%s"
        % (", ".join(PROGIDS), last)
    )


def find_doc(app, name):
    for i in range(app.Documents.Count):
        d = app.Documents.Item(i + 1)
        if d.Name == name or d.FullName == name:
            return d
    return None


def active_name(app):
    try:
        return app.ActiveDocument.Name
    except Exception:                   # noqa: BLE001
        return None


def foreground():
    try:
        import win32gui
        return win32gui.GetWindowText(win32gui.GetForegroundWindow()) or "(无标题)"
    except Exception as e:              # noqa: BLE001
        return "ERR:%s" % e


def para_items(doc):
    """返回去重后的 (文本, 起始位置)，覆盖正文段落和表格单元格内的段落。"""
    seen, items = set(), []
    for i in range(1, doc.Paragraphs.Count + 1):
        r = doc.Paragraphs(i).Range
        if r.Start not in seen:
            seen.add(r.Start)
            items.append((r.Text, r.Start))
    for ti in range(1, doc.Tables.Count + 1):
        t = doc.Tables(ti)
        for ci in range(1, t.Range.Cells.Count + 1):
            cell = t.Range.Cells(ci)
            for pi in range(1, cell.Range.Paragraphs.Count + 1):
                r = cell.Range.Paragraphs(pi).Range
                if r.Start not in seen:
                    seen.add(r.Start)
                    items.append((r.Text, r.Start))
    return items


def clean(txt):
    return txt.replace("\r", "").replace("\x07", "").strip()


def replace_all(doc, old, new, cell_only=False):
    """字符定位法替换。WPS 的 Find.Execute(Replace=2) 会假成功，因此不用它。

    返回命中数量。默认从后往前替换，避免位置偏移。
    """
    hits = []
    if cell_only:
        ranges = []
        for ti in range(1, doc.Tables.Count + 1):
            t = doc.Tables(ti)
            for ci in range(1, t.Range.Cells.Count + 1):
                ranges.append(t.Range.Cells(ci).Range)
    else:
        ranges = [doc.Range(st, st + len(tx)) for tx, st in para_items(doc)]

    for rng in ranges:
        txt = rng.Text
        j = txt.find(old)
        while j >= 0:
            hits.append(rng.Start + j)
            j = txt.find(old, j + 1)

    hits = sorted(set(hits), reverse=True)
    for pos in hits:
        doc.Range(pos, pos + len(old)).Text = new
    return len(hits)


# ---------------------------------------------------------------- 命令

def cmd_list(app, args):
    a = active_name(app)
    print("已打开文档 %d 个（★ = 用户当前正在操作的）" % app.Documents.Count)
    for i in range(app.Documents.Count):
        d = app.Documents.Item(i + 1)
        print("  %s %s | 已保存=%s" % ("★" if d.Name == a else " ", d.Name, d.Saved))
    return 0


def cmd_info(app, args):
    doc = find_doc(app, args.doc)
    if doc is None:
        print("!! 未找到已打开文档：%s" % args.doc)
        return 1
    print("文档：%s" % doc.Name)
    print("路径：%s" % doc.FullName)
    print("段落数：%d | 表格数：%d" % (doc.Paragraphs.Count, doc.Tables.Count))
    print("页数：%s | 字数：%s" % (doc.ComputeStatistics(2), doc.ComputeStatistics(0)))
    print("是否在前台：%s" % (doc.Name == active_name(app)))
    print("用户前台窗口：%s" % foreground())
    return 0


def cmd_dump(app, args):
    doc = find_doc(app, args.doc)
    if doc is None:
        print("!! 未找到已打开文档：%s" % args.doc)
        return 1
    n = 0
    for idx, (txt, st) in enumerate(para_items(doc), 1):
        s = clean(txt)
        if not s:
            continue
        r = doc.Range(st, st + len(txt))
        try:
            font = "%s %s" % (r.Font.Name, r.Font.Size)
        except Exception:               # noqa: BLE001
            font = "?"
        print("[%03d] (%s) %s" % (idx, font, s))
        n += 1
        if args.limit and n >= args.limit:
            print("...(已到 --limit 上限)")
            break
    print("--- 共 %d 个非空段落 ---" % n)
    return 0


def cmd_find(app, args):
    doc = find_doc(app, args.doc)
    if doc is None:
        print("!! 未找到已打开文档：%s" % args.doc)
        return 1
    total = 0
    for idx, (txt, st) in enumerate(para_items(doc), 1):
        cnt = txt.count(args.needle)
        if cnt:
            total += cnt
            print("[%03d] 命中 %d 次 | %s" % (idx, cnt, clean(txt)[:80]))
    print("--- 共命中 %d 次 ---" % total)
    return 0


def cmd_replace(app, args):
    doc = find_doc(app, args.doc)
    if doc is None:
        print("!! 未找到已打开文档：%s" % args.doc)
        return 1

    fg_before = foreground()
    print("前台窗口(改动前)：%s" % fg_before)

    alerts_bak = app.DisplayAlerts
    app.DisplayAlerts = False
    t0 = time.time()
    try:
        n = replace_all(doc, args.old, args.new, cell_only=args.cell_only)
    finally:
        app.DisplayAlerts = alerts_bak

    print("命中并替换 %d 处（%.2fs）" % (n, time.time() - t0))
    print("前台窗口(改动后)：%s" % foreground())

    if n == 0:
        print("没有命中，未做任何修改。")
        return 1

    if not args.save:
        print("--- 预演模式，未保存。确认无误后加 --save 重新执行 ---")
        print("（注意：文档内容已变，未保存；如不想保留请自行撤销）")
        return 0

    doc.Save()
    print("已保存：%s" % doc.FullName)

    body = doc.Content.Text
    print("旧文本剩余：%d | 新文本出现：%d" % (body.count(args.old), body.count(args.new)))
    try:
        data = open(doc.FullName, "rb").read()
        print("磁盘校验：旧文本 %s | 新文本 %s" % (
            args.old.encode("utf-16-le") in data,
            args.new.encode("utf-16-le") in data))
    except Exception as e:              # noqa: BLE001
        print("磁盘校验跳过：%s" % e)
    print("页数：%s | 字数：%s" % (doc.ComputeStatistics(2), doc.ComputeStatistics(0)))
    return 0


# ---------------------------------------------------------------- 入口

def build_parser():
    p = argparse.ArgumentParser(description="挂载已打开的 WPS 文字实例，后台读写文档")
    sub = p.add_subparsers(dest="cmd", required=True)

    sub.add_parser("list", help="列出已打开文档")

    sp = sub.add_parser("info", help="文档概况")
    sp.add_argument("doc", help="文档名或完整路径")

    sp = sub.add_parser("dump", help="导出全部段落（含表格）")
    sp.add_argument("doc", help="文档名或完整路径")
    sp.add_argument("--limit", type=int, default=0, help="最多输出多少个非空段落")

    sp = sub.add_parser("find", help="定位关键词")
    sp.add_argument("doc", help="文档名或完整路径")
    sp.add_argument("needle", help="要查找的文本")

    sp = sub.add_parser("replace", help="后台替换文本")
    sp.add_argument("doc", help="文档名或完整路径")
    sp.add_argument("old", help="旧文本")
    sp.add_argument("new", help="新文本")
    sp.add_argument("--save", action="store_true", help="替换后保存（默认只预演）")
    sp.add_argument("--cell-only", action="store_true", dest="cell_only",
                    help="只处理表格单元格内的文本")

    return p


def main(argv=None):
    args = build_parser().parse_args(argv)
    app, _pid = attach()
    handler = {
        "list": cmd_list,
        "info": cmd_info,
        "dump": cmd_dump,
        "find": cmd_find,
        "replace": cmd_replace,
    }[args.cmd]
    return handler(app, args)


if __name__ == "__main__":
    sys.exit(main())
