# -*- coding: utf-8 -*-
"""把多份 PDF 合成一册（按给定顺序，带书签）。

书签标题默认取文件名里的编号，方便在阅读器里点书签定位。

用法:
    # 1) 按台账顺序合并原始记录（推荐）
    python merge_pdfs.py --ledger <台账xlsx> --folder <单份PDF目录> --out 合并.pdf

    # 2) 按指定编号清单合并
    python merge_pdfs.py --codes DWYHY...001,DWYHY...002 --folder <目录> --out 合并.pdf

    # 3) 直接合并目录下所有 pdf（按文件名排序）
    python merge_pdfs.py --folder <目录> --glob "*.pdf" --out 合并.pdf

参数:
    --title-len  书签标题取文件名末尾几位(默认 5, 即编号后5位)
    --no-bookmark  不加书签
    --pattern     单份文件名模板, 默认 "<编号>_原始记录.pdf";
                   --kind report 时用 "<编号>_电子报告.pdf"

排序说明: 组号后缀 '-10' 在字符串排序里会排到 '-2' 前面,
所以 --codes 走台账时会用 fix_and_run.sort_samples 同样的规则重排组号,
避免合出来的册子组号乱序。见 sort_samples。
"""
import argparse
import glob
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def load_codes_from_ledger(path):
    """从台账 xlsx 的『任务列表』sheet 读报告编号(A列), 去重保序。"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    name = "任务列表" if "任务列表" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[name]
    out, seen = [], set()
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if not v:
            continue
        c = str(v).strip()
        if c and c not in seen:
            seen.add(c)
            out.append(c)
    return out


def group_no(sample):
    m = re.search(r"-(\d+)$", str(sample or "").strip())
    return int(m.group(1)) if m else None


def sort_samples(samples):
    """按组号【数值】升序（'10' 不能排在 '2' 前面）。"""
    keyed = [(group_no(s) is None, group_no(s) or 0, i)
             for i, s in enumerate(samples)]
    keyed.sort()
    return [samples[i] for _, _, i in keyed]


def bookmark_title(path, n=5):
    """书签标题：优先取文件名里的报告编号(去扩展名)，否则退回末 n 位。

    文件名形如 DWYHY20260000010813_原始记录.pdf -> 标题 DWYHY20260000010813
    （比只取末5位好用：既能定位，又看得出是哪份）
    """
    stem = os.path.splitext(os.path.basename(path))[0]
    m = re.search(r"[A-Za-z]{2,}\d{6,}", stem)
    if m:
        return m.group(0)
    return stem[-n:] if n else stem


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ledger", default=os.getenv("DYW_LEDGER", ""),
                    help="台账 xlsx；给了就按 A 列顺序取编号")
    ap.add_argument("--codes", default="", help="逗号分隔编号，与 --ledger 二选一")
    ap.add_argument("--folder", required=True, help="单份 PDF 所在目录")
    ap.add_argument("--out", required=True, help="输出 PDF 路径")
    ap.add_argument("--pattern", default="", help='文件名模板，默认 "<编号>_原始记录.pdf"')
    ap.add_argument("--kind", default="record", choices=["record", "report"],
                    help="record=原始记录(默认) / report=电子报告")
    ap.add_argument("--glob", default="", help="不用编号，直接匹配目录下 pdf")
    ap.add_argument("--title-len", type=int, default=5, help="书签标题取文件名末几位")
    ap.add_argument("--no-bookmark", action="store_true")
    a = ap.parse_args()

    try:
        from pypdf import PdfWriter, PdfReader
    except ImportError:
        print("缺少 pypdf，请先安装: python -m pip install pypdf")
        return 1

    tpl = a.pattern or ("%s_原始记录.pdf" if a.kind == "record" else "%s_电子报告.pdf")

    files = []
    if a.glob:
        files = sorted(glob.glob(os.path.join(a.folder, a.glob)))
    else:
        if a.ledger and os.path.exists(a.ledger):
            codes = load_codes_from_ledger(a.ledger)
        elif a.codes:
            codes = [x.strip() for x in a.codes.replace("\n", ",").split(",") if x.strip()]
        else:
            print("请给 --ledger 或 --codes（或用 --glob 匹配目录）")
            return 1
        # 组号重排：同编号下若有多组，保证 -1..-10 的数值序
        from collections import defaultdict
        g = defaultdict(list)
        for c in codes:
            g[c.rsplit("-", 1)[0] if "-" in c else c].append(c)
        files = []
        for _, group in g.items():
            for c in sort_samples(group):
                p = os.path.join(a.folder, tpl % c)
                files.append(p)

    missing = [p for p in files if not os.path.exists(p)]
    files = [p for p in files if os.path.exists(p)]
    if not files:
        print("没找到可合并的 PDF（目录: %s）" % a.folder)
        return 1

    w = PdfWriter()
    pages = 0
    multi = []
    for p in files:
        n = len(PdfReader(p).pages)
        if n > 1:
            multi.append((os.path.basename(p), n))
        w.append(p)
        if not a.no_bookmark:
            w.add_outline_item(bookmark_title(p, a.title_len), pages)
        pages += n

    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    w.write(a.out)

    print("=" * 58)
    print("合并 %d 份 / 总 %d 页 / %.1f MB" % (
        len(files), pages, os.path.getsize(a.out) / 1048576))
    print("书签: %s" % ("关闭" if a.no_bookmark else "%d 个" % len(files)))
    if missing:
        print("缺失 %d 个:" % len(missing))
        for p in missing[:10]:
            print("   ", os.path.basename(p))
    if multi:
        print("多页文件: %s" % ", ".join("%s(%d页)" % (b, n) for b, n in multi))
    print("输出:", a.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())