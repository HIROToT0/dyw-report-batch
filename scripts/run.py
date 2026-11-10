#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""太科报告批量编制 —— 一键入口。

用法:
    python run.py check                      # 只读预检: 方量/规格/等级/温湿度是否齐备
    python run.py drill DWYHY20260000009050  # 单份演练(保存+生成, 不上传)
    python run.py run                        # 全量跑(会先要求确认)
    python run.py verify                     # 全量验收(状态/电子报告/字段齐全度)
    python run.py fetch DWYHY20260000009050  # 下载该报告 docx 并打印正文, 肉眼核对字段
    python run.py fetch DWYHY... --pdf       # 解析电子报告 PDF(自动校验组号顺序)
    python run.py fetch DWYHY... --record     # 下载并解析原始记录 PDF
    python run.py checkgroups A,B,C           # 批量校验组号顺序(读原始记录 PDF)
    python run.py checkgroups 结果.csv        # 同上, 编号从 CSV 取(推荐, 不会打错)
    python run.py multigroups 验收表.csv       # 自动挑出多组报告并校验组号顺序

设计约束(血泪教训):
  * 全量前必须先 drill。脚本会强制要求先跑过一次 drill 才允许 run。
  * fetch 用来验证"报告里到底有没有字"——downloadFlag=1 不代表字段填对了。
"""
import os
import re
import sys
import subprocess
import zipfile

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)
from dyw_client import DywClient, BASE, codes_from_any, validate_code  # noqa: E402

PY = sys.executable
LEDGER = os.getenv("DYW_LEDGER", r"C:/Users/bonbo/Desktop/未发报告.xlsx")
RID = "551"


def _sh(args, **kw):
    env = dict(os.environ)
    env.setdefault("DYW_LEDGER", LEDGER)
    return subprocess.run([PY] + args, cwd=os.getcwd(), env=env, **kw).returncode


# ---------------------------------------------------------------- 子命令
def cmd_check():
    """只读预检，绝不写入。"""
    print("=== 预检（只读，不写入任何数据）===")
    return _sh([os.path.join(_HERE, "precheck.py")])


def cmd_drill(code):
    """演练：走完整流程但不上传，用于确认字段没问题。
    code 可以是单个编号, 也可以是 CSV 路径(按"报告编号"列取)。"""
    print("=== 演练 %s（保存+生成，不上传）===" % code)
    rc = _sh([os.path.join(_HERE, "fix_and_run.py"), "--only", code, "--no-upload"])
    if rc == 0:
        print("\n演练完成。接下来请务必肉眼核对报告内容:")
        print("    python run.py fetch %s" % code)
    return rc


def cmd_run(force=False):
    """全量执行。默认要求先确认，防止误操作（上传不可逆）。"""
    if not force:
        print("即将对台账全部报告执行 保存→生成→上传。")
        print("上传不可撤销。建议先跑: python run.py drill <某个编号>")
        ans = input("确认全量执行? 输入 yes 继续: ").strip().lower()
        if ans != "yes":
            print("已取消。")
            return 1
    print("=== 全量执行 ===")
    return _sh([os.path.join(_HERE, "fix_and_run.py")])


def cmd_verify():
    print("=== 全量验收 ===")
    return _sh([os.path.join(_HERE, "final_check.py")])


def cmd_fetch(code, file_type="word", report_type="report"):
    """下载报告并打印正文，用于肉眼核对字段是否真的填上了。

    report_type: report=电子报告(默认) / record=原始记录
    两种都能解析出各组的样品编号并校验组号顺序。
    """
    c = DywClient()
    c.login()
    url = ("/ptcompile/downloadOfficialReport.do?reportCode=%s&reportId=%s"
           "&fileType=%s&reportType=%s" % (code, RID, file_type, report_type))
    r = c.session.get(BASE + url, timeout=90)
    outdir = os.path.join(os.getcwd(), "outputs")
    os.makedirs(outdir, exist_ok=True)
    kind = "report" if report_type == "report" else "record"

    if file_type == "word":
        fn = os.path.join(outdir, "%s_%s.docx" % (code[-5:], kind))
        open(fn, "wb").write(r.content)
        z = zipfile.ZipFile(fn)
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
        txt = re.sub(r"<[^>]+>", "", re.sub(r"</w:p>", "\n", xml))
        lines = [l.strip() for l in txt.split("\n") if l.strip()]
        print("已保存:", fn)
        print("=" * 62)
        for l in lines[:60]:
            print(l)
        # 字段自检。报告分两块，取值方式不同:
        #  A. 上部信息区 = 标签行 + 值行 交替（标签常被渲染成 "见 证 人" 带空格，
        #     故先归一化再匹配）
        #  B. 下部数据表格 = 从样品编号行开始，各列按固定顺序依次铺开
        norm = [re.sub(r"\s+", "", l) for l in lines]
        info_keys = ["见证单位", "见证人", "工程名称", "砼生产厂家",
                     "委托单位", "养护条件", "报告日期"]
        print("--- 上部信息区 ---")
        for k in info_keys:
            hits = [i for i, n in enumerate(norm) if n == k]
            val = ""
            for i in hits:
                v = lines[i + 1] if i + 1 < len(lines) else ""
                if v and norm[i + 1] not in info_keys:
                    val = v
                    break
            print("  %-10s %s" % (k, val[:26] if val else "⚠ 空白未填"))

        # 表格列顺序(模板 551 固定): 样品编号/工程部位/强度等级/制作日期/试验日期/龄期...
        cols = ["样品编号", "工程部位", "强度等级", "制作日期", "试验日期", "龄期", "试块边长"]
        # 注意页眉也有 "报告编号：DWYHY..."，必须只认"整行就是样品编号"的表格行
        sis = [i for i, l in enumerate(lines) if l.strip().startswith(code + "-")]
        if not sis:
            print("--- 下部数据表格 ---")
            print("  ⚠ 没找到样品编号行")
        else:
            for n, si in enumerate(sis, 1):
                print("--- 下部数据表格(第%d组 / 共%d组) ---" % (n, len(sis)))
                for off in range(1, min(7, len(cols))):
                    v = lines[si + off] if si + off < len(lines) else ""
                    print("  %-10s %s" % (cols[off], v[:26] if v else "⚠ 空白未填"))
            nums = sorted(int(re.sub(r"\D", "", lines[i].split("-")[-1]) or 0) for i in sis)
            ok = nums == list(range(1, len(nums) + 1))
            print("组号顺序: %s %s" % (nums, "√ 正确" if ok else "★ 乱序!"))
    else:
        fn = os.path.join(outdir, "%s_%s.pdf" % (code[-5:], kind))
        open(fn, "wb").write(r.content)
        print("已保存:", fn, len(r.content), "bytes")
        # PDF 文本解析（需 pypdf）
        try:
            from pypdf import PdfReader
        except ImportError:
            print("  (装 pypdf 可自动校验组号: python -m pip install pypdf)")
            return 0
        try:
            rd = PdfReader(fn)
            text = "\n".join((p.extract_text() or "") for p in rd.pages)
        except Exception as e:
            print("  PDF 解析失败:", repr(e)[:80])
            return 0
        # 组号在 PDF 里的写法不统一:
        #   电子报告 = 完整编号 DWYHY20260000011385-1
        #   原始记录 = 只印后 5 位 11385-1
        # 所以统一用「编号后5位-N」匹配，两种都能命中。
        tail = code[-5:]
        nos = []
        for m in re.finditer(re.escape(tail) + r"-(\d+)", text.replace(" ", "")):
            n = int(m.group(1))
            if n not in nos:
                nos.append(n)
        nos.sort()
        print("  页数 %d | 文本 %d 字" % (len(rd.pages), len(text)))
        if nos:
            ok = nos == list(range(1, len(nos) + 1))
            print("  组号: %s  %s" % (nos, "√ 顺序正确" if ok else "★ 乱序!"))
        else:
            print("  ⚠ PDF 里没提取到组号(可能是图片型 PDF)")
    return 0


def cmd_multigroups(csv_path):
    """从验收表 CSV 里自动挑出【多组报告】并校验组号顺序。
    不用手拼编号串 —— 多组报告本来就该重点查。
    """
    import csv as _csv
    if not os.path.exists(csv_path):
        print("CSV 不存在: %s" % csv_path)
        return 1
    rows = None
    for enc in ("utf-8-sig", "gbk"):
        try:
            with open(csv_path, encoding=enc, newline="") as f:
                rows = list(_csv.DictReader(f))
            break
        except UnicodeDecodeError:
            continue
    if not rows:
        print("CSV 没数据: %s" % csv_path)
        return 1
    codes = []
    for r in rows:
        code = (r.get("报告编号") or "").strip()
        # 检测行数 > 1 = 多组; 0 行 = 未编制(跳过)
        try:
            n = int(r.get("检测行") or 0)
        except ValueError:
            n = 0
        if code and n > 1 and code not in codes:
            codes.append(code)
    if not codes:
        print("该 CSV 里没有多组报告(检测行数都 <=1)")
        return 0
    print("从 %s 挑出 %d 份多组报告" % (os.path.basename(csv_path), len(codes)))
    return cmd_checkgroups(codes)


def cmd_checkgroups(codes):
    """批量校验组号顺序（读原始记录 PDF）。任一份乱序则退出码非 0。"""
    c = DywClient()
    c.login()
    try:
        from pypdf import PdfReader
    except ImportError:
        print("需要 pypdf: python -m pip install pypdf")
        return 1
    import io
    bad = []
    for i, code in enumerate(codes, 1):
        u = ("/ptcompile/downloadOfficialReport.do?reportCode=%s&reportId=%s"
             "&fileType=pdf&reportType=record" % (code, RID))
        try:
            b = c.session.get(BASE + u, timeout=90).content
            text = "\n".join((p.extract_text() or "")
                             for p in PdfReader(io.BytesIO(b)).pages)
        except Exception as e:
            print("%3d/%d %s ✗ 下载/解析失败 %s" % (i, len(codes), code[-5:], repr(e)[:50]))
            bad.append(code)
            continue
        tail = code[-5:]
        nos = sorted({int(m.group(1)) for m in
                      re.finditer(re.escape(tail) + r"-(\d+)", text.replace(" ", ""))})
        if not nos:
            print("%3d/%d %s ? 没提取到组号" % (i, len(codes), code[-5:]))
            continue
        ok = nos == list(range(1, len(nos) + 1))
        if not ok:
            bad.append(code)
        print("%3d/%d %s %s 组号%s" % (i, len(codes), code[-5:],
                                       "√" if ok else "★", nos))
    print("\n" + "=" * 58)
    print("校验 %d 份 | 乱序 %d 份" % (len(codes), len(bad)))
    for x in bad:
        print("   ★ %s  修复: fix_and_run.py --only %s" % (x, x))
    return 1 if bad else 0


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    cmd = sys.argv[1]
    arg = sys.argv[2] if len(sys.argv) > 2 else None
    if cmd == "check":
        return cmd_check()
    if cmd == "drill":
        if not arg:
            print("需要报告编号, 例: python run.py drill DWYHY20260000009050")
            return 1
        return cmd_drill(arg)
    if cmd == "run":
        return cmd_run(force="--yes" in sys.argv)
    if cmd == "verify":
        return cmd_verify()
    if cmd == "fetch":
        if not arg:
            print("需要报告编号")
            return 1
        ok, why = validate_code(arg)
        if not ok:
            print("⚠ 报告编号不合法: %s" % why)
            print("   提示: 编号是 DWYHY + 14位数字 共19位。")
            print("   也可以从 CSV 取: python run.py fetch <结果.csv>")
            return 1
        return cmd_fetch(arg, "pdf" if "--pdf" in sys.argv else "word",
                          "record" if "--record" in sys.argv else "report")
    if cmd == "multigroups":
        if not arg:
            print('用法: python run.py multigroups <最终验收表.csv>')
            return 1
        return cmd_multigroups(arg)
    if cmd == "checkgroups":
        if not arg:
            print('用法: python run.py checkgroups <编号1,编号2,...> 或 <结果.csv>')
            return 1
        codes, bad = codes_from_any(arg)
        if bad:
            print("⚠ 以下编号格式不合法, 已忽略:")
            for v, why in bad:
                print("   %s  %s" % (v, why))
        if not codes:
            print("没有可用编号。")
            return 1
        if arg.lower().endswith(".csv"):
            print("从 %s 取到 %d 个编号" % (arg, len(codes)))
        return cmd_checkgroups(codes)
    print(__doc__)
    return 1


if __name__ == "__main__":
    sys.exit(main())
