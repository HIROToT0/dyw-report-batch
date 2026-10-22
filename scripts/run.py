#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""太科报告批量编制 —— 一键入口。

用法:
    python run.py check                      # 只读预检: 方量/规格/等级/温湿度是否齐备
    python run.py drill DWYHY20260000009050  # 单份演练(保存+生成, 不上传)
    python run.py run                        # 全量跑(会先要求确认)
    python run.py verify                     # 全量验收(状态/电子报告/字段齐全度)
    python run.py fetch DWYHY20260000009050  # 下载该报告 docx 并打印正文, 肉眼核对字段

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
from dyw_client import DywClient, BASE  # noqa: E402

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
    """单份演练：走完整流程但不上传，用于确认字段没问题。"""
    print("=== 单份演练 %s（保存+生成，不上传）===" % code)
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


def cmd_fetch(code, file_type="word"):
    """下载报告并打印正文，用于肉眼核对字段是否真的填上了。"""
    c = DywClient()
    c.login()
    url = ("/ptcompile/downloadOfficialReport.do?reportCode=%s&reportId=%s"
           "&fileType=%s&reportType=report" % (code, RID, file_type))
    r = c.session.get(BASE + url, timeout=90)
    outdir = os.path.join(os.getcwd(), "outputs")
    os.makedirs(outdir, exist_ok=True)

    if file_type == "word":
        fn = os.path.join(outdir, "%s_report.docx" % code[-5:])
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
        si = next((i for i, l in enumerate(lines) if l.strip().startswith(code + "-")), None)
        print("--- 下部数据表格(第1组) ---")
        if si is None:
            print("  ⚠ 没找到样品编号行")
        else:
            for off in range(1, min(7, len(cols))):
                v = lines[si + off] if si + off < len(lines) else ""
                print("  %-10s %s" % (cols[off], v[:26] if v else "⚠ 空白未填"))
    else:
        fn = os.path.join(outdir, "%s_report.pdf" % code[-5:])
        open(fn, "wb").write(r.content)
        print("已保存:", fn, len(r.content), "bytes")
    return 0


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
        return cmd_fetch(arg, "pdf" if "--pdf" in sys.argv else "word")
    print(__doc__)
    return 1


if __name__ == "__main__":
    sys.exit(main())
