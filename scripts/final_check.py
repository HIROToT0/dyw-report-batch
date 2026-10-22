# -*- coding: utf-8 -*-
"""最终验收: 扫描台账全部报告, 核对状态/字段/文件, 输出验收表。"""
import sys, os, csv, datetime, re
from collections import Counter
import openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dyw_client import DywClient

XLSX = os.getenv("DYW_LEDGER", r'C:/Users/bonbo/Desktop/未发报告.xlsx')
OUT = os.getenv("DYW_OUT", os.path.join(os.getcwd(), "docs", "最终验收表.csv"))
os.makedirs(os.path.dirname(OUT), exist_ok=True)
REQ = ["gong_cheng_ming_chen_2", "jian_zheng_ren_dan_wei", "jian_zheng_ren_yuan",
       "chang_jia", "gong_cheng_bu_wei", "ling_qi"]

def col_of(ws, *names, default=None):
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(1, c).value or "").strip() in names:
            return c
    return default


ws = openpyxl.load_workbook(XLSX, data_only=True)['任务列表']
c_temp = col_of(ws, "温度", default=6)
c_hum = col_of(ws, "湿度", default=7)
tasks, order = {}, []
for r in range(2, ws.max_row + 1):
    code = ws.cell(r, 1).value
    if not code: continue
    code = str(code).strip()
    t = tasks.setdefault(code, {"n": 0, "temp": None, "hum": None})
    if code not in order: order.append(code)
    t["n"] += 1
    t["temp"], t["hum"] = ws.cell(r, c_temp).value, ws.cell(r, c_hum).value

STATUS = {"TEST": "报告待编制", "SUBMIT": "已编制待审批", "APPOK": "批准通过",
          "REJECT": "驳回", "CN": "取消", "EXPOK": "试验员通过",
          "CMPOK": "编制人通过", "CHKOK": "审核通过", "NEW": "新建"}

c = DywClient(); c.login()


def sample_key(row):
    """从检测数据行里取样品编号（模板字段名各版本不一，兜底搜值）。"""
    for k in ("yang_pin_bian_hao", "yPBH", "sampleCode", "yangpinbianhao"):
        if row.get(k):
            return str(row[k]).strip()
    for k, v in row.items():
        if isinstance(v, str) and re.match(r"^DWYHY\d+-", v):
            return v
    return ""


def group_no(s):
    m = re.search(r"-(\d+)$", str(s or "").strip())
    return int(m.group(1)) if m else None


rows = []
for i, code in enumerate(order, 1):
    j = c.post('/ptcompile/getReportInfo.do', data={'reportCode': code, 'reportId': '551'}).json()
    d = j.get('data') or {}
    ss = d.get('savedSingleField') or {}
    oi = d.get('orderInfo') or {}
    st = oi.get('reportStatus')
    miss = [k for k in REQ if ss.get(k) in (None, "")]
    brows = d.get('savedBatchFieldList') or []
    # 组号顺序校验：服务端按添加顺序生成报告，乱序会导致报告显示错乱
    gnos = [group_no(sample_key(r)) for r in brows]
    gnos = [x for x in gnos if x is not None]
    seq_ok = "√" if gnos == sorted(gnos) else "★乱序"
    rows.append({
        "报告编号": code, "组数": tasks[code]["n"],
        "状态": STATUS.get(st, st), "电子报告": d.get('downloadFlag'),
        "已上传日": (oi.get('uploadDate') or "")[:10],
        "缺失渲染字段": ",".join(miss) or "无",
        "报告日期": ss.get('bao_gao_ri_qi') or "", "试验日期": ss.get('shi_yan_ri_qi') or "",
        "温度": ss.get('wen_du') or "", "湿度": ss.get('shi_du') or "",
        "备注A": ss.get('bei_zhu_A') or "", "检测行": len(brows), "组号顺序": seq_ok,
    })
    if i % 15 == 0: print("  ...已核 %d/%d" % (i, len(order)))

with open(OUT, 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0].keys())); w.writeheader(); w.writerows(rows)

print("=" * 62)
print("报告总数 %d" % len(rows))
print("状态分布:", dict(Counter(r["状态"] for r in rows)),
      "  (注: 未走完审批流程一律显示'报告待编制', 不代表失败)")
print("电子报告已上传(downloadFlag=1): %d" % sum(1 for r in rows if r["电子报告"] == "1"))
print("渲染字段齐全: %d / %d" % (sum(1 for r in rows if r["缺失渲染字段"] == "无"), len(rows)))
print("备注A 已填: %d" % sum(1 for r in rows if r["备注A"]))
print("检测行数分布:", dict(Counter(r["检测行"] for r in rows)))
print("\n仍缺渲染字段的报告:")
for r in rows:
    if r["缺失渲染字段"] != "无":
        print("  ", r["报告编号"][-5:], r["缺失渲染字段"])
bad_seq = [r for r in rows if r["组号顺序"] != "√"]
if bad_seq:
    print("\n★ 组号显示乱序(需重新跑这 %d 份): %s" % (
        len(bad_seq), ", ".join("%s(%d组)" % (r["报告编号"][-5:], r["检测行"]) for r in bad_seq)))
    print("   修复: fix_and_run.py --only <编号>   (现在会按组号数值升序重排)")
else:
    print("\n组号顺序: 全部正确")
print("\n输出:", OUT)
