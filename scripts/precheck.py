# -*- coding: utf-8 -*-
"""批量预检: 不写入任何数据, 判断每份报告能否自动跑完。

输出三类清单(2026-11-02 升级):
  1. 可自动跑   —— 字段齐全、规格等级受支持, 脚本能直接编制上传
  2. 缺数据待补 —— 系统里字段本身是空的(方量/累计温度), 需人工补录后再跑
  3. 已有成果   —— downloadFlag=1 已经上传过, 跳过避免覆盖人工成果
     (脚本上传的备注A 是"总方量：xxx m³"格式; 人工做的通常是别的文字,
      据此区分并在清单里标注)

可执行:
    python precheck.py                 # 只看统计
    python precheck.py --list          # 打印三类编号清单(可直接喂给 --exclude)
    python precheck.py --csv           # 导出 CSV 明细
"""
import sys, os, csv, argparse
from collections import Counter, defaultdict
import openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dyw_client import DywClient, parse_volume, parse_temperature, format_remark_a
from report_engine import SIZE_RANGE, LOAD_RANGE

XLSX = os.getenv("DYW_LEDGER", r'C:/Users/bonbo/Desktop/未发报告.xlsx')
OUT_CSV = os.path.join(os.getcwd(), "docs", "预检清单.csv")


def col_of(ws, *names, default=None):
    """按表头名定位列(1-based)，兼容各单位导出模板列序不同。"""
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(1, c).value or "").strip()
        if h in names:
            return c
    return default


def load():
    ws = openpyxl.load_workbook(XLSX, data_only=True)['任务列表']
    c_temp = col_of(ws, "温度", default=6)
    c_hum = col_of(ws, "湿度", default=7)
    print("台账列定位: 温度=第%d列 湿度=第%d列" % (c_temp, c_hum))
    g = defaultdict(lambda: {"samples": [], "temp": None, "hum": None})
    for r in range(2, ws.max_row + 1):
        code = ws.cell(r, 1).value
        if not code: continue
        t = g[str(code).strip()]
        s = ws.cell(r, 2).value
        if s: t["samples"].append(str(s).strip())
        t["temp"] = ws.cell(r, c_temp).value
        t["hum"] = ws.cell(r, c_hum).value
    return g


def is_script_remark(text):
    """备注A 是否为脚本生成的格式(总方量：<n>m³ ...)。"""
    s = (text or "").strip()
    return s.startswith("总方量：") or s.startswith("总方量:")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--list", action="store_true", help="打印三类编号清单")
    ap.add_argument("--csv", action="store_true", help="导出 CSV 明细")
    a = ap.parse_args()

    c = DywClient(); c.login()
    tasks = load()
    ok, need, done, notready = [], [], [], []

    for i, (code, t) in enumerate(sorted(tasks.items()), 1):
        j = c.post('/ptcompile/getReportInfo.do',
                   data={'reportCode': code, 'reportId': '551'}).json()
        if not j.get('isSuccess'):
            need.append((code, "查询失败", str(j.get('responseMsg')), ""))
            continue
        d = j['data']
        ss = d.get('savedSingleField') or {}
        oi = d.get('orderInfo') or {}
        gui, grade = ss.get('gui_ge'), ss.get('chanpin_guige')
        vol = parse_volume(ss.get('qi_ta'))
        yh = (ss.get('yang_hu_tiao_jian') or '').strip()
        remark = (ss.get('bei_zhu_A') or '').strip()
        flag = d.get('downloadFlag')

        # ---- 已有成果判定：已上传过就跳过，绝不覆盖 ----
        if flag == "1":
            who = "脚本上传" if is_script_remark(remark) else "人工上传"
            done.append((code, "已上传", remark or "(空)", who))
            continue

        # ---- 状态未就绪：报告记录还没生成，编不了 ----
        # IN = 业务系统已推委托单但报告记录未初始化（2026-11-10 实测 12 份）
        #     此时加不进检测数据，保存报"当前报告状态不能编制！"
        stt = (oi.get('reportStatus') or '').strip()
        if stt and stt != 'TEST':
            notready.append((code, "状态=%s 未就绪" % stt, "", ""))
            continue

        probs = []
        if gui not in SIZE_RANGE: probs.append("规格不支持:%s" % gui)
        if grade not in LOAD_RANGE: probs.append("等级不支持:%s" % grade)
        if vol is None: probs.append("方量解析失败:%r" % (ss.get('qi_ta') or ''))
        temp = None
        if yh == '同条件养护':
            temp = parse_temperature(ss.get('lei_ji_yang_hu_wen_du'))
            if temp is None:
                probs.append("累计温度解析失败:%r" % (ss.get('lei_ji_yang_hu_wen_du')))
        # 委托单缺"试验日期"/"成型日期" -> 报告表格该行会缺失, 后面列全部错位。
        # shi_yan_ri_qi 是只读字段(来自委托单 mustAttribute), saveAllData 写不进去,
        # 所以只能人工去业务系统把委托单补上。
        if not (ss.get('shi_yan_ri_qi') or '').strip():
            probs.append("委托单缺试验日期(只读字段,需业务系统补)")
        if not (ss.get('cheng_xing_ri_qi') or '').strip():
            probs.append("委托单缺成型日期(龄期算不出)")
        if not t['samples']: probs.append("台账无样品编号")
        if t['temp'] is None: probs.append("台账无温度")
        if t['hum'] is None: probs.append("台账无湿度")

        rec = (code, len(t['samples']), gui, grade, vol, yh, temp,
               t['temp'], t['hum'], "; ".join(probs))
        if probs:
            need.append((code, "; ".join(probs), format_remark_a(vol, temp) or "", ""))
        else:
            ok.append(rec)
        if probs:
            print("%3d %s ✗ %s" % (i, code[-5:], "; ".join(probs)))

    print()
    print("=" * 62)
    print("可自动跑: %d 份 | 缺数据待补: %d 份 | 已有成果跳过: %d 份 | 状态未就绪: %d 份 | 合计 %d"
          % (len(ok), len(need), len(done), len(notready), len(tasks)))
    print("养护分布:", dict(Counter(r[5] for r in ok)))
    print("规格分布:", dict(Counter(r[2] for r in ok)))
    print("等级分布:", dict(Counter(r[3] for r in ok)))

    if notready:
        print("\n--- 状态未就绪(%d 份, 报告记录还没生成, 编不了) ---" % len(notready))
        for code, why, _, _ in notready:
            print("  %s  %s  (等业务系统推进到 TEST 后再跑)" % (code[-5:], why))

    if done:
        manual = [x for x in done if x[3] == "人工上传"]
        print("\n--- 已有成果(%d 份, 已上传跳过) ---" % len(done))
        for code, _, remark, who in done:
            print("  %s  %s  备注A=%s" % (code[-5:], who, repr(remark)[:40]))
        if manual:
            print("  ↑ 其中 %d 份是人工做的, 备注A 非脚本格式" % len(manual))

    if need:
        print("\n--- 缺数据待补(%d 份) ---" % len(need))
        for code, why, _, _ in need:
            print("  %s  %s" % (code[-5:], why))

    if a.list:
        print("\n### 可直接复制给 fix_and_run.py --exclude ###")
        print(",".join(x[0] for x in need) if need else "(无)")

    if a.csv:
        os.makedirs(os.path.dirname(OUT_CSV), exist_ok=True)
        rows = []
        for r in ok:
            rows.append({"报告编号": r[0], "分类": "可自动跑", "组数": r[1], "规格": r[2],
                         "等级": r[3], "方量": r[4] or "", "养护": r[5],
                         "累计温度": r[6] if r[6] is not None else "",
                         "台账温度": r[7], "台账湿度": r[8],
                         "备注A预览": format_remark_a(r[4], r[6]) or "", "问题": ""})
        for code, why, remark, who in need:
            rows.append({"报告编号": code, "分类": "缺数据待补", "组数": "", "规格": "",
                         "等级": "", "方量": "", "养护": "", "累计温度": "", "台账温度": "",
                         "台账湿度": "", "备注A预览": remark, "问题": why})
        for code, why, remark, who in done:
            rows.append({"报告编号": code, "分类": "已有成果跳过", "组数": "", "规格": "",
                         "等级": "", "方量": "", "养护": "", "累计温度": "", "台账温度": "",
                         "台账湿度": "", "备注A预览": remark, "问题": who})
        for code, why, remark, who in notready:
            rows.append({"报告编号": code, "分类": "状态未就绪", "组数": "", "规格": "",
                         "等级": "", "方量": "", "养护": "", "累计温度": "", "台账温度": "",
                         "台账湿度": "", "备注A预览": "", "问题": why})
        # 状态未就绪的单独导一份, 方便直接转给业务/系统同事跟进
        if notready:
            pend = os.path.join(os.getcwd(), "docs", "待跟进_未就绪.csv")
            os.makedirs(os.path.dirname(pend), exist_ok=True)
            prows = []
            for code, why, _, _ in notready:
                j2 = c.post('/ptcompile/getReportInfo.do',
                            data={'reportCode': code, 'reportId': '551'}).json()
                oi2 = (j2.get('data') or {}).get('orderInfo') or {}
                prows.append({"报告编号": code, "委托单号": oi2.get('orderNo') or "",
                              "报告状态": oi2.get('reportStatus') or "", "问题": why})
            with open(pend, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=list(prows[0].keys()))
                w.writeheader(); w.writerows(prows)
            print("待跟进清单: %s" % pend)
        with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)
        print("\n明细:", OUT_CSV)

    print("\n--- 备注A 生成样例 ---")
    for r in ok[:6]:
        print("  %s  %s" % (r[0][-5:], repr(format_remark_a(r[4], r[6]))))
    print("--- 同条件样例 ---")
    n = 0
    for r in ok:
        if r[5] == '同条件养护' and n < 4:
            n += 1
            print("  %s  %s" % (r[0][-5:], repr(format_remark_a(r[4], r[6]))))


if __name__ == "__main__":
    main()
