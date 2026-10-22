# -*- coding: utf-8 -*-
"""按台账批量编制报告 —— 完整人工流程版。

修正了此前遗漏的关键步骤（2026-09-28 晚）：
  1. **必须逐行"添加"检测数据**（addReportMultiData.do）。
     仅用 saveAllData 传 multiDataList，数据能读回，但服务端不会生成报告文件
     （downloadFlag=0、previewPdfUrl=None），上传必失败"未获取到报告文件信息"。
  2. **必须补齐 8 个"报告渲染用"字段**。模板把业务字段分两组：
     只读组（从订单带出）gongcheng_mingcheng / jianzheng_danwei / jianzheng_ren /
     sheng_chan_chang_jia / jie_gou_bu_wei ；
     可填写组（报告实际渲染）gong_cheng_ming_chen_2 / jian_zheng_ren_dan_wei /
     jian_zheng_ren_yuan / chang_jia / gong_cheng_bu_wei / ling_qi / shi_yan_yi_qi。
     缺了后者，报告里"工程名称/见证单位/见证人/生产厂家/工程部位/龄期"全是空白。

流程: 删旧行 -> 逐行添加 -> 保存 -> 生成 -> 上传

用法:
    python fix_and_run.py --only DWYHY20260000009050
    python fix_and_run.py --limit 3
    python fix_and_run.py                  # 全量
"""
import sys, os, csv, time, argparse, re, datetime
from collections import defaultdict
import openpyxl

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)
from dyw_client import DywClient, BASE, parse_volume, parse_temperature, format_remark_a
from report_engine import make_row, SIZE_RANGE, LOAD_RANGE

XLSX = os.getenv("DYW_LEDGER", r'C:/Users/bonbo/Desktop/未发报告.xlsx')
_OUTDIR = os.path.join(os.getcwd(), "docs")
OUT = os.getenv("DYW_OUT", os.path.join(_OUTDIR, "批量执行结果_完整流程.csv"))
RID, RKEY = "551", "t551"
for _d in (_OUTDIR, os.path.join(os.getcwd(), "outputs")):
    os.makedirs(_d, exist_ok=True)

# 只读字段 -> 报告渲染字段 的映射（值一致，来自人工报告交叉验证）
FIELD_MAP = {
    "gong_cheng_ming_chen_2": "gongcheng_mingcheng",   # 工程名称二 <- 工程名称
    "jian_zheng_ren_dan_wei": "jianzheng_danwei",      # 见证人单位 <- 见证单位
    "jian_zheng_ren_yuan":    "jianzheng_ren",         # 见证人员   <- 见证人
    "chang_jia":              "sheng_chan_chang_jia",  # 厂家       <- 生产厂家
    "gong_cheng_bu_wei":      "jie_gou_bu_wei",        # 工程部位   <- 结构部位
}


def col_of(ws, *names, default=None):
    """按表头名定位列(1-based)，兼容各单位导出模板列序不同。"""
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(1, c).value or "").strip()
        if h in names:
            return c
    return default


def group_no(sample):
    """取样品编号的组号(末段数字)。取不到返回 None。"""
    s = str(sample or "").strip()
    m = re.search(r"-(\d+)$", s)
    return int(m.group(1)) if m else None


def sort_samples(samples):
    """按组号【数值】升序排。

    必须排！台账导出时 Excel 把组号当字符串处理，10 组会排成
    1,10,2,3,4...9 —— 服务端按添加顺序生成报告表格，
    不排的话报告里这 10 组显示顺序就是乱的（使用方 2026-10-21 反馈）。
    组号取不到数字的排最后，保持相对次序。
    """
    idx = list(range(len(samples)))
    keyed = [(group_no(s) is None, group_no(s) if group_no(s) is not None else 0, i)
             for i, s in enumerate(samples)]
    keyed.sort()
    return [samples[i] for _, _, i in keyed]


def load_tasks():
    ws = openpyxl.load_workbook(XLSX, data_only=True)['任务列表']
    c_temp = col_of(ws, "温度", default=6)
    c_hum = col_of(ws, "湿度", default=7)
    print("台账列定位: 温度=第%d列 湿度=第%d列" % (c_temp, c_hum))
    g, order = {}, []
    reordered = []
    for r in range(2, ws.max_row + 1):
        code = ws.cell(r, 1).value
        if not code:
            continue
        code = str(code).strip()
        t = g.setdefault(code, {"samples": [], "temp": None, "hum": None, "rows": []})
        if code not in order:
            order.append(code)
        s = ws.cell(r, 2).value
        if s:
            t["samples"].append(str(s).strip())
        t["temp"], t["hum"] = ws.cell(r, c_temp).value, ws.cell(r, c_hum).value
        t["rows"].append(r)
    # 组号按数值排序（10 组及以上必排，见 sort_samples 说明）
    for code, t in g.items():
        if len(t["samples"]) > 1:
            fixed = sort_samples(t["samples"])
            if fixed != t["samples"]:
                reordered.append(code)
                t["samples"] = fixed
    if reordered:
        print("组号已重排 %d 份(10 组及以上台账乱序): %s" % (
            len(reordered), ", ".join(x[-5:] for x in reordered)))
    return g, order


def must_attr(oi, key):
    """从 orderInfo.mustAttribute 里取 键&^*值 的片段。"""
    ma = oi.get("mustAttribute") or ""
    m = re.search(re.escape(key) + r"&\^\*(.*?)(?:~\$#|$)", ma)
    return (m.group(1) or "").strip() if m else ""


def days_between(d1, d2):
    try:
        return str((datetime.date.fromisoformat(str(d2)[:10]) -
                    datetime.date.fromisoformat(str(d1)[:10])).days)
    except Exception:
        return ""


def fmt_temp(v):
    try:
        return "%.1f" % float(v)
    except (TypeError, ValueError):
        return None


def fmt_hum(v):
    try:
        f = float(v)
        return str(int(f)) if f == int(f) else str(f)
    except (TypeError, ValueError):
        return None


def process(c, code, t, do_upload=True):
    res = {"报告编号": code, "组数": len(t["samples"]), "备注A": "", "温度": "", "湿度": "",
           "补齐字段": "", "添加": "", "保存": "", "生成": "", "上传": "", "状态": "", "说明": ""}

    j = c.post('/ptcompile/getReportInfo.do', data={'reportCode': code, 'reportId': RID}).json()
    if not j.get('isSuccess'):
        res["说明"] = "查询失败:" + str(j.get('responseMsg'))
        return False, res
    d = j['data']
    ss, oi = d['savedSingleField'], d['orderInfo']
    gui_ge, grade = ss.get('gui_ge'), ss.get('chanpin_guige')
    if gui_ge not in SIZE_RANGE or grade not in LOAD_RANGE:
        res["说明"] = "规格/等级不支持: %s/%s" % (gui_ge, grade)
        return False, res

    # ---- 1. 删掉旧检测数据行 ----
    for row in (d.get('savedBatchFieldList') or []):
        if row.get('id'):
            try:
                c.delete_batch_row(code, row['id'], report_id=RID, report_key=RKEY)
            except Exception:
                pass

    # ---- 2. 逐行添加检测数据 ----
    samples = t["samples"] or ["%s-%d" % (code, i) for i in range(1, len(samples) + 1)]
    added = 0
    for s in samples:
        r = c.add_batch_row(code, make_row(s, gui_ge, grade), report_id=RID, report_key=RKEY)
        if isinstance(r, dict) and r.get('isSuccess'):
            added += 1
    res["添加"] = "%d/%d" % (added, len(samples))
    if added != len(samples):
        res["说明"] += " 添加数据不完整;"

    # ---- 3. 补齐报告渲染字段 ----
    filled = []
    for dst, src in FIELD_MAP.items():
        if ss.get(dst) in (None, ""):
            v = ss.get(src)
            if v in (None, ""):
                # 回落: 从必备属性取
                fallback = {"gong_cheng_bu_wei": "结构部位",
                            "chang_jia": "生产厂家"}
                if dst in fallback:
                    v = must_attr(oi, fallback[dst])
            if v not in (None, ""):
                ss[dst] = v
                filled.append(dst)
    if ss.get("ling_qi") in (None, ""):
        lq = days_between(ss.get("cheng_xing_ri_qi"), ss.get("shi_yan_ri_qi"))
        if lq:
            ss["ling_qi"] = lq
            filled.append("ling_qi")
    if ss.get("shi_yan_yi_qi") in (None, ""):
        ss["shi_yan_yi_qi"] = "\u2460\u2461"      # ①②
        filled.append("shi_yan_yi_qi")
    res["补齐字段"] = ",".join(filled)

    # ---- 4. 单值: 备注A / 报告日期 / 温度 / 湿度 ----
    vol = parse_volume(ss.get('qi_ta'))
    ta = parse_temperature(ss.get('lei_ji_yang_hu_wen_du')) \
        if (ss.get('yang_hu_tiao_jian') or '').strip() == '同条件养护' else None
    if vol is not None:
        ss['bei_zhu_A'] = format_remark_a(vol, ta)
    else:
        res["说明"] += " 方量缺失(备注A未填);"
    res["备注A"] = ss.get('bei_zhu_A') or ''
    ss['bao_gao_ri_qi'] = ss.get('shi_yan_ri_qi') or ss.get('bao_gao_ri_qi')
    tv, hv = fmt_temp(t["temp"]), fmt_hum(t["hum"])
    if tv:
        ss['wen_du'] = tv
    if hv:
        ss['shi_du'] = hv
    res["温度"], res["湿度"] = ss.get('wen_du'), ss.get('shi_du')

    # ---- 5. 保存 ----
    d2 = c.get_full_report(code)
    ss = {k: v for k, v in ss.items() if not k.startswith('sys_')}
    ml = [{k: v for k, v in x.items() if not k.startswith('sys_') and k != 'id'}
          for x in (d2.get('savedBatchFieldList') or [])]
    r = c.session.post(BASE + "/ptcompile/saveAllData.do", timeout=30,
                       json={"reportId": RID, "reportKey": RKEY, "reportCode": code,
                             "singleData": ss, "multiDataList": ml})
    try:
        rj = r.json()
    except Exception:
        rj = {"isSuccess": False, "responseMsg": r.text[:120]}
    res["保存"] = "✓" if rj.get('isSuccess') else "✗"
    if not rj.get('isSuccess'):
        res["说明"] += " 保存失败:" + str(rj.get('responseMsg'))[:60]
        return False, res

    # ---- 6. 生成 ----
    r = c.session.post(BASE + "/ptcompile/generateReport.do", timeout=90,
                       json={"reportId": RID, "reportKey": RKEY, "reportCode": code})
    try:
        rj = r.json()
    except Exception:
        rj = {"isSuccess": False, "responseMsg": r.text[:120]}
    res["生成"] = "✓" if rj.get('isSuccess') else "✗"
    if not rj.get('isSuccess'):
        res["说明"] += " 生成失败:" + str(rj.get('responseMsg'))[:60]
        return False, res

    # 等报告文件产出（downloadFlag=1 且 previewPdfUrl 有值）
    ok_file = False
    for _ in range(6):
        time.sleep(2)
        jj = c.post('/ptcompile/getReportInfo.do',
                    data={'reportCode': code, 'reportId': RID}).json().get('data') or {}
        if (jj.get('orderInfo') or {}).get('previewPdfUrl'):
            ok_file = True
            break
    if not ok_file:
        res["说明"] += " 报告文件未产出;"

    # ---- 7. 上传 ----
    if not do_upload:
        res["上传"] = "跳过"
        res["状态"] = "已生成(未上传)"
        return True, res
    r = c.session.post(BASE + "/ptcompile/uploadReport.do", timeout=90,
                       data={"reportCode": code, "reportId": RID})
    try:
        rj = r.json()
    except Exception:
        rj = {"isSuccess": False, "responseMsg": r.text[:120]}
    res["上传"] = "✓" if rj.get('isSuccess') else "✗"
    if not rj.get('isSuccess'):
        res["说明"] += " 上传失败:" + str(rj.get('responseMsg'))[:60]
        return False, res

    time.sleep(1)
    fin = c.get_full_report(code)
    res["状态"] = fin['orderInfo'].get('reportStatus')
    return True, res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--only'); ap.add_argument('--limit', type=int, default=0)
    ap.add_argument('--exclude', default=''); ap.add_argument('--no-upload', action='store_true')
    a = ap.parse_args()

    tasks, order = load_tasks()
    codes = [a.only] if a.only else order
    excl = set(x.strip() for x in a.exclude.split(',') if x.strip())
    codes = [x for x in codes if x not in excl]
    if a.limit:
        codes = codes[:a.limit]

    # --only 指定的编号可能不在台账里(如临时补跑某一份)：
    # 台账没有就现场构造一条空任务(样品编号按 编号-1 生成)，而不是直接 KeyError。
    for code in codes:
        if code not in tasks:
            print("⚠ %s 不在台账中，按样品编号 %s-1 现场构造任务" % (code[-5:], code))
            tasks[code] = {"samples": ["%s-1" % code], "temp": None, "hum": None,
                           "rows": []}

    c = DywClient(); c.login(); c.get_template_info(551)
    results, t0 = [], time.time()
    for i, code in enumerate(codes, 1):
        try:
            ok, res = process(c, code, tasks[code], do_upload=not a.no_upload)
        except Exception as e:
            ok, res = False, {"报告编号": code, "组数": len(tasks[code]["samples"]), "备注A": "",
                              "温度": "", "湿度": "", "补齐字段": "", "添加": "", "保存": "✗",
                              "生成": "✗", "上传": "✗", "状态": "", "说明": "异常:" + repr(e)[:120]}
        results.append(res)
        print("[%2d/%d] %s 组%s 补%s 加%s 存%s 生%s 传%s %s" % (
            i, len(codes), code[-5:], res["组数"], res["补齐字段"] or "-", res["添加"],
            res["保存"], res["生成"], res["上传"], res["说明"]))

    with open(OUT, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=list(results[0].keys()))
        w.writeheader(); w.writerows(results)

    # 回写台账进度。演练(--no-upload)不回写: 一份报告可能先前已上传成功,
    # 演练又跑一遍会把它错误地改标成"已生成(未上传)"。
    if a.no_upload:
        print("演练模式, 不回写台账进度列")
    else:
        wb = openpyxl.load_workbook(XLSX)
        ws = wb['任务列表']
        # 进度列: 有"进度"表头就复用, 否则追加到最后一列之后(不能硬编码第8列,
        # 新台账第8列是湿度, 会覆盖业务数据)
        c_prog = col_of(ws, "进度", "备注进度", default=ws.max_column + 1)
        if c_prog > ws.max_column or not str(ws.cell(1, c_prog).value or "").strip():
            ws.cell(1, c_prog).value = "进度"
        print("回写进度列: 第%d列" % c_prog)
        for res in results:
            mark = "已完成" if res["上传"] == "✓" else (
                "已生成(未上传)" if res["上传"] == "跳过" else "失败:" + res["说明"][:40])
            for r in tasks[res["报告编号"]]["rows"]:
                ws.cell(r, c_prog).value = mark
        try:
            wb.save(XLSX)
            print("台账进度列已回写")
        except PermissionError:
            # 副本按台账原名命名, 避免多批台账互相覆盖同名输出
            stem = os.path.splitext(os.path.basename(XLSX))[0]
            alt = os.path.join(os.getcwd(), "outputs", "%s_进度已更新.xlsx" % stem)
            os.makedirs(os.path.dirname(alt), exist_ok=True)
            wb.save(alt)
            print("⚠ 台账被占用(可能正被 Excel 打开), 进度已写入 %s" % alt)

    print("\n" + "=" * 62)
    print("耗时 %.1fs | %d 份" % (time.time() - t0, len(results)))
    print("  上传成功: %d   未上传: %d   失败: %d" % (
        sum(1 for r in results if r["上传"] == "✓"),
        sum(1 for r in results if r["上传"] == "跳过"),
        sum(1 for r in results if r["上传"] == "✗")))
    print("结果: %s" % OUT)
    for r in results:
        if r["上传"] == "✗":
            print("  ✗ %s %s" % (r["报告编号"], r["说明"]))


if __name__ == '__main__':
    main()
