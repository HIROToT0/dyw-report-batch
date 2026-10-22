# -*- coding: utf-8 -*-
"""模板 551（混凝土试块抗压--GB/T 50081）本地公式引擎。

来源复盘（2026-09-28）：getReportTableInfo.do?id=551 返回的 callFunctionList（29 条）
里，长/宽/极限荷载这 9 个字段**本身就是公式字段**，参数之一是 yang_pin_bian_hao ——
这就是"输入样品编号回车 → 自动带出数据"的真身：样品编号变化会触发公式重算。

    chang_A/B/C, kuan_A/B/C      <- 尺寸填写(gui_ge, yang_pin_bian_hao)
        100×100×100 -> uniform(99.7, 100.4).toFixed(1)
        150×150×150 -> uniform(149.7, 150.4).toFixed(1)
    ji_xian_he_zai_yi/er/san     <- 抗压极限荷载(chanpin_guige, yang_pin_bian_hao)
        C20:235~275  C25:300~345  C30:357~410  C35:420~480  C40:480~550  C45:540~620
        -> .toFixed(3)

派生链（与页面一致，逐级读取上一级的**显示值**）：
    chi_cun_*   = "长×宽"  (各 toFixed(1))
    shou_ya_mian_ji_* = 长*宽                        round=2 acc=0
    kang_ya_*   = 1000*荷载*系数/(长*宽)              round=2 acc=1
        系数: 100 立方 -> 0.95 ; 150 立方 -> 1.0 ; 200 立方 -> 1.05
    ping_jun_zhi: 三值降序 a>=b>=c
        (a-b)<=0.15b 且 (b-c)<=0.15b   -> (a+b+c)/3
        (a-b)>0.15b  且 (b-c)>0.15b    -> "无效"
        否则                            -> b
        round=2 acc=1
    da_dao_zhi = ping_jun_zhi / int(等级去首字母) * 100   round=2 acc=0
    jian_hao_A/B/C = 组号 + "-1/-2/-3"

round=2 是 common.js 的 liuShuang()：**四舍六入五留双**（不是四舍五入）。
"""
import random
from decimal import Decimal, InvalidOperation

# ---------- 修约规则 ----------
def liu_shuang(value, acc):
    """四舍六入五留双（复刻 common.js liuShuang），返回字符串（保留 acc 位小数）。

    注意 JS 侧先做了 Number(value.toFixed(7))，这里同样处理，保证浮点尾数一致。
    """
    try:
        d = Decimal(repr(round(float(value), 7)))
    except (TypeError, ValueError, InvalidOperation):
        return str(value)
    scale = 10 ** acc
    scaled = d * scale
    floor_part = int(scaled)          # 正数，等价于向下取整
    frac = scaled - floor_part
    if frac > Decimal("0.5"):
        res = floor_part + 1
    elif frac < Decimal("0.5"):
        res = floor_part
    else:                              # 恰好 5
        res = floor_part if floor_part % 2 == 0 else floor_part + 1
    out = Decimal(res) / scale
    return ("%%.%df" % acc) % out


def fmt(value, rule, acc):
    """按 resultRoundRole / resultDataAccuracy 格式化。rule/acc 为 '99' 表示不处理。"""
    if str(rule) == "99" or str(acc) == "99":
        return value if isinstance(value, str) else str(value)
    try:
        if float(value) != float(value):     # NaN
            return str(value)
    except (TypeError, ValueError):
        return str(value)
    return liu_shuang(float(value), int(acc))


# ---------- 基础测量值生成 ----------
SIZE_RANGE = {"100×100×100": (99.7, 100.4), "150×150×150": (149.7, 150.4)}
LOAD_RANGE = {"C20": (235, 275), "C25": (300, 345), "C30": (357, 410),
              "C35": (420, 480), "C40": (480, 550), "C45": (540, 620)}


def gen_size(gui_ge, rng=random):
    lo, hi = SIZE_RANGE[gui_ge]
    return "%.1f" % rng.uniform(lo, hi)


def gen_load(grade, rng=random):
    lo, hi = LOAD_RANGE[grade]
    return "%.3f" % rng.uniform(lo, hi)


# ---------- 派生计算 ----------
def _chi_cun(chang, kuan):
    return "%.1f×%.1f" % (float(chang), float(kuan))


def _kang_ya(load, chang, kuan):
    c = float(chang)
    if 99 <= c <= 101:
        k = 0.95
    elif 149 <= c <= 151:
        k = 1.0
    elif 199 <= c <= 201:
        k = 1.05
    else:
        return None
    return fmt(1000 * float(load) * k / (c * float(kuan)), 2, 1)


def _ping_jun(a, b, c):
    a, b, c = sorted([float(a), float(b), float(c)], reverse=True)
    if (a - b) <= b * 0.15 and (b - c) <= b * 0.15:
        return fmt((a + b + c) / 3, 2, 1)
    if (a - b) > b * 0.15 and (b - c) > b * 0.15:
        return "无效"
    return fmt(b, 2, 1)


def _da_dao(ping_jun, grade):
    try:
        denom = int(str(grade)[1:])
    except ValueError:
        return "---"
    return fmt(float(ping_jun) / denom * 100, 2, 0)


def group_no(sample_code):
    """样品编号 DWYHY20260000009050-2 -> 组号 '2'"""
    return sample_code.split("-")[-1] if "-" in sample_code else "1"


# ---------- 生成一行完整检测数据 ----------
BLANK = ("he_zai_zhi__k_N__", "chang__m_m__", "kuan__m_m_",
         "shou_ya_mian_ji__m_m__", "yang_pin_chi_cun")


def make_row(sample_code, gui_ge, grade, chang=None, kuan=None, load=None, rng=random):
    """生成一行检测数据 dict（含全部派生字段）。

    chang / kuan / load 均为三元组，传入表示**已知值**（复算/验证用），
    缺省则按 gui_ge / grade 随机生成（与页面公式行为一致）。
    """
    _c = list(chang) if chang else [gen_size(gui_ge, rng) for _ in range(3)]
    _k = list(kuan) if kuan else [gen_size(gui_ge, rng) for _ in range(3)]
    _l = list(load) if load else [gen_load(grade, rng) for _ in range(3)]
    chang_a, chang_b, chang_c = _c
    kuan_a, kuan_b, kuan_c = _k
    g = group_no(sample_code)

    row = {
        "yang_pin_bian_hao": sample_code,
        "jian_hao_A": "%s-1" % g, "jian_hao_B": "%s-2" % g, "jian_hao_C": "%s-3" % g,
        "chang_A": chang_a, "chang_B": chang_b, "chang_C": chang_c,
        "kuan_A": kuan_a, "kuan_B": kuan_b, "kuan_C": kuan_c,
        "ji_xian_he_zai_yi": _l[0], "ji_xian_he_zai_er": _l[1], "ji_xian_he_zai_san": _l[2],
    }
    row["chi_cun_A"] = _chi_cun(chang_a, kuan_a)
    row["chi_cun_B"] = _chi_cun(chang_b, kuan_b)
    row["chi_cun_C"] = _chi_cun(chang_c, kuan_c)
    row["shou_ya_mian_ji_A"] = fmt(float(chang_a) * float(kuan_a), 2, 0)
    row["shou_ya_mian_ji_B"] = fmt(float(chang_b) * float(kuan_b), 2, 0)
    row["shou_ya_mian_ji_C"] = fmt(float(chang_c) * float(kuan_c), 2, 0)
    row["kang_ya_yi"] = _kang_ya(_l[0], chang_a, kuan_a)
    row["kang_ya_er"] = _kang_ya(_l[1], chang_b, kuan_b)
    row["kang_ya_san"] = _kang_ya(_l[2], chang_c, kuan_c)
    row["ping_jun_zhi"] = _ping_jun(row["kang_ya_yi"], row["kang_ya_er"], row["kang_ya_san"])
    row["da_dao_zhi"] = _da_dao(row["ping_jun_zhi"], grade)
    for b in BLANK:
        row[b] = ""
    return row


def _norm3(v):
    if isinstance(v, dict):
        return [v.get("A"), v.get("B"), v.get("C")]
    return list(v)
