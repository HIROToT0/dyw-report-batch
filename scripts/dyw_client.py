# -*- coding: utf-8 -*-
"""
太科技术有限公司 - 报告系统客户端
站点: 通过环境变量 DYW_BASE 指定, 例如
      DYW_BASE=http://<你的站点>/Report
      该系统有多个部署环境, 本技能不写死具体地址。
      注意: **不同环境/不同账号对同一批报告的编制权限可能不同**,
            遇到"用户无权限编制该报告"先换账号试(见 SKILL.md 排错表)。
登录接口: POST <BASE>/ptsysuser/signin.do  参数 userId / password
会话标识: Cookie ptkeyName = pt_<host>_<companyId>_<siteId>_<userId>

用法:
    python dyw_client.py                 # 登录 + 会话自检
    python dyw_client.py --force         # 强制重新登录(忽略已存会话)
    DYW_BASE=http://<你的站点>/Report python dyw_client.py   # 指定站点

    from dyw_client import DywClient
    c = DywClient()
    c.login()                            # 会话失效时自动重登
    c.get_report_info("报告编号")

已探明接口 (2026-09-28 实测):
    POST <BASE>/ptsysuser/signin.do                登录
    POST <BASE>/ptcompile/getLastCompileReport.do  最近编制的报告ID
    POST <BASE>/ptcompile/getReportInfo.do         报告主信息 (reportCode 必填, 可加 reportId)
    POST <BASE>/ptcompile/getReportTableInfo.do    报告表格/检测数据
    GET  <BASE>/ptcompile/getCompile.do            编制页
    GET  <BASE>/ptcompile/toLookImage.do?id=       查看原始记录图片
    POST <BASE>/ptcompile/{add,update,delete}ReportMultiData.do  检测数据增删改

重要机制:
    1. getCompile.do 的 URL 参数会被**服务端忽略** —— 带 reportId / reportCode 与不带,
       返回的 HTML 完全相同(实测 len 均为 42951)。页面是空壳, 数据在浏览器端由 JS 拉取。
       所以**直接抓 HTML 拿不到任何报告内容**, 必须走 getReportInfo.do。
    2. 真实参数名是 reportCode(报告编号, 扫任务单得到), 不是 reportId。
       校验顺序: reportCode 空 -> "未选择报告编号!" ; 缺模板 -> "未选择报告模板!" ;
       两个都传才进入查询 -> 查不到则返回 "未查询到此报告!"
    3. 未登录时业务接口返回 302 重定向(非 JSON), 可据此判定会话失效。
"""
import os
import re
import sys
import pickle
import argparse
import requests

# 站点地址**不写死在本技能里**（避免泄露内网地址）：必须通过环境变量指定。
#   export DYW_BASE=http://<你的站点>/Report
# 也可把该域名加进 hosts 文件指向实际 IP，然后直接用域名。
BASE = (os.getenv("DYW_BASE") or "http://REPLACE_ME/Report").rstrip("/")
# 凭证文件按域名是否含 test. 自动分 prod/test
LOGIN_URL = BASE + "/ptsysuser/signin.do"
CHECK_URL = BASE + "/ptcompile/getLastCompileReport.do"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
# 会话与凭证一律存在用户目录，绝不写进技能目录（技能可能被同步/分享出去）
_HOME = os.path.expanduser("~")
_WB = os.path.join(_HOME, ".workbuddy")
# Cookie 文件按站点隔离: 测试站/正式站各存一份, 避免串号
_SITE = "test" if "test." in BASE else "prod"
COOKIE_FILE = os.path.join(_WB, ".dyw_cookies_%s.pkl" % _SITE)
CRED_FILE = os.path.join(_WB, "dyw_credentials_%s.json" % _SITE)


def load_credentials(user=None, pwd=None):
    """按优先级取凭证: 显式传参 > 环境变量 > 外部配置文件 > 交互输入。

    配置文件格式 (~/.workbuddy/dyw_credentials_prod.json 或 _test.json):
        {"user": "你的账号", "pwd": "你的密码"}
    文件名后缀按站点自动选(prod/test), 不含密码的技能目录。
    """
    if user and pwd:
        return user, pwd
    env_u, env_p = os.getenv("DYW_USER"), os.getenv("DYW_PWD")
    if env_u and env_p:
        return env_u, env_p
    if os.path.exists(CRED_FILE):
        try:
            import json
            with open(CRED_FILE, encoding="utf-8") as f:
                c = json.load(f)
            if c.get("user") and c.get("pwd"):
                return c["user"], c["pwd"]
        except Exception:
            pass
    if user or pwd:  # 只传了一个，另一个补环境变量
        return user or env_u, pwd or env_p
    raise RuntimeError(
        "未配置太科系统账号。任选一种:\n"
        "  1) 环境变量: export DYW_USER=xxx; export DYW_PWD=yyy\n"
        "  2) 配置文件: %s 内容 {\"user\":\"xxx\",\"pwd\":\"yyy\"}\n"
        "  3) 代码传参: DywClient(user='xxx', pwd='yyy')" % CRED_FILE
    )


class DywClient:
    _template_defaults_cache = {}  # report_id -> list of field dicts

    def __init__(self, user=None, pwd=None, cookie_file=COOKIE_FILE):
        self.user, self.pwd = load_credentials(user, pwd)
        self.cookie_file = cookie_file
        self.user_info = None
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": UA,
            "X-Requested-With": "XMLHttpRequest",
            "Referer": BASE + "/",
        })
        self._load_cookies()

    # ---------- 会话持久化 ----------
    def _load_cookies(self):
        if os.path.exists(self.cookie_file):
            try:
                with open(self.cookie_file, "rb") as f:
                    self.session.cookies.update(pickle.load(f))
            except Exception:
                pass  # 损坏则忽略，走重新登录

    def _save_cookies(self):
        with open(self.cookie_file, "wb") as f:
            pickle.dump(self.session.cookies, f)

    # ---------- 核心 ----------
    def check(self):
        """会话是否仍有效(不触发登录)"""
        try:
            r = self.session.post(CHECK_URL, timeout=20)
            return bool(r.json().get("isSuccess"))
        except Exception:
            return False

    def login(self, force=False):
        """登录。force=True 时忽略已有会话。返回用户信息 dict。"""
        if not force and self.check():
            return self.user_info or {"_reused": True}

        r = self.session.post(LOGIN_URL,
                              data={"userId": self.user, "password": self.pwd},
                              timeout=25)
        r.raise_for_status()
        j = r.json()
        if not j.get("isSuccess"):
            raise RuntimeError("登录失败: %s" % j.get("responseMsg"))
        self.user_info = j.get("data") or {}
        self._save_cookies()
        return self.user_info

    # ---------- 便捷请求 ----------
    def post(self, path, data=None, **kw):
        return self.session.post(BASE + path, data=data or {}, timeout=30, **kw)

    def get(self, path, **kw):
        return self.session.get(BASE + path, timeout=30, **kw)

    # ---------- 报告编制业务 ----------
    def get_last_report_id(self):
        """最近编制的报告ID"""
        j = self.post("/ptcompile/getLastCompileReport.do").json()
        return j.get("data") if j.get("isSuccess") else None

    def get_report_info(self, report_code, report_id=""):
        """报告主信息。report_code=报告编号(扫任务单得到, 必填); report_id=模板ID(如 551)。

        两者都传才会真正执行查询(common.js 的 searchPl 就是这么传的)。
        常见 responseMsg:
            未选择报告编号!  -> reportCode 为空
            未选择报告模板!  -> 只传了 reportCode, 未传 reportId
            未查询到此报告!  -> 参数齐全但该报告编号不存在
        """
        return self.post("/ptcompile/getReportInfo.do",
                         data={"reportCode": report_code, "reportId": report_id}).json()

    def get_template_info(self, template_id):
        """加载模板结构(字段/公式/表格配置)。template_id 就是 URL 上的 reportId, 如 551。

        实测: POST /ptcompile/getReportTableInfo.do {'id': '551'} -> 约 97KB 模板定义
        副作用: 把字段定义缓存到 _template_defaults_cache,供 save_full_report 套默认值用。
        """
        j = self.post("/ptcompile/getReportTableInfo.do",
                      data={"id": template_id}).json()
        if j.get("isSuccess"):
            d = j.get("data") or {}
            self._template_defaults_cache[template_id] = d.get("customFieldList") or []
        return j

    def query_templates(self, page_size=100, **kw):
        """查询模板列表(模板设计模块)。

        注意: 该接口**只接受 JSON body**, 用表单编码会返回 HTTP 415。
        返回字段含 id / reportName / reportKey / status / reportCount / departmentName 等。
        kw 可传 reportName / status / departmentId 等搜索条件。
        """
        payload = {"searchObj": kw, "pageSize": page_size}
        return self.session.post(BASE + "/ptdesign/queryReportTable.do",
                                 json=payload, timeout=30).json()

    def open_compile_page(self, report_code=None, report_id=None):
        """打开编制页。注意: 服务端忽略这些参数, 返回的永远是同一个空壳页面,
        报告内容需通过 get_report_info() 获取。
        """
        q = []
        if report_code:
            q.append("reportCode=%s" % report_code)
        if report_id:
            q.append("reportId=%s" % report_id)
        url = "/ptcompile/getCompile.do" + ("?" + "&".join(q) if q else "")
        return self.get(url)

    # ---------- 高层 API (解析+写入辅助) ----------
    def get_full_report(self, report_code, report_id="551"):
        """取报告的完整原始 JSON 数据(包含 orderInfo / savedSingleField / savedBatchFieldList)。
        返回原始 dict。
        """
        j = self.post("/ptcompile/getReportInfo.do",
                      data={"reportCode": report_code, "reportId": report_id}).json()
        return j.get("data") or {} if j.get("isSuccess") else {}

    def get_status(self, report_code, report_id="551"):
        """报告当前状态(用于演练/监控):reportStatus + 关键字段"""
        d = self.get_full_report(report_code, report_id)
        if not d:
            return None
        ss = d.get("savedSingleField") or {}
        return {
            "reportStatus": (d.get("orderInfo") or {}).get("reportStatus"),
            "bei_zhu_A": ss.get("bei_zhu_A"),
            "wen_du": ss.get("wen_du"),
            "qi_ta": ss.get("qi_ta"),
            "batch_rows": len(d.get("savedBatchFieldList") or []),
        }

    def download_informal(self, report_code, report_id="551"):
        """下载非正式版报告(用于演练后验证)。
        返回 Response 对象;若 404 说明文件还没生成。"""
        return self.session.get(
            BASE + "/ptcompile/downloadInformalReport.do?reportCode=%s&reportId=%s" % (report_code, report_id),
            timeout=30)

    def save_full_report(self, data, report_id="551", report_key="t551"):
        """把一份完整数据写回(单值+检测数据)。
        data 需含 savedSingleField(单值dict) 与 savedBatchFieldList(检测数据列表, 可为空)。
        实现要点(从实际演练中得来):
          1. 整体 JSON body 提交(非 form)
          2. 剔除所有 sys_* 字段(毫秒戳会触发 MysqlDataTruncation)
          3. 默认值补齐(仅数值字段): wen_du 等若为 None 套用模板 fieldDefaultValue
          4. 报告日期 = 试验日期(使用方规则);**忽略模板里的 "now" 默认值**
          5. multiDataList 每行也要剔除 sys_* 和 id(主键冲突)
        返回 (status, body) 元组; body 是 dict 或 str(出错时)。
        """
        ss = dict(data["savedSingleField"])
        ss = {k: v for k, v in ss.items() if not k.startswith("sys_")}

        # 使用方规则:报告日期 = 试验日期(若报告日期为空,取试验日期)
        if not ss.get("bao_gao_ri_qi") or ss.get("bao_gao_ri_qi") == "now":
            ss["bao_gao_ri_qi"] = ss.get("shi_yan_ri_qi") or ""

        # 套用模板默认值(仅数值字段,跳过日期类含 'now' 等占位符)
        for f in self._template_defaults_cache.get(report_id, []):
            v = ss.get(f["fieldKey"])
            default = f.get("fieldDefaultValue")
            if v in (None, "") and default and default != "now":
                ss[f["fieldKey"]] = default

        # 检测数据: 剔除 sys_* 和 id
        ml = []
        for row in (data.get("savedBatchFieldList") or []):
            ml.append({k: v for k, v in row.items()
                       if not k.startswith("sys_") and k != "id"})

        payload = {
            "reportId": report_id,
            "reportKey": report_key,
            "reportCode": ss["baogao_bianhao"],
            "singleData": ss,
            "multiDataList": ml,
        }
        r = self.session.post(BASE + "/ptcompile/saveAllData.do",
                              json=payload, timeout=30)
        try:
            return True, r.json()
        except Exception:
            return False, r.text[:300]

    def generate_report(self, report_code, report_id="551", report_key="t551"):
        """生成报告(只需要三个参数, 数据须先存好)。
        实现要点: JSON body(dict 直传)"""
        r = self.session.post(BASE + "/ptcompile/generateReport.do",
                              json={"reportId": report_id, "reportKey": report_key,
                                    "reportCode": report_code}, timeout=30)
        try:
            return True, r.json()
        except Exception:
            return False, r.text[:300]

    def upload_report(self, report_code, report_id="551"):
        """上传报告(注意: 会推送到业务系统, 谨慎调用)。

        实测(2026-09-28): 必须**表单编码**。layTable.js:2628 reportUpload() 走的是
        enablePublicNo() -> $.ajax 默认 x-www-form-urlencoded；用 JSON body 会返回
        {"responseMsg":"报告编号为空！"}。
        """
        r = self.session.post(BASE + "/ptcompile/uploadReport.do",
                              data={"reportCode": report_code, "reportId": report_id},
                              timeout=30)
        try:
            return True, r.json()
        except Exception:
            return False, r.text[:300]

    def add_batch_row(self, report_code, multi_data, report_id="551", report_key="t551"):
        """追加一行检测数据(对应弹窗里的"添加"按钮)。

        实测(2026-09-28): 表单编码 + **jQuery 深序列化**格式才会被接受 ——
        参数名必须写成 multiData[<fieldKey>], 否则:
            - multiData=<json字符串>  -> HTTP 400 (Tomcat 解析失败)
            - 整体 JSON body          -> {"responseMsg":"请传入模板key！"}
        multi_data 传 dict, 本方法自动展开。返回服务端 JSON。
        """
        form = {"reportKey": report_key, "reportCode": report_code}
        for k, v in (multi_data or {}).items():
            form["multiData[%s]" % k] = "" if v is None else v
        return self.post("/ptcompile/addReportMultiData.do", data=form).json()

    def delete_batch_row(self, report_code, row_id, report_id="551", report_key="t551"):
        """删除一行检测数据(对应表格里的"删除")。row_id 取 batch 行的 id 字段。"""
        return self.post("/ptcompile/deleteReportMultiData.do",
                         data={"reportKey": report_key,
                               "reportCode": report_code,
                               "id": row_id}).json()

    def fill_remark_a(self, data, volume=None, temperature=None, override_existing=False):
        """从 savedSingleField['qi_ta'] 识别方量, 同时根据 'yang_hu_tiao_jian'
        是否为'同条件养护'决定是否附加累计温度信息, 写入 savedSingleField['bei_zhu_A']。

        规则(来自使用方, 2026-09-28):
          - 养护条件 = '同条件养护' -> 备注A = '总方量：<m³>' + 10空格 + '日累计温度值达到<t>℃▪d'
            t 取自 'lei_ji_yang_hu_wen_du' (累计养护温度) 字段的数值(剥掉单位)
          - 其它养护条件 -> 备注A = '总方量：<m³>' (无累计温度)
          - 字符精确匹配已发报告样本(U+25AA ▪)

        volume/temperature 显式传入时优先使用, 否则自动从 qi_ta / lei_ji_... 解析。
        override_existing=False 时, 若备注A已有内容则跳过(避免覆盖人工填的累计温度)。
        返回是否做了修改。"""
        ss = data.get("savedSingleField") or {}
        existing = (ss.get("bei_zhu_A") or "").strip()
        if existing and not override_existing:
            return False
        v = volume if volume is not None else parse_volume(ss.get("qi_ta"))
        if v is None:
            return False
        # 是否需要累计温度
        yh = (ss.get("yang_hu_tiao_jian") or "").strip()
        t = None
        if yh == "\u540c\u6761\u4ef6\u517b\u62a4":  # 同条件养护
            t = temperature if temperature is not None else parse_temperature(ss.get("lei_ji_yang_hu_wen_du"))
        ss["bei_zhu_A"] = format_remark_a(v, t)
        return True


# ---------- 方量识别工具函数 ----------
# 单位按优先级匹配: m³ → 方 → 立方
_VOL_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(?:m\u00b3|\u65b9|\u7acb\u65b9)")  # 数字 + m³/方/立方
# 累计温度数值提取: 可能是 "875℃" 或纯数字 "641" 或小数 "710.5"
_TEMP_RE = re.compile(r"(\d+(?:\.\d+)?)")
_NUM_RE = re.compile(r"(\d+(?:\.\d+)?)")

# 已发报告样本(2026-09-28 验证 3 份): 总方量:<num>m³ + 10空格 + 日累计温度值达到<t>℃▪d
# 字符精确序列: U+2103 摄氏度 + U+25AA 小黑方块 + U+0064 ASCII 'd'
TEMPERATURE_SUFFIX_FMT = "\u603b\u65b9\u91cf\uff1a{n}m\u00b3" + (" " * 10) + "\u65e5\u7d2f\u8ba1\u6e29\u5ea6\u503c\u8fbe\u5230{t}\u2103\u25aa\u0064"


def parse_volume(text):
    """从"其它"字段解析方量(立方米数)。
    已覆盖: 26m³ / 2.5m³ / 470m³ / 105方 / 1.5立方 / 95立方 / 代表数量470m³ / 共浇筑：7m³
    返回 None 表示无法解析(如 '共浇筑：' 后面无数字)。"""
    if text is None:
        return None
    s = str(text)
    m = _VOL_RE.search(s)
    if m:
        return float(m.group(1))
    m = _NUM_RE.search(s)
    return float(m.group(1)) if m else None


def parse_temperature(text):
    """从累计养护温度字段提取数值。
    已覆盖: '875℃' / '710.5' / '641' / 空(返回 None)。"""
    if text is None:
        return None
    s = str(text).strip()
    if not s or s in ("/", "-"):
        return None
    m = _TEMP_RE.search(s)
    return float(m.group(1)) if m else None


def format_remark_a(volume, temperature=None):
    """构造备注A.
    - 仅方量: '总方量：<num>m³' (整数不带 .0, 小数保留原样)
    - 方量 + 累计温度: '总方量：<num>m³' + 10空格 + '日累计温度值达到<t>℃▪d'
      字符格式按已发报告样本(08750 / 08268 / 08370 / 08700 验证)。
    返回 None 表示 volume 也是 None。"""
    if volume is None:
        return None
    v = int(volume) if float(volume) == int(volume) else volume
    head = "\u603b\u65b9\u91cf\uff1a%sm\u00b3" % v
    if temperature is not None:
        t = int(temperature) if float(temperature) == int(temperature) else temperature
        return head + (" " * 10) + "\u65e5\u7d2f\u8ba1\u6e29\u5ea6\u503c\u8fbe\u5230%s\u2103\u25aa\u0064" % t
    return head


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true", help="强制重新登录")
    args = ap.parse_args()

    c = DywClient()
    reused = (not args.force) and c.check()
    info = c.login(force=args.force)

    print("=" * 46)
    print("站点      :", BASE)
    print("会话状态  :", "有效(复用已存Cookie)" if reused else "有效(本次新登录)")
    if reused:
        print("账号      :", c.user, "(详情需 --force 重新登录获取)")
    else:
        u = info.get("kfsysUser", {})
        print("账号      :", info.get("kfuserId"))
        print("姓名      :", info.get("kfuserName"))
        print("手机      :", u.get("mobile"))
        print("companyId :", u.get("companyId"), " siteId:", u.get("siteId"))
        print("是否管理员:", u.get("isSupervisor"))
        print("菜单权限  :", u.get("menuCodes"))
    ck = {x.name: x.value for x in c.session.cookies}
    print("会话Cookie:", ck.get("ptkeyName"))
    print("=" * 46)
    return 0


if __name__ == "__main__":
    sys.exit(main())
