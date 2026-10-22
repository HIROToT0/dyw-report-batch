---
name: dyw-report-batch
description: 检测报告系统（taikec 报告编制平台，混凝土试块抗压模板 551）的批量自动化——按 Excel 台账逐份补齐报告渲染字段、逐行添加检测数据、填备注A/温湿度，然后保存→生成→上传，回写台账进度；并支持批量下载原始记录/电子报告 PDF、合并成带书签的一册。当用户要求"按台账跑报告""批量编制/上传报告""下载原始记录""合并PDF""未发报告自动化"时使用。
agent_created: true
---

# 检测报告系统批量编制（混凝土试块抗压 · 模板 551）

一套完整跑通的人工流程自动化：**预检 → 演练核对 → 批量编制上传 → 验收**，
外加**原始记录/电子报告 PDF 批量下载与合订**。

- 站点通过 `DYW_BASE` 指定（仅 HTTP，HTTPS 握手失败）。
  该系统有多个部署环境（正式/旧站），**用哪个站点由使用者配置**，本技能不写死。
- 模板 `reportId=551` / `reportKey=t551` = 混凝土试块抗压--GB/T 50081。
  若目标环境模板 ID 不同，改脚本里的 `RID/RKEY` 与 `dyw_client` 默认参数。

```bash
export DYW_BASE=http://<你的站点>/Report   # 必填, 技能里不内置地址
python <skill>/scripts/run.py run --yes
```

## 前置

依赖：`requests`、`openpyxl`；合并 PDF 另需 `pypdf`（`python -m pip install pypdf`）。

**凭证**（**技能文件里不含任何密码**，按优先级取）：

| 优先级 | 方式 |
|---|---|
| 1 | 环境变量 `DYW_USER` / `DYW_PWD` |
| 2 | 配置文件 `~/.workbuddy/dyw_credentials_{prod\|test}.json` → `{"user":"xxx","pwd":"yyy"}`（按站点自动选：URL 含 `test.` → `_test`） |
| 3 | 代码传参 `DywClient(user=..., pwd=...)` |

会话 Cookie 存 `~/.workbuddy/.dyw_cookies_{prod\|test}.pkl`（**多站隔离，不串号**），失效自动重登。

> ⚠️ 报 **"用户无权限编制该报告！"** 时，先分清两种情况（用**不存在的编号**做对照，报错文案不同）：
>
> | 探测编号 | 返回 | 结论 |
> |---|---|---|
> | 乱填的（如 `XXXX0000000000`） | `未查询到此报告！` | 接口正常，**报告不在站上** → 业务系统未推送 |
> | 台账里的编号 | `用户无权限编制该报告！` | **报告在站上**，只是**当前账号没被授权编制** → 换账号或找管理员开权限 |
>
> 实测教训：**管理员身份 ≠ 有编制权**。某环境中管理员账号对全部报告无编制权，
> 换一个普通账号反而 129/129 全部通过。所以遇到"无权限"**先换账号试**，
> 别过早断定"报告没推送"。
> 报告系统没有"领取任务"入口，报告由业务系统推入。
> 所以换账号是有效的解法，别过早断定"没推送"。

台账路径：`export DYW_LEDGER=<xlsx 路径>`，默认 `C:/Users/bonbo/Desktop/未发报告.xlsx`。
结果写到**当前工作目录**的 `./docs` 和 `./outputs`（自动创建）。

```bash
cd <项目目录>    # 换任何目录都能跑，输出目录会自动建
```

## 端到端示例（从台账到合并 PDF）

```bash
# 0) 环境
export DYW_BASE=http://<你的站点>/Report   # 必填, 技能里不内置地址
export DYW_LEDGER="C:/path/未发报告.xlsx"

# 1) 预检（只读，不写任何数据）
python <skill>/scripts/precheck.py

# 2) 单份演练 + 下载核对（确认字段无误再批量）
python <skill>/scripts/run.py drill <某个报告编号>
python <skill>/scripts/run.py fetch <某个报告编号>

# 3) 全量编制上传（会上传，需 yes 确认）
python <skill>/scripts/fix_and_run.py

# 4) 验收（含组号顺序核对）
python <skill>/scripts/final_check.py

# 5) 下载原始记录 + 合并成一册
python <skill>/scripts/download_records.py --kind record --type pdf --out outputs/本次
python <skill>/scripts/merge_pdfs.py --ledger "$DYW_LEDGER" \
    --folder outputs/本次/原始记录_pdf --out outputs/原始记录_合并.pdf
```

参考耗时（129 份台账、171 行、含 1 份 10 组）：预检 21s、演练 4s、全量 551s、
验收 60s、下载 20s、合并 3s。

## 标准流程（**顺序不能变，少一步就出问题**）

统一入口 `run.py`：

```bash
python <skill>/scripts/run.py check                      # 1. 只读预检(不写入)
python <skill>/scripts/run.py drill DWYHY20260000009050  # 2. 单份演练(保存+生成, 不上传)
python <skill>/scripts/run.py fetch DWYHY20260000009050  # 3. 下载报告, 肉眼核对字段
python <skill>/scripts/run.py run                        # 4. 全量(需输入 yes 确认)
python <skill>/scripts/run.py verify                     # 5. 全量验收
```

> `run` 会要求手输 `yes` 才执行——上传不可逆。非交互场景用 `run --yes`。

**多批台账建议用 `DYW_OUT` 分开存结果**，否则验收表/执行结果会被下一次覆盖：

```bash
DYW_OUT=<cwd>/docs/最终验收表_<批次>.csv python <skill>/scripts/final_check.py
DYW_OUT=<cwd>/docs/批量执行结果_<批次>.csv python <skill>/scripts/fix_and_run.py
```

`run.py` **不支持 `--exclude`**；要跳过某几份（如缺方量的）直接调底层脚本：

```bash
python <skill>/scripts/fix_and_run.py --exclude DWYHY20260000009347,DWYHY20260000009348
```

也可以直接调底层脚本（参数同下）：

```bash
python <skill>/scripts/precheck.py
python <skill>/scripts/fix_and_run.py --only <编号> --no-upload   # 演练
python <skill>/scripts/fix_and_run.py                            # 全量
python <skill>/scripts/final_check.py
```

`fix_and_run.py` 参数：`--only` / `--limit N` / `--exclude 编号,编号` / `--no-upload`。
**`--no-upload`（演练）不会回写台账进度列**——否则会把已完成的报告误标成"未上传"。

对每份报告执行的动作序列：

| # | 动作 | 接口 |
|---|---|---|
| 1 | 删除已有检测数据行 | `deleteReportMultiData.do`（逐行，参数 `id`） |
| 2 | **逐行"添加"检测数据** | `addReportMultiData.do`（`multiData[字段]=值`） |
| 3 | 补齐报告渲染字段 + 备注A/日期/温湿度 | 本地计算 |
| 4 | 保存 | `saveAllData.do`（JSON body） |
| 5 | 生成报告 | `generateReport.do`（JSON body） |
| 6 | 等 `previewPdfUrl` 出现（约 2~6 秒） | 轮询 `getReportInfo.do` |
| 7 | 上传报告 | `uploadReport.do`（**表单编码**） |

## ⚠️ 三个致命坑

### 坑 1：必须"逐行添加"检测数据，不能只用 saveAllData

用 `saveAllData.do` 的 `multiDataList` 提交数据，**能读回、保存也返回成功**，
但服务端不会产出报告文件（`downloadFlag=0`、`previewPdfUrl=None`），
上传必失败：**`未获取到报告文件信息！`**

必须先用 `addReportMultiData.do` **逐行添加**（每行一个请求），
参数是 jQuery 深序列化形式 `multiData[fieldKey]=value`，然后才 `saveAllData`。

### 坑 2：必须补齐 8 个"报告渲染字段"

模板 551 的业务字段分**两组**，报告 Word/PDF 渲染用的是**第二组（可填写）**：

| 报告里显示为 | 渲染字段（必须填） | 值来源 |
|---|---|---|
| 工程名称 | `gong_cheng_ming_chen_2` | `gongcheng_mingcheng` |
| 见证单位 | `jian_zheng_ren_dan_wei` | `jianzheng_danwei` |
| 见证人 | `jian_zheng_ren_yuan` | `jianzheng_ren` |
| 砼生产厂家 | `chang_jia` | `sheng_chan_chang_jia` |
| 工程部位 | `gong_cheng_bu_wei` | `jie_gou_bu_wei` |
| 龄期(天) | `ling_qi` | 试验日期 − 成型日期 |
| 试验仪器 | `shi_yan_yi_qi` | `①②`（默认值） |
| 累计养护温度 | `lei_ji_yang_hu_wen_du` | 同条件养护才有 |

只填第一组（只读组）时，报告里这些栏位**全是空白**——这就是"生成的报告缺字段"。

### 坑 3：10 组及以上必须按组号【数值】升序添加

台账导出时 Excel 把组号当**字符串**排，10 组会排成 `1,10,2,3,4...9`。
而服务端是**按添加顺序**生成报告表格行的，不重排的话，
报告里这 10 组的**显示顺序就是错的**（用户 2026-10-21 反馈）。

`load_tasks()` 里已内置 `sort_samples()`：正则取编号末段组号 → 转 int → 升序。
运行时会打印 `组号已重排 N 份`，便于确认命中。

验收表 `final_check.py` 有 `组号顺序` 列（`√` / `★乱序`），
会读回 `savedBatchFieldList` 的实际行序逐份核对，
乱序的用 `fix_and_run.py --only <编号>` 重跑即可（会先删旧行再按新序添加）。

> 2~9 组不受影响（字符串序 == 数值序），只有跨过 10 才暴露，所以很容易漏。

## 判断成功的客观指标

| 字段 | 含义 |
|---|---|
| `downloadFlag == "1"` | **已上传电子报告**（核心指标） |
| `previewPdfUrl` 有值 | 报告文件已产出 |
| `uploadDate` / `uploadBy` | 上传时间与账号 |

`orderInfo.reportStatus` 含义：
`TEST`=报告待编制 / `SUBMIT`=已编制待审批 / `EXPOK`=试验员通过 /
`CMPOK`=编制人通过 / `CHKOK`=审核通过 / `APPOK`=批准通过 / `REJECT`=驳回。

> **用户确认（2026-09-30）：只要没走完报告审批流程，界面就一直显示"报告待编制"。**
> 也就是说 `reportStatus=TEST` 是**正常态**，不是脚本没跑成功的信号。
> 判断"这份报告编制+上传好了没"，唯一硬指标是 `downloadFlag == "1"`
> （外加报告文件能下载、内容字段齐全），**不要拿状态列去判定成败**。

## 下载报告 / 原始记录文件

```python
# 两个参数都不能猜错：
#   reportId 必须传【模板 ID 551】—— 传记录级 ID(orderInfo.reportId) 会报 "报告编制数据未保存！"
#   reportType 是【字符串】: report=电子报告 / record=原始记录
GET /ptcompile/downloadOfficialReport.do
    ?reportCode=<完整编号>&reportId=551&fileType=pdf|word&reportType=report|record
```

| reportType | 文档 | 标题（据此判断下对了没） |
|---|---|---|
| `report` | 电子报告 | 混凝土试块抗压强度试验**报告**（编号 `DWYHY…`）|
| `record` | 原始记录 | 混凝土试块抗压试验**原始记录**（编号 `JLDWYHY…`）|

- 返回体是**裸字节**（纯 PDF / 纯 DOCX），`Content-Disposition` 里带真实文件名。
- docx 解 `word/document.xml` 即可读文本核对。
- 批量下载：`python <skill>/scripts/download_records.py --codes A,B`
  （不给 `--codes` 则按 `DYW_LEDGER` 台账里的报告编号全量下；`--out` 指定根目录，
  文件落在 `<out>/原始记录_pdf/<编号>_原始记录.pdf`）。

### 多份 PDF 合成一册

```bash
# 按台账顺序合并（推荐；组号会按数值重排，避免 -10 排到 -2 前）
python <skill>/scripts/merge_pdfs.py \
  --ledger <台账xlsx> --folder outputs/2026-10-22/原始记录_pdf \
  --out outputs/原始记录_合并.pdf

# 或按编号清单
python <skill>/scripts/merge_pdfs.py --codes A,B --folder <目录> --out 合并.pdf
```

- 自动加书签（标题取编号后 5 位），输出合并份数/总页数/体积，并列出多页文件
- 页数不固定：1 组≈1 页，但 **10 组也可能只占 4 页**（每页 2~3 组），
  **不能按组数估页数**，要按实际页数分卷

> `pip` 要用 `python -m pip`（某些隔离环境没有 `Scripts/pip.exe`）。

用 `python <skill>/scripts/run.py verify` 一次核对全部：状态、downloadFlag、
渲染字段齐全度、备注A、检测行数、**组号顺序**。

### 报告字段自检（`run.py fetch`）

`downloadFlag=1` **只代表文件产出了，不代表字段填对**——必须拆开 docx 看正文。
`fetch` 会下载 docx、打印正文，并自动核对两部分：

- **上部信息区**：见证单位/见证人/工程名称/砼生产厂家/委托单位/养护条件/报告日期
  （标签行与值行交替；标签常被渲染成 `见 证 人` 带空格，故先归一化再匹配）
- **下部数据表格**：从"整行就是样品编号"的那行开始取各列
  （样品编号/工程部位/强度等级/制作日期/试验日期/龄期/试块边长）

两个坑：**页眉也有 `报告编号：DWYHY...`**，匹配样品编号必须限定"整行以 `编号-` 开头"；
**表格表头也含"工程部位/龄期"**，不能把它们当信息区字段取。

## 台账格式（任务列表 sheet）

**列位置不固定，脚本一律按「表头名」定位**，
因为不同批次导出的台账列序会变（有的把 温度/湿度 放在 F/G，有的放到 G/H，
中间还可能多出"受理人/委托日期/要求试验日期"）。

| 列（按识别） | 内容 |
|---|---|
| A 报告编号 | **完整 18 位串** `DWYHY20260000009050`（用后 5 位查不到） |
| B 样品编号 | `报告编号-N`，N 为组号；多组占多行 |
| 表头"温度" | 22 / 23 / 23.5 → `wen_du`（1 位小数） |
| 表头"湿度" | 68 / 70 / 74 → `shi_du`（整数字符串） |
| 表头"进度" | 脚本回写：已完成 / 已生成(未上传) / 失败:原因；**没有该列就追加到末列之后** |

> ⚠️ 曾经硬编码第 6/7/8 列，遇到新台账会**把"湿度"当成"进度"覆盖写掉**——务必保持表头自适应。
> <｜hy_place▁holder▁no▁813｜>回写失败（台账被 Excel / 企微占用）时，副本按台账原名存
> `outputs/<台账名>_进度已更新.xlsx`，不会覆盖其它批次的副本。

## 其它接口约定（编码方式各不相同）

| 接口 | 编码 |
|---|---|
| `signin.do` / `getReportInfo.do` / `getReportTableInfo.do` | form |
| `saveAllData.do` / `generateReport.do` | **JSON body**（form 提交会 fastjson syntax error） |
| `uploadReport.do` | **表单编码**（JSON 会报 `报告编号为空！`） |
| `addReportMultiData.do` | form + `multiData[fieldKey]=value` |
| `deleteReportMultiData.do` / `delAlCompileData.do` | form |

硬约束：
- `saveAllData` 必须剔除所有 `sys_*` 字段（毫秒戳触发 MysqlDataTruncation）
- `multiDataList` 每行必须剔除 `id`（否则 Duplicate entry 主键冲突）
- 报告日期 `bao_gao_ri_qi` = 试验日期；忽略模板默认值 `'now'`
- `wen_du` / `shi_du` 必填；标准养护也要填温度（默认 22.0）
- **`delAlCompileData.do` 会清掉单值数据**（含温度），非必要不要用

## 检测数据从哪来

长/宽/极限荷载在模板里是**公式字段**，参数含样品编号，所以界面上"输入样品编号回车就带出数据"是本地公式重算。
`report_engine.py` 复刻了这套公式（含 common.js 的**四舍六入五留双** `liuShuang()`）：

- 长/宽 = `uniform(99.7, 100.4)`（100 立方）/ `(149.7, 150.4)`（150 立方）
- 极限荷载：C20 235~275 / C25 300~345 / C30 357~410 / C35 420~480 / C40 480~550 / C45 540~620
- 派生链：尺寸 → 受压面积 → 抗压（`1000*荷载*系数/(长*宽)`，100 立方系数 0.95）→ 平均值 → 达到值

已用已发报告 08750 验证 **42/42 字段 100% 一致**。改公式后务必重跑验证。

## 备注A 规则

- `总方量：<v>m³`；方量从 `qi_ta`（其它）解析（正则覆盖 `m³/方/立方`）
- 同条件养护 → 追加 **10 个空格** + `日累计温度值达到<t>℃▪d`（`▪`=U+25AA，`d`=ASCII）
- 备注A 是单行 input，**不能换行**

## 常见失败

| 报错 | 原因 |
|---|---|
| `未获取到报告文件信息！` | 没走"逐行添加"，报告文件未产出 |
| `该报告编制数据未保存！` | 上传时 reportId 传了记录级 ID（必须传模板 551） |
| `当前报告状态不能编制！` | 该报告已上传（SUBMIT 等），台账冗余，剔除 |
| `用户无权限编制该报告！` | 报告没推送到当前站点，或当前账号无该报告编制权限（换站/新账号时最常见） |
| `未获取到报告模板信息！` | 保存时 reportId 传了记录级 ID，必须传 551 |
| `温度不能为空！` | 用了 `delAlCompileData` 清掉了单值 |
| `极限荷载一不能为空！` | 检测数据行没填 |
| `MysqlDataTruncation` | 提交了 `sys_*` 毫秒戳 |
| `Duplicate entry` | multiDataList 带了 `id` |
| `报告编号为空！` | uploadReport 用了 JSON，改表单 |
| `报告编制数据未保存！` | 下载接口 reportId 传成了记录级 ID，必须用模板 ID `551` |
| 台账保存 PermissionError | 台账被 Excel 打开，脚本会自动改写到 outputs 副本 |
| `ConnectionResetError(10054)` 登录失败 | 站点**偶发**重置，先 `GET /Report/` 探活（200 正常）后重跑即可，不是凭证问题 |
| 预检报"方量解析失败" | 系统里 `qi_ta`（其它）字段本身为空，委托时没填方量 → 没法算备注A，**该份跳过不上传**，交人工补 |

## 文件

`scripts/`：

| 文件 | 作用 |
|---|---|
| `run.py` | **一键入口**（check/drill/fetch/run/verify） |
| `fix_and_run.py` | 主流程（删行→逐行添加→补字段→保存→生成→上传） |
| `final_check.py` | 全量验收 |
| `precheck.py` | 只读预检 |
| `dyw_client.py` | 接口客户端（含凭证加载、Cookie 持久化） |
| `report_engine.py` | 本地公式引擎（长/宽/荷载及派生值） |
| `download_records.py` | 批量下载原始记录/电子报告（`--kind record\|report`） |
| `merge_pdfs.py` | 合并 PDF 成带书签的一册（`--ledger` / `--codes` / `--glob`） |

已删除 `run_batch.py`：它缺"逐行添加"这一步，会导致**假成功**，留着是误导。
