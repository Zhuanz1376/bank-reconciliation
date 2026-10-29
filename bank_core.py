# -*- coding: utf-8 -*-
"""
================================================================
【bank_core.py —— 银行对账系统·核心对账逻辑层】
================================================================
设计原则：核心逻辑与图形界面完全解耦。
  - 本模块只做数据处理，绝不 import tkinter；
  - GUI（银行对账系统.py）只负责"收数据 → 调 core → 展示结果"；
  - 好处：核心算法可以脱离界面做单元测试（test_bank_core.py），
    日后换 Web 界面 / 命令行界面时，这一层一行都不用改。

对账口径：
  第1轮【严格匹配】结算方式 + 票号 + 借方金额 + 贷方金额 四字段全等

  第2轮【容错匹配】票号 + 金额 + 同方向（借对借 / 贷对贷）
           —— 解决"同一笔业务两边日期写法不一致"导致的漏配；
  第3轮【容错匹配】日期 + 金额 + 同方向
           —— 兜底无票号 / 票号写法不同的流水（需人工复核）。

企业方待对账 = 期初企业方未达账（结转） + 本期单位银行存款日记账
银行方待对账 = 期初银行方未达账（结转） + 本期银行对账单

余额口径：
  每笔流水滚动余额 = 上一笔余额 + 借方金额 - 贷方金额（双方同公式）
  企业期末账面余额 = 企业期初 + 本期借方合计 - 本期贷方合计
  银行期末对账单余额 = 银行期初 + 本期借方合计 - 本期贷方合计

期末余额调节表（交叉调节口径）：
  企业侧调整后 = 企业期末账面 + (银行已收企业未收) - (银行已付企业未付)
  银行侧调整后 = 银行期末对账单 + (企业已收银行未收) - (企业已付银行未付)
  两端调整后余额相等 ⇔ 对账平衡
================================================================
"""

import csv
import datetime

# ------------------------------------------------------------
# 常量
# ------------------------------------------------------------
SRC_ENTERPRISE = "企业日记账"
SRC_BANK = "银行对账单"

# 未达账项四大类别
CAT_BANK_RECV_ENT_NOT = "银行已收 企业未收"
CAT_BANK_PAY_ENT_NOT = "银行已付 企业未付"
CAT_ENT_RECV_BANK_NOT = "企业已收 银行未收"
CAT_ENT_PAY_BANK_NOT = "企业已付 银行未付"

# 匹配等级
MATCH_STRICT = "严格"   # 第1轮：四字段全等
MATCH_LOOSE = "容错"    # 第2/3轮：票号+金额 / 日期+金额 兜底

# 金额比较精度：小于 0.005 元视为相等（规避浮点误差）
EPS = 0.005


# ------------------------------------------------------------
# 工具函数
# ------------------------------------------------------------
def validate_date(s):
    """任意输入 -> 'YYYY-MM-DD'；无法解析返回 None。兼容 2025-12-08 / 2025.12.8 / 2025/12/8。"""
    if not s:
        return None
    s = str(s).strip().replace("/", "-").replace(".", "-").replace(" ", "")
    for fmt in ("%Y-%m-%d", "%y-%m-%d"):
        try:
            return datetime.datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def validate_amount(s):
    """任意输入 -> float（保留2位）；空串按 0；非数字返回 None。"""
    if s is None or s == "":
        return 0.0
    t = str(s).strip().replace(",", "").replace("，", "").replace("￥", "").replace("¥", "")
    if t == "":
        return 0.0
    try:
        return round(float(t), 2)
    except (ValueError, TypeError):
        return None


def to_float(val):
    """openpyxl 单元格 / 任意值 -> float，空值返回 0。"""
    v = validate_amount(val)
    return 0.0 if v is None else v


def same_amount(a, b):
    """金额相等判断（考虑浮点精度）。None 与 None 视为相等。"""
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return abs(a - b) < EPS


# ------------------------------------------------------------
# 数据类：一笔银行流水
# ------------------------------------------------------------
class FlowRecord:
    """一笔银行流水（企业日记账 / 银行对账单通用）。

    matched 字段含义：
      ""        —— 未匹配（未达账）
      "严格"    —— 四字段全等匹配成功，界面上显示 √
      "容错"    —— 票号+金额 或 日期+金额 兜底匹配，界面上显示 ？（建议人工复核）
    """

    def __init__(self, date, settlement_method, ticket_no,
                 debit, credit, balance=0.0, source="",
                 is_opening=False, matched="", match_key=None):
        self.date = validate_date(date) or str(date or "")
        self.settlement_method = str(settlement_method or "").strip()
        self.ticket_no = str(ticket_no or "").strip()
        self.debit = round(float(debit or 0), 2)    # 借方金额
        self.credit = round(float(credit or 0), 2)  # 贷方金额
        self.balance = round(float(balance or 0), 2)
        self.source = source                  # 企业日记账 / 银行对账单
        self.is_opening = bool(is_opening)    # 是否为期初结转的未达账
        self.matched = matched                # "" / "严格" / "容错"
        self.match_key = match_key            # 命中对方的票号（调试用）

    # ---- 派生信息 ----
    def amount(self):
        """本笔业务金额（借方或贷方，其一非零）。"""
        return self.debit if self.debit > 0 else self.credit

    def direction(self):
        """记账方向：'借' 或 '贷'。"""
        return "借" if self.debit > 0 else "贷"

    def is_income(self, source=None):
        """按来源判断是否为收入方向：企业借方=收入；银行贷方=收入。"""
        src = source or self.source
        return self.debit > 0 if src == SRC_ENTERPRISE else self.credit > 0

    def mark_label(self):
        """界面对账标记：√（严格）/ ？（容错）/ 空。"""
        if self.matched == MATCH_STRICT:
            return "√"
        if self.matched == MATCH_LOOSE:
            return "？"
        return ""

    def tuple_for_tree(self):
        """转成 ttk.Treeview 一行的 values 元组。"""
        return (
            "期初" if self.is_opening else "本期",
            self.source,
            self.date,
            self.settlement_method,
            self.ticket_no,
            "{:,.2f}".format(self.debit) if self.debit else "",
            "{:,.2f}".format(self.credit) if self.credit else "",
            "{:,.2f}".format(self.balance) if self.balance else "",
            self.mark_label(),
        )

    # ---- JSON 持久化 ----
    def to_dict(self):
        return {
            "date": self.date, "settlement_method": self.settlement_method,
            "ticket_no": self.ticket_no, "debit": self.debit, "credit": self.credit,
            "balance": self.balance, "source": self.source,
            "is_opening": self.is_opening, "matched": self.matched,
            "match_key": self.match_key,
        }

    @classmethod
    def from_dict(cls, d):
        rec = cls(
            date=d.get("date", ""), settlement_method=d.get("settlement_method", ""),
            ticket_no=d.get("ticket_no", ""), debit=d.get("debit", 0) or 0,
            credit=d.get("credit", 0) or 0, balance=d.get("balance", 0) or 0,
            source=d.get("source", ""), is_opening=d.get("is_opening", False),
        )
        rec.matched = d.get("matched", "") or ""
        rec.match_key = d.get("match_key")
        return rec

    def __repr__(self):
        return ("FlowRecord({} {} {} {} 借{} 贷{}{})".format(
            self.source, self.date, self.settlement_method, self.ticket_no,
            self.debit, self.credit, " 期初结转" if self.is_opening else ""))


# ------------------------------------------------------------
# 匹配键
# ------------------------------------------------------------
def match_key_strict(rec):
    """严格匹配键：结算方式 + 票号 + 借方金额 + 贷方金额（四字段全等）。
    """
    return (rec.settlement_method, rec.ticket_no,
            round(rec.debit, 2), round(rec.credit, 2))


# ------------------------------------------------------------
# 滚动余额
# ------------------------------------------------------------
def roll_balance(records, opening):
    """从期初余额开始按列表顺序逐笔滚动：余额 = 上一笔余额 + 借 - 贷。

    期初结转记录（is_opening=True）不参与滚动，余额置 0。
    直接写回每个 FlowRecord.balance，并返回最末余额。
    """
    bal = float(opening or 0)
    for r in records:
        if r.is_opening:
            r.balance = 0.0
            continue
        bal = round(bal + r.debit - r.credit, 2)
        r.balance = bal
    return bal


# ------------------------------------------------------------
# 核心：三轮自动对账
# ------------------------------------------------------------
def _match_round1_strict(ent_recs, bank_recs, used_ent, used_bank, pairs):
    """第1轮：严格四字段全等。用字典索引，O(n) 复杂度。"""
    index = {}
    for j, br in enumerate(bank_recs):
        index.setdefault(match_key_strict(br), []).append(j)
    for i, er in enumerate(ent_recs):
        if used_ent[i]:
            continue
        for j in index.get(match_key_strict(er), []):
            if used_bank[j]:
                continue
            _pair_up(er, bank_recs[j], MATCH_STRICT, pairs)
            used_ent[i] = True
            used_bank[j] = True
            break


def _match_round2_ticket(ent_recs, bank_recs, used_ent, used_bank, pairs):
    """第2轮（容错）：票号 + 金额 + 同方向。解决两边日期写法不一致导致的漏配。"""
    for i, er in enumerate(ent_recs):
        if used_ent[i] or not er.ticket_no:
            continue
        for j, br in enumerate(bank_recs):
            if used_bank[j] or not br.ticket_no:
                continue
            if er.ticket_no != br.ticket_no or not same_amount(er.amount(), br.amount()):
                continue
            if (er.debit > 0 and br.debit > 0) or (er.credit > 0 and br.credit > 0):
                _pair_up(er, br, MATCH_LOOSE, pairs)
                used_ent[i] = True
                used_bank[j] = True
                break


def _match_round3_date(ent_recs, bank_recs, used_ent, used_bank, pairs):
    """第3轮（容错）：日期 + 金额 + 同方向。兜底无票号 / 票号写法不同的流水。"""
    for i, er in enumerate(ent_recs):
        if used_ent[i]:
            continue
        for j, br in enumerate(bank_recs):
            if used_bank[j]:
                continue
            if er.date != br.date or not same_amount(er.amount(), br.amount()):
                continue
            if (er.debit > 0 and br.debit > 0) or (er.credit > 0 and br.credit > 0):
                _pair_up(er, br, MATCH_LOOSE, pairs)
                used_ent[i] = True
                used_bank[j] = True
                break


def _pair_up(er, br, level, pairs):
    """登记一对已达账项，并给双方打上匹配标记。"""
    er.matched = br.matched = level
    er.match_key, br.match_key = br.ticket_no, er.ticket_no
    pairs.append((er, br, level))


def categorize_unreached(ent_recs, bank_recs, used_ent, used_bank):
    """把未匹配流水归入未达账四大类。"""
    unreached = []
    for i, er in enumerate(ent_recs):
        if used_ent[i]:
            continue
        cat = CAT_ENT_RECV_BANK_NOT if er.debit > 0 else CAT_ENT_PAY_BANK_NOT
        unreached.append({"category": cat, "record": er})
    for j, br in enumerate(bank_recs):
        if used_bank[j]:
            continue
        # 银行对账单：借方=企业账户存款增加（=银行已收），贷方=减少（=银行已付）
        cat = CAT_BANK_RECV_ENT_NOT if br.debit > 0 else CAT_BANK_PAY_ENT_NOT
        unreached.append({"category": cat, "record": br})
    order = {CAT_ENT_RECV_BANK_NOT: 0, CAT_ENT_PAY_BANK_NOT: 1,
             CAT_BANK_RECV_ENT_NOT: 2, CAT_BANK_PAY_ENT_NOT: 3}
    unreached.sort(key=lambda x: (order.get(x["category"], 9), x["record"].date))
    return unreached


def run_reconciliation(records, ent_open, bank_open):
    """一键对账入口：分组 → 三轮匹配 → 未达分类 → 滚动余额 → 生成调节表。

    参数：
      records   : 全部 FlowRecord 列表（含期初结转 is_opening=True 的记录）
      ent_open  : 企业日记账期初（调整前）余额
      bank_open : 银行对账单期初（调整前）余额
    返回 result 字典（字段见函数内注释），GUI 与单元测试共用。
    """
    ent_recs = [r for r in records if r.source == SRC_ENTERPRISE]
    bank_recs = [r for r in records if r.source == SRC_BANK]
    if not ent_recs or not bank_recs:
        raise ValueError("请同时录入【企业日记账】和【银行对账单】两类数据后再对账。")

    # 重置匹配状态
    for r in records:
        r.matched = ""
        r.match_key = None

    used_ent = [False] * len(ent_recs)
    used_bank = [False] * len(bank_recs)
    pairs = []                       # [(企业FlowRecord, 银行FlowRecord, 等级)]

    _match_round1_strict(ent_recs, bank_recs, used_ent, used_bank, pairs)
    _match_round2_ticket(ent_recs, bank_recs, used_ent, used_bank, pairs)
    _match_round3_date(ent_recs, bank_recs, used_ent, used_bank, pairs)

    unreached = categorize_unreached(ent_recs, bank_recs, used_ent, used_bank)
    ent_only = [u["record"] for u in unreached if u["record"].source == SRC_ENTERPRISE]
    bank_only = [u["record"] for u in unreached if u["record"].source == SRC_BANK]

    # 滚动余额（期初结转记录跳过）
    ent_end = roll_balance(ent_recs, ent_open)
    bank_end = roll_balance(bank_recs, bank_open)

    statement = build_statement(ent_open, bank_open, ent_only, bank_only,
                                [r for r in ent_recs if not r.is_opening],
                                [r for r in bank_recs if not r.is_opening])

    return {
        "pairs": pairs,                      # 已达账项 [(企业, 银行, 等级)]
        "unreached": unreached,              # 未达账项 [{"category", "record"}]
        "ent_only": ent_only,                # 期末企业方未达账
        "bank_only": bank_only,              # 期末银行方未达账
        "ent_end": ent_end,                  # 企业日记账期末余额（滚动）
        "bank_end": bank_end,                # 银行对账单期末余额（滚动）
        "statement": statement,              # 期末余额调节表数字
        "stats": {
            "企业方总笔数": len(ent_recs), "银行方总笔数": len(bank_recs),
            "严格匹配对数": sum(1 for _, _, lv in pairs if lv == MATCH_STRICT),
            "容错匹配对数": sum(1 for _, _, lv in pairs if lv == MATCH_LOOSE),
            "企业方未达笔数": len(ent_only), "银行方未达笔数": len(bank_only),
        },
    }


# ------------------------------------------------------------
# 期末余额调节表（数字部分）
# ------------------------------------------------------------
def build_statement(ent_open, bank_open, ent_only, bank_only, ent_period, bank_period):
    """按交叉调节口径计算期末余额调节表的全部数字。

      企业期末账面   = 企业期初 + 本期企业借方合计 - 本期企业贷方合计
      银行期末对账单 = 银行期初 + 本期银行借方合计 - 本期银行贷方合计
      企业侧调整后   = 企业期末账面   + 银行已收企业未收 - 银行已付企业未付
      银行侧调整后   = 银行期末对账单 + 企业已收银行未收 - 企业已付银行未付
    （未达账合计同时包含期初结转与本期新产生的未达账。）
    """
    bank_recv = round(sum(r.debit for r in bank_only), 2)    # 银行已收 企业未收
    bank_pay = round(sum(r.credit for r in bank_only), 2)    # 银行已付 企业未付
    ent_recv = round(sum(r.debit for r in ent_only), 2)      # 企业已收 银行未收
    ent_pay = round(sum(r.credit for r in ent_only), 2)      # 企业已付 银行未付

    ent_end = round(float(ent_open or 0)
                    + sum(r.debit for r in ent_period)
                    - sum(r.credit for r in ent_period), 2)
    bank_end = round(float(bank_open or 0)
                     + sum(r.debit for r in bank_period)
                     - sum(r.credit for r in bank_period), 2)

    adj_ent = round(ent_end + bank_recv - bank_pay, 2)
    adj_bank = round(bank_end + ent_recv - ent_pay, 2)

    return {
        "企业期末余额": ent_end, "银行期末余额": bank_end,
        "企业_银行已收企业未收": bank_recv, "企业_银行已付企业未付": bank_pay,
        "银行_企业已收银行未收": ent_recv, "银行_企业已付银行未付": ent_pay,
        "企业调整后余额": adj_ent, "银行调整后余额": adj_bank,
        "平衡": abs(adj_ent - adj_bank) < EPS,
        "差异": round(adj_ent - adj_bank, 2),
    }


def statement_text(st, date_str=None):
    """把调节表数字渲染成等宽文本（GUI 里 Text 控件直接显示）。"""
    if date_str is None:
        date_str = datetime.date.today().strftime("%Y-%m-%d")
    sep = "=" * 58
    line = "-" * 58
    txt = [sep,
           "                   银 行 存 款 余 额 调 节 表",
           "                   编制日期：" + date_str,
           sep, "",
           "【企业日记账账面余额】             {:>14,.2f} 元".format(st["企业期末余额"]),
           "  加：银行已收，企业未收款           {:>14,.2f} 元".format(st["企业_银行已收企业未收"]),
           "  减：银行已付，企业未付款           {:>14,.2f} 元".format(st["企业_银行已付企业未付"]),
           line,
           "  调节后余额                        {:>14,.2f} 元".format(st["企业调整后余额"]),
           "",
           "【银行对账单余额】                 {:>14,.2f} 元".format(st["银行期末余额"]),
           "  加：企业已收，银行未收款           {:>14,.2f} 元".format(st["银行_企业已收银行未收"]),
           "  减：企业已付，银行未付款           {:>14,.2f} 元".format(st["银行_企业已付银行未付"]),
           line,
           "  调节后余额                        {:>14,.2f} 元".format(st["银行调整后余额"]),
           "", sep]
    if st["平衡"]:
        txt.append("★ 校验结果：两边调节后余额相等，对账平衡！")
    else:
        txt.append("✗ 校验结果：两边调节后余额不相等，相差 {:,.2f} 元，请核查。".format(st["差异"]))
    txt.append(sep)
    return "\n".join(txt)


# ------------------------------------------------------------
# 数据导入（CSV / openpyxl 工作表）
# ------------------------------------------------------------
def parse_flow_rows(rows, source, has_header=True):
    """把 [日期,结算方式,票号,借方,贷方(,余额)] 行列表转成 FlowRecord（余额列忽略，统一滚动计算）。"""
    out = []
    for row in rows:
        cells = [("" if c is None else str(c).strip()) for c in (row if row is not None else [])]
        if has_header and cells and any(k in cells[0] for k in ("日期", "date", "Date", "表3")):
            has_header = False
            continue
        if len(cells) < 5:
            continue
        d = validate_date(cells[0])
        debit = validate_amount(cells[3])
        credit = validate_amount(cells[4])
        if d is None or debit is None or credit is None or (debit == 0 and credit == 0):
            continue
        way, no = cells[1], cells[2]
        if any(k in way + no for k in ("合计", "小计", "总计")):
            continue
        out.append(FlowRecord(date=d, settlement_method=way, ticket_no=no,
                              debit=debit, credit=credit, source=source))
    return out


def read_flow_csv(path, source):
    """读取流水 CSV（utf-8-sig 兼容 Excel 导出的 BOM）。"""
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    return parse_flow_rows(rows, source)


def parse_opening_rows(rows):
    """解析期初余额调节表（节名格式或 Excel Sheet1 行）。

    支持两种输入：
      1) [节名] 分节的 CSV 文本行（readlines() 的结果，# 开头为注释）；
      2) Excel Sheet1 的原始行（靠关键字定位）。
    返回 dict(ent_open, bank_open, ent_unreach[], bank_unreach[])
    """
    result = {"ent_open": 0.0, "bank_open": 0.0, "ent_unreach": [], "bank_unreach": []}
    section = None
    ent_flag = bank_flag = False
    for row in rows:
        if isinstance(row, str):                      # 文本行：拆成单元格
            line = row.strip()
            if not line or line.startswith("#"):      # 空行 / 注释
                continue
            row = line.split(",")
        cells = [("" if c is None else str(c).strip()) for c in (row if row is not None else [])]
        first = cells[0] if cells else ""
        # ---- [节名] 分节格式 ----
        if first.startswith("[") and first.endswith("]"):
            section = first[1:-1]
            # 节名同时决定后面的明细行归属哪一方未达账
            ent_flag = section == "期初企业方未达账"
            bank_flag = section == "期初银行方未达账"
            continue
        if section == "企业日记账期初余额" and first and first.isdigit():
            result["ent_open"] = to_float(first)
            continue
        if section == "银行对账单期初余额" and first and first.isdigit():
            result["bank_open"] = to_float(first)
            continue
        # ---- Excel Sheet1 关键字定位格式 ----
        joined = " ".join(cells)
        if "期初企业方未达账" in joined:
            ent_flag, bank_flag, section = True, False, None
            continue
        if "期初银行方未达账" in joined:
            bank_flag, ent_flag, section = True, False, None
            continue
        if "调整前余额" in joined:
            for k, c in enumerate(cells):
                if "调整前余额" in c and k + 1 < len(cells):
                    v = to_float(cells[k + 1])
                    if v > 0:
                        if "日记账" in c or k <= 1:
                            result["ent_open"] = result["ent_open"] or v
                        else:
                            result["bank_open"] = result["bank_open"] or v
            continue
        # ---- 未达账明细行 ----
        if ent_flag or bank_flag:
            if first in ("日期", "date", ""):
                continue
            if len(cells) >= 5:
                d = validate_date(cells[0])
                debit, credit = to_float(cells[3]), to_float(cells[4])
                if d and (abs(debit) > EPS or abs(credit) > EPS):
                    item = {"date": d, "settlement_method": cells[1],
                            "ticket_no": cells[2], "debit": debit, "credit": credit}
                    (result["ent_unreach"] if ent_flag else result["bank_unreach"]).append(item)
    return result


def read_opening_csv(path):
    """读取 [节名] 分节的期初余额调节表 CSV。"""
    with open(path, "r", encoding="utf-8") as f:
        return parse_opening_rows(f.readlines())


def opening_items_to_records(opening):
    """期初未达账明细 -> FlowRecord（is_opening=True，source 由方向决定）。

    期初企业方未达账 = 企业已记、银行未记 → 记在企业方；
    期初银行方未达账 = 银行已记、企业未记 → 记在银行方。
    """
    out = []
    for it in opening.get("ent_unreach", []):
        out.append(FlowRecord(date=it["date"], settlement_method=it["settlement_method"],
                              ticket_no=it["ticket_no"], debit=it["debit"], credit=it["credit"],
                              source=SRC_ENTERPRISE, is_opening=True))
    for it in opening.get("bank_unreach", []):
        out.append(FlowRecord(date=it["date"], settlement_method=it["settlement_method"],
                              ticket_no=it["ticket_no"], debit=it["debit"], credit=it["credit"],
                              source=SRC_BANK, is_opening=True))
    return out


# ------------------------------------------------------------
# 内置示例账套（打包成 exe 后无需外部文件即可运行）
# ------------------------------------------------------------
def load_builtin_sample():
    """返回一份内置示例账套，与随仓库提供的 CSV 示例数据一致。

    返回 dict(ent_open, bank_open, records[])
    """
    ent_rows = [
        ("2025-12-01", "202", "Z1001", None, 61440), ("2025-12-01", "202", "Z2001", 163800, None),
        ("2025-12-01", "202", "Z1003", 17000, None), ("2025-12-01", "202", "Z1002", None, 20000),
        ("2025-12-01", "202", "Z1004", None, 57600), ("2025-12-03", "202", "Z1005", None, 93600),
        ("2025-12-05", "4", "H4001", None, 35200), ("2025-12-07", "202", "Z1007", None, 100),
        ("2025-12-09", "202", "Z1008", None, 6400), ("2025-12-09", "202", "Z1009", None, 100),
        ("2025-12-09", "202", "Z2003", 175500, None), ("2025-12-10", "202", "Z1011", None, 2520),
        ("2025-12-11", "202", "Z1012", None, 20000), ("2025-12-11", "202", "Z1013", None, 31650),
        ("2025-12-12", "202", "Z1015", 126000, None), ("2025-12-13", "202", "Z1014", None, 35844),
        ("2025-12-19", "202", "Z2004", 270000, None), ("2025-12-20", "202", "Z1010", None, 98520),
        ("2025-12-22", "5", "J5001", 3624.78, None), ("2025-12-23", "202", "Z2005", 8000, None),
        ("2025-12-24", "202", "Z1016", None, 100), ("2025-12-25", "202", "Z2008", 39600, None),
        ("2025-12-27", "202", "Z2009", 462000, None), ("2025-12-27", "202", "Z2002", None, 46500),
        ("2025-12-27", "202", "Z1006", None, 616590),
    ]
    bank_rows = [
        ("2025-12-01", "202", "Z1001", None, 61440), ("2025-12-02", "202", "Z1152", None, 13800),
        ("2025-12-02", "202", "Z2001", 163800, None), ("2025-12-04", "202", "Z2002", 240000, None),
        ("2025-12-04", "202", "Z1002", None, 20000), ("2025-12-05", "202", "Z1004", None, 57600),
        ("2025-12-05", "202", "Z1005", None, 93600), ("2025-12-05", "4", "H4001", None, 35200),
        ("2025-12-05", "202", "Z1007", None, 100), ("2025-12-08", "202", "Z1008", None, 6400),
        ("2025-12-08", "202", "Z1009", None, 100), ("2025-12-08", "202", "Z2003", 175500, None),
        ("2025-12-10", "202", "Z1011", None, 2520), ("2025-12-10", "202", "Z1012", None, 20000),
        ("2025-12-10", "202", "Z1013", None, 31650), ("2025-12-13", "202", "Z1010", None, 178000),
        ("2025-12-16", "202", "Z1014", None, 35844), ("2025-12-18", "202", "Z2004", 270000, None),
        ("2025-12-20", "202", "Z2006", 2808, None), ("2025-12-26", "5", "J5001", 3624.78, None),
        ("2025-12-27", "202", "Z2005", 8000, None), ("2025-12-27", "202", "Z1016", None, 100),
        ("2025-12-29", "202", "Z2008", 39600, None), ("2025-12-30", "202", "Z2009", 462000, None),
        ("2025-12-30", "201", "X1001", None, 358450), ("2025-12-30", "202", "Z1006", None, 616590),
    ]
    records = []
    for d, w, no, dr, cr in ent_rows:
        records.append(FlowRecord(date=d, settlement_method=w, ticket_no=no,
                                  debit=dr or 0, credit=cr or 0, source=SRC_ENTERPRISE))
    for d, w, no, dr, cr in bank_rows:
        records.append(FlowRecord(date=d, settlement_method=w, ticket_no=no,
                                  debit=dr or 0, credit=cr or 0, source=SRC_BANK))
    # 期初未达账：企业方一笔（企业已付 13800，银行未记 → 银行 12-02 已记 Z1152，期初配对）
    records.append(FlowRecord(date="2025-12-02", settlement_method="202", ticket_no="Z1152",
                              debit=0, credit=13800, source=SRC_ENTERPRISE, is_opening=True))
    return {"ent_open": 602349.0, "bank_open": 616149.0, "records": records}
