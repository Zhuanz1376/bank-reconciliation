# -*- coding: utf-8 -*-
"""
================================================================
【出纳管理之银行对账系统】 Bank Reconciliation System
================================================================
严格按照《企业会计基础工作规范》与用户提供的模板 Excel 设计，实现：
  1. 期初余额调节表 加载与 未达账项 自动结转；
  2. 企业银行存款日记账 / 银行对账单 逐笔录入或 Excel 导入；
  3. 自动对账（以「票号 + 金额」为主键，日期为辅助），区分已达 / 未达；
  4. 自动生成与模板 Sheet1/Sheet4 一致的《银行存款余额调节表》；
  5. 结果导出为 .xlsx（多 Sheet）或 .csv / .txt。

依赖：openpyxl（Excel 读写） + tkinter（GUI，Python 内置）。
运行：python3 银行对账系统.py
================================================================
"""

import csv
import os
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================
# 常量
# ============================================================
SRC_ENTERPRISE = "企业日记账"
SRC_BANK = "银行对账单"

# 未达账项四大类别
CAT_BANK_RECV_ENT_NOT = "银行已收 企业未收"
CAT_BANK_PAY_ENT_NOT = "银行已付 企业未付"
CAT_ENT_RECV_BANK_NOT = "企业已收 银行未收"
CAT_ENT_PAY_BANK_NOT = "企业已付 银行未付"


# ============================================================
# 工具函数
# ============================================================
def validate_date(s):
    if not s:
        return None
    s = str(s).strip().replace("/", "-").replace(".", "-").replace(" ", "")
    try:
        d = datetime.datetime.strptime(s, "%Y-%m-%d")
    except ValueError:
        try:
            d = datetime.datetime.strptime(s, "%y-%m-%d")
        except ValueError:
            return None
    return d.strftime("%Y-%m-%d")


def validate_amount(s):
    if s is None or s == "":
        return 0.0
    try:
        v = float(s)
    except (ValueError, TypeError):
        return None
    return round(v, 2)


def to_float(val):
    """openpyxl 单元格 → float，空值返回 0"""
    if val is None or val == "":
        return 0.0
    try:
        return round(float(val), 2)
    except (ValueError, TypeError):
        return 0.0


# ============================================================
# 数据类 —— 对齐模板字段
# ============================================================
class FlowRecord:
    """一笔银行流水"""

    def __init__(self, date, settlement_method, ticket_no,
                 debit, credit, balance=0.0, source="",
                 is_opening=False, matched=False, match_key=None):
        self.date = date                      # YYYY-MM-DD
        self.settlement_method = str(settlement_method or "").strip()
        self.ticket_no = str(ticket_no or "").strip()
        self.debit = round(float(debit or 0), 2)   # 借方金额（企业=收入；银行=支出）
        self.credit = round(float(credit or 0), 2)  # 贷方金额（企业=支出；银行=收入）
        self.balance = round(float(balance or 0), 2)
        self.source = source                  # 企业日记账 / 银行对账单
        self.is_opening = is_opening          # 是否为期初结转过来的未达账
        self.matched = matched                # 是否已对上
        self.match_key = match_key            # 对账时的票号键

    def amount(self):
        """本笔业务金额（借方或贷方，其一非零）"""
        return self.debit if self.debit > 0 else self.credit

    def direction(self):
        """企业侧方向：'收入' 或 '支出'"""
        return "收入" if self.debit > 0 else "支出"

    def is_income(self, source=None):
        """按来源判断是否为收入方向的金额"""
        src = source or self.source
        if src == SRC_ENTERPRISE:
            return self.debit > 0
        else:
            return self.credit > 0

    def tuple_for_tree(self):
        return (
            "期初" if self.is_opening else "本期",
            self.source,
            self.date,
            self.settlement_method,
            self.ticket_no,
            "{:.2f}".format(self.debit) if self.debit else "",
            "{:.2f}".format(self.credit) if self.credit else "",
            "{:.2f}".format(self.balance) if self.balance else "",
            "✓" if self.matched else "",
        )


# ============================================================
# 银行对账系统 GUI
# ============================================================
class BankReconciliationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("出纳管理之银行对账系统  （企业会计基础工作规范）")
        self.root.geometry("1380x840")
        self.root.minsize(1200, 720)

        # ---------- 业务数据 ----------
        self.opening_ent_balance = tk.StringVar(value="602349.00")
        self.opening_bank_balance = tk.StringVar(value="616149.00")

        self.records = []        # 全部流水（含期初结转）：FlowRecord
        self.opening_unreach_ent = []   # 期初企业方未达账明细
        self.opening_unreach_bank = []  # 期初银行方未达账明细

        self.reconciled_pairs = []  # [(FlowRecord(企业), FlowRecord(银行))]
        self.unreached = []         # [{"category": str, "record": FlowRecord}]

        self._build_ui()
        self._refresh_trees()

    # --------------------------------------------------------
    # UI 构建
    # --------------------------------------------------------
    def _build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        # ========== 顶部：余额区 ==========
        top = ttk.LabelFrame(
            self.root,
            text="① 余额调节起点（期初余额可载入模板；期末余额可从 期初+本期发生额 自动计算）",
            padding=8)
        top.pack(fill="x", padx=8, pady=(8, 4))

        ttk.Label(top, text="企业日记账余额：").grid(row=0, column=0, sticky="e")
        ttk.Entry(top, textvariable=self.opening_ent_balance, width=14).grid(row=0, column=1, padx=2)
        ttk.Label(top, text="元").grid(row=0, column=2, padx=(0, 18))

        ttk.Label(top, text="银行对账单余额：").grid(row=0, column=3, sticky="e")
        ttk.Entry(top, textvariable=self.opening_bank_balance, width=14).grid(row=0, column=4, padx=2)
        ttk.Label(top, text="元").grid(row=0, column=5, padx=(0, 18))

        ttk.Button(top, text="从模板 Excel 载入期初",
                   command=self.load_opening_excel).grid(row=0, column=6, padx=4)
        ttk.Button(top, text="自动推导期末余额",
                   command=self._compute_ending_balances).grid(row=0, column=7, padx=4)
        ttk.Button(top, text="载入示例数据",
                   command=self.load_sample_data).grid(row=0, column=8, padx=4)

        # ========== 中部：流水录入 ==========
        mid = ttk.LabelFrame(self.root, text="② 本期流水录入 / 导入（企业日记账 与 银行对账单）", padding=8)
        mid.pack(fill="x", padx=8, pady=4)

        self.src_var = tk.StringVar(value=SRC_ENTERPRISE)
        ttk.Label(mid, text="数据来源：").grid(row=0, column=0, sticky="e")
        ttk.Radiobutton(mid, text=SRC_ENTERPRISE, variable=self.src_var,
                        value=SRC_ENTERPRISE).grid(row=0, column=1, padx=4)
        ttk.Radiobutton(mid, text=SRC_BANK, variable=self.src_var,
                        value=SRC_BANK).grid(row=0, column=2, padx=4)

        ttk.Label(mid, text="日期：").grid(row=0, column=3, sticky="e")
        self.ent_date = ttk.Entry(mid, width=10)
        self.ent_date.insert(0, datetime.date.today().strftime("%Y-%m-%d"))
        self.ent_date.grid(row=0, column=4, padx=2)

        ttk.Label(mid, text="结算方式：").grid(row=0, column=5, sticky="e")
        self.ent_method = ttk.Entry(mid, width=8)
        self.ent_method.grid(row=0, column=6, padx=2)

        ttk.Label(mid, text="票号：").grid(row=0, column=7, sticky="e")
        self.ent_ticket = ttk.Entry(mid, width=10)
        self.ent_ticket.grid(row=0, column=8, padx=2)

        ttk.Label(mid, text="借方金额：").grid(row=1, column=0, sticky="e", pady=(4, 0))
        self.ent_debit = ttk.Entry(mid, width=11)
        self.ent_debit.grid(row=1, column=1, padx=2, pady=(4, 0))

        ttk.Label(mid, text="贷方金额：").grid(row=1, column=2, sticky="e", pady=(4, 0))
        self.ent_credit = ttk.Entry(mid, width=11)
        self.ent_credit.grid(row=1, column=3, padx=2, pady=(4, 0))

        ttk.Label(mid, text="期末余额：").grid(row=1, column=4, sticky="e", pady=(4, 0))
        self.ent_balance = ttk.Entry(mid, width=11)
        self.ent_balance.grid(row=1, column=5, padx=2, pady=(4, 0))

        ttk.Button(mid, text="添加一笔", command=self.add_record).grid(row=1, column=6, padx=6, pady=(4, 0))
        ttk.Button(mid, text="导入 Excel", command=self.import_excel).grid(row=1, column=7, padx=4, pady=(4, 0))
        ttk.Button(mid, text="导入 CSV/TXT", command=self.import_csv).grid(row=1, column=8, padx=4, pady=(4, 0))
        ttk.Button(mid, text="删除选中行", command=self.delete_selected).grid(row=1, column=9, padx=4, pady=(4, 0))

        hint = ("说明：企业日记账——借方=存入(收入)、贷方=取出(支出)；"
                "银行对账单——借方=存入(收入)、贷方=取出(支出)。"
                "两边借贷位置相同即为同一笔业务。")
        ttk.Label(mid, text=hint, foreground="#666").grid(
            row=2, column=0, columnspan=10, sticky="w", pady=(6, 0))

        # ========== 主体 PanedWindow：上=流水；下左=未达账 下右=余额调节表 ==========
        body = ttk.PanedWindow(self.root, orient=tk.VERTICAL)
        body.pack(fill="both", expand=True, padx=8, pady=4)

        # --- 上：全部流水 ---
        top_body = ttk.LabelFrame(body, text="③ 全部流水清单（含期初结转 + 本期录入，共 {} 笔）".format(len(self.records)), padding=4)
        body.add(top_body, weight=3)

        cols = ("phase", "source", "date", "settlement_method", "ticket_no",
                "debit", "credit", "balance", "matched")
        self.flow_tree = ttk.Treeview(top_body, columns=cols, show="headings", height=10)
        headings = {
            "phase": "期间", "source": "数据来源", "date": "日期",
            "settlement_method": "结算方式", "ticket_no": "票号",
            "debit": "借方金额", "credit": "贷方金额", "balance": "期末余额",
            "matched": "已达",
        }
        widths = {"phase": 60, "source": 110, "date": 110, "settlement_method": 90,
                  "ticket_no": 110, "debit": 120, "credit": 120, "balance": 120, "matched": 60}
        for c in cols:
            self.flow_tree.heading(c, text=headings[c])
            self.flow_tree.column(c, width=widths[c], anchor="center")
        vsb = ttk.Scrollbar(top_body, orient="vertical", command=self.flow_tree.yview)
        self.flow_tree.configure(yscrollcommand=vsb.set)
        self.flow_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # --- 下：未达账 + 余额调节表 ---
        bottom = ttk.PanedWindow(body, orient=tk.HORIZONTAL)
        body.add(bottom, weight=2)

        left = ttk.LabelFrame(bottom, text="④ 未达账项明细（共 {} 笔）".format(len(self.unreached)), padding=4)
        bottom.add(left, weight=1)
        ucols = ("category", "source", "date", "settlement_method", "ticket_no", "amount", "phase")
        self.unreach_tree = ttk.Treeview(left, columns=ucols, show="headings", height=8)
        uhead = {"category": "类别", "source": "来源", "date": "日期", "settlement_method": "结算方式",
                 "ticket_no": "票号", "amount": "金额", "phase": "期间"}
        uwidth = {"category": 160, "source": 90, "date": 100, "settlement_method": 80,
                  "ticket_no": 100, "amount": 110, "phase": 55}
        for c in ucols:
            self.unreach_tree.heading(c, text=uhead[c])
            self.unreach_tree.column(c, width=uwidth[c], anchor="center")
        self.unreach_tree.pack(fill="both", expand=True)

        right = ttk.LabelFrame(bottom, text="⑤ 银行存款余额调节表", padding=6)
        bottom.add(right, weight=1)
        self.statement_text = tk.Text(right, height=12, width=60, font=("Consolas", 11))
        self.statement_text.pack(fill="both", expand=True)
        op = ttk.Frame(right)
        op.pack(fill="x", pady=(6, 0))
        ttk.Button(op, text="开始自动对账", command=self.do_reconcile).pack(side="left", padx=4)
        ttk.Button(op, text="导出 Excel (.xlsx)", command=self.export_excel).pack(side="left", padx=4)
        ttk.Button(op, text="导出 CSV", command=self.export_csv).pack(side="left", padx=4)
        ttk.Button(op, text="清空全部", command=self.clear_all).pack(side="left", padx=4)
        self.check_label = ttk.Label(op, text="校验结果：尚未对账", foreground="#1565c0",
                                     font=("", 11, "bold"))
        self.check_label.pack(side="left", padx=16)

    # --------------------------------------------------------
    # 录入 / 导入
    # --------------------------------------------------------
    def add_record(self):
        date = validate_date(self.ent_date.get())
        if date is None:
            messagebox.showwarning("日期格式错误", "日期应为 YYYY-MM-DD（如 2025-12-08）")
            return
        debit = validate_amount(self.ent_debit.get())
        credit = validate_amount(self.ent_credit.get())
        balance = validate_amount(self.ent_balance.get())
        if debit is None or credit is None or balance is None:
            messagebox.showwarning("金额错误", "金额必须是数字")
            return
        if debit > 0 and credit > 0:
            messagebox.showwarning("金额冲突", "同一笔流水只能填借方或贷方其中之一")
            return
        if debit == 0 and credit == 0:
            messagebox.showwarning("金额为空", "请填写借方或贷方金额（必须 > 0）")
            return
        rec = FlowRecord(
            date=date,
            settlement_method=self.ent_method.get().strip(),
            ticket_no=self.ent_ticket.get().strip(),
            debit=debit, credit=credit,
            balance=balance,
            source=self.src_var.get(),
        )
        self.records.append(rec)
        # 清空便于连续录入
        for w in (self.ent_method, self.ent_ticket, self.ent_debit, self.ent_credit, self.ent_balance):
            w.delete(0, tk.END)
        self._refresh_trees()

    def import_csv(self):
        path = filedialog.askopenfilename(
            title="选择 CSV / TXT 流水文件",
            filetypes=[("数据文件", "*.csv *.txt"), ("CSV", "*.csv"), ("TXT", "*.txt")])
        if not path:
            return
        added, skipped = 0, 0
        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                sample = f.read(2048)
                f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample)
                    reader = csv.reader(f, dialect)
                except csv.Error:
                    reader = csv.reader(f)
                rows = list(reader)
            start = 0
            if rows and any(k in str(rows[0]) for k in ("日期", "date", "Date")):
                start = 1
            for row in rows[start:]:
                row = [str(c).strip() for c in row]
                # 列兼容：日期/结算方式/票号/借方/贷方[/期末余额] [/来源]
                if len(row) < 5:
                    skipped += 1
                    continue
                date = validate_date(row[0])
                method = row[1] if len(row) > 1 else ""
                ticket = row[2] if len(row) > 2 else ""
                debit = validate_amount(row[3] if len(row) > 3 else "")
                credit = validate_amount(row[4] if len(row) > 4 else "")
                balance = validate_amount(row[5] if len(row) > 5 else "")
                source = row[6] if len(row) >= 7 and row[6] in (SRC_ENTERPRISE, SRC_BANK) \
                    else self.src_var.get()
                if date is None or debit is None or credit is None \
                        or (debit == 0 and credit == 0):
                    skipped += 1
                    continue
                self.records.append(FlowRecord(
                    date=date, settlement_method=method, ticket_no=ticket,
                    debit=debit, credit=credit, balance=balance, source=source))
                added += 1
        except Exception as e:
            messagebox.showerror("导入失败", str(e))
            return
        self._refresh_trees()
        messagebox.showinfo("导入完成", f"成功导入 {added} 笔，跳过 {skipped} 笔（格式不符）。")

    def import_excel(self):
        """从模板 Excel 导入（4 个 Sheet）：期初余额调节表 / 企业日记账 / 银行对账单 / 期末调节表(空)"""
        path = filedialog.askopenfilename(
            title="选择模板 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            wb = load_workbook(path, data_only=True)
        except Exception as e:
            messagebox.showerror("打开失败", str(e))
            return

        opened = 0
        for sname in ("Sheet1", "Sheet2", "Sheet3", "Sheet4"):
            if sname not in wb.sheetnames:
                continue
            ws = wb[sname]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue

            # Sheet1 = 期初余额调节表（含期初未达账明细）
            if sname == "Sheet1":
                self._parse_opening_sheet(rows)
                opened += 1

            # Sheet2 = 单位银行存款日记账（本期）
            elif sname == "Sheet2":
                added = self._parse_flow_sheet(rows, SRC_ENTERPRISE)
                opened += 1 if added else 0

            # Sheet3 = 银行对账单（本期）
            elif sname == "Sheet3":
                added = self._parse_flow_sheet(rows, SRC_BANK)
                opened += 1 if added else 0

        if opened == 0:
            messagebox.showwarning("未识别到有效数据",
                                   "请确认 Excel 是否包含 Sheet1（期初余额调节表）、Sheet2（企业日记账）或 Sheet3（银行对账单）。")
        else:
            # 将期初未达账项结转进入主流水表（打上 is_opening=True 标记）
            self._carry_forward_opening()
            self._refresh_trees()
            messagebox.showinfo("导入完成",
                                f"已成功载入期初余额调节表和 {opened - 1} 张本期流水表（若有）。\n"
                                f"点击【开始自动对账】即可继续。")

    def _parse_opening_sheet(self, rows):
        """解析 Sheet1（期初余额调节表）格式：
           A3=调整前余额：  B3=企业余额
           C3=调整前余额：  D3=银行余额
           区域 A8..D10 或 A11..D13 = 期初未达账明细
        """
        for row in rows:
            if not row:
                continue
            # 期初企业余额
            if row[0] and "调整前余额" in str(row[0]) and len(row) > 1 and row[1] is not None:
                v = to_float(row[1])
                if v > 0 and self.opening_ent_balance.get() == "0.00":
                    self.opening_ent_balance.set("{:.2f}".format(v))
                elif v > 0:
                    self.opening_ent_balance.set("{:.2f}".format(v))
            # 期初银行余额
            if len(row) > 2 and row[2] and "调整前余额" in str(row[2]) and len(row) > 3 and row[3] is not None:
                v = to_float(row[3])
                if v > 0:
                    self.opening_bank_balance.set("{:.2f}".format(v))

        # 扫描"期初企业方未达账"区域
        ent_flag = False
        bank_flag = False
        for row in rows:
            if not row:
                continue
            text = str(row[0] or "")
            if "期初企业方未达账" in text:
                ent_flag, bank_flag = True, False
                continue
            if "期初银行方未达账" in text:
                ent_flag, bank_flag = False, True
                continue
            # 跳过表头（日期/结算方式/票号/借方/贷方）
            if text in ("日期", "date"):
                continue
            if ent_flag:
                # 列：日期、结算方式、票号、借方金额、贷方金额
                if len(row) >= 5:
                    d = validate_date(row[0])
                    debt = to_float(row[3])
                    cred = to_float(row[4])
                    if d and (debt > 0 or cred > 0):
                        self.opening_unreach_ent.append({
                            "date": d,
                            "settlement_method": str(row[1] or "").strip(),
                            "ticket_no": str(row[2] or "").strip(),
                            "debit": debt,
                            "credit": cred,
                        })
            elif bank_flag:
                if len(row) >= 5:
                    d = validate_date(row[0])
                    debt = to_float(row[3])
                    cred = to_float(row[4])
                    if d and (debt > 0 or cred > 0):
                        self.opening_unreach_bank.append({
                            "date": d,
                            "settlement_method": str(row[1] or "").strip(),
                            "ticket_no": str(row[2] or "").strip(),
                            "debit": debt,
                            "credit": cred,
                        })

    def _parse_flow_sheet(self, rows, source):
        """解析 Sheet2 / Sheet3 本期流水表"""
        added = 0
        header_seen = False
        for row in rows:
            if not row or all(c is None or c == "" for c in row):
                continue
            if any(k in str(row[0] or "") for k in ("日期", "date", "Date", "表3")):
                header_seen = True
                continue
            if not header_seen:
                continue
            if len(row) < 5:
                continue
            date = validate_date(row[0])
            if date is None:
                continue
            method = str(row[1] or "").strip()
            ticket = str(row[2] or "").strip()
            debit = to_float(row[3])
            credit = to_float(row[4])
            balance = to_float(row[5]) if len(row) > 5 else 0.0
            if debit == 0 and credit == 0:
                continue
            self.records.append(FlowRecord(
                date=date, settlement_method=method, ticket_no=ticket,
                debit=debit, credit=credit, balance=balance, source=source))
            added += 1
        return added

    def _carry_forward_opening(self):
        """把期初未达账项转成 is_opening=True 的 FlowRecord，方便统一对账"""
        # 先移除之前结转过的期初（避免重复）
        self.records = [r for r in self.records if not r.is_opening]

        for item in self.opening_unreach_ent:
            self.records.append(FlowRecord(
                date=item["date"], settlement_method=item["settlement_method"],
                ticket_no=item["ticket_no"],
                debit=item["debit"], credit=item["credit"], balance=0.0,
                source=SRC_ENTERPRISE, is_opening=True))
        for item in self.opening_unreach_bank:
            self.records.append(FlowRecord(
                date=item["date"], settlement_method=item["settlement_method"],
                ticket_no=item["ticket_no"],
                debit=item["debit"], credit=item["credit"], balance=0.0,
                source=SRC_BANK, is_opening=True))

    def delete_selected(self):
        sel = self.flow_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在流水清单中选中要删除的行")
            return
        idxs = sorted([int(i) for i in sel], reverse=True)
        for i in idxs:
            if 0 <= i < len(self.records):
                del self.records[i]
        self._refresh_trees()

    def clear_all(self):
        if not messagebox.askyesno("确认清空", "确定要清空全部数据，重新开始吗？"):
            return
        self.records.clear()
        self.opening_unreach_ent.clear()
        self.opening_unreach_bank.clear()
        self.reconciled_pairs.clear()
        self.unreached.clear()
        self.opening_ent_balance.set("0.00")
        self.opening_bank_balance.set("0.00")
        self.statement_text.delete("1.0", tk.END)
        self.check_label.config(text="校验结果：尚未对账", foreground="#1565c0")
        self._refresh_trees()

    # --------------------------------------------------------
    # 核心：自动对账（票号 + 金额 为主键）
    # --------------------------------------------------------
    def do_reconcile(self):
        # 确保期初已结转
        self._carry_forward_opening()

        ent_recs = [r for r in self.records if r.source == SRC_ENTERPRISE]
        bank_recs = [r for r in self.records if r.source == SRC_BANK]

        if not ent_recs or not bank_recs:
            messagebox.showwarning("数据不足",
                                   "请同时录入【企业日记账】和【银行对账单】两类数据后再对账。\n"
                                   "可点击【载入示例数据】或【载入模板 Excel】快速生成。")
            return

        # 重置 matched 标记
        for r in self.records:
            r.matched = False
            r.match_key = None

        self.reconciled_pairs.clear()
        self.unreached.clear()

        # 1) 先按「票号」建索引（票号可能重复时，允许后续按金额再匹配）
        ent_by_ticket = {}
        for r in ent_recs:
            key = (r.ticket_no, r.amount())
            ent_by_ticket.setdefault(key, []).append(r)
        bank_by_ticket = {}
        for r in bank_recs:
            key = (r.ticket_no, r.amount())
            bank_by_ticket.setdefault(key, []).append(r)

        used_ent = [False] * len(ent_recs)
        used_bank = [False] * len(bank_recs)

        # 2) 主匹配：票号 + 金额 + 同方向（同借方或同贷方）
        #    注：银行对账单的借贷含义与企业日记账相反，因此数据层面"同方向"即业务匹配。
        for i, er in enumerate(ent_recs):
            if used_ent[i] or not er.ticket_no:
                continue
            for j, br in enumerate(bank_recs):
                if used_bank[j]:
                    continue
                if er.ticket_no != br.ticket_no or er.amount() != br.amount():
                    continue
                # 同方向匹配：企业借方 ↔ 银行借方，企业贷方 ↔ 银行贷方
                if (er.debit > 0 and br.debit > 0) or (er.credit > 0 and br.credit > 0):
                    self.reconciled_pairs.append((er, br))
                    er.matched = True
                    br.matched = True
                    er.match_key = br.ticket_no
                    br.match_key = er.ticket_no
                    used_ent[i] = True
                    used_bank[j] = True
                    break

        # 3) 辅助匹配：相同日期 + 相同金额 + 同方向（无票号 或 票号没对上时兜底）
        for i, er in enumerate(ent_recs):
            if used_ent[i]:
                continue
            for j, br in enumerate(bank_recs):
                if used_bank[j]:
                    continue
                if er.date != br.date or er.amount() != br.amount():
                    continue
                if (er.debit > 0 and br.debit > 0) or (er.credit > 0 and br.credit > 0):
                    self.reconciled_pairs.append((er, br))
                    er.matched = True
                    br.matched = True
                    er.match_key = br.ticket_no or "(日期+金额匹配)"
                    br.match_key = er.ticket_no or "(日期+金额匹配)"
                    used_ent[i] = True
                    used_bank[j] = True
                    break

        # 4) 归类未达账项
        for i, er in enumerate(ent_recs):
            if used_ent[i]:
                continue
            if er.debit > 0:
                self.unreached.append({"category": CAT_ENT_RECV_BANK_NOT, "record": er})
            else:
                self.unreached.append({"category": CAT_ENT_PAY_BANK_NOT, "record": er})
        for j, br in enumerate(bank_recs):
            if used_bank[j]:
                continue
            # 模板数据：银行借方=收入（企业账户增加），银行贷方=支出（企业账户减少）
            if br.debit > 0:
                self.unreached.append({"category": CAT_BANK_RECV_ENT_NOT, "record": br})
            else:
                self.unreached.append({"category": CAT_BANK_PAY_ENT_NOT, "record": br})

        # 排序：按类别 → 日期
        cat_order = {CAT_ENT_RECV_BANK_NOT: 0, CAT_ENT_PAY_BANK_NOT: 1,
                     CAT_BANK_RECV_ENT_NOT: 2, CAT_BANK_PAY_ENT_NOT: 3}
        self.unreached.sort(key=lambda x: (cat_order.get(x["category"], 9), x["record"].date))

        self._refresh_trees()
        self._generate_statement()

    # --------------------------------------------------------
    # 生成标准余额调节表（对齐模板 Sheet1/Sheet4 格式）
    # --------------------------------------------------------
    def _generate_statement(self):
        try:
            ent_bal = round(float(self.opening_ent_balance.get() or 0), 2)
            bank_bal = round(float(self.opening_bank_balance.get() or 0), 2)
        except ValueError:
            messagebox.showwarning("余额错误", "期初余额必须是数字")
            return

        # 注意：银行对账单借贷含义与企业相反——借方=收入（企业账户增加），贷方=支出（企业账户减少）
        bank_recv = round(sum(u["record"].debit for u in self.unreached
                              if u["category"] == CAT_BANK_RECV_ENT_NOT), 2)
        bank_pay = round(sum(u["record"].credit for u in self.unreached
                             if u["category"] == CAT_BANK_PAY_ENT_NOT), 2)
        ent_recv = round(sum(u["record"].debit for u in self.unreached
                             if u["category"] == CAT_ENT_RECV_BANK_NOT), 2)
        ent_pay = round(sum(u["record"].credit for u in self.unreached
                            if u["category"] == CAT_ENT_PAY_BANK_NOT), 2)

        adj_ent = round(ent_bal + bank_recv - bank_pay, 2)
        adj_bank = round(bank_bal + ent_recv - ent_pay, 2)
        balanced = abs(adj_ent - adj_bank) < 0.005

        sep = "=" * 56
        line = "-" * 56
        txt = []
        txt.append(sep)
        txt.append("                 银 行 存 款 余 额 调 节 表")
        txt.append("                 编制日期：" + datetime.date.today().strftime("%Y-%m-%d"))
        txt.append(sep)
        txt.append("")
        txt.append("【企业日记账账面余额】             {:>14,.2f} 元".format(ent_bal))
        txt.append("  加：银行已收，企业未收款           {:>14,.2f} 元".format(bank_recv))
        txt.append("  减：银行已付，企业未付款           {:>14,.2f} 元".format(bank_pay))
        txt.append(line)
        txt.append("  调节后余额                        {:>14,.2f} 元".format(adj_ent))
        txt.append("")
        txt.append("【银行对账单余额】                 {:>14,.2f} 元".format(bank_bal))
        txt.append("  加：企业已收，银行未收款           {:>14,.2f} 元".format(ent_recv))
        txt.append("  减：企业已付，银行未付款           {:>14,.2f} 元".format(ent_pay))
        txt.append(line)
        txt.append("  调节后余额                        {:>14,.2f} 元".format(adj_bank))
        txt.append("")
        txt.append(sep)
        if balanced:
            txt.append("★ 校验结果：两边调节后余额相等，对账平衡！")
        else:
            txt.append("✗ 校验结果：两边调节后余额不相等，相差 {:,.2f} 元，请核查。".format(adj_bank - adj_ent))
        txt.append(sep)

        self.statement_text.delete("1.0", tk.END)
        self.statement_text.insert("1.0", "\n".join(txt))

        if balanced:
            self.check_label.config(text="校验结果：✔ 调节后余额相等，对账平衡", foreground="#2e7d32")
        else:
            self.check_label.config(text="校验结果：✘ 调节后余额不相等，请核查", foreground="#c62828")

        self._last_statement = "\n".join(txt)
        self._last_balanced = balanced

    # --------------------------------------------------------
    # 导出
    # --------------------------------------------------------
    def export_excel(self):
        if not hasattr(self, "_last_statement"):
            messagebox.showinfo("提示", "请先点击【开始自动对账】生成结果后再导出。")
            return

        path = filedialog.asksaveasfilename(
            title="保存为 Excel",
            defaultextension=".xlsx",
            initialfile="银行对账结果_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S"),
            filetypes=[("Excel 文件", "*.xlsx")])
        if not path:
            return

        try:
            wb = Workbook()

            # --- Sheet1：期初余额调节表（对齐模板） ---
            ws1 = wb.active
            ws1.title = "期初余额调节表"
            self._write_reconciliation_sheet(ws1, mode="opening")

            # --- Sheet2：本期企业日记账 ---
            ws2 = wb.create_sheet("单位银行存款日记账")
            self._write_flow_sheet(ws2, SRC_ENTERPRISE)

            # --- Sheet3：本期银行对账单 ---
            ws3 = wb.create_sheet("银行对账单")
            self._write_flow_sheet(ws3, SRC_BANK)

            # --- Sheet4：期末余额调节表（主输出） ---
            ws4 = wb.create_sheet("期末余额调节表")
            self._write_reconciliation_sheet(ws4, mode="ending")

            # --- Sheet5：已达账项 ---
            ws5 = wb.create_sheet("已达账项")
            self._write_reconciled_sheet(ws5)

            # --- Sheet6：未达账项明细 ---
            ws6 = wb.create_sheet("未达账项明细")
            self._write_unreached_sheet(ws6)

            wb.save(path)
        except Exception as e:
            messagebox.showerror("导出失败", str(e))
            return

        messagebox.showinfo("导出成功", f"结果已保存至：\n{path}\n\n包含以下 6 个 Sheet：\n"
                                       f"Sheet1 期初余额调节表（结转）\n"
                                       f"Sheet2 单位银行存款日记账\n"
                                       f"Sheet3 银行对账单\n"
                                       f"Sheet4 期末余额调节表（主表）\n"
                                       f"Sheet5 已达账项明细\n"
                                       f"Sheet6 未达账项明细")

    def _write_reconciliation_sheet(self, ws, mode="ending"):
        title = "期初余额调节表" if mode == "opening" else "期末余额调节表"
        ent_bal = to_float(self.opening_ent_balance.get())
        bank_bal = to_float(self.opening_bank_balance.get())

        # 注意：银行对账单借贷含义与企业相反——借方=收入（企业账户增加），贷方=支出（企业账户减少）
        bank_recv = round(sum(u["record"].debit for u in self.unreached
                              if u["category"] == CAT_BANK_RECV_ENT_NOT), 2)
        bank_pay = round(sum(u["record"].credit for u in self.unreached
                             if u["category"] == CAT_BANK_PAY_ENT_NOT), 2)
        ent_recv = round(sum(u["record"].debit for u in self.unreached
                             if u["category"] == CAT_ENT_RECV_BANK_NOT), 2)
        ent_pay = round(sum(u["record"].credit for u in self.unreached
                            if u["category"] == CAT_ENT_PAY_BANK_NOT), 2)
        adj_ent = round(ent_bal + bank_recv - bank_pay, 2)
        adj_bank = round(bank_bal + ent_recv - ent_pay, 2)

        ws.merge_cells("A1:E1")
        ws["A1"] = "表3-14  " + title
        ws["A1"].font = Font(name="黑体", size=14, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

        thin = Side(border_style="thin", color="333333")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        fill = PatternFill("solid", fgColor="E8EEF4")

        header_row = ["单位日记账项目", "日记账余额", "银行对账单项目", "对账单余额"]
        for i, h in enumerate(header_row, 1):
            c = ws.cell(row=2, column=i, value=h)
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center")
            c.border = border
            c.fill = fill

        rows_data = [
            ["调整前余额：", ent_bal, "调整前余额：", bank_bal],
            ["加：银行已收，企业未收", bank_recv, "加：企业已收，银行未收", ent_recv],
            ["减：银行已付，企业未付", bank_pay, "减：企业已付，银行未付", ent_pay],
            ["调整后余额：", adj_ent, "调整后余额：", adj_bank],
        ]
        for r, row in enumerate(rows_data, 3):
            for i, val in enumerate(row, 1):
                c = ws.cell(row=r, column=i, value=val)
                c.border = border
                c.alignment = Alignment(horizontal="center" if i in (2, 4) else "left")
                if i in (2, 4) and isinstance(val, (int, float)):
                    c.number_format = "#,##0.00"
                if r == 6:  # 调整后余额加粗
                    c.font = Font(bold=True)

        # 未达账明细区域
        if mode == "ending":
            ent_start = 9
            ws.cell(row=ent_start, column=1, value="期末企业方未达账：").font = Font(bold=True)
            ent_start += 1
            headers = ["日期", "结算方式", "票号", "借方金额", "贷方金额"]
            for i, h in enumerate(headers, 1):
                c = ws.cell(row=ent_start, column=i, value=h)
                c.font = Font(bold=True)
                c.alignment = Alignment(horizontal="center")
                c.border = border

            ent_unreach = [u for u in self.unreached
                           if u["category"] in (CAT_ENT_RECV_BANK_NOT, CAT_ENT_PAY_BANK_NOT)]
            for k, u in enumerate(ent_unreach, 1):
                r = u["record"]
                vals = [r.date, r.settlement_method, r.ticket_no,
                        r.debit if r.debit else "", r.credit if r.credit else ""]
                for i, v in enumerate(vals, 1):
                    c = ws.cell(row=ent_start + k, column=i, value=v)
                    c.border = border
                    if i in (4, 5) and isinstance(v, (int, float)):
                        c.number_format = "#,##0.00"

            bank_start = ent_start + 1 + len(ent_unreach) + 1
            ws.cell(row=bank_start, column=1, value="期末银行方未达账：").font = Font(bold=True)
            bank_start += 1
            for i, h in enumerate(headers, 1):
                c = ws.cell(row=bank_start, column=i, value=h)
                c.font = Font(bold=True)
                c.alignment = Alignment(horizontal="center")
                c.border = border
            bank_unreach = [u for u in self.unreached
                            if u["category"] in (CAT_BANK_RECV_ENT_NOT, CAT_BANK_PAY_ENT_NOT)]
            for k, u in enumerate(bank_unreach, 1):
                r = u["record"]
                vals = [r.date, r.settlement_method, r.ticket_no,
                        r.debit if r.debit else "", r.credit if r.credit else ""]
                for i, v in enumerate(vals, 1):
                    c = ws.cell(row=bank_start + k, column=i, value=v)
                    c.border = border
                    if i in (4, 5) and isinstance(v, (int, float)):
                        c.number_format = "#,##0.00"

        # 列宽
        widths = [18, 16, 18, 16, 12]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    def _write_flow_sheet(self, ws, source):
        ws.merge_cells("A1:F1")
        ws["A1"] = ("表3-15  单位银行存款日记账" if source == SRC_ENTERPRISE
                    else "表3-16  银行对账单")
        ws["A1"].font = Font(name="黑体", size=13, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center")

        headers = ["日期", "结算方式", "票号", "借方金额", "贷方金额", "期末余额"]
        thin = Side(border_style="thin", color="333333")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        fill = PatternFill("solid", fgColor="E8EEF4")
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=2, column=i, value=h)
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center")
            c.border = border
            c.fill = fill

        rows = [r for r in self.records if r.source == source and not r.is_opening]
        rows.sort(key=lambda r: r.date)
        for k, r in enumerate(rows, 3):
            vals = [r.date, r.settlement_method, r.ticket_no,
                    r.debit if r.debit else None,
                    r.credit if r.credit else None,
                    r.balance if r.balance else None]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=k, column=i, value=v)
                c.border = border
                if i in (4, 5, 6) and isinstance(v, (int, float)):
                    c.number_format = "#,##0.00"
        widths = [12, 10, 12, 14, 14, 14]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    def _write_reconciled_sheet(self, ws):
        ws["A1"] = "已达账项（票号/日期 + 金额 匹配成功）"
        ws["A1"].font = Font(name="黑体", size=13, bold=True)
        thin = Side(border_style="thin", color="333333")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        fill = PatternFill("solid", fgColor="E8EEF4")

        headers = ["企业日期", "企业票号", "企业方向", "金额",
                   "银行日期", "银行票号", "银行方向", "金额"]
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=2, column=i, value=h)
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center")
            c.border = border
            c.fill = fill
        for k, (er, br) in enumerate(self.reconciled_pairs, 3):
            vals = [er.date, er.ticket_no,
                    "收入" if er.debit > 0 else "支出", er.amount(),
                    br.date, br.ticket_no,
                    "收入" if br.credit > 0 else "支出", br.amount()]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=k, column=i, value=v)
                c.border = border
                if i in (4, 8) and isinstance(v, (int, float)):
                    c.number_format = "#,##0.00"
        widths = [12, 12, 10, 14, 12, 12, 10, 14]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    def _write_unreached_sheet(self, ws):
        ws["A1"] = "未达账项明细（按类别分列）"
        ws["A1"].font = Font(name="黑体", size=13, bold=True)
        thin = Side(border_style="thin", color="333333")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        fill = PatternFill("solid", fgColor="E8EEF4")

        headers = ["类别", "来源", "日期", "结算方式", "票号",
                   "借方金额", "贷方金额", "期间", "金额"]
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=2, column=i, value=h)
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center")
            c.border = border
            c.fill = fill
        for k, u in enumerate(self.unreached, 3):
            r = u["record"]
            amount = r.amount()
            vals = [u["category"], r.source, r.date, r.settlement_method, r.ticket_no,
                    r.debit if r.debit else None, r.credit if r.credit else None,
                    "期初" if r.is_opening else "本期", amount]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=k, column=i, value=v)
                c.border = border
                if i in (6, 7, 9) and isinstance(v, (int, float)):
                    c.number_format = "#,##0.00"
        widths = [20, 11, 12, 10, 12, 14, 14, 8, 14]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    def export_csv(self):
        if not hasattr(self, "_last_statement"):
            messagebox.showinfo("提示", "请先点击【开始自动对账】生成结果后再导出。")
            return
        path = filedialog.asksaveasfilename(
            title="保存为 CSV",
            defaultextension=".csv",
            initialfile="银行对账结果_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S"),
            filetypes=[("CSV 文件", "*.csv")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f)
                w.writerow(["银行对账系统 导出时间", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
                w.writerow([])
                w.writerow(["【一、全部流水清单】"])
                w.writerow(["期间", "来源", "日期", "结算方式", "票号", "借方金额", "贷方金额", "期末余额", "已达"])
                for r in self.records:
                    w.writerow(["期初" if r.is_opening else "本期", r.source, r.date,
                                r.settlement_method, r.ticket_no,
                                r.debit or "", r.credit or "", r.balance or "",
                                "是" if r.matched else ""])
                w.writerow([])
                w.writerow(["【二、已达账项】"])
                w.writerow(["企业日期", "企业票号", "企业金额", "银行日期", "银行票号", "银行金额"])
                for er, br in self.reconciled_pairs:
                    w.writerow([er.date, er.ticket_no, er.amount(),
                                br.date, br.ticket_no, br.amount()])
                w.writerow([])
                w.writerow(["【三、未达账项】"])
                w.writerow(["类别", "来源", "日期", "结算方式", "票号", "金额", "期间"])
                for u in self.unreached:
                    r = u["record"]
                    w.writerow([u["category"], r.source, r.date, r.settlement_method, r.ticket_no,
                                r.amount(), "期初" if r.is_opening else "本期"])
                w.writerow([])
                w.writerow(["【四、银行存款余额调节表】"])
                for ln in self._last_statement.splitlines():
                    w.writerow([ln])
        except Exception as e:
            messagebox.showerror("导出失败", str(e))
            return
        messagebox.showinfo("导出成功", f"已保存至：\n{path}")

    # --------------------------------------------------------
    # 自动推导期末余额
    # --------------------------------------------------------
    def _compute_ending_balances(self):
        """从期初余额 + 本期借贷方发生额 推导期末余额。
        企业期末 = 期初 + 本期借方（收入）- 本期贷方（支出）
        银行期末 = 期初 + 本期借方（支出）- 本期贷方（收入）
        （注意：is_opening=True 的期初结转记录不计入本期发生额）
        """
        try:
            ent_0 = float(self.opening_ent_balance.get() or 0)
            bank_0 = float(self.opening_bank_balance.get() or 0)
        except ValueError:
            messagebox.showwarning("余额错误", "请先填入正确的期初余额（数字）")
            return

        ent_period = [r for r in self.records
                      if r.source == SRC_ENTERPRISE and not r.is_opening]
        bank_period = [r for r in self.records
                       if r.source == SRC_BANK and not r.is_opening]

        ent_debit = sum(r.debit for r in ent_period)
        ent_credit = sum(r.credit for r in ent_period)
        bank_debit = sum(r.debit for r in bank_period)
        bank_credit = sum(r.credit for r in bank_period)

        if not ent_period and not bank_period:
            messagebox.showinfo("提示",
                                "还没有本期流水数据，无法推导期末余额。\n"
                                "请先录入或导入本期企业日记账 / 银行对账单。")
            return

        ent_ending = round(ent_0 + ent_debit - ent_credit, 2)
        bank_ending = round(bank_0 + bank_debit - bank_credit, 2)

        self.opening_ent_balance.set("{:.2f}".format(ent_ending))
        self.opening_bank_balance.set("{:.2f}".format(bank_ending))

        messagebox.showinfo("期末余额已推导",
                            f"基于期初余额 + 本期发生额：\n\n"
                            f"企业日记账期末余额 = {ent_0:,.2f}（期初）\n"
                            f"                   + {ent_debit:,.2f}（借方/收入）\n"
                            f"                   - {ent_credit:,.2f}（贷方/支出）\n"
                            f"                   = {ent_ending:,.2f} 元\n\n"
                            f"银行对账单期末余额 = {bank_0:,.2f}（期初）\n"
                            f"                   + {bank_debit:,.2f}（借方/支出）\n"
                            f"                   - {bank_credit:,.2f}（贷方/收入）\n"
                            f"                   = {bank_ending:,.2f} 元\n\n"
                            f"（公式：期末 = 期初 + 借方合计 - 贷方合计）")

    def _compute_ending_balances_silent(self):
        """无弹窗版本，供 load_sample_data 等内部调用。返回 (ent_ending, bank_ending) 或 None。"""
        try:
            ent_0 = float(self.opening_ent_balance.get() or 0)
            bank_0 = float(self.opening_bank_balance.get() or 0)
        except ValueError:
            return None
        ent_period = [r for r in self.records
                      if r.source == SRC_ENTERPRISE and not r.is_opening]
        bank_period = [r for r in self.records
                       if r.source == SRC_BANK and not r.is_opening]
        if not ent_period and not bank_period:
            return None
        ent_ending = round(ent_0
                           + sum(r.debit for r in ent_period)
                           - sum(r.credit for r in ent_period), 2)
        bank_ending = round(bank_0
                            + sum(r.debit for r in bank_period)
                            - sum(r.credit for r in bank_period), 2)
        self.opening_ent_balance.set("{:.2f}".format(ent_ending))
        self.opening_bank_balance.set("{:.2f}".format(bank_ending))
        return ent_ending, bank_ending

    # --------------------------------------------------------
    # 载入示例数据（便于快速演示）
    # --------------------------------------------------------
    def load_sample_data(self):
        folder = os.path.dirname(os.path.abspath(__file__))
        ent_csv = os.path.join(folder, "示例数据_单位银行存款日记账.csv")
        bank_csv = os.path.join(folder, "示例数据_银行对账单.csv")
        opening_csv = os.path.join(folder, "示例数据_期初余额调节表.csv")
        template_xlsx = os.path.join(folder, "期初银行存款余额调节表.xlsx")

        self.clear_all()

        # 1. 期初余额（从模板 Excel 载入）
        if os.path.exists(template_xlsx):
            try:
                wb = load_workbook(template_xlsx, data_only=True)
                if "Sheet1" in wb.sheetnames:
                    self._parse_opening_sheet(list(wb["Sheet1"].iter_rows(values_only=True)))
            except Exception:
                pass
        elif os.path.exists(opening_csv):
            try:
                self._load_opening_csv(opening_csv)
            except Exception:
                pass

        # 2. 本期流水
        loaded_count = 0
        if os.path.exists(ent_csv):
            loaded_count += self._load_flow_csv(ent_csv, SRC_ENTERPRISE)
        if os.path.exists(bank_csv):
            loaded_count += self._load_flow_csv(bank_csv, SRC_BANK)

        self._carry_forward_opening()
        # 自动推导期末余额（从期初+本期发生额）
        self._compute_ending_balances_silent()
        self._refresh_trees()
        # 自动对账
        self.do_reconcile()
        ent_0 = float(self.opening_ent_balance.get() or 0)
        bank_0 = float(self.opening_bank_balance.get() or 0)
        messagebox.showinfo("示例数据已载入并完成对账",
                            f"已载入：\n"
                            f"  • 期初余额（企业 / 银行）\n"
                            f"  • 本期企业日记账 / 银行对账单\n"
                            f"  • 期初未达账自动结转\n\n"
                            f"已自动计算：\n"
                            f"  • 企业期末余额 {ent_0:,.2f} 元\n"
                            f"  • 银行期末余额 {bank_0:,.2f} 元\n\n"
                            f"已自动对账：\n"
                            f"  • 已达账项 {len(self.reconciled_pairs)} 对\n"
                            f"  • 未达账项 {len(self.unreached)} 笔\n\n"
                            f"（导出 Excel 可获得标准格式的余额调节表）")

    def _load_opening_csv(self, path):
        section = None
        with open(path, "r", encoding="utf-8") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#"):
                    continue
                if line.startswith("[") and line.endswith("]"):
                    section = line[1:-1]
                    continue
                if section == "企业日记账期初余额":
                    self.opening_ent_balance.set(line)
                elif section == "银行对账单期初余额":
                    self.opening_bank_balance.set(line)
                elif section == "期初企业方未达账":
                    parts = [p.strip() for p in line.split(",")]
                    if parts[0] == "日期":
                        continue
                    if len(parts) >= 5:
                        d = validate_date(parts[0])
                        if d:
                            self.opening_unreach_ent.append({
                                "date": d, "settlement_method": parts[1],
                                "ticket_no": parts[2],
                                "debit": validate_amount(parts[3]),
                                "credit": validate_amount(parts[4]),
                            })
                elif section == "期初银行方未达账":
                    parts = [p.strip() for p in line.split(",")]
                    if parts[0] == "日期":
                        continue
                    if len(parts) >= 5:
                        d = validate_date(parts[0])
                        if d:
                            self.opening_unreach_bank.append({
                                "date": d, "settlement_method": parts[1],
                                "ticket_no": parts[2],
                                "debit": validate_amount(parts[3]),
                                "credit": validate_amount(parts[4]),
                            })

    def _load_flow_csv(self, path, source):
        added = 0
        with open(path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)
        for row in rows[1:]:
            row = [str(c).strip() for c in row]
            if len(row) < 5:
                continue
            date = validate_date(row[0])
            if date is None:
                continue
            debit = validate_amount(row[3])
            credit = validate_amount(row[4])
            balance = validate_amount(row[5]) if len(row) > 5 else 0.0
            if debit == 0 and credit == 0:
                continue
            self.records.append(FlowRecord(
                date=date, settlement_method=row[1], ticket_no=row[2],
                debit=debit, credit=credit, balance=balance, source=source))
            added += 1
        return added

    def load_opening_excel(self):
        """仅载入期初余额调节表"""
        path = filedialog.askopenfilename(
            title="选择模板 Excel 文件（载入期初）",
            filetypes=[("Excel 文件", "*.xlsx")])
        if not path:
            return
        try:
            wb = load_workbook(path, data_only=True)
            if "Sheet1" in wb.sheetnames:
                self._parse_opening_sheet(list(wb["Sheet1"].iter_rows(values_only=True)))
                self._carry_forward_opening()
                self._refresh_trees()
                messagebox.showinfo("期初数据已载入",
                                    f"企业期初余额：{self.opening_ent_balance.get()} 元\n"
                                    f"银行期初余额：{self.opening_bank_balance.get()} 元\n"
                                    f"企业方期初未达账 {len(self.opening_unreach_ent)} 笔\n"
                                    f"银行方期初未达账 {len(self.opening_unreach_bank)} 笔\n"
                                    f"（已自动结转进入本期流水）")
            else:
                messagebox.showwarning("未找到 Sheet1", "模板 Excel 中应包含 Sheet1（期初余额调节表）")
        except Exception as e:
            messagebox.showerror("载入失败", str(e))

    # --------------------------------------------------------
    # 表格刷新
    # --------------------------------------------------------
    def _refresh_trees(self):
        # 流水表
        for item in self.flow_tree.get_children():
            self.flow_tree.delete(item)
        for idx, r in enumerate(self.records):
            self.flow_tree.insert("", "end", iid=str(idx), values=r.tuple_for_tree())
        # 更新计数（通过 window title 不刷新控件标题，改成追加到提示 label 由用户手动看）
        top_body = None
        for child in self.root.winfo_children():
            if isinstance(child, ttk.LabelFrame) and "全部流水清单" in str(child.cget("text")):
                top_body = child
                break
        if top_body is not None:
            top_body.config(text="③ 全部流水清单（含期初结转 + 本期录入，共 {} 笔）".format(len(self.records)))

        # 未达账表
        for item in self.unreach_tree.get_children():
            self.unreach_tree.delete(item)
        for u in self.unreached:
            r = u["record"]
            self.unreach_tree.insert("", "end", values=(
                u["category"], r.source, r.date, r.settlement_method, r.ticket_no,
                "{:.2f}".format(r.amount()),
                "期初" if r.is_opening else "本期"))


# ============================================================
def main():
    root = tk.Tk()
    BankReconciliationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
