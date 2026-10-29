# -*- coding: utf-8 -*-
"""跨进程持久化验证（一次性取证脚本，可删）。
用法：
    python3 _persist_demo_check.py save   # 阶段A：录数据 → 对账 → 存 JSON → 进程退出
    python3 _persist_demo_check.py load   # 阶段B：全新进程，空内存 → 打开项目 → 比对指纹
"""
import sys, os, json, time, hashlib

PROJ = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJ)
os.chdir(PROJ)

from tkinter import filedialog, messagebox
import 银行对账系统 as app_mod

JSON_PATH = os.path.join(PROJ, "_persist_demo_项目.json")
REPORT = os.path.join(PROJ, "_persist_demo_报告.txt")

filedialog.asksaveasfilename = lambda **k: JSON_PATH
filedialog.askopenfilename = lambda **k: JSON_PATH
messagebox.askyesno = lambda *a, **k: True
for n in ("showinfo", "showwarning", "showerror"):
    setattr(messagebox, n, lambda *a, **k: None)

import tkinter as tk
import bank_core as bc


def fingerprint(records, opening):
    """把账套内容标准化后取指纹；两边指纹一致 = 数据一模一样。"""
    norm = sorted(
        (r.source, r.date, r.settlement_method, r.ticket_no, round(r.debit or 0, 2),
         round(r.credit or 0, 2), bool(r.is_opening))
        for r in records)
    payload = json.dumps({"n": len(records), "norm": norm,
                          "open": [opening.get("ent_open"), opening.get("bank_open")]},
                         ensure_ascii=False)
    return hashlib.sha256(payload.encode()).hexdigest()


def report(lines):
    text = "\n".join(lines)
    print(text)
    with open(REPORT, "a", encoding="utf-8") as f:
        f.write(text + "\n")


mode = sys.argv[1] if len(sys.argv) > 1 else "save"
if mode == "save" and os.path.exists(REPORT):
    os.remove(REPORT)
t0 = time.time()

root = tk.Tk()
root.withdraw()

if mode == "save":
    a = app_mod.BankReconciliationApp(root)
    a.load_sample_data()
    a.do_reconcile()
    t_work = time.time() - t0
    a.project_file = JSON_PATH                 # 等价于用户在对话框里选了路径
    a.save_project()
    st, sp = a.result["statement"], a.result["stats"]
    report([
        "[阶段A] 进程 %d —— 录数据 + 对账 + 保存" % os.getpid(),
        "  流水笔数            : %d" % len(a.records),
        "  严格 / 容错匹配      : %d / %d" % (sp["严格匹配对数"], sp["容错匹配对数"]),
        "  期末余额 企业/银行   : %.2f / %.2f" % (st["企业期末余额"], st["银行期末余额"]),
        "  调整后余额 企业/银行  : %.2f / %.2f  平衡=%s"
        % (st["企业调整后余额"], st["银行调整后余额"], st["平衡"]),
        "  账套指纹            : %s" % fingerprint(a.records, a.opening)[:16],
        "  JSON 大小           : %d 字节（%d 行）"
        % (os.path.getsize(JSON_PATH), sum(1 for _ in open(JSON_PATH, encoding="utf-8"))),
        "  耗时               : 录入+对账 %.2fs，保存 %.2fs" % (t_work, time.time() - t0 - t_work),
        "  → 进程 A 退出，内存清零（这就是老师版关掉窗口的状态）",
        "",
    ])
else:
    a = app_mod.BankReconciliationApp(root)    # 全新进程：内存里 0 笔
    t_empty = time.time() - t0
    old = open(REPORT, encoding="utf-8").read() if os.path.exists(REPORT) else "(缺阶段A报告)"
    a.open_project()
    t_open = time.time() - t0
    a.do_reconcile()                            # 重算一次，验证结论也能复现
    st, sp = a.result["statement"], a.result["stats"]
    fp = fingerprint(a.records, a.opening)[:16]
    report([
        "[阶段B] 进程 %d —— 独立新进程，启动时内存为空" % os.getpid(),
        "  空程序启动耗时        : %.2fs" % t_empty,
        "  点【打开项目】后      : %d 笔流水，耗时 %.2fs" % (len(a.records), t_open),
        "  严格 / 容错匹配      : %d / %d" % (sp["严格匹配对数"], sp["容错匹配对数"]),
        "  期末余额 企业/银行   : %.2f / %.2f" % (st["企业期末余额"], st["银行期末余额"]),
        "  调整后余额 企业/银行  : %.2f / %.2f  平衡=%s"
        % (st["企业调整后余额"], st["银行调整后余额"], st["平衡"]),
        "  账套指纹            : %s" % fp,
        "",
        "  >> A/B 指纹对比      : %s"
        % ("一致 —— 跨进程零丢失" if fp in old else "不一致 —— 请核查"),
    ])

root.destroy()
