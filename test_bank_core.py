# -*- coding: utf-8 -*-
"""
================================================================
【test_bank_core.py —— 核心对账逻辑单元测试】
================================================================
运行方式（无需安装 pytest，用 Python 自带的 unittest）：
    cd 银行对账系统
    python3 test_bank_core.py

测试数据：内置示例账套，标准答案：
  企业期初 602,349.00    银行期初 616,149.00
  已达账项 22 对（全部严格匹配）
  企业方未达 4 笔：Z1003(借17,000) Z1015(借126,000) Z2002(贷46,500) Z1010(贷98,520)
  银行方未达 4 笔：Z2002(借240,000) Z2006(借2,808) X1001(贷358,450) Z1010(贷178,000)
  企业期末账面   741,709.78
  银行期末对账单 450,087.78
  四项未达：银行已收企业未收 242,808 / 银行已付企业未付 536,450
            企业已收银行未收 143,000 / 企业已付银行未付 145,020
  两端调整后余额 = 448,067.78（平衡）
================================================================
"""

import json
import os
import unittest

import bank_core as bc


def load_sample_records():
    sample = bc.load_builtin_sample()
    return sample["ent_open"], sample["bank_open"], sample["records"]


class TestSampleReconciliation(unittest.TestCase):
    """示例账套整链对账：结果与预期标准答案一致。"""

    @classmethod
    def setUpClass(cls):
        cls.ent_open, cls.bank_open, cls.records = load_sample_records()
        cls.result = bc.run_reconciliation(cls.records, cls.ent_open, cls.bank_open)

    def test_match_count(self):
        st = self.result["stats"]
        self.assertEqual(st["严格匹配对数"], 22, "严格匹配应为 22 对")
        self.assertEqual(st["容错匹配对数"], 0, "原始示例数据无容错匹配")
        self.assertEqual(st["企业方未达笔数"], 4, "企业方未达应为 4 笔")
        self.assertEqual(st["银行方未达笔数"], 4, "银行方未达应为 4 笔")

    def test_unreached_tickets(self):
        ent_tickets = sorted(r.ticket_no for r in self.result["ent_only"])
        bank_tickets = sorted(r.ticket_no for r in self.result["bank_only"])
        self.assertEqual(ent_tickets, ["Z1003", "Z1010", "Z1015", "Z2002"])
        self.assertEqual(bank_tickets, ["X1001", "Z1010", "Z2002", "Z2006"])

    def test_four_categories(self):
        st = self.result["statement"]
        self.assertEqual(st["企业_银行已收企业未收"], 242808.0)
        self.assertEqual(st["企业_银行已付企业未付"], 536450.0)
        self.assertEqual(st["银行_企业已收银行未收"], 143000.0)
        self.assertEqual(st["银行_企业已付银行未付"], 145020.0)

    def test_ending_balances(self):
        self.assertEqual(self.result["ent_end"], 741709.78)
        self.assertEqual(self.result["bank_end"], 450087.78)

    def test_statement_balanced(self):
        st = self.result["statement"]
        self.assertTrue(st["平衡"], "两端调整后余额应相等")
        self.assertEqual(st["企业调整后余额"], 448067.78)
        self.assertEqual(st["银行调整后余额"], 448067.78)

    def test_opening_unreach_carried_and_matched(self):
        """期初结转的 Z1152（企业方未达账）应与银行 12-02 Z1152 严格匹配。"""
        z1152 = [r for r in self.records if r.ticket_no == "Z1152"]
        self.assertEqual(len(z1152), 2)
        self.assertTrue(all(r.matched == bc.MATCH_STRICT for r in z1152),
                        "期初未达账 Z1152 应被严格匹配")
        self.assertTrue(any(r.is_opening for r in z1152), "一笔应标记为期初结转")

    def test_rolling_balance_last_row(self):
        """企业日记账最后一笔 Z1006 的滚动余额应为 741,709.78。"""
        z1006 = [r for r in self.records
                 if r.ticket_no == "Z1006" and r.source == bc.SRC_ENTERPRISE][0]
        self.assertEqual(z1006.balance, 741709.78)


class TestStrictMatching(unittest.TestCase):
    """严格口径：四字段任一不同都不能算严格匹配。"""

    def _reconcile_two(self, ent_kwargs, bank_kwargs):
        er = bc.FlowRecord(source=bc.SRC_ENTERPRISE, **ent_kwargs)
        br = bc.FlowRecord(source=bc.SRC_BANK, **bank_kwargs)
        res = bc.run_reconciliation([er, br], 1000.0, 1000.0)
        return res

    def test_identical_is_strict(self):
        kw = dict(date="2025-12-01", settlement_method="202", ticket_no="Z1",
                  debit=0, credit=100)
        res = self._reconcile_two(kw, kw)
        self.assertEqual(res["stats"]["严格匹配对数"], 1)

    def test_method_differs_not_strict(self):
        """结算方式不同 → 严格不匹配，票号+金额容错兜底。"""
        res = self._reconcile_two(
            dict(date="2025-12-01", settlement_method="202", ticket_no="Z1", debit=0, credit=100),
            dict(date="2025-12-01", settlement_method="电汇", ticket_no="Z1", debit=0, credit=100))
        self.assertEqual(res["stats"]["严格匹配对数"], 0)
        self.assertEqual(res["stats"]["容错匹配对数"], 1, "应由第2轮票号+金额容错兜底")

    def test_amount_differs_not_matched(self):
        """票号相同但金额不同（Z1010 案例）→ 双方均未达，不强行配对。"""
        res = self._reconcile_two(
            dict(date="2025-12-20", settlement_method="202", ticket_no="Z1010", debit=0, credit=98520),
            dict(date="2025-12-13", settlement_method="202", ticket_no="Z1010", debit=0, credit=178000))
        self.assertEqual(res["stats"]["企业方未达笔数"], 1)
        self.assertEqual(res["stats"]["银行方未达笔数"], 1)

    def test_loose_round3_date_amount(self):
        """无票号流水 → 第3轮日期+金额兜底。"""
        res = self._reconcile_two(
            dict(date="2025-12-05", settlement_method="", ticket_no="", debit=500, credit=0),
            dict(date="2025-12-05", settlement_method="202", ticket_no="", debit=500, credit=0))
        self.assertEqual(res["stats"]["容错匹配对数"], 1)
        er = res["pairs"][0][0]
        self.assertEqual(er.mark_label(), "？", "容错匹配界面上应显示 ？")

    def test_cross_direction_never_matches(self):
        """借贷方向相反的同样金额不允许配对。"""
        res = self._reconcile_two(
            dict(date="2025-12-01", settlement_method="202", ticket_no="Z9", debit=300, credit=0),
            dict(date="2025-12-01", settlement_method="202", ticket_no="Z9", debit=0, credit=300))
        self.assertEqual(res["stats"]["严格匹配对数"], 0)
        self.assertEqual(res["stats"]["容错匹配对数"], 0)


class TestRollBalance(unittest.TestCase):
    def test_rolling(self):
        rs = [bc.FlowRecord(date="2025-12-01", settlement_method="202", ticket_no="A",
                            debit=100, credit=0, source=bc.SRC_ENTERPRISE),
              bc.FlowRecord(date="2025-12-02", settlement_method="202", ticket_no="B",
                            debit=0, credit=30, source=bc.SRC_ENTERPRISE)]
        end = bc.roll_balance(rs, 1000)
        self.assertEqual(rs[0].balance, 1100)
        self.assertEqual(rs[1].balance, 1070)
        self.assertEqual(end, 1070)

    def test_opening_records_skipped(self):
        rs = [bc.FlowRecord(date="2025-12-01", settlement_method="202", ticket_no="A",
                            debit=0, credit=50, source=bc.SRC_ENTERPRISE, is_opening=True),
              bc.FlowRecord(date="2025-12-02", settlement_method="202", ticket_no="B",
                            debit=200, credit=0, source=bc.SRC_ENTERPRISE)]
        end = bc.roll_balance(rs, 1000)
        self.assertEqual(rs[0].balance, 0, "期初结转记录不参与滚动")
        self.assertEqual(end, 1200)


class TestJsonPersistence(unittest.TestCase):
    """项目保存 / 打开：FlowRecord 必须能无损序列化往返。"""

    def test_roundtrip(self):
        _, _, records = load_sample_records()
        data = json.dumps([r.to_dict() for r in records], ensure_ascii=False)
        back = [bc.FlowRecord.from_dict(d) for d in json.loads(data)]
        self.assertEqual(len(back), len(records))
        for a, b in zip(records, back):
            self.assertEqual(a.date, b.date)
            self.assertEqual(a.ticket_no, b.ticket_no)
            self.assertEqual(a.debit, b.debit)
            self.assertEqual(a.credit, b.credit)
            self.assertEqual(a.source, b.source)
            self.assertEqual(a.is_opening, b.is_opening)
        # 序列化恢复后重新对账，结果必须一致
        res = bc.run_reconciliation(back, 602349.0, 616149.0)
        self.assertEqual(res["statement"]["企业调整后余额"], 448067.78)


class TestCsvImport(unittest.TestCase):
    """CSV 导入：应能读取仓库自带的示例数据文件。"""

    BASE = os.path.dirname(os.path.abspath(__file__))

    def test_read_flow_csv(self):
        path = os.path.join(self.BASE, "示例数据_单位银行存款日记账.csv")
        if not os.path.exists(path):
            self.skipTest("示例 CSV 不在当前目录")
        rows = bc.read_flow_csv(path, bc.SRC_ENTERPRISE)
        self.assertEqual(len(rows), 25)

    def test_read_opening_csv(self):
        path = os.path.join(self.BASE, "示例数据_期初余额调节表.csv")
        if not os.path.exists(path):
            self.skipTest("示例 CSV 不在当前目录")
        opening = bc.read_opening_csv(path)
        self.assertEqual(opening["ent_open"], 602349.0)
        self.assertEqual(opening["bank_open"], 616149.0)
        self.assertEqual(len(opening["ent_unreach"]), 1)

    def test_csv_matches_builtin(self):
        """CSV 导入后对账结果应与内置示例完全一致（同一套账）。"""
        path_e = os.path.join(self.BASE, "示例数据_单位银行存款日记账.csv")
        path_b = os.path.join(self.BASE, "示例数据_银行对账单.csv")
        path_o = os.path.join(self.BASE, "示例数据_期初余额调节表.csv")
        if not all(os.path.exists(p) for p in (path_e, path_b, path_o)):
            self.skipTest("示例 CSV 不在当前目录")
        opening = bc.read_opening_csv(path_o)
        records = bc.opening_items_to_records(opening)
        records += bc.read_flow_csv(path_e, bc.SRC_ENTERPRISE)
        records += bc.read_flow_csv(path_b, bc.SRC_BANK)
        res = bc.run_reconciliation(records, opening["ent_open"], opening["bank_open"])
        self.assertEqual(res["stats"]["严格匹配对数"], 22)
        self.assertTrue(res["statement"]["平衡"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
