import unittest

import formatter as f


class TemplateTest(unittest.TestCase):
    def test_chat_override_and_default(self):
        self.assertEqual(f.parse_template("整形して 列: 日付, 氏名、住所"), ["日付", "氏名", "住所"])
        self.assertEqual(f.parse_template("整形して"), f.FORMAT_TEMPLATE)


class RuleTest(unittest.TestCase):
    def test_split_dates(self):
        row = f.SourceRow(1, {"日付": "7/1A7/3B", "日付2": "8・5C", "氏名": "山田太郎"})
        self.assertEqual(f.split_dates(row), ("7/1", ["A", "7/3B", "8/5C"]))

    def test_normalize_and_validate(self):
        self.assertEqual(f.normalize("1000001", "postal"), "100-0001")
        self.assertEqual(f.normalize("０３－１２３４－５６７８", "phone"), "03-1234-5678")
        self.assertIsNone(f.invalid("03-1234-5678", "phone"))
        self.assertEqual(f.invalid("090-12345678", "phone"), "電話番号の形式ではありません")
        self.assertEqual(f.invalid("13/40", "date"), "日付の形式ではありません")


class FinalizeTest(unittest.TestCase):
    def test_flags_corrections_and_omissions(self):
        template = ["日付", "姓", "名", "住所", "備考"]
        src = [f.SourceRow(1, {"日付": "7/1A", "氏名": "山田太郎", "住所": "千代田区丸の内1-1", "列4": "X9"})]
        shaped = {0: {"src": 0, "姓": "山田", "名": "太郎", "住所": "千代田区丸ノ内1-1", "備考": ""}}
        result = f.finalize(template, src, shaped, ["日付", "氏名", "住所", "列4"])
        self.assertEqual(result.rows[0][:4], ["7/1", "山田", "太郎", "千代田区丸ノ内1-1"])
        self.assertEqual(result.rows[0][4], "A")
        kinds = {(i.column, i.kind) for i in result.issues}
        self.assertIn(("住所", "補正"), kinds)            # value not found in the source row
        self.assertIn(("備考", "欠落の可能性"), kinds)     # "X9" was dropped

    def test_missing_llm_row_falls_back(self):
        src = [f.SourceRow(1, {"住所": "千代田区丸の内1-1"})]
        result = f.finalize(["住所"], src, {}, ["住所"])
        self.assertEqual(result.rows, [["千代田区丸の内1-1"]])
        self.assertEqual(result.issues[0].kind, "未整形")


if __name__ == "__main__":
    unittest.main()
