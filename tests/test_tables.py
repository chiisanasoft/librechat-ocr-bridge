"""Run inside the proxy image:
docker run --rm -v "$PWD":/w -w /w/proxy librechat-ollama-proxy:local python -m unittest discover -s ../tests
"""
import unittest

import headers
import tables


class HeaderInferenceTest(unittest.TestCase):
    def test_ledger_columns(self):
        grid = [
            ["7/12", "信金", "山田", "太郎", "100-0001", "千代田区丸の内1-1", "03-1234-5678"],
            ["7/15", "信金", "佐藤", "花子", "530-0001", "大阪市北区梅田2-3", "06-2345-6789"],
            ["8/02", "信金", "鈴木", "一郎", "460-0002", "名古屋市中区栄3-4", "052-345-6789"],
        ]
        self.assertEqual(headers.infer_headers(grid),
                         ["日付", "列2", "姓", "名", "郵便番号", "住所", "電話番号"])

    def test_existing_header_is_kept(self):
        self.assertIsNone(headers.infer_headers([["日付", "氏名", "住所"], ["1/2", "山田太郎", "港区1-1"]]))

    def test_full_name_and_wareki(self):
        grid = [["山田 太郎", "R8.7.6"], ["佐藤 花子", "令和8年7月8日"]]
        self.assertEqual(headers.infer_headers(grid), ["氏名", "日付"])

    def test_unrecognised_table(self):
        self.assertIsNone(headers.infer_headers([["apple", "3"], ["pear", "5"]]))


class TableConversionTest(unittest.TestCase):
    def test_spans_and_escaping(self):
        html = ('<table><tr><th colspan="2">氏名</th><th>電話</th></tr>'
                '<tr><td rowspan="2">A|B</td><td>x<br>y</td><td>028-000-0000</td></tr>'
                '<tr><td>z</td><td>1</td></tr></table>')
        md, found = tables.convert_html_tables(html)
        self.assertEqual(len(found), 1)
        self.assertIn("| A\\|B | x<br>y | 028-000-0000 |", md)
        grid, merges = found[0].grid()
        self.assertEqual(merges, [(0, 0, 0, 1), (1, 0, 2, 0)])


if __name__ == "__main__":
    unittest.main()
