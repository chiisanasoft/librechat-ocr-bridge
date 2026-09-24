"""Infers column headers (日付, 姓, 名, 郵便番号, 住所, ...) for OCR'd tables that have none.

Each column is classified from its non-empty values; a column gets a type when
most values match it. A header row is only added when the first row does not
already look like one.
"""

import re

KANJI = r"一-鿿々ヶ"  # CJK ideographs, 々, ヶ

# Checked in order; the first matching type wins for a cell
CELL_TYPES = [
    ("郵便番号", re.compile(r"^〒?\s*\d{3}\s*-\s*\d{4}$")),
    ("電話番号", re.compile(r"(?<!\d)0\d{1,4}\s*[-(（)）]\s*\d{1,4}\s*-?\s*\d{3,8}(?!\d)|(?<!\d)0\d{9,10}(?!\d)")),
    ("日付", re.compile(r"^(?:(?:\d{4}|[RHSrhs令平昭]\w{0,2}\d{1,2})\s*[/／.．年-]\s*)?\d{1,2}\s*[/／・.．月]\s*\d{1,2}(?!\d)")),
    ("住所", re.compile(rf"^(?:東京都|北海道|(?:京都|大阪)府|[{KANJI}]{{2,3}}県)?[{KANJI}ぁ-んァ-ン]{{1,6}}[市区郡町村].*\d")),
    ("氏名", re.compile(rf"^[{KANJI}]{{2,3}}\s+[{KANJI}]{{1,3}}$|^[{KANJI}]{{4,6}}$")),
    ("名前", re.compile(rf"^[{KANJI}]{{1,3}}$")),
]

HEADER_WORDS = re.compile(r"日付|年月日|氏名|名前|姓|名|郵便|〒|住所|所在地|電話|TEL|tel|番号|備考|No\.?$")

MATCH_RATIO = 0.6
MIN_VALUES = 2
NAME_MIN_UNIQUE_RATIO = 0.5  # rejects repeated labels such as "信金" posing as names


def _cell_type(value: str) -> str | None:
    v = value.strip()
    for name, pattern in CELL_TYPES:
        if pattern.search(v):
            return name
    return None


def classify_column(values: list[str]) -> str | None:
    values = [v.strip() for v in values if v and v.strip()]
    if len(values) < MIN_VALUES:
        return None
    counts: dict[str, int] = {}
    for v in values:
        t = _cell_type(v)
        if t:
            counts[t] = counts.get(t, 0) + 1
    if not counts:
        return None
    best, n = max(counts.items(), key=lambda kv: kv[1])
    if n / len(values) < MATCH_RATIO:
        return None
    if best in ("名前", "氏名") and len(set(values)) / len(values) < NAME_MIN_UNIQUE_RATIO:
        return None
    return best


def looks_like_header(row: list[str]) -> bool:
    filled = [c for c in row if c.strip()]
    if not filled:
        return False
    hits = sum(1 for c in filled if len(c) <= 8 and HEADER_WORDS.search(c))
    return hits / len(filled) >= 0.5


def infer_headers(grid: list[list[str]]) -> list[str] | None:
    """Header labels for a grid without a header row, or None when it already has one."""
    if not grid or looks_like_header(grid[0]):
        return None
    n_cols = len(grid[0])
    types = [classify_column([row[c] for row in grid]) for c in range(n_cols)]
    if not any(types):
        return None

    labels: list[str] = []
    for c, t in enumerate(types):
        if t == "名前":
            # Two adjacent short-name columns are 姓 + 名; a lone one is a full name
            if c + 1 < n_cols and types[c + 1] == "名前" and (c == 0 or types[c - 1] != "名前"):
                t = "姓"
            elif c > 0 and types[c - 1] == "名前" and labels[-1] == "姓":
                t = "名"
            else:
                t = "氏名"
        labels.append(t or f"列{c + 1}")

    seen: dict[str, int] = {}
    for i, label in enumerate(labels):
        seen[label] = seen.get(label, 0) + 1
        if seen[label] > 1:
            labels[i] = f"{label}{seen[label]}"
    return labels
