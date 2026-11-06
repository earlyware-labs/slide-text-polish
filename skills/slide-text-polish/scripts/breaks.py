# /// script
# requires-python = ">=3.10"
# dependencies = ["budoux"]
# ///
"""改行してよい位置（文節の切れ目）を求める。

BudouX があれば文節区切りに使い、無ければ文字種の切れ目による簡易版で代用する。
用語辞書の語の内側と「数字＋単位」の内側は改行候補から除く。
"""
import re

HEAD_NG = set('、。，．,.)）」』】〕〉》！？!?ー・：；:;ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ々〜')
TAIL_NG = set('(（「『【〔〈《')
NUM_UNIT = re.compile(r'[\d０-９][\d０-９,.，．]*\s*(?:%|％|円|万円|億円|千円|万|億|千|件|人|名|社|個|台|枚|回|点|年|か月|ヶ月|月|日|時間|分|秒|倍|pt|GB|MB|kg|km|m)?')

try:
    import budoux
    _parser = budoux.load_default_japanese_parser()
except Exception:  # BudouX が無い環境
    _parser = None


def _kind(ch):
    o = ord(ch)
    if 0x3040 <= o <= 0x309f:
        return 'hira'
    if 0x30a0 <= o <= 0x30ff:
        return 'kata'
    if 0x4e00 <= o <= 0x9fff or ch == '々':
        return 'kanji'
    if ch.isascii() and ch.isalnum():
        return 'alnum'
    return 'other'


def _simple_chunks(text):
    """簡易文節分割: ひらがな→非ひらがな の境目と、句読点・空白の後で区切る"""
    chunks, cur = [], ''
    for k, ch in enumerate(text):
        if cur:
            prev = text[k - 1]
            if (_kind(prev) == 'hira' and _kind(ch) in ('kanji', 'kata', 'alnum')) \
                    or prev in '、。・ 　/／' or (_kind(prev) == 'other' and prev not in TAIL_NG and _kind(ch) != 'other'):
                chunks.append(cur)
                cur = ''
        cur += ch
    if cur:
        chunks.append(cur)
    return chunks


def chunks(text):
    return _parser.parse(text) if _parser else _simple_chunks(text)


def break_points(text, glossary=()):
    pos, points = 0, []
    for c in chunks(text)[:-1]:
        pos += len(c)
        points.append(pos)
    protected = []
    for g in glossary:
        for m in re.finditer(re.escape(g), text):
            protected.append((m.start(), m.end()))
    for m in NUM_UNIT.finditer(text):
        protected.append((m.start(), m.end()))
    for m in re.finditer(r'[A-Za-z][A-Za-z0-9.+\-]*', text):
        protected.append((m.start(), m.end()))
    return [p for p in points
            if not any(s < p < e for s, e in protected)
            and not (p < len(text) and text[p] in HEAD_NG)
            and not (p > 0 and text[p - 1] in TAIL_NG)]


def load_glossary(path):
    if not path:
        return []
    with open(path, encoding='utf-8') as f:
        return [ln.strip() for ln in f if ln.strip() and not ln.startswith('#')]


def suggest_break(text, line_end, glossary=()):
    """実際の改行位置 line_end の手前で一番近い改行候補を返す（無ければ None）"""
    cands = [p for p in break_points(text, glossary) if p <= line_end]
    return max(cands) if cands else None


if __name__ == '__main__':
    import sys
    s = sys.argv[1] if len(sys.argv) > 1 else '品番マスタと自動照合し、年間約860万円の削減効果を見込む'
    print('BudouX' if _parser else '簡易分割', chunks(s))
    print(break_points(s))
