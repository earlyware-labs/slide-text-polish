# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx", "budoux"]
# ///
"""文章を短くした後、大事な情報が消えていないかを修正前後で照合する。

使い方: uv run --with python-pptx --with budoux python check_loss.py 修正前.pptx 修正後.pptx [--glossary glossary.txt]
- 段落単位: 否定表現の数の変化
- 資料全体: 数字（単位込み）・カタカナ語/英字の用語・辞書の語が完全に消えたもの
  （別のスライドに移した場合は消失扱いにしない）
終了コード: 指摘があれば 1
"""
import argparse
import re
import sys
from collections import Counter

from pptx import Presentation

from breaks import load_glossary
from common import chart_labels, text_frames

NUM = re.compile(r'[\d][\d,.]*\s*(?:%|万円|億円|千円|円|万|億|千|件|人|名|社|個|台|枚|回|点|年|か月|ヶ月|月|日|時間|分|秒|倍|pt|GB|MB)?')
NEG = re.compile(r'ない|なく|せず|ません|ず[、。]|不[一-鿿]|未[一-鿿]|非[一-鿿]|無[一-鿿]|禁止|不要|不可|除く|以外')
TERM = re.compile(r'[゠-ヿ]{3,}|[A-Za-z][A-Za-z0-9.+\-]{1,}')
Z2H = str.maketrans('０１２３４５６７８９，．％', '0123456789,.%')


def norm(s):
    return s.translate(Z2H).replace(' ', '').replace('　', '')


def collect(path):
    prs = Presentation(path)
    paras = {}
    for i, where, _, tf, _ in text_frames(prs):
        for k, p in enumerate(tf.paragraphs):
            if p.text.strip():
                paras[(i, where, k)] = p.text.replace('\v', '')
    for i, name, kind, label in chart_labels(prs):
        paras[(i, f'{name}:{kind}', label)] = label
    return paras


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('before')
    ap.add_argument('after')
    ap.add_argument('--glossary')
    a = ap.parse_args()
    gl = load_glossary(a.glossary)
    b, af = collect(a.before), collect(a.after)
    whole_b = norm('\n'.join(b.values()))
    whole_a = norm('\n'.join(af.values()))
    issues = []

    for m in sorted(set(NUM.findall(whole_b))):
        m = m.strip()
        if m and m not in whole_a:
            issues.append(f'数字が資料から消えた: {m}')
    for t in sorted(set(TERM.findall(whole_b))):
        if t not in whole_a:
            issues.append(f'用語が資料から消えた: {t}')
    for g in gl:
        if norm(g) in whole_b and norm(g) not in whole_a:
            issues.append(f'辞書の語が資料から消えた: {g}')
    # 段落単位の否定表現（同じ場所の段落どうし）
    for key, tb in b.items():
        ta = af.get(key)
        if ta is None or ta == tb:
            continue
        nb, na = Counter(NEG.findall(tb)), Counter(NEG.findall(ta))
        if nb != na:
            issues.append(f'S{key[0]} [{key[1]}] 否定表現が変わった: 「{tb[:40]}」→「{ta[:40]}」')
    # 否定表現の総数
    nb, na = len(NEG.findall(whole_b)), len(NEG.findall(whole_a))
    if nb != na:
        issues.append(f'資料全体の否定表現の数: {nb} → {na}')

    if issues:
        print(f'要確認 {len(issues)} 件')
        for x in issues:
            print(' -', x)
        sys.exit(1)
    print('改悪チェック: 指摘なし')


if __name__ == '__main__':
    main()
