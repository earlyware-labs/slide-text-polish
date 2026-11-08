# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx", "budoux"]
# ///
"""スライドの文章量・フォント・言語・改行をチェックして一覧にする。

使い方:
  uv run --with python-pptx --with budoux python lint.py deck.pptx                      # 改行は文字幅からの推定
  uv run --with python-pptx --with budoux python lint.py deck.pptx --lines lines.tsv    # PowerPoint 実機の改行（check_ppt.ps1 の出力）を使う
  uv run --with python-pptx --with budoux python lint.py deck.pptx --glossary glossary.txt --json out.json
"""
import argparse
import csv
import json
from collections import defaultdict

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Pt

from bullets import bullet_issues, label_without_hanging
from breaks import HEAD_NG, TAIL_NG, break_points, load_glossary
from common import JA, WESTERN, chart_labels, text_frames, width

LIMIT = {'title': 20, 'text': 25, 'cell': 15, 'chart': 15, 'category': 8, 'series': 8}
MIN_PT = {'title': 28, 'text': 18, 'cell': 12, 'chart': 12}
MAX_ITEMS = 5


def is_title(sh):
    try:
        return sh.is_placeholder and sh.placeholder_format.type in (1, 3)  # TITLE, CENTER_TITLE
    except Exception:
        return False


def para_size(p):
    for r in p.runs:
        if r.font.size:
            return r.font.size
    return None


def estimate_lines(text, sh, tf, size, marL=0):
    """枠の幅と文字サイズから折り返しを推定（目安）"""
    try:
        inner = sh.width - (tf.margin_left or 91440) - (tf.margin_right or 91440) - marL
    except Exception:
        return [text]
    per_line = max(1.0, inner / (size or Pt(18)))
    lines, cur, cw = [], '', 0
    for ch in text:
        if cw + (1 if ord(ch) > 0xff else 0.5) > per_line:
            # 簡易禁則: 行頭禁則文字は前の行にぶら下げる
            if ch in HEAD_NG:
                cur += ch
                lines.append(cur)
                cur, cw = '', 0
                continue
            lines.append(cur)
            cur, cw = '', 0
        cur += ch
        cw += 1 if ord(ch) > 0xff else 0.5
    if cur:
        lines.append(cur)
    return lines


def load_lines(path):
    d = defaultdict(list)
    with open(path, encoding='utf-8-sig') as f:
        for row in csv.DictReader(f, delimiter='\t'):
            d[(int(row['slide']), row['id'], int(row['para']))].append(row['text'])
    return d


def line_issues(lines, glossary):
    issues = []
    if len(lines) < 2:
        return issues
    full = ''.join(lines)
    ok = set(break_points(full, glossary))
    pos = 0
    for k, ln in enumerate(lines[:-1]):
        pos += len(ln)
        nxt = lines[k + 1]
        if pos not in ok:
            issues.append(f'泣き別れ: 「{ln[-4:]}｜{nxt[:4]}」')
        if nxt and nxt[0] in HEAD_NG:
            issues.append(f'行頭禁則: 「{nxt[:4]}」')
        if ln and ln[-1] in TAIL_NG:
            issues.append(f'行末禁則: 「{ln[-4:]}」')
    if width(lines[-1].strip()) <= 2:
        issues.append(f'最終行が{len(lines[-1].strip())}文字だけ')
    return issues


def shape_key(sh, where, kind):
    """lines.tsv と照合する図形のキー（id。表のセルは id[行,列]）"""
    key = str(sh.shape_id)
    return key + where[where.rfind('['):] if kind == 'cell' else key


def lint(path, lines_tsv=None, glossary=()):
    prs = Presentation(path)
    real = load_lines(lines_tsv) if lines_tsv else None
    out = []
    stale = []  # lines.tsv と本文が一致しない段落
    ea_fonts = defaultdict(int)

    def add(i, where, kind, msg, text):
        out.append({'slide': i, 'where': where, 'kind': kind, 'issue': msg, 'text': text[:60]})

    for i, where, kind, tf, sh in text_frames(prs):
        k = 'title' if kind == 'text' and is_title(sh) else kind
        paras = [p for p in tf.paragraphs if p.text.strip()]
        if k == 'text' and len(paras) > MAX_ITEMS + 1:
            add(i, where, k, f'項目が多い({len(paras)})', paras[0].text)
        if k in ('text', 'cell'):
            for pi, msg in bullet_issues(tf, sh):
                add(i, where, k, msg, tf.paragraphs[pi - 1].text.replace('\v', '⏎'))
        for pi, p in enumerate(tf.paragraphs, 1):
            text = p.text.replace('\v', '')
            if not text.strip():
                continue
            if width(text) > LIMIT[k]:
                add(i, where, k, f'長い({width(text):.0f}字>{LIMIT[k]})', text)
            size = para_size(p)
            if size and size.pt < MIN_PT[k]:
                add(i, where, k, f'文字が小さい({size.pt:.0f}pt<{MIN_PT[k]})', text)
            if k == 'text' and text.rstrip().endswith('。') and len(paras) > 1:
                add(i, where, k, '箇条書きの文末に「。」', text)
            for r in p.runs:
                if not JA.search(r.text):
                    continue
                rPr = r._r.find(qn('a:rPr'))
                lang = rPr.get('lang') if rPr is not None else None
                if lang != 'ja-JP':
                    add(i, where, k, f'言語が{lang or "未指定"}（禁則が効かない）', text)
                ea = rPr.find(qn('a:ea')) if rPr is not None else None
                face = ea.get('typeface') if ea is not None else None
                if face:
                    ea_fonts[face] += 1
                if face in WESTERN:
                    add(i, where, k, f'和文に欧文フォント({face})', text)
                elif not face:
                    latin = rPr.find(qn('a:latin')) if rPr is not None else None
                    lf = latin.get('typeface') if latin is not None else None
                    if lf in WESTERN:
                        add(i, where, k, f'和文フォント未指定({lf}のみ)', text)
                if rPr is not None and rPr.get('i') == '1':
                    add(i, where, k, '日本語が斜体', text)
                break
            # 改行
            if real is not None:
                lines = real.get((i, shape_key(sh, where, kind), pi))
                # lines.tsv がこのファイルのものか（修正前のファイルで書き出したものではないか）
                if kind in ('text', 'cell') and (
                        lines is None or ''.join(lines) != text.replace('\t', ' ')):
                    stale.append(f'S{i} [{where}] {text[:30]}')
                    lines = None
            else:
                # 段落内改行ごとに分け、ぶら下げ（marL）の分だけ幅を狭めて推定する
                marL = int(p._p.pPr.get('marL') or 0) if p._p.pPr is not None else 0
                lines = [ln for seg in p.text.split('\v')
                         for ln in estimate_lines(seg, sh, tf, size, marL)] if k in ('text', 'title') else None
            if lines:
                if k == 'title' and len(lines) > 1:
                    add(i, where, k, f'タイトルが{len(lines)}行', text)
                if k in ('text', 'cell') and len(lines) > 1 and label_without_hanging(p._p):
                    add(i, where, k, 'ラベル付きの段落が折り返すが、ぶら下げが無い（2行目がラベルの下に来る）', text)
                if k == 'text' and len(lines) > 2:
                    add(i, where, k, f'{len(lines)}行に折り返し', text)
                if k == 'cell' and len(lines) > 1:
                    add(i, where, k, f'セルが{len(lines)}行', text)
                for m in line_issues(lines, glossary):
                    add(i, where, k, m, text)
    for i, name, kind, label in chart_labels(prs):
        if width(label) > LIMIT[kind]:
            add(i, name, kind, f'グラフの{"項目名" if kind == "category" else "系列名"}が長い({width(label):.0f}字)', label)
    return out, dict(ea_fonts), 'PowerPoint実機' if real is not None else '推定', stale


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx')
    ap.add_argument('--lines')
    ap.add_argument('--glossary')
    ap.add_argument('--json')
    a = ap.parse_args()
    gl = load_glossary(a.glossary)
    issues, fonts, mode, stale = lint(a.pptx, a.lines, gl)
    print(f'改行の判定: {mode} / 指摘 {len(issues)} 件 / 和文フォント: {fonts or "指定なし"}')
    if mode == '推定':
        print('  ※改行は推定。実機での確認（SKILL.md 手順6）が済むまで完了にしない')
    elif stale:
        print(f'  ※lines.tsv と本文が一致しない段落が {len(stale)} 件（その段落は改行を判定していない）。'
              'このファイルで check_ppt.ps1 を実行し直す')
        for s in stale[:5]:
            print(f'    {s}')
    if len(fonts) > 2:
        print('  和文フォントが3種類以上。統一を検討')
    for x in issues:
        print(f"S{x['slide']} [{x['where']}] {x['issue']}: {x['text']}")
    if a.json:
        with open(a.json, 'w', encoding='utf-8') as f:
            json.dump(issues, f, ensure_ascii=False, indent=1)


if __name__ == '__main__':
    main()
