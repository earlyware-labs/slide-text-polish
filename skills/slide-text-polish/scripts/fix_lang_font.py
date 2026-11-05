"""日本語テキストの言語指定・和文フォント・斜体を一括修正する。

使い方: python fix_lang_font.py 入力.pptx 出力.pptx [--font 游ゴシック]
- 日本語を含む run と段落末を lang="ja-JP" に（PowerPoint の日本語禁則処理を効かせる）
- 英語辞書のスペルミス印 err="1" を削除
- 和文フォント <a:ea> が未指定・空・欧文なら指定フォントを設定
- 日本語の斜体を解除
- テーマの東アジアフォントが空・欧文なら同じく設定
"""
import argparse
import sys

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

from common import JA, WESTERN, text_frames

# a:rPr 内の子要素の順序（スキーマ順）
RPR_ORDER = ['ln', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill',
             'effectLst', 'effectDag', 'highlight', 'uLnTx', 'uLn', 'uFillTx', 'uFill',
             'latin', 'ea', 'cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst']


def set_ea(rPr, font):
    ea = rPr.find(qn('a:ea'))
    if ea is not None:
        face = ea.get('typeface') or ''
        if face and face not in WESTERN:
            return False
        ea.set('typeface', font)
        return True
    ea = etree.Element(qn('a:ea'))
    ea.set('typeface', font)
    idx = RPR_ORDER.index('ea')
    for child in rPr:
        tag = etree.QName(child).localname
        if tag in RPR_ORDER and RPR_ORDER.index(tag) > idx:
            child.addprevious(ea)
            return True
    rPr.append(ea)
    return True


def fix(prs, font):
    stats = {'lang': 0, 'err': 0, 'font': 0, 'italic': 0}
    for _, _, _, tf, _ in text_frames(prs):
        for p in tf.paragraphs:
            if not JA.search(p.text):
                continue
            targets = [r._r.get_or_add_rPr() for r in p.runs if JA.search(r.text)]
            end = p._p.find(qn('a:endParaRPr'))
            if end is not None:
                targets.append(end)
            for rPr in targets:
                if rPr.get('lang') != 'ja-JP':
                    rPr.set('lang', 'ja-JP')
                    stats['lang'] += 1
                if 'err' in rPr.attrib:
                    del rPr.attrib['err']
                    stats['err'] += 1
                if rPr.get('i') == '1':
                    rPr.set('i', '0')
                    stats['italic'] += 1
                if set_ea(rPr, font):
                    stats['font'] += 1
    # テーマの東アジアフォント
    for master in prs.slide_masters:
        for rel in master.part.rels.values():
            if rel.reltype.endswith('/theme'):
                theme = etree.fromstring(rel.target_part.blob)
                changed = False
                for ea in theme.iter(qn('a:ea')):
                    if ea.getparent().tag in (qn('a:majorFont'), qn('a:minorFont')):
                        face = ea.get('typeface') or ''
                        if not face or face in WESTERN:
                            ea.set('typeface', font)
                            changed = True
                if changed:
                    rel.target_part._blob = etree.tostring(theme, xml_declaration=True,
                                                           encoding='UTF-8', standalone=True)
                    stats['theme'] = stats.get('theme', 0) + 1
    return stats


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src')
    ap.add_argument('dst')
    ap.add_argument('--font', default='游ゴシック')
    a = ap.parse_args()
    if a.src == a.dst:
        sys.exit('元ファイルは上書きしません。出力先を別名にしてください')
    prs = Presentation(a.src)
    stats = fix(prs, a.font)
    prs.save(a.dst)
    Presentation(a.dst)  # 壊れていないか開き直して確認
    print('修正件数:', stats)


if __name__ == '__main__':
    main()
