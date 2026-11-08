"""箇条書きの記号とぶら下げインデント。

AI製の資料では「・」を文字で打ち、改行を段落の区切り（Enter）や空白で揃えていることが多い。
この形だと2行目以降が記号の下（段落の左端）に来るため、本物の箇条書き記号＋ぶら下げインデント
（marL＝本文の開始位置、indent＝-記号の幅）に直し、続きの段落は段落内改行（<a:br/>）でつなぐ。
"""
import copy
import re

from lxml import etree
from pptx.oxml.ns import qn

# 文字で打たれた箇条書き記号（- と * は直後に空白があるときだけ）
TEXT_BULLET = re.compile(r'^[ \t　]*(?:([・•●○■□◆◇▪▸►‣])[ \t　]*|([-‐－*＊])[ \t　]+)')
LEAD_SPACE = re.compile(r'^[ \t　]+')
BU_TAGS = ('buChar', 'buAutoNum', 'buBlip')
# a:pPr 内の子要素の順序（スキーマ順）
PPR_ORDER = ['lnSpc', 'spcBef', 'spcAft', 'buClrTx', 'buClr', 'buSzTx', 'buSzPct', 'buSzPts',
             'buFontTx', 'buFont', 'buNone', 'buAutoNum', 'buChar', 'buBlip', 'tabLst', 'defRPr', 'extLst']
DEFAULT_SZ = 1800  # 1/100 pt


def para_text(p_el):
    """段落の文字列（段落内改行は \\v）"""
    out = []
    for ch in p_el:
        tag = etree.QName(ch).localname
        if tag in ('r', 'fld'):
            out.append(''.join(t.text or '' for t in ch.iter(qn('a:t'))))
        elif tag == 'br':
            out.append('\v')
    return ''.join(out)


def text_bullet(text):
    """行頭の文字の記号を返す（無ければ None）"""
    m = TEXT_BULLET.match(text)
    return (m.group(1) or m.group(2)) if m else None


def level(p_el):
    pPr = p_el.find(qn('a:pPr'))
    return int(pPr.get('lvl', '0')) if pPr is not None else 0


def explicit_bullet(p_el):
    """段落に直接書かれた箇条書き指定: True（記号あり）/ False（buNone）/ None（指定なし＝継承）"""
    pPr = p_el.find(qn('a:pPr'))
    if pPr is None:
        return None
    for ch in pPr:
        tag = etree.QName(ch).localname
        if tag in BU_TAGS:
            return True
        if tag == 'buNone':
            return False
    return None


def has_hanging(p_el, sh):
    """ぶら下げインデントがあるか。値が継承されるプレースホルダーはマスターの設定を信じる"""
    pPr = p_el.find(qn('a:pPr'))
    marL = pPr.get('marL') if pPr is not None else None
    indent = pPr.get('indent') if pPr is not None else None
    if marL is None and indent is None and getattr(sh, 'is_placeholder', False):
        return True
    marL, indent = int(marL or 0), int(indent or 0)
    return indent < 0 and marL >= -indent


def _lvl_sz(container, lvl):
    if container is None:
        return None
    p = container.find(qn(f'a:lvl{lvl + 1}pPr'))
    d = p.find(qn('a:defRPr')) if p is not None else None
    return int(d.get('sz')) if d is not None and d.get('sz') else None


def font_size(p_el, sh):
    """段落の文字サイズ（1/100 pt）。run → 段落末 → 図形・レイアウト・マスターの書式 の順に探す"""
    for r in p_el.iter(qn('a:rPr'), qn('a:endParaRPr')):
        if r.get('sz'):
            return int(r.get('sz'))
    lvl = level(p_el)
    chain, s = [], sh
    while s is not None:
        chain.append(s)
        s = getattr(s, '_base_placeholder', None) if getattr(s, 'is_placeholder', False) else None
    for s in chain:
        sz = _lvl_sz(s._element.find('.//' + qn('a:lstStyle')), lvl)
        if sz:
            return sz
    try:
        master = sh.part.slide_layout.slide_master
        styles = master._element.find(qn('p:txStyles'))
        if getattr(sh, 'is_placeholder', False):
            name = 'titleStyle' if sh.placeholder_format.type in (1, 3) else 'bodyStyle'
        else:
            name = 'otherStyle'
        sz = _lvl_sz(styles.find(qn(f'p:{name}')), lvl)
        if sz:
            return sz
    except Exception:
        pass
    return DEFAULT_SZ


def _insert_ordered(pPr, el):
    idx = PPR_ORDER.index(etree.QName(el).localname)
    for ch in pPr:
        tag = etree.QName(ch).localname
        if tag in PPR_ORDER and PPR_ORDER.index(tag) > idx:
            ch.addprevious(el)
            return
    pPr.append(el)


def set_hanging(p_el, sh, char=None, font=None):
    """ぶら下げインデントを設定する。char を渡すとその記号の箇条書きにする（記号の位置は元のまま）"""
    pPr = p_el.get_or_add_pPr()
    w = font_size(p_el, sh) * 127  # 1/100 pt → EMU。全角の記号1文字分
    marL = int(pPr.get('marL') or 0)
    indent = int(pPr.get('indent') or 0)
    bullet_x = marL + min(indent, 0)
    pPr.set('marL', str(bullet_x + w))
    pPr.set('indent', str(-w))
    if char:
        for ch in list(pPr):
            if etree.QName(ch).localname in ('buNone', 'buFontTx', 'buFont') + BU_TAGS:
                pPr.remove(ch)
        if font:
            f = etree.Element(qn('a:buFont'))
            f.set('typeface', font)
            _insert_ordered(pPr, f)
        c = etree.Element(qn('a:buChar'))
        c.set('char', char)
        _insert_ordered(pPr, c)


def _strip_leading(nodes, pattern):
    """先頭の run から pattern を取り除く（空になった run は消す）"""
    for r in nodes:
        if etree.QName(r).localname != 'r':
            return
        t = r.find(qn('a:t'))
        s = t.text or ''
        new = pattern.sub('', s, count=1) if pattern is TEXT_BULLET else pattern.sub('', s)
        t.text = new
        if new:
            return
        r.getparent().remove(r)
        if pattern is TEXT_BULLET and s.strip():
            return  # 記号だけの run を消した。続く run の空白は次で処理
    return


def _strip_after_breaks(p_el):
    """段落内改行の直後の空白（字下げのつもりの空白）を取り除く"""
    kids = list(p_el)
    for k, ch in enumerate(kids):
        if etree.QName(ch).localname == 'br':
            _strip_leading(kids[k + 1:], LEAD_SPACE)


def _content(p_el):
    return [ch for ch in p_el if etree.QName(ch).localname in ('r', 'br', 'fld')]


def _merge(p_el, q_el):
    """q の内容を段落内改行で p の後ろにつなげ、q を消す"""
    end = p_el.find(qn('a:endParaRPr'))
    runs = p_el.findall(qn('a:r'))
    br = etree.Element(qn('a:br'))
    if runs and runs[-1].find(qn('a:rPr')) is not None:
        br.append(copy.deepcopy(runs[-1].find(qn('a:rPr'))))
    moved = [br] + _content(q_el)
    for el in moved:
        if end is not None:
            end.addprevious(el)
        else:
            p_el.append(el)
    _strip_leading(moved[1:], LEAD_SPACE)
    q_el.getparent().remove(q_el)


def fix_bullets(tf, sh, font):
    """文字の記号を箇条書きに、続きの段落を段落内改行に、ぶら下げの無い箇条書きを直す。
    戻り値: (記号を直した数, ぶら下げを付けた数, つないだ段落の一覧[(前, 後)])"""
    converted, hanged, merged = 0, 0, []
    txBody = tf._txBody
    k = 0
    while k < len(txBody.findall(qn('a:p'))):
        p = txBody.findall(qn('a:p'))[k]
        k += 1
        text = para_text(p)
        char = text_bullet(text)
        if char:
            _strip_leading(_content(p), LEAD_SPACE)
            _strip_leading(_content(p), TEXT_BULLET)
            _strip_leading(_content(p), LEAD_SPACE)
            set_hanging(p, sh, char, font)
            converted += 1
            # 直後の記号の無い段落は、この項目の続きとみなしてつなぐ
            while k < len(txBody.findall(qn('a:p'))):
                q = txBody.findall(qn('a:p'))[k]
                qt = para_text(q)
                if not qt.strip() or text_bullet(qt) or explicit_bullet(q) or level(q) != level(p):
                    break
                merged.append((para_text(p).replace('\v', '⏎'), qt.strip()))
                _merge(p, q)
            _strip_after_breaks(p)
        elif explicit_bullet(p) and text.strip() and not has_hanging(p, sh):
            set_hanging(p, sh)
            _strip_after_breaks(p)
            hanged += 1
    return converted, hanged, merged


def bullet_issues(tf, sh):
    """lint 用: (段落番号, 指摘) を返す"""
    out = []
    prev_text_bullet = False
    for n, p in enumerate(tf._txBody.findall(qn('a:p')), 1):
        text = para_text(p)
        if not text.strip():
            prev_text_bullet = False
            continue
        if text_bullet(text):
            out.append((n, f'箇条書きの「{text_bullet(text)}」が文字（2行目以降が揃わない）'))
            prev_text_bullet = True
            continue
        if prev_text_bullet and not explicit_bullet(p):
            out.append((n, '箇条書きの続きが別の段落（段落内改行にする）'))
            continue
        prev_text_bullet = False
        if explicit_bullet(p) and not has_hanging(p, sh):
            out.append((n, 'ぶら下げインデントが無い（2行目以降が記号の下に来る）'))
        if re.search(r'\v[ \t　]+', text):
            out.append((n, '改行後を空白で字下げしている'))
    return out
