"""箇条書きの記号・ラベル付きの段落と、ぶら下げインデント。

AI製の資料では「・」を文字で打ち、改行を段落の区切り（Enter）や空白で揃えていることが多い。
この形だと2行目以降が記号の下（段落の左端）に来るため、本物の箇条書き記号＋ぶら下げインデント
（marL＝本文の開始位置、indent＝-記号の幅）に直し、続きの段落は段落内改行（<a:br/>）でつなぐ。
「共通の注意 本文」のように強調したラベルで始まる段落も、ラベルの後をタブにしてラベルの幅でぶら下げる
（タブは本文をぶら下げ位置まで送るので、ラベル幅の見積もりがずれても2行目と本文の先頭は揃う）。
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


def set_hanging(p_el, sh, char=None, font=None, w=None):
    """ぶら下げインデントを設定する。char を渡すとその記号の箇条書きにする（記号の位置は元のまま）。
    w はぶら下げの幅（EMU）。省略時は全角の記号1文字分"""
    pPr = p_el.get_or_add_pPr()
    if w is None:
        w = font_size(p_el, sh) * 127  # 1/100 pt → EMU
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


def _sig(r):
    """run の強調の書式（太字・斜体・下線・色）"""
    rPr = r.find(qn('a:rPr'))
    if rPr is None:
        return ('0', '0', 'none', b'')
    fill = rPr.find(qn('a:solidFill'))
    return (rPr.get('b', '0'), rPr.get('i', '0'), rPr.get('u', 'none'),
            etree.tostring(fill) if fill is not None else b'')


def _label_width(label):
    """ラベルの幅（全角=1）の見積もり。欧文はプロポーショナルなので大文字・数字を広めに"""
    n = 0.0
    for ch in label:
        if ord(ch) > 0xff:
            n += 1
        elif ch.isupper() or ch.isdigit():
            n += 0.65
        else:
            n += 0.55
    return n


def _rtext(r):
    return ''.join(t.text or '' for t in r.iter(qn('a:t')))


def label_split(p_el):
    """「共通の注意 本文」のように、強調したラベル（太字や色違い）の後に本文が続く段落なら
    (ラベルの run の一覧, 本文の最初の run, ラベル文字列) を返す。違えば None"""
    kids = _content(p_el)
    if not kids or etree.QName(kids[0]).localname != 'r':
        return None
    first = _sig(kids[0])
    k = 0
    while k < len(kids) and etree.QName(kids[k]).localname == 'r' and _sig(kids[k]) == first:
        k += 1
    if k >= len(kids) or etree.QName(kids[k]).localname != 'r':
        return None
    body_sig = _sig(kids[k])
    if not ((first[0] == '1' and body_sig[0] != '1') or first[3] != body_sig[3]):
        return None
    raw = ''.join(_rtext(r) for r in kids[:k])
    label = raw.strip(' \t　')
    body = ''.join(_rtext(r) if etree.QName(r).localname == 'r' else '\v' for r in kids[k:])
    if not label or _label_width(label) > 8 or not body.strip():
        return None
    if not (raw[-1:] in (' ', '\t', '　') or body[:1] in (' ', '\t', '　') or label[-1] in '：:'):
        return None
    return kids[:k], kids[k], label


def set_label_hanging(p_el, sh):
    """ラベルの後をタブにし、ラベルの幅でぶら下げる（2行目以降が本文の先頭に揃う）"""
    labels, body, label = label_split(p_el)
    for r in reversed(labels):  # ラベル末尾の空白を消す
        t = r.find(qn('a:t'))
        t.text = (t.text or '').rstrip(' \t　')
        if t.text:
            break
    t = body.find(qn('a:t'))
    t.text = '\t' + (t.text or '').lstrip(' \t　')
    rPr = labels[0].find(qn('a:rPr'))
    sz = int(rPr.get('sz')) if rPr is not None and rPr.get('sz') else font_size(p_el, sh)
    bold = 1.05 if rPr is not None and rPr.get('b') == '1' else 1.0
    w = int((_label_width(label) * bold + 0.5) * sz * 127)  # 0.5字分は本文との間隔
    set_hanging(p_el, sh, w=w)


def fix_bullets(tf, sh, font):
    """文字の記号を箇条書きに、ラベル付きの段落をラベル幅のぶら下げに、続きの段落を段落内改行に、
    ぶら下げの無い箇条書きを直す。
    戻り値: (記号を直した数, ぶら下げを付けた数, つないだ段落の一覧[(前, 後)])"""
    converted, hanged, merged = 0, 0, []

    def paras():
        return tf._txBody.findall(qn('a:p'))

    k = 0
    while k < len(paras()):
        p = paras()[k]
        k += 1
        text = para_text(p)
        char = text_bullet(text)
        if char:
            _strip_leading(_content(p), LEAD_SPACE)
            _strip_leading(_content(p), TEXT_BULLET)
            _strip_leading(_content(p), LEAD_SPACE)
            set_hanging(p, sh, char, font)
            converted += 1
        elif not explicit_bullet(p) and label_split(p):
            if label_without_hanging(p) or not _rtext(label_split(p)[1]).startswith('\t'):
                set_label_hanging(p, sh)
                hanged += 1
        else:
            if explicit_bullet(p) and text.strip() and not has_hanging(p, sh):
                set_hanging(p, sh)
                _strip_after_breaks(p)
                hanged += 1
            continue
        # 直後の、記号もラベルも無い段落は、この項目の続きとみなしてつなぐ
        while k < len(paras()):
            q = paras()[k]
            qt = para_text(q)
            if (not qt.strip() or text_bullet(qt) or explicit_bullet(q) or label_split(q)
                    or level(q) != level(p)):
                break
            merged.append((para_text(p).replace('\v', '⏎').replace('\t', ' '), qt.strip()))
            _merge(p, q)
        _strip_after_breaks(p)
    return converted, hanged, merged


def label_without_hanging(p_el):
    """ラベル付きの段落で、ぶら下げが無いもの（折り返すと2行目がラベルの下に来る）"""
    if explicit_bullet(p_el) or not label_split(p_el):
        return False
    pPr = p_el.find(qn('a:pPr'))
    return int(pPr.get('indent') or 0) >= 0 if pPr is not None else True


def bullet_issues(tf, sh):
    """lint 用: (段落番号, 指摘) を返す"""
    out = []
    prev = None  # 直前の段落が「文字の記号の箇条書き」「ラベル付きの段落」なら、その種類
    for n, p in enumerate(tf._txBody.findall(qn('a:p')), 1):
        text = para_text(p)
        if not text.strip():
            prev = None
            continue
        if text_bullet(text):
            out.append((n, f'箇条書きの「{text_bullet(text)}」が文字（2行目以降が揃わない）'))
            prev = '箇条書き'
            continue
        if not explicit_bullet(p) and label_split(p):
            if label_without_hanging(p) and '\v' in text:
                out.append((n, 'ラベル付きの段落にぶら下げが無い（2行目がラベルの下に来る）'))
            prev = 'ラベル付きの段落'
            continue
        if prev and not explicit_bullet(p):
            out.append((n, f'{prev}の続きが別の段落（段落内改行にする）'))
            continue
        prev = None
        if explicit_bullet(p) and not has_hanging(p, sh):
            out.append((n, 'ぶら下げインデントが無い（2行目以降が記号の下に来る）'))
        if re.search(r'\v[ \t　]+', text):
            out.append((n, '改行後を空白で字下げしている'))
    return out

