"""slide-text-polish 共通関数: スライド内の全テキスト（グループ・表・グラフ含む）を走査する。"""
import re
import unicodedata

from pptx import Presentation  # noqa: F401  (他スクリプトから再利用)
from pptx.enum.shapes import MSO_SHAPE_TYPE

JA = re.compile(r'[぀-ヿ一-鿿]')
WESTERN = {
    'Arial', 'Calibri', 'Helvetica', 'Segoe UI', 'Times New Roman', 'Verdana',
    'Tahoma', 'Roboto', 'Inter', 'Aptos', 'Calibri Light', 'Georgia',
}


def w(ch):
    """全角=1, 半角=0.5"""
    return 1 if unicodedata.east_asian_width(ch) in 'FWA' else 0.5


def width(text):
    return sum(w(c) for c in text)


def text_frames(prs):
    """(スライド番号, 場所, 種類, text_frame, 図形) を返す。種類: text / cell / chart"""

    def walk(shapes, i, path):
        for sh in shapes:
            name = f'{path}{sh.name}'
            if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                yield from walk(sh.shapes, i, name + '/')
            elif getattr(sh, 'has_table', False) and sh.has_table:
                for r, row in enumerate(sh.table.rows):
                    for c, cell in enumerate(row.cells):
                        yield i, f'{name}[{r},{c}]', 'cell', cell.text_frame, sh
            elif getattr(sh, 'has_chart', False) and sh.has_chart:
                ch = sh.chart
                if ch.has_title and ch.chart_title.has_text_frame:
                    yield i, f'{name}:title', 'chart', ch.chart_title.text_frame, sh
                for ax_name in ('category_axis', 'value_axis'):
                    try:
                        ax = getattr(ch, ax_name)
                        if ax.has_title and ax.axis_title.has_text_frame:
                            yield i, f'{name}:{ax_name}', 'chart', ax.axis_title.text_frame, sh
                    except Exception:
                        pass
            elif sh.has_text_frame:
                yield i, name, 'text', sh.text_frame, sh

    for i, slide in enumerate(prs.slides, 1):
        yield from walk(slide.shapes, i, '')


def chart_labels(prs):
    """グラフの項目名・系列名（軸ラベル・凡例になる文字）"""
    for i, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if getattr(sh, 'has_chart', False) and sh.has_chart:
                for plot in sh.chart.plots:
                    for cat in plot.categories:
                        yield i, sh.name, 'category', str(cat)
                    for s in plot.series:
                        yield i, sh.name, 'series', s.name or ''
