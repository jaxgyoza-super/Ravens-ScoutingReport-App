"""
chart_generator.py
フィールド型ヒートマップ生成モジュール
"""
import io
import warnings

import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.patches as mpatches
import matplotlib.font_manager as fm
import matplotlib.colors as mcolors
import matplotlib.cm as mcm
from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg

import os as _os

# ── 日本語フォント設定 ─────────────────────────────────────────
# japanize_matplotlib の TTF パスを直接取得して FontProperties を構築。
# rcParams も同時に設定するため set_title / suptitle も自動で日本語になる。
_JP_FONT = None   # ax.text / tick labels 用 FontProperties

def _setup_jp_font():
    global _JP_FONT
    # 1. japanize_matplotlib のフォントファイルを探してパス直指定
    try:
        import japanize_matplotlib as _jm
        import os as _o
        _fonts_dir = _o.join(_o.dirname(_jm.__file__), 'fonts')
        _ttf = next(
            (_o.join(_fonts_dir, f) for f in sorted(_o.listdir(_fonts_dir))
             if f.endswith('.ttf')),
            None,
        )
        if _ttf:
            fm.fontManager.addfont(_ttf)
            _prop = fm.FontProperties(fname=_ttf)
            _name = _prop.get_name()
            matplotlib.rcParams['font.family'] = [_name, 'sans-serif']
            _JP_FONT = _prop
            return
    except Exception:
        pass
    # 2. Windows 環境フォールバック
    _available = {f.name for f in fm.fontManager.ttflist}
    for _name in ('Yu Gothic', 'Meiryo', 'MS Gothic', 'Hiragino Sans'):
        if _name in _available:
            matplotlib.rcParams['font.family'] = [_name, 'sans-serif']
            _JP_FONT = fm.FontProperties(family=_name)
            return
    _JP_FONT = fm.FontProperties()

_setup_jp_font()

def _ensure_font():
    """後方互換のため残す"""
    pass


# ── カラー定数 ─────────────────────────────────────────────────
FIELD_BG   = '#ffffff'
CELL_EMPTY = '#dddddd'

# セル幅（COVERAGE/PRESSURE は狭め、SIGN(D)/BLITZ は広め）
CW_NARROW = 1.80
CW_WIDE   = 3.20   # SIGN(D)/BLITZ: SOLDIERなど長い値に対応
CW_MAP = {
    'coverage': CW_NARROW,
    'pressure': CW_NARROW,
    'sign_d':   CW_WIDE,
    'blitz':    CW_WIDE,
}
METRIC_TITLES = {
    'coverage': 'COVERAGE',
    'pressure': 'ラッシュ人数',
    'sign_d':   'Blitz詳細',
    'blitz':    'Blitz種類',
}


def _heat_color(ratio: float) -> str:
    """
    0〜20%  : 白 (#ffffff)
    20〜100%: 白 → 濃い赤 (#990000) のグラデーション
    """
    if ratio <= 0.20:
        return '#ffffff'
    t = (ratio - 0.20) / 0.80
    t = max(0.0, min(1.0, t))
    r = int(0xFF + t * (0x99 - 0xFF))
    g = int(0xFF + t * (0x00 - 0xFF))
    b = int(0xFF + t * (0x00 - 0xFF))
    return f'#{r:02x}{g:02x}{b:02x}'


def _build_heatmap_cmap():
    """_heat_color と同じグラデーションの ListedColormap を生成"""
    c = []
    for i in range(256):
        ratio = i / 255.0
        if ratio <= 0.20:
            c.append((1.0, 1.0, 1.0, 1.0))
        else:
            t = (ratio - 0.20) / 0.80
            c.append((1.0 + t * (0x99 / 255 - 1.0),
                       1.0 - t,
                       1.0 - t,
                       1.0))
    return mcolors.ListedColormap(c)

_HEATMAP_CMAP = _build_heatmap_cmap()


def _text_color(bg_hex: str) -> str:
    r = int(bg_hex[1:3], 16)
    g = int(bg_hex[3:5], 16)
    b = int(bg_hex[5:7], 16)
    return '#111111' if 0.299 * r + 0.587 * g + 0.114 * b > 100 else '#ffffff'


def _wrap_val(text: str) -> str:
    """8文字超の場合、+ で折り返して複数行に"""
    if len(text) <= 8 or '+' not in text:
        return text
    parts = text.split('+')
    lines, cur = [], ''
    for p in parts:
        candidate = cur + ('+' if cur else '') + p
        if len(candidate) > 9 and cur:
            lines.append(cur)
            cur = p
        else:
            cur = candidate
    if cur:
        lines.append(cur)
    return '\n'.join(lines)


def _val_fontsize(text: str) -> int:
    """折り返し後の最長行文字数に応じてフォントサイズを決定"""
    max_len = max(len(line) for line in text.split('\n'))
    if max_len <= 2:  return 48
    if max_len <= 4:  return 42
    if max_len <= 6:  return 36
    if max_len <= 8:  return 33
    if max_len <= 10: return 27
    if max_len <= 13: return 21
    return 17


def generate_field_heatmap(grid_data: dict, situation_label: str = '', show_colorbar: bool = False,
                           show_zone_shading: bool = True, redzone_mode: bool = False):
    """
    COVERAGE / PRESSURE / SIGN(D) / BLITZ の4指標を 2×2 で並べた
    フィールド型ヒートマップを生成し (PNG BytesIO, tie_notes) を返す。

    tie_notes: list of (note_id, panel_title, col_label, row_label, main_val, tied_vals)
               tied_vals = [(val_str, count), ...]

    plt.subplots ではなく Figure を直接使用することでグローバル pyplot
    状態を回避し、ThreadPoolExecutor での並列呼び出しを安全にする。
    """
    _ensure_font()
    col_labels = grid_data['col_labels']
    row_labels  = grid_data['row_labels']
    cells       = grid_data['cells']

    nc = len(col_labels)
    nr = len(row_labels)

    # ── レイアウト定数 ─────────────────────────────────────────
    CH  = 3.00
    RZH = 1.60
    EZH = RZH  # 自陣Endzoneの高さをRedzoneと統一
    PAD = 0.10

    metrics = [
        ('coverage', 'cov_n',    'cov_others',   'cov_allvals'),
        ('pressure', 'press_n',  'press_others',  'press_allvals'),
        ('sign_d',   'sign_d_n', 'sign_others',   'sign_allvals'),
        ('blitz',    'blitz_n',  'blitz_others',  'blitz_allvals'),
    ]

    HALFWAY_Y = (nr - 1) * CH

    def _fmt_row(label):
        if '（' in label:
            idx = label.index('（')
            return label[:idx] + '\n' + label[idx:]
        return label

    row_labels_disp = [_fmt_row(row_labels[nr - 1 - i]) for i in range(nr)]

    max_panel_w = max(nc * CW_MAP[mk] for mk, *_ in metrics)
    panel_h = nr * CH + RZH + EZH

    # plt.subplots の代わりに Figure を直接生成（スレッドセーフ）
    fig = Figure(
        figsize=(max_panel_w * 2 + 9.0, panel_h * 2 + 3.0),
        dpi=50,
    )
    FigureCanvasAgg(fig)  # バックエンドをアタッチ（canvas.draw() は呼ばない）
    fig.patch.set_facecolor(FIELD_BG)

    axes = fig.subplots(
        2, 2,
        gridspec_kw={'wspace': 0.55, 'hspace': 0.65}
    )

    # タイ注釈収集
    tie_notes = []   # (note_id, panel_title, col_label, row_label, main_val, tied_vals)
    note_counter = [0]

    for ax, (metric_key, metric_n_key, metric_others_key, metric_allvals_key) in zip(axes.flatten(), metrics):
        CW = CW_MAP[metric_key]
        ax.set_facecolor(FIELD_BG)
        y_red = nr * CH

        # ── RED ZONE / 敵陣Endzone ────────────────────────────────
        if redzone_mode:
            # 敵陣Endzone：自陣Endzoneと同じ青で網掛け
            for ci in range(nc):
                ax.add_patch(mpatches.Rectangle(
                    (ci * CW, y_red), CW, RZH,
                    lw=1.2, ec='#0000cc', fc='#dde8ff', hatch='\\\\\\',
                    alpha=0.9, zorder=2))
            ax.text(nc * CW / 2, y_red + RZH / 2, '敵陣Endzone',
                    ha='center', va='center', fontsize=50,
                    color='#0000cc', fontweight='bold', zorder=3,
                    fontproperties=_JP_FONT)
        else:
            # 通常モード：Redzone（緑）
            for ci in range(nc):
                ax.add_patch(mpatches.Rectangle(
                    (ci * CW, y_red), CW, RZH,
                    lw=1.2, ec='#006600',
                    fc='#ddffdd' if show_zone_shading else 'none',
                    hatch='///' if show_zone_shading else None,
                    alpha=0.9, zorder=2))
            ax.text(nc * CW / 2, y_red + RZH / 2, 'Redzone',
                    ha='center', va='center', fontsize=50,
                    color='#006600', fontweight='bold', zorder=3,
                    fontproperties=_JP_FONT)

        # ── データセル ─────────────────────────────────────────
        for ri_disp in range(nr):
            ri = nr - 1 - ri_disp
            y  = (nr - 1 - ri_disp) * CH
            for ci in range(nc):
                cell  = cells.get((ri, ci), {'n': 0})
                n     = cell.get('n', 0)
                val_n = cell.get(metric_n_key, 0) if n > 0 else 0
                ratio = val_n / n if n > 0 else 0.0
                bg    = _heat_color(ratio) if n > 0 else CELL_EMPTY
                tc    = _text_color(bg)

                ax.add_patch(mpatches.FancyBboxPatch(
                    (ci * CW + PAD, y + PAD),
                    CW - 2 * PAD, CH - 2 * PAD,
                    boxstyle='round,pad=0.04',
                    lw=1.0, ec='#bbbbbb', fc=bg, zorder=2))

                if n > 0:
                    val      = cell.get(metric_key, '-')
                    others   = cell.get(metric_others_key, [])
                    allvals  = cell.get(metric_allvals_key, [])

                    if allvals:
                        # ── 縦並び表示（セル合計n≤2プレー）※なし ──────
                        n_items = len(allvals)
                        # フォントサイズ：項目数に応じて縮小
                        fs_stack = max(24, 44 - (n_items - 1) * 6)
                        # Y位置：下部に「(全Nプレー)」用のスペースを確保
                        bottom_label_h = CH * 0.18
                        margin = CH * 0.10
                        usable = CH - margin - bottom_label_h
                        for i, (av, ac) in enumerate(allvals):
                            frac = (i + 0.5) / n_items
                            yp = y + bottom_label_h + usable * (1.0 - frac)
                            wrapped_av = _wrap_val(str(av))
                            ax.text(ci * CW + CW / 2, yp, wrapped_av,
                                    ha='center', va='center',
                                    fontsize=fs_stack, fontweight='bold',
                                    color=tc, zorder=3, linespacing=1.2,
                                    fontproperties=_JP_FONT)
                        # 下部に合計プレー数を表示
                        ax.text(ci * CW + CW / 2, y + CH * 0.08,
                                f'(全{n}プレー)',
                                ha='center', va='center',
                                fontsize=28, color=tc, zorder=3,
                                fontproperties=_JP_FONT)
                    else:
                        # ── 通常表示 ────────────────────────────────────
                        wrapped = _wrap_val(str(val))
                        fs      = _val_fontsize(wrapped)
                        n_lines = wrapped.count('\n') + 1

                        # タイがある場合は ※N を割り当て
                        has_others = bool(others)
                        if has_others:
                            note_counter[0] += 1
                            nid = note_counter[0]
                            tie_notes.append((
                                nid,
                                METRIC_TITLES[metric_key],
                                col_labels[ci].replace('\n', ' '),
                                row_labels[ri],
                                str(val),
                                val_n,
                                n,
                                others,
                            ))

                        top_y = y + CH * (0.72 if n_lines >= 3 else 0.65 if n_lines == 2 else 0.63)
                        cnt_y = y + CH * 0.18

                        ax.text(ci * CW + CW / 2, top_y, wrapped,
                                ha='center', va='center',
                                fontsize=fs, fontweight='bold', color=tc,
                                zorder=3, linespacing=1.3,
                                fontproperties=_JP_FONT)
                        ax.text(ci * CW + CW / 2, cnt_y,
                                f'({val_n}/{n})',
                                ha='center', va='center',
                                fontsize=36, color=tc, zorder=3,
                                fontproperties=_JP_FONT)

                        # タイ：※N を右下に小字表示
                        if has_others:
                            ax.text(ci * CW + CW - PAD * 2, y + PAD * 2,
                                    f'※{nid}',
                                    ha='right', va='bottom',
                                    fontsize=38, color=tc, alpha=0.90,
                                    zorder=3, fontproperties=_JP_FONT)
                else:
                    ax.text(ci * CW + CW / 2, y + CH / 2, '─',
                            ha='center', va='center',
                            fontsize=48, color='#aaaaaa', zorder=3,
                            fontproperties=_JP_FONT)

        # ── 自陣 END ZONE（redzone_mode では非表示）─────────────
        if redzone_mode:
            y_end = 0  # 下端を 0 に
        else:
            y_end = -EZH
            for ci in range(nc):
                ax.add_patch(mpatches.Rectangle(
                    (ci * CW, y_end), CW, EZH,
                    lw=1.2, ec='#006600',
                    fc='#ddffdd' if show_zone_shading else 'none',
                    hatch='///' if show_zone_shading else None,
                    alpha=0.9, zorder=2))
            ax.text(nc * CW / 2, y_end + EZH / 2, '自陣Endzone',
                    ha='center', va='center', fontsize=50,
                    color='#006600', fontweight='bold', zorder=3,
                    fontproperties=_JP_FONT)

        # ── 区切り線 ───────────────────────────────────────────
        for i in range(nr + 1):
            ax.axhline(y=i * CH, color='#aaaaaa', lw=1.0, alpha=0.8, zorder=3)

        # ── ハーフライン（redzone_mode では非表示）─────────────
        if not redzone_mode:
            ax.axhline(y=HALFWAY_Y, color='#000000', lw=3.5, zorder=5)
            ax.text(nc * CW + PAD * 2, HALFWAY_Y, '← ハーフ',
                    ha='left', va='center', fontsize=41,
                    color='#000000', fontweight='bold', zorder=6,
                    clip_on=False, fontproperties=_JP_FONT)

        # ── 軸設定 ─────────────────────────────────────────────
        ax.set_xlim(0, nc * CW)
        ax.set_ylim(y_end, y_red + RZH)

        ax.set_xticks([(ci + 0.5) * CW for ci in range(nc)])
        ax.set_xticklabels(col_labels, multialignment='center')
        for _tl in ax.get_xticklabels():
            _tl.set_fontproperties(_JP_FONT)
            _tl.set_fontsize(42)
            _tl.set_color('#222222')
        ax.xaxis.set_label_position('top')
        ax.xaxis.tick_top()
        ax.tick_params(axis='x', colors='#222222', length=0, pad=6)

        ax.set_yticks([(nr - 0.5 - i) * CH for i in range(nr)])
        ax.set_yticklabels(row_labels_disp, multialignment='center')
        for _tl in ax.get_yticklabels():
            _tl.set_fontproperties(_JP_FONT)
            _tl.set_fontsize(39)
            _tl.set_color('#222222')
        ax.tick_params(axis='y', colors='#222222', length=0, pad=6)

        for spine in ax.spines.values():
            spine.set_edgecolor('#aaaaaa')
            spine.set_linewidth(1.0)

        ax.set_title(METRIC_TITLES[metric_key], fontsize=52,
                     fontweight='bold', color='#111111', pad=36)

    fig.suptitle(situation_label, fontsize=54, fontweight='bold',
                 color='#111111', y=0.99)

    # tight_layout の代わりに固定余白（canvas.draw() 不要）
    fig.subplots_adjust(left=0.10, right=0.92, top=0.87, bottom=0.05,
                        wspace=0.65, hspace=0.80)

    buf = io.BytesIO()
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        fig.savefig(buf, format='png', facecolor=FIELD_BG, dpi=50)
    buf.seek(0)
    return buf, tie_notes
