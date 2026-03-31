import pandas as pd
import re

# ── DISTゾーン定義 ──────────────────────────────────────────
DIST_ZONES = [
    ('① Short (1〜3)',        1,   3),
    ('② Medium-Short (4〜6)', 4,   6),
    ('③ Medium-Long (7〜10)', 7,  10),
    ('④ Long (11〜)',        11, 999),
]

# ── OFFフォームグルーピング ─────────────────────────────────
SLOT_FORMS  = {'SLOT A', 'SLOT B', 'SLOT VEER'}
BUNCH_FORMS = {'BUNCH', 'NEAR BUNCH', 'TIGHT BUNCH'}
PRO_FORMS   = {'PRO I', 'PRO B', 'PRO A'}
TWIN_FORMS  = {'TWIN A', 'TWIN B', 'TWIN I'}

DEFINED_GROUPS_ORDER = [
    'ACE', 'SUPER', 'WAC', 'WAT', 'SPREAD', 'DISH',
    'SLOT（A/B/VEER合算）', 'EMPTY',
    'BUNCH合算', 'PRO合算（I/B/A）', 'PRO FAR', 'PRO NEAR',
    'TWIN合算（A/B/I）', 'TWIN FAR', 'TWIN NEAR', 'UMB系（合算）',
]


def _get_form_group(form):
    f = str(form).strip()
    if f in ('ACE',):       return 'ACE'
    if f in ('SUPER',):     return 'SUPER'
    if f == 'WAC':          return 'WAC'
    if f == 'WAT':          return 'WAT'
    if f in ('SPREAD',):    return 'SPREAD'
    if f in ('DISH',):      return 'DISH'
    if f in SLOT_FORMS:     return 'SLOT（A/B/VEER合算）'
    if 'EMPTY' in f:        return 'EMPTY'
    if f in BUNCH_FORMS:    return 'BUNCH合算'
    if f in PRO_FORMS:      return 'PRO合算（I/B/A）'
    if f == 'PRO FAR':      return 'PRO FAR'
    if f == 'PRO NEAR':     return 'PRO NEAR'
    if f in TWIN_FORMS:     return 'TWIN合算（A/B/I）'
    if f == 'TWIN FAR':     return 'TWIN FAR'
    if f == 'TWIN NEAR':    return 'TWIN NEAR'
    if f.startswith('UMB'): return 'UMB系（合算）'
    return None  # 未分類


# ── ヘルパー ────────────────────────────────────────────────
def _notes(sub_df):
    parts = []
    for _, row in sub_df.iterrows():
        try:
            play = int(row['PLAY #'])
        except (ValueError, TypeError):
            play = row['PLAY #']
        parts.append(f"{row['OPPONENT']} #{play}")
    return ', '.join(parts)


def _pct_str(cnt, total):
    pct = cnt / total * 100 if total > 0 else 0
    return f"{pct:.0f}% ({cnt})"


def _extract_count(pct_str):
    m = re.search(r'\((\d+)\)', str(pct_str))
    return int(m.group(1)) if m else 0


# ── フィルタ関数 ─────────────────────────────────────────────
def filter_normal(df):
    """DN=1,2 / 2MIN除外 / RedZone除外"""
    return df[
        df['DN'].isin([1, 2]) &
        ~df['IS_2MIN'] &
        ~df['IS_REDZONE']
    ].copy()


def filter_3rd(df):
    """DN=3,4 / 2MIN除外 / RedZone除外"""
    return df[
        df['DN'].isin([3, 4]) &
        ~df['IS_2MIN'] &
        ~df['IS_REDZONE']
    ].copy()


def filter_redzone(df):
    """YARD LN 1〜25 / 2MIN除外"""
    return df[
        df['IS_REDZONE'] &
        ~df['IS_2MIN']
    ].copy()


def filter_2min(df):
    """2MIN = Y"""
    return df[df['IS_2MIN']].copy()


# ── 分析関数 ────────────────────────────────────────────────
def analyze_front(df):
    """DEF FRONT割合テーブル"""
    total = len(df)
    rows = []
    if total == 0:
        return pd.DataFrame(columns=['DEF FRONT', '割合（実数）', '備考'])
    counts = df['DEF_FRONT'].value_counts()
    for cat, cnt in counts.items():
        if not cat or cat == 'nan':
            continue
        sub = df[df['DEF_FRONT'] == cat]
        rows.append({
            'DEF FRONT': cat,
            '割合（実数）': _pct_str(cnt, total),
            '備考': _notes(sub) if cnt <= 5 else '',
        })
    return pd.DataFrame(rows)


def analyze_sign_d(df):
    """SIGN(D)割合テーブル"""
    total = len(df)
    rows = []
    if total == 0:
        return pd.DataFrame(columns=['SIGN(D)', '割合（実数）', '備考'])
    counts = df['SIGN_D'].value_counts()
    for cat, cnt in counts.items():
        if not cat or cat == 'nan':
            continue
        sub = df[df['SIGN_D'] == cat]
        rows.append({
            'SIGN(D)': cat,
            '割合（実数）': _pct_str(cnt, total),
            '備考': _notes(sub) if cnt <= 5 else '',
        })
    return pd.DataFrame(rows)


def analyze_coverage(df):
    """COVERAGE割合テーブル + Cover3内訳テーブルを返す"""
    total = len(df)
    rows = []
    if total == 0:
        return (
            pd.DataFrame(columns=['COVERAGE', '割合（実数）', '備考']),
            pd.DataFrame(columns=['COMPONENT', '割合（実数）', '備考']),
        )
    counts = df['COVERAGE_NORM'].value_counts()
    for cat, cnt in counts.items():
        if not cat or cat == 'nan':
            continue
        sub = df[df['COVERAGE_NORM'] == cat]
        rows.append({
            'COVERAGE': cat,
            '割合（実数）': _pct_str(cnt, total),
            '備考': _notes(sub) if cnt <= 5 else '',
        })
    main_tbl = pd.DataFrame(rows)

    # Cover 3 内訳
    cover3 = df[df['COVERAGE_NORM'] == '3']
    if len(cover3) > 0:
        comp_rows = []
        c3_total = len(cover3)
        comp_counts = cover3['COMPONENT_DISPLAY'].value_counts()
        for cat, cnt in comp_counts.items():
            if not cat or cat == 'nan':
                continue
            sub = cover3[cover3['COMPONENT_DISPLAY'] == cat]
            comp_rows.append({
                'COMPONENT': cat,
                '割合（実数）': _pct_str(cnt, c3_total),
                '備考': _notes(sub) if cnt <= 5 else '',
            })
        comp_tbl = pd.DataFrame(comp_rows)
    else:
        comp_tbl = pd.DataFrame(columns=['COMPONENT', '割合（実数）', '備考'])

    return main_tbl, comp_tbl


def analyze_off_form(df):
    """OFF FORM グルーピング割合テーブル"""
    total = len(df)
    if total == 0:
        return pd.DataFrame(columns=['フォーメーション', '割合（実数）', '備考'])

    df = df.copy()
    df['_GROUP'] = df['OFF_FORM_NORM'].apply(_get_form_group)

    rows = []
    # 定義済みグループ（順序保持、後でソート）
    for grp in DEFINED_GROUPS_ORDER:
        sub = df[df['_GROUP'] == grp]
        cnt = len(sub)
        if cnt == 0:
            continue
        rows.append({
            'フォーメーション': grp,
            '割合（実数）': _pct_str(cnt, total),
            '備考': _notes(sub) if cnt <= 5 else '',
        })

    # 未分類で3プレー以上
    unc = df[df['_GROUP'].isna()]
    for form, sub in unc.groupby('OFF_FORM_NORM'):
        cnt = len(sub)
        if cnt >= 3:
            rows.append({
                'フォーメーション': f'{form}（その他）',
                '割合（実数）': _pct_str(cnt, total),
                '備考': _notes(sub) if cnt <= 5 else '',
            })

    rows.sort(key=lambda r: _extract_count(r['割合（実数）']), reverse=True)
    return pd.DataFrame(rows)


def analyze_packages(df):
    """DEF FRONT + SIGN(D) + COVERAGE の組み合わせが3回以上のパッケージ"""
    if len(df) == 0:
        return pd.DataFrame(columns=['DEF FRONT', 'SIGN(D)', 'COVERAGE', '出現数', '備考（大学・PLAY#）'])

    df = df.copy()
    df['_PKG'] = (
        df['DEF_FRONT'].fillna('') + ' | ' +
        df['SIGN_D'].fillna('') + ' | ' +
        df['COVERAGE_NORM'].fillna('')
    )
    counts = df['_PKG'].value_counts()
    rows = []
    for pkg, cnt in counts.items():
        if cnt < 3:
            continue
        sub = df[df['_PKG'] == pkg]
        parts = pkg.split(' | ')
        rows.append({
            'DEF FRONT': parts[0] if len(parts) > 0 else '',
            'SIGN(D)':   parts[1] if len(parts) > 1 else '',
            'COVERAGE':  parts[2] if len(parts) > 2 else '',
            '出現数': cnt,
            '備考（大学・PLAY#）': _notes(sub),
        })
    return pd.DataFrame(rows)


def analyze_3rd_zones(df):
    """DISTゾーン別に分析結果をまとめて返す"""
    results = {}
    for zone_name, dist_min, dist_max in DIST_ZONES:
        sub = df[(df['DIST'] >= dist_min) & (df['DIST'] <= dist_max)].copy()
        cov_tbl, comp3_tbl = analyze_coverage(sub)
        results[zone_name] = {
            'n':        len(sub),
            'front':    analyze_front(sub),
            'coverage': cov_tbl,
            'comp3':    comp3_tbl,
            'packages': analyze_packages(sub),
        }
    return results
