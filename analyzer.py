import pandas as pd
import re

# ── SIGN / BLITZ 表示名変換 ──────────────────────────────────
_SIGN_EXACT = {
    'FO':  '広RIP',
    'N':   'read',
    'MB':  'LB中ブリッツ',
    'CS':      'クロス',
    'BOTH.A':  '両A-LB Blitz',
    'OSC': '外クロス',
    'OCS': '外クロス',
    'ISC': '中クロス',
    'ICS': '中クロス',
    'TO':  'Tアウチャ',
    'EI':  'Eインチャ',
}
# F/B + A/B gap + LB種別 パターン
# 例: F.A-M → 広A-Mike Blitz、B.B-FS → 狭B-FS Blitz
_FB_GAP_RE  = re.compile(r'^(F|B)\.(A|B)-(M|W|FS|S)$')
_LB_NAMES   = {'M': 'Mike', 'W': 'Willy', 'FS': 'FS', 'S': 'Sam'}
_MB_NUM_RE  = re.compile(r'^MB(\d+)$')

def _sign_token(t: str) -> str:
    """SIGN/BLITZ の単一トークンを表示名に変換"""
    t = t.strip()
    if t in _SIGN_EXACT:
        return _SIGN_EXACT[t]
    # F/B・A/B gap パターン
    m = _FB_GAP_RE.match(t)
    if m:
        width = '広' if m.group(1) == 'F' else '狭'
        gap   = m.group(2)           # A or B
        lb    = _LB_NAMES[m.group(3)]
        return f'{width}{gap}-{lb} Blitz'
    # MB + 数字
    m = _MB_NUM_RE.match(t)
    if m:
        return f'LB中ブリッツ{m.group(1)}枚'
    if t.upper().startswith('BOTH.'):
        return '両' + _sign_token(t[5:])
    return t

def _hm_sign_display(v: str) -> str:
    """ヒートマップ専用：sign_display の結果から末尾の ' Blitz' を除去"""
    d = sign_display(v)
    return d[:-6] if d.endswith(' Blitz') else d


def sign_display(v: str) -> str:
    """SIGN_D / BLITZ 値を表示名に変換（複合値は + で結合）"""
    if not v or str(v).lower() in ('nan', ''):
        return v
    v = str(v).strip()
    # + または , で複合トークンに分割
    sep = '+' if '+' in v else (',' if ',' in v else None)
    if sep:
        parts = [_sign_token(p.strip()) for p in v.split(sep)]
        return (' + ' if sep == '+' else ', ').join(parts)
    return _sign_token(v)

# ── YARD LNゾーン定義 ─────────────────────────────────────
YARD_LN_ZONES = [
    ('①自陣深く（-1〜-9yds）',    -9,  -1),
    ('②自陣（-10〜-49yds）',      -49, -10),
    ('③敵陣（26〜50yds）',         26,  50),
]

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


# ── PERSONNELグループ定義 ────────────────────────────────────
PERSONNEL_GROUPS = [
    ('①10',        ['10']),
    ('②11',        ['11']),
    ('③12/21/22',  ['12', '21', '22']),
]


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


def filter_personnel(df, personnel_values: list):
    """指定PERSONNELグループに絞り込む"""
    return df[df['PERSONNEL_NORM'].isin(personnel_values)].copy()


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
            '備考': _notes(sub) if cnt <= 3 else '',
        })
    return pd.DataFrame(rows)


def analyze_sign_d(df, min_count: int = 3):
    """SIGN(D)割合テーブル。min_count未満のエントリは除外する（デフォルト3プレー未満除外）"""
    total = len(df)
    rows = []
    if total == 0:
        return pd.DataFrame(columns=['SIGN', '割合（実数）', '備考'])
    counts = df['SIGN_D'].value_counts()
    for cat, cnt in counts.items():
        if not cat or cat == 'nan':
            continue
        if cnt < min_count:
            continue
        sub = df[df['SIGN_D'] == cat]
        rows.append({
            'SIGN': sign_display(cat),
            '割合（実数）': _pct_str(cnt, total),
            '備考': _notes(sub) if cnt <= 3 else '',
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
        # Cover3は内訳行（COMPONENT）に備考を記載するので、ここは空欄
        notes = '' if cat == '3' else (_notes(sub) if cnt <= 3 else '')
        rows.append({
            'COVERAGE': cat,
            '割合（実数）': _pct_str(cnt, total),
            '備考': notes,
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
                '備考': _notes(sub) if cnt <= 3 else '',
            })
        comp_tbl = pd.DataFrame(comp_rows)
    else:
        comp_tbl = pd.DataFrame(columns=['COMPONENT', '割合（実数）', '備考'])

    return main_tbl, comp_tbl


def analyze_off_form_coverage(df):
    """OFF FORMグループごとにCOVERAGE割合をまとめる。
    戻り値: [(group_name, n, cov_tbl, comp3_tbl), ...]  プレー数降順
    """
    if len(df) == 0:
        return []

    df = df.copy()
    df['_GROUP'] = df['OFF_FORM_NORM'].apply(_get_form_group)

    results = []

    # 定義済みグループ
    for grp in DEFINED_GROUPS_ORDER:
        sub = df[df['_GROUP'] == grp]
        n = len(sub)
        if n == 0:
            continue
        cov_tbl, comp3_tbl = analyze_coverage(sub)
        results.append((grp, n, cov_tbl, comp3_tbl))

    # 未分類で3プレー以上
    unc = df[df['_GROUP'].isna()]
    for form, sub in unc.groupby('OFF_FORM_NORM'):
        n = len(sub)
        if n >= 3:
            cov_tbl, comp3_tbl = analyze_coverage(sub)
            results.append((f'{form}（その他）', n, cov_tbl, comp3_tbl))

    # 順番は定義済みグループ順 → その他（ソートしない）
    return results


def analyze_off_form_sign_d(df):
    """OFF FORMグループごとにSIGN(D)割合をまとめる。
    戻り値: [(group_name, n, sign_d_tbl), ...]  定義済み順
    """
    if len(df) == 0:
        return []

    df = df.copy()
    df['_GROUP'] = df['OFF_FORM_NORM'].apply(_get_form_group)

    results = []

    def _sign_d_for_form(sub):
        """2プレー以上のサインが2種類以上あればそれだけ使い、なければ1プレーも含める"""
        tbl_2plus = analyze_sign_d(sub, min_count=2)
        if len(tbl_2plus) >= 2:
            return tbl_2plus
        return analyze_sign_d(sub, min_count=1)

    for grp in DEFINED_GROUPS_ORDER:
        sub = df[df['_GROUP'] == grp]
        n = len(sub)
        if n == 0:
            continue
        results.append((grp, n, _sign_d_for_form(sub)))

    # 未分類で3プレー以上
    unc = df[df['_GROUP'].isna()]
    for form, sub in unc.groupby('OFF_FORM_NORM'):
        n = len(sub)
        if n >= 3:
            results.append((f'{form}（その他）', n, _sign_d_for_form(sub)))

    return results


def analyze_packages(df):
    """DEF FRONT + SIGN(D) + COVERAGE の組み合わせが3回以上のパッケージ"""
    if len(df) == 0:
        return pd.DataFrame(columns=['DEF FRONT', 'SIGN', 'COVERAGE', '出現数', '備考（大学・PLAY#）'])

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
            'SIGN':      parts[1] if len(parts) > 1 else '',
            'COVERAGE':  parts[2] if len(parts) > 2 else '',
            '出現数': cnt,
            '備考（大学・PLAY#）': _notes(sub),
        })
    return pd.DataFrame(rows)


def _group_front_rz(v):
    """Red Zone 用 FRONT グループ化（Over / Under は大文字小文字問わず統合）"""
    v = str(v).strip()
    if v.lower() in ('over', 'under'):
        return 'Over/Under'
    return v


# ── LBブリッツ解析用定数 ─────────────────────────────────────────────
# DL要素（LBカウント対象外）: OCS系/ICS系/BCS/数字CS/SL数字/N-T/TEC/T-E/NTT/NO/TO/BO/COUNTER/CS単体
_RZ_DL_RE = re.compile(
    r'^(?:OCS\d*|ICS\d*|BCS|\d+CS|SL\d+|N-T|TEC|T-E|NTT|NO|TO|BO|COUNTER|CS)$',
    re.IGNORECASE,
)


def _rz_lb_count(v: str):
    """
    SIGN(D)/BLITZ値を解析して (fo_n, mb_n, has_dl) を返す。
      FO系（外ギャップ, 1人）: SOLDIER, WARRIOR, FO
      MB系（A/Bギャップ, 1人）: -W/-M/.W/.M含む、先頭W/M+記号、単体W/M
      MB\d: d人の内側LBブリッツ
      DL要素: _RZ_DL_RE に合致する部分（LBカウント対象外）
    """
    fo_n = mb_n = 0
    has_dl = False
    for part in re.split(r'[+,]', v):
        part = part.strip()
        if not part:
            continue
        if _RZ_DL_RE.match(part):
            has_dl = True
            continue
        # MB\d（例: MB2, MB3）→ 数字ぶんのLBカウント
        m = re.fullmatch(r'MB(\d+)', part)
        if m:
            mb_n += int(m.group(1))
            continue
        # FO系（外ギャップ LB）
        if re.fullmatch(r'SOLDIER|WARRIOR|FO', part):
            fo_n += 1
            continue
        # MB単体（A/Bギャップ LB 1人）
        if part == 'MB':
            mb_n += 1
            continue
        # A/Bギャップ blitz: -W/-M, .W/.M, 先頭W/M+記号（M.RIP等）, 単体W/M
        if (re.search(r'[-\.][WM]', part)
                or re.fullmatch(r'[WM]', part)
                or re.match(r'^[WM][\.\-]', part)):
            mb_n += 1
    return fo_n, mb_n, has_dl


def _group_sign_rz(v: str) -> str:
    """
    Red Zone 用 SIGN(D) グループ化（LBブリッツ人数優先）
      0人 + DL要素あり → Cross系
      1人 FO系         → FO系
      1人 MB系         → W or M Blitz
      2人              → LB2人ブリッツ
      3人以上          → LB{n}人ブリッツ
      その他           → raw（BOTH.TI, TLL 等）
    """
    v = str(v).strip()
    if not v or v.lower() == 'nan':
        return ''
    if v == 'N':
        return 'N'
    if 'PINCH' in v:
        return 'PINCH'

    fo_n, mb_n, has_dl = _rz_lb_count(v)
    total = fo_n + mb_n

    if total == 0:
        if has_dl or re.search(
                r'OCS|ICS|BCS|\dCS|SL\d|N-T|TEC|T-E|NTT|\bNO\b|\bTO\b|\bBO\b|COUNTER|\bCS\b', v):
            return 'Cross系'
        return v  # BOTH.TI, TLL, TEN 等はそのまま個別
    if total == 1:
        return 'FO系' if fo_n == 1 else 'W or M Blitz'
    if total == 2:
        return 'LB2人ブリッツ'
    return f'LB{total}人ブリッツ'


def _group_blitz_rz(v: str) -> str:
    """
    Red Zone 用 BLITZ グループ化（LBブリッツ人数優先、SIGN(D)と同ロジック）
    INCHARGE は独立グループ。
    """
    v = str(v).strip()
    if not v or v.lower() == 'nan':
        return ''
    if v == 'N':
        return 'N'
    if 'INCHARGE' in v:
        return 'INCHARGE'

    fo_n, mb_n, has_dl = _rz_lb_count(v)
    total = fo_n + mb_n

    if total == 0:
        if has_dl or re.search(
                r'OCS|ICS|BCS|\dCS|SL\d|N-T|TEC|T-E|NTT|\bNO\b|\bTO\b|\bBO\b|COUNTER|\bCS\b', v):
            return 'Cross系'
        return v
    if total == 1:
        return 'FO系' if fo_n == 1 else 'W or M Blitz'
    if total == 2:
        return 'LB2人ブリッツ'
    return f'LB{total}人ブリッツ'


def _group_cov_rz(v):
    """Red Zone 用 COVERAGE グループ化"""
    v = str(v).strip()
    if '3' in v:
        return 'Cover3'
    return v


def _pkg_label(*parts: str) -> str:
    """パッケージ表示文字列を ' + ' で結合する（SIGN値は sign_display で変換）"""
    return ' + '.join(sign_display(p) for p in parts)


def analyze_redzone_packages(df):
    """
    Red Zone パッケージ割合表。
    ・3プレー以上のパッケージ → FRONT + SIGN + COVERAGE の3次元でそのまま維持
      （グループ内の全プレーが同一の具体値ならその値を表示。SIGN='N' は 'read'）
    ・2プレー以下のパッケージ → FRONT 別に集約
        COVERAGE のみで集計した最大 vs SIGN のみで集計した最大 を比較し、
        多い方の次元を採用して 'FRONT + 採用次元値' として再集計。
        再集計後も 2 プレー以下になる行は非表示。
    """
    total = len(df)
    if total == 0:
        return pd.DataFrame(columns=['パッケージ', '割合（実数）'])

    tmp = df.copy()
    tmp['_fg'] = tmp['DEF_FRONT'].apply(_group_front_rz)
    tmp['_sg'] = tmp['SIGN_D'].apply(_group_sign_rz)
    tmp['_cg'] = tmp['COVERAGE_NORM'].apply(_group_cov_rz)

    # グループサイズで stable（≥3）と small（≤2）を分離
    grp_sizes = tmp.groupby(['_fg', '_sg', '_cg'])['_fg'].transform('size')
    stable_df = tmp[grp_sizes >= 3]
    small_df  = tmp[grp_sizes <  3]

    rows = []  # (count, pkg_str)

    # ── stable：FRONT + SIGN + COVERAGE（3次元）──
    for (f, s, c), grp in stable_df.groupby(['_fg', '_sg', '_cg']):
        cnt = len(grp)
        raw_signs = grp['SIGN_D'].astype(str).str.strip().unique()
        sign_disp = raw_signs[0] if len(raw_signs) == 1 else s
        rows.append((cnt, _pkg_label(f, sign_disp, c)))

    # ── small：FRONT 別に次元選択して再集約 ──
    for f, f_grp in small_df.groupby('_fg'):
        cov_counts  = f_grp.groupby('_cg').size().sort_values(ascending=False)
        sign_counts = f_grp.groupby('_sg').size().sort_values(ascending=False)

        max_cov  = int(cov_counts.iloc[0])  if len(cov_counts)  > 0 else 0
        max_sign = int(sign_counts.iloc[0]) if len(sign_counts) > 0 else 0

        if max_cov >= max_sign:
            # COVERAGE 次元を採用（SIGN を無視）
            for cv, cnt in cov_counts.items():
                if cnt > 2:
                    rows.append((int(cnt), _pkg_label(f, cv)))
        else:
            # SIGN 次元を採用（COVERAGE を無視）
            for sv, cnt in sign_counts.items():
                if cnt > 2:
                    rows.append((int(cnt), _pkg_label(f, sv)))

    rows.sort(key=lambda x: -x[0])
    return pd.DataFrame([
        {'パッケージ': pkg, '割合（実数）': _pct_str(cnt, total)}
        for cnt, pkg in rows
    ])


def build_redzone_pkg_map(df):
    """
    pkg_label → 該当行 DataFrame の辞書を返す（consider_redzone_package_item 用）。
    analyze_redzone_packages と同じグループ化ロジックを使用。
    """
    if len(df) == 0:
        return {}
    tmp = df.copy()
    tmp['_fg'] = tmp['DEF_FRONT'].apply(_group_front_rz)
    tmp['_sg'] = tmp['SIGN_D'].apply(_group_sign_rz)
    tmp['_cg'] = tmp['COVERAGE_NORM'].apply(_group_cov_rz)

    grp_sizes = tmp.groupby(['_fg', '_sg', '_cg'])['_fg'].transform('size')
    result = {}   # pkg_label → list of df chunks

    # stable（≥3プレー）: 3次元グループ
    for (f, s, c), grp in tmp[grp_sizes >= 3].groupby(['_fg', '_sg', '_cg']):
        raw_signs = grp['SIGN_D'].astype(str).str.strip().unique()
        sign_disp = raw_signs[0] if len(raw_signs) == 1 else s
        label = _pkg_label(f, sign_disp, c)
        result.setdefault(label, []).append(grp)

    # small（<3プレー）: FRONT別に次元選択して2次元グループ
    small = tmp[grp_sizes < 3]
    for f, f_grp in small.groupby('_fg'):
        cov_counts  = f_grp.groupby('_cg').size().sort_values(ascending=False)
        sign_counts = f_grp.groupby('_sg').size().sort_values(ascending=False)
        max_cov  = int(cov_counts.iloc[0])  if len(cov_counts)  > 0 else 0
        max_sign = int(sign_counts.iloc[0]) if len(sign_counts) > 0 else 0
        if max_cov >= max_sign:
            for cv, cnt in cov_counts.items():
                if cnt > 2:
                    label = _pkg_label(f, cv)
                    result.setdefault(label, []).append(f_grp[f_grp['_cg'] == cv])
        else:
            for sv, cnt in sign_counts.items():
                if cnt > 2:
                    label = _pkg_label(f, sv)
                    result.setdefault(label, []).append(f_grp[f_grp['_sg'] == sv])

    return {k: pd.concat(v, ignore_index=True) for k, v in result.items()}


def analyze_manzono(df):
    """
    Man系 / Zone系 比率テーブル。Red Zone COVERAGE表の下に表示。
    Man系 : COVERAGE_NORM の先頭数値が 0 または 1 (Cover 0, Cover 1, 1 FREE 等)
    Zone系: COVERAGE_NORM の先頭数値が 2 以上 (Cover 2〜7 等)
    """
    total = len(df)
    if total == 0:
        return pd.DataFrame(columns=['タイプ', '割合（実数）'])

    def _classify(v):
        s = str(v).strip()
        if not s or s.lower() == 'nan':
            return None
        try:
            n = int(float(s.split()[0]))
            if n <= 1:
                return 'Man系（Cover 0・1系）'
            return 'Zone系（Cover 2以上）'
        except ValueError:
            return None

    classified = df['COVERAGE_NORM'].apply(_classify)
    rows = []
    for label in ['Man系（Cover 0・1系）', 'Zone系（Cover 2以上）']:
        cnt = (classified == label).sum()
        if cnt == 0:
            continue
        rows.append({
            'タイプ': label,
            '割合（実数）': _pct_str(cnt, total),
        })
    return pd.DataFrame(rows)


def analyze_3rd_zones(df, dist_zones=None):
    """DISTゾーン別に分析結果をまとめて返す"""
    zones = dist_zones if dist_zones is not None else DIST_ZONES
    results = {}
    for zone_name, dist_min, dist_max in zones:
        sub = df[(df['DIST'] >= dist_min) & (df['DIST'] <= dist_max)].copy()
        cov_tbl, comp3_tbl = analyze_coverage(sub)
        results[zone_name] = {
            'n':        len(sub),
            'df':       sub,
            'sign':     analyze_sign_d(sub, min_count=1),
            'front':    analyze_front(sub),
            'coverage': cov_tbl,
            'comp3':    comp3_tbl,
            'packages': analyze_packages(sub),
        }
    return results


# ── BLITZ / PRESSURE 分析 ────────────────────────────────────
def analyze_blitz(df):
    """BLITZ割合テーブル（2プレー以下は表示しない）"""
    total = len(df)
    rows = []
    if total == 0:
        return pd.DataFrame(columns=['BLITZ', '割合（実数）', '備考'])
    counts = df['BLITZ_CLEAN'].value_counts()
    for cat, cnt in counts.items():
        if not cat or cat == 'nan':
            continue
        if cnt <= 2:
            continue
        sub = df[df['BLITZ_CLEAN'] == cat]
        rows.append({
            'BLITZ': sign_display(cat),
            '割合（実数）': _pct_str(cnt, total),
            '備考': _notes(sub) if cnt <= 3 else '',
        })
    return pd.DataFrame(rows)


def analyze_pressure(df):
    """PRESSURE（ラッシュ人数）分布テーブル"""
    total = len(df)
    rows = []
    if total == 0:
        return pd.DataFrame(columns=['PRESSURE', '割合（実数）', '備考'])
    counts = df['PRESSURE_NUM'].dropna().astype(int).value_counts().sort_index()
    for val, cnt in counts.items():
        sub = df[df['PRESSURE_NUM'] == val]
        rows.append({
            'PRESSURE': str(val),
            '割合（実数）': _pct_str(cnt, total),
            '備考': _notes(sub) if cnt <= 3 else '',
        })
    return pd.DataFrame(rows)


def _top1(series):
    """最頻値と件数を '値 (N)' 形式で返す。データなしは '-'"""
    s = series[series.astype(str).str.strip().ne('')]
    if len(s) == 0:
        return '-'
    top = s.value_counts().index[0]
    cnt = s.value_counts().iloc[0]
    return f'{top} ({cnt})'


# ── 多次元分析：ネスト構造 ───────────────────────────────────
def analyze_multidim_nested(df):
    """YARD LN × DN × DIST の入れ子構造で各指標を分析
    戻り値: { yln_name: { dn_label: { dist_name: { stats } } } }
    """
    results = {}
    for yln_name, yln_min, yln_max in YARD_LN_ZONES:
        yln_df = df[(df['YARD LN'] >= yln_min) & (df['YARD LN'] <= yln_max)].copy()
        if len(yln_df) == 0:
            continue
        dn_dict = {}
        for dn in sorted(yln_df['DN'].dropna().unique()):
            dn_df = yln_df[yln_df['DN'] == dn]
            dn_label = f'DN={int(dn)}'
            dist_dict = {}
            for zone_name, dist_min, dist_max in DIST_ZONES:
                sub = dn_df[(dn_df['DIST'] >= dist_min) & (dn_df['DIST'] <= dist_max)]
                n = len(sub)
                if n == 0:
                    continue
                cov_tbl, comp3_tbl = analyze_coverage(sub)
                dist_dict[zone_name] = {
                    'n':        n,
                    'front':    analyze_front(sub),
                    'coverage': cov_tbl,
                    'comp3':    comp3_tbl,
                    'sign_d':   analyze_sign_d(sub),
                    'blitz':    analyze_blitz(sub),
                    'pressure': analyze_pressure(sub),
                    'packages': analyze_packages(sub),
                }
            if dist_dict:
                dn_dict[dn_label] = dist_dict
        if dn_dict:
            results[yln_name] = dn_dict
    return results


# ── フィールド型グリッド列定義 ─────────────────────────────────
# (ラベル, DN, DIST_min, DIST_max)  DIST_max=999 は上限なし
NORMAL_FIELD_COLS = [
    ('1st\n10',    1, 1, 999),
    ('2nd\n1~3',   2, 1,   3),
    ('2nd\n4~6',   2, 4,   6),
    ('2nd\n7~10',  2, 7,  10),
    ('2nd\n11~',   2, 11, 999),
]

THIRD_FIELD_COLS = [
    ('3rd/4th\n1~3',  (3, 4), 1,   3),
    ('3rd/4th\n4~6',  (3, 4), 4,   6),
    ('3rd/4th\n7~10', (3, 4), 7,  10),
    ('3rd/4th\n11~',  (3, 4), 11, 999),
]

# Red Zone 用ヤードゾーン定義
REDZONE_YARD_LN_ZONES = [
    ('①Redzone（11〜25yds）',    11, 25),
    ('②Redzone中盤（6〜10yds）',  6, 10),
    ('③G前（1〜5yds）',           1,  5),
]

# Red Zone 用列定義（Normal と 3rd でそれぞれ NORMAL/THIRD_FIELD_COLS を流用）
# REDZONE_NORMAL_FIELD_COLS → NORMAL_FIELD_COLS と同じ（1st / 2nd）
# REDZONE_THIRD_FIELD_COLS  → THIRD_FIELD_COLS  と同じ（3rd/4th）


def analyze_field_grid(df, situation: str = 'normal', field_cols=None, yard_zones=None) -> dict:
    """フィールド型ヒートマップ用グリッドデータを生成する。

    Parameters
    ----------
    df : フィルタ済みの DataFrame（filter_normal / filter_3rd 適用後）
    situation : 'normal' または 'third'
    field_cols : カスタム列定義リスト。None の場合はデフォルトを使用
    yard_zones : カスタム YARD LN ゾーン定義リスト。None の場合は YARD_LN_ZONES を使用

    Returns
    -------
    dict
        col_labels : list[str]  列ラベル（DN+DIST 組み合わせ）
        row_labels : list[str]  行ラベル（YARD LN ゾーン; 0=Backed Up）
        cells      : dict       (ri, ci) → {n, front, front_n, coverage, cov_n}
    """
    if field_cols is not None:
        cols = field_cols
    else:
        cols = NORMAL_FIELD_COLS if situation == 'normal' else THIRD_FIELD_COLS
    zones = yard_zones if yard_zones is not None else YARD_LN_ZONES
    col_labels = [c[0] for c in cols]
    row_labels  = [z[0] for z in zones]

    cells = {}
    for ri, (_, yln_min, yln_max) in enumerate(zones):
        yln_df = df[(df['YARD LN'] >= yln_min) & (df['YARD LN'] <= yln_max)]
        for ci, col_def in enumerate(cols):
            _, dn, dist_min, dist_max = col_def
            dn_filter = (
                yln_df['DN'].isin(dn)
                if isinstance(dn, (list, tuple))
                else (yln_df['DN'] == dn)
            )
            sub = yln_df[
                dn_filter &
                (yln_df['DIST'] >= dist_min) &
                (yln_df['DIST'] <= dist_max)
            ]
            n = len(sub)
            cov_val, cov_n     = '-', 0
            press_val, press_n = '-', 0
            sign_val, sign_n   = '-', 0
            blitz_val, blitz_n = '-', 0
            cov_others = press_others = sign_others = blitz_others = []
            cov_allvals = press_allvals = sign_allvals = blitz_allvals = []

            def _others(vc, top_n):
                """最頻値と同数（タイ）の2位以降の値リストをすべて返す（上限なし）"""
                result = []
                for i in range(1, len(vc)):
                    cnt = int(vc.iloc[i])
                    if cnt == top_n:
                        result.append((str(vc.index[i]), cnt))
                    else:
                        break  # value_counts は降順なので同数でなければ以降も該当しない
                return result

            def _allvals(vc):
                """セル合計プレー数が==2かつ値が2種類のとき全ユニーク値リストを返す（縦並び表示用）"""
                if len(vc) == 0 or int(vc.sum()) != 2 or len(vc) < 2:
                    return []
                return [(str(v), int(c)) for v, c in vc.items()]

            if n > 0:
                # COVERAGE
                cc = sub['COVERAGE_NORM'].replace('', pd.NA).dropna().value_counts()
                if len(cc) > 0:
                    cov_val    = cc.index[0]
                    cov_n      = int(cc.iloc[0])
                    cov_others = _others(cc, cov_n)
                    cov_allvals = _allvals(cc)

                # PRESSURE（最頻ラッシュ人数）
                pn = sub['PRESSURE_NUM'].dropna().astype(int).value_counts()
                if len(pn) > 0:
                    press_val    = str(pn.index[0])
                    press_n      = int(pn.iloc[0])
                    press_others = _others(pn, press_n)
                    press_allvals = _allvals(pn)

                # SIGN(D)
                sc = sub['SIGN_D'].replace('', pd.NA).dropna().value_counts()
                if len(sc) > 0:
                    sign_val    = _hm_sign_display(sc.index[0])
                    sign_n      = int(sc.iloc[0])
                    sign_others = [(_hm_sign_display(v), c) for v, c in _others(sc, sign_n)]
                    sign_allvals = [(_hm_sign_display(v), c) for v, c in _allvals(sc)]

                # BLITZ
                bc = sub['BLITZ_CLEAN'].replace('', pd.NA).dropna().value_counts()
                if len(bc) > 0:
                    blitz_val    = _hm_sign_display(bc.index[0])
                    blitz_n      = int(bc.iloc[0])
                    blitz_others = [(_hm_sign_display(v), c) for v, c in _others(bc, blitz_n)]
                    blitz_allvals = [(_hm_sign_display(v), c) for v, c in _allvals(bc)]

            cells[(ri, ci)] = {
                'n':              n,
                'coverage':       cov_val,   'cov_n':      cov_n,    'cov_others':    cov_others,   'cov_allvals':   cov_allvals,
                'pressure':       press_val,  'press_n':    press_n,  'press_others':  press_others, 'press_allvals': press_allvals,
                'sign_d':         sign_val,   'sign_d_n':   sign_n,   'sign_others':   sign_others,  'sign_allvals':  sign_allvals,
                'blitz':          blitz_val,  'blitz_n':    blitz_n,  'blitz_others':  blitz_others, 'blitz_allvals': blitz_allvals,
            }

    return {
        'col_labels': col_labels,
        'row_labels': row_labels,
        'cells':      cells,
    }


# ── 多次元分析：サマリーグリッド ─────────────────────────────
def analyze_multidim_grid(df):
    """YARD LN（行）× DIST（列）のグリッド表を返す
    各セル = COVERAGE / PRESSURE / SIGN(D) / BLITZ の最頻値 (N=件数)
    戻り値: { 'coverage': DataFrame, 'pressure': DataFrame,
              'sign_d': DataFrame, 'blitz': DataFrame }
    """
    yln_labels  = [z[0] for z in YARD_LN_ZONES]
    dist_labels = [z[0] for z in DIST_ZONES]

    cov_grid   = {d: [] for d in dist_labels}
    press_grid = {d: [] for d in dist_labels}
    sign_grid  = {d: [] for d in dist_labels}
    blitz_grid = {d: [] for d in dist_labels}

    for yln_name, yln_min, yln_max in YARD_LN_ZONES:
        yln_df = df[(df['YARD LN'] >= yln_min) & (df['YARD LN'] <= yln_max)]
        for zone_name, dist_min, dist_max in DIST_ZONES:
            sub = yln_df[(yln_df['DIST'] >= dist_min) & (yln_df['DIST'] <= dist_max)]
            cov_grid[zone_name].append(_top1(sub['COVERAGE_NORM']))
            press_series = sub['PRESSURE_NUM'].dropna().astype(int).astype(str)
            press_grid[zone_name].append(_top1(press_series))
            sign_grid[zone_name].append(_top1(sub['SIGN_D']))
            blitz_grid[zone_name].append(_top1(sub['BLITZ_CLEAN']))

    def _make_df(grid):
        out = pd.DataFrame(grid, index=yln_labels)
        out.index.name = 'YARD LN \\ DIST'
        return out

    return {
        'coverage': _make_df(cov_grid),
        'pressure': _make_df(press_grid),
        'sign_d':   _make_df(sign_grid),
        'blitz':    _make_df(blitz_grid),
    }
