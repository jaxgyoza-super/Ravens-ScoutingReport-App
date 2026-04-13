import pandas as pd
import re

# ── バリデーション用定数 ──────────────────────────────────────────
REQUIRED_COLS = [
    'DN', 'DIST', 'YARD LN', 'PLAY #', '2 MIN',
    'COVERAGE', 'COMPONENT', 'OFF FORM', 'DEF FRONT',
    'SIGN(D)', 'BLITZ', 'PRESSURE', 'PERSONNEL',
]

COMPONENT_MAP = {
    'FS BUZZ': 'FS Buzz',
    'FS SKY':  'FS Sky',
    'R BUZZ':  'R Buzz',
    'R SKY':   'R Sky',
    # レガシー表記（旧マッピング名が残っている場合も統一）
    'SILVER':  'FS Buzz',
    'GREEN':   'FS Sky',
    'GOLD':    'R Buzz',
    'YELLOW':  'R Sky',
}


def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file, header=0)

    # --- 対戦校名の付与 ---
    # 1列目のヘッダー名が最初の対戦校。その後、同列に値が入ったら対戦校が切り替わる
    first_opp = str(df.columns[0]).strip()
    opps = []
    current = first_opp
    for val in df.iloc[:, 0]:
        s = str(val).strip()
        if s and s.lower() != 'nan':
            current = s
        opps.append(current)
    df['OPPONENT'] = opps

    # --- 数値変換 ---
    for col in ['DN', 'DIST', 'YARD LN', 'PLAY #']:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # --- 2MIN フラグ ---
    df['IS_2MIN'] = df['2 MIN'].astype(str).str.strip().str.upper() == 'Y'

    # --- RedZone フラグ (YARD LN 1〜25) ---
    df['IS_REDZONE'] = (df['YARD LN'] >= 1) & (df['YARD LN'] <= 25)

    # --- COVERAGE 正規化 (33→3, 数値の.0除去) ---
    def norm_coverage(v):
        s = str(v).strip()
        if s.lower() == 'nan' or s == '':
            return ''
        # 数値で来る場合 (3.0, 33.0 など)
        s = re.sub(r'\.0$', '', s)
        # Cover 3 に統合するバリエーション
        if s in ('33', 'ROTE 3', 'ROTE3', '3?'):
            s = '3'
        # 1FREE 表記ゆれ統一
        if s == '1FREE':
            s = '1 FREE'
        return s

    df['COVERAGE_NORM'] = df['COVERAGE'].apply(norm_coverage)

    # --- COMPONENT 表示名マッピング ---
    def map_component(v):
        s = str(v).strip()
        if s.lower() == 'nan' or s == '':
            return ''
        return COMPONENT_MAP.get(s, s)

    df['COMPONENT_DISPLAY'] = df['COMPONENT'].apply(map_component)

    # --- OFF FORM 正規化 (SG プレフィックス除去) ---
    df['OFF_FORM_NORM'] = (
        df['OFF FORM'].astype(str).str.strip()
        .str.replace(r'^SG\s+', '', regex=True)
    )
    df.loc[df['OFF_FORM_NORM'].str.lower() == 'nan', 'OFF_FORM_NORM'] = ''

    # --- DEF FRONT / SIGN(D) / BLITZ クリーニング ---
    for src, dst in [('DEF FRONT', 'DEF_FRONT'), ('SIGN(D)', 'SIGN_D'), ('BLITZ', 'BLITZ_CLEAN')]:
        df[dst] = df[src].astype(str).str.strip()
        df.loc[df[dst].str.lower() == 'nan', dst] = ''

    # --- PRESSURE 数値化 ---
    df['PRESSURE_NUM'] = pd.to_numeric(df['PRESSURE'], errors='coerce')

    # --- PERSONNEL 正規化（10.0 → '10' など）---
    def norm_personnel(v):
        s = str(v).strip()
        if s.lower() == 'nan' or s == '':
            return ''
        return re.sub(r'\.0$', '', s)

    df['PERSONNEL_NORM'] = df['PERSONNEL'].apply(norm_personnel)

    # --- HASH / OFFSTR / BACKFIELD 正規化 ---
    def _clean(v):
        s = str(v).strip().upper()
        return '' if s in ('NAN', '') else s

    # 列が存在しない場合は空文字列で補完
    for raw, norm in [('HASH', 'HASH_NORM'), ('OFF STR', 'OFFSTR_NORM'), ('BACKFIELD', 'BACKFIELD_NORM')]:
        if raw in df.columns:
            df[norm] = df[raw].apply(_clean)
        else:
            df[norm] = ''

    # --- RB_WIDE_SIDE 派生列 ---
    # 広い側: HASH=L → 右(R)が広い、HASH=R → 左(L)が広い
    # BACKFIELD=ST → RBはOFFSTRと同じ側、WE → 逆側、PISTOL → QB後ろ、空白 → RBなし
    def calc_rb_wide(row):
        h = row['HASH_NORM']
        o = row['OFFSTR_NORM']
        b = row['BACKFIELD_NORM']
        if not b:
            return 'なし'
        if b == 'PISTOL':
            return 'PISTOL'
        if not h or not o:
            return ''
        wide = 'R' if h == 'L' else ('L' if h == 'R' else '')
        if not wide:
            return ''
        rb_side = o if b == 'ST' else ('R' if o == 'L' else 'L')
        if rb_side == wide:
            return '広い'
        return '狭い'

    df['RB_WIDE_SIDE'] = df.apply(calc_rb_wide, axis=1)

    return df


def get_opponents(df):
    return sorted(df['OPPONENT'].unique().tolist())


def validate_excel(file_bytes) -> list:
    """
    アップロードされた Excel ファイルの形式チェック。
    エラーメッセージのリストを返す（空リスト = 問題なし）。
    """
    import io
    errors = []

    # ── 読み込みチェック ──────────────────────────────────────────
    try:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), header=0)
    except Exception as e:
        return [f'Excelファイルが読み込めなかったよ…：{e}']

    # ── データ行数チェック ────────────────────────────────────────
    if len(df_raw) == 0:
        errors.append('データが1件もないよ！ヘッダー行だけのファイルかも？')
        return errors  # 以降のチェックは意味がないので早期リターン

    # ── 必須列チェック ────────────────────────────────────────────
    missing = [col for col in REQUIRED_COLS if col not in df_raw.columns]
    if missing:
        errors.append(f'以下の列が見つからないよっ：{", ".join(missing)}')

    # ── 1列目（大学名列）チェック ─────────────────────────────────
    # NG パターン：
    #   ① 1列目ヘッダーが既存のデータ列名（大学名列を追加し忘れ）
    #   ② 1列目の値が全部空
    #   ③ 1列目の値が全部数値（PLAY # などの数字が入っている）
    first_col_name = str(df_raw.columns[0]).strip()
    first_col_vals = df_raw.iloc[:, 0].astype(str).str.strip()
    non_empty = first_col_vals[~first_col_vals.isin(['', 'nan', 'NaN'])]

    def _all_numeric(series):
        def _is_num(s):
            try:
                float(s)
                return True
            except ValueError:
                return False
        return len(series) > 0 and all(_is_num(v) for v in series)

    if (first_col_name in REQUIRED_COLS
            or len(non_empty) == 0
            or _all_numeric(non_empty)):
        errors.append(
            '1列目に大学名が入ってないよ！'
            '「使い方」の手順③を見てね。'
        )

    return errors
