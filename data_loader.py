import pandas as pd
import re

COMPONENT_MAP = {
    'FS BUZZ': 'SILVER',
    'FS SKY': 'GREEN',
    'R BUZZ': 'GOLD',
    'R SKY': 'YELLOW',
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
        if s == '33':
            s = '3'
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

    # --- DEF FRONT / SIGN(D) クリーニング ---
    for src, dst in [('DEF FRONT', 'DEF_FRONT'), ('SIGN(D)', 'SIGN_D')]:
        df[dst] = df[src].astype(str).str.strip()
        df.loc[df[dst].str.lower() == 'nan', dst] = ''

    return df


def get_opponents(df):
    return sorted(df['OPPONENT'].unique().tolist())
