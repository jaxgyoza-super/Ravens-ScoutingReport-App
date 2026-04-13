import os
import subprocess
import streamlit as st
import streamlit.components.v1 as components



def _notify_pc(title: str, message: str) -> None:
    """Windows のバルーン通知を表示する（PowerShell 経由、ウィンドウ非表示）。非Windows環境では何もしない。"""
    if os.name != 'nt':
        return
    ps = (
        'Add-Type -AssemblyName System.Windows.Forms;'
        '$n=New-Object System.Windows.Forms.NotifyIcon;'
        '$n.Icon=[System.Drawing.SystemIcons]::Information;'
        '$n.Visible=$true;'
        f'$n.ShowBalloonTip(6000,"{title}","{message}",'
        '[System.Windows.Forms.ToolTipIcon]::Info);'
        'Start-Sleep -Seconds 7;'
        '$n.Dispose()'
    )
    subprocess.Popen(
        ['powershell', '-WindowStyle', 'Hidden', '-Command', ps],
        creationflags=subprocess.CREATE_NO_WINDOW,
    )


# ── ローディング GIF の base64 読み込み ──────────────────────────
@st.cache_resource
def _loading_gif_html() -> str:
    """起動時に1回だけ base64 エンコード。"""
    import base64 as _b64
    fpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'loading.gif')
    if os.path.exists(fpath):
        with open(fpath, 'rb') as f:
            b64 = _b64.b64encode(f.read()).decode()
        return (
            f'<div style="text-align:center;padding:8px 0 8px;">'
            f'<img src="data:image/gif;base64,{b64}" '
            f'style="width:288px;height:288px;object-fit:contain;">'
            f'<div style="font-size:16px;font-weight:900;color:#D32F2F;'
            f'letter-spacing:0.25em;text-transform:uppercase;margin-top:10px;">'
            f'生成中...</div>'
            f'</div>'
        )
    # フォールバック
    return (
        '<div style="text-align:center;padding:28px 0 20px;">'
        '<div style="font-size:28px;">⏳</div>'
        '<div style="font-size:16px;font-weight:900;color:#D32F2F;'
        'letter-spacing:0.25em;text-transform:uppercase;margin-top:10px;">'
        '生成中...</div>'
        '</div>'
    )


st.set_page_config(
    page_title='Ravens Auto Scout Kit',
    page_icon='🏈',
    layout='wide',
)


# ── グローバル CSS ────────────────────────────────────────────
st.markdown("""
<style>
/* Google Fonts — Barlow Condensed (スタイリッシュな縦長フォント) */
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:ital,wght@1,800&display=swap');

/* ベース背景・ページスクロール無効化 */
html, body {
    overflow: hidden !important;
    height: 100% !important;
}
[data-testid="stAppViewContainer"] {
    background: #EFEFEF;
    height: 100vh !important;
    overflow: hidden !important;
}
[data-testid="stMain"] {
    overflow: hidden !important;
}

/* Streamlit 固有のヘッダーバーをすべて非表示 */
[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
header[data-testid="stHeader"] {
    display: none !important;
    height: 0 !important;
}
.block-container {
    padding: 0.4rem 2.0rem 0.2rem !important;
    max-width: 100% !important;
}

/* ナビ行（タイトル + 使い方ボタン）コンパクト化 */
.rask-topbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    height: 40px;
    border-bottom: 1px solid #F0F0F0;
    margin-bottom: 6px;
    padding: 0 2px;
}

.rask-nav-title {
    font-family: 'Barlow Condensed', 'Arial Narrow', sans-serif;
    font-size: 32px;
    font-weight: 800;
    font-style: italic;
    letter-spacing: -0.5px;
    line-height: 1;
    display: flex;
    align-items: center;
    gap: 0;
}
.rask-nav-title .r { color: #D32F2F; }
.rask-nav-title .b { color: #1A1A1A; }
.rask-nav-title .sp { margin-left: 8px; }

/* カード（st.container(border=True)） */
[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 20px !important;
    border-color: #F0F0F0 !important;
    padding: 12px 16px !important;
    background: white !important;
    box-shadow: 0 4px 24px -4px rgba(0,0,0,0.04) !important;
}

/* ファイルアップローダーのデフォルトDropzoneを非表示
   （<input type="file"> はDOMに残り、JS から .click() できる） */
[data-testid="stFileUploaderDropzone"] {
    display: none !important;
}
/* ファイル選択後に出るファイル名バッジも非表示（Python側でカスタム表示） */
[data-testid="stFileUploader"] [data-testid="stMarkdownContainer"],
[data-testid="stFileUploader"] small {
    display: none !important;
}

/* 大学チェックボックスを角丸ボタン風に（超コンパクト） */
[data-testid="stCheckbox"] {
    background: white;
    border: 1px solid #F1F5F9;
    border-radius: 5px;
    padding: 0px 3px !important;
    margin-bottom: 0px !important;
    transition: border-color 0.2s;
    min-height: 0 !important;
}
[data-testid="stCheckbox"]:hover {
    border-color: #E2E8F0;
}
[data-testid="stCheckbox"] label {
    font-size: 8px !important;
    font-weight: 600 !important;
    color: #475569;
    line-height: 1.2 !important;
    cursor: pointer;
}
[data-testid="stCheckbox"] label span { color: #475569 !important; }
[data-testid="stCheckbox"] input:checked ~ label span { color: #D32F2F !important; }
/* チェックボックス本体（四角）も小さく */
[data-testid="stCheckbox"] input[type="checkbox"] {
    width: 11px !important;
    height: 11px !important;
}
/* チェックボックス要素間の余白を完全ゼロに */
[data-testid="stCheckbox"] > label {
    gap: 4px !important;
}

/* 大学選択タブをコンパクトに */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    gap: 0 !important;
    min-height: 26px !important;
    padding-bottom: 0 !important;
}
[data-testid="stTabs"] [data-baseweb="tab"] {
    padding: 2px 8px !important;
    font-size: 10px !important;
    font-weight: 700 !important;
    min-height: 26px !important;
    height: 26px !important;
}
[data-testid="stTabs"] [data-baseweb="tab-panel"] {
    padding: 2px 0 0 0 !important;
}
[data-testid="stTabs"] [data-baseweb="tab-border"] {
    display: none !important;
}

/* st.container(height=N) が生成するスクロール領域のスクロールバーを完全非表示 */
[data-testid="stVerticalBlockBorderWrapper"] > div {
    scrollbar-width: none !important;
    -ms-overflow-style: none !important;
}
[data-testid="stVerticalBlockBorderWrapper"] > div::-webkit-scrollbar {
    display: none !important;
    width: 0px !important;
    height: 0px !important;
}
div[style*="overflow: auto"],
div[style*="overflow:auto"] {
    scrollbar-width: none !important;
    -ms-overflow-style: none !important;
}
div[style*="overflow: auto"]::-webkit-scrollbar,
div[style*="overflow:auto"]::-webkit-scrollbar {
    display: none !important;
    width: 0px !important;
}

/* ヤード設定 数値入力 */
[data-testid="stNumberInput"] input {
    text-align: center !important;
    border-radius: 6px !important;
    font-size: 14px !important;
    font-weight: 600 !important;
    padding: 2px !important;
}

/* START ANALYSIS ボタン（primary） */
[data-testid="stButton"]:has(button[data-testid="baseButton-primary"]) {
    width: 100% !important;
}
[data-testid="stButton"]:has(button[data-testid="baseButton-primary"]) > button,
button[data-testid="baseButton-primary"] {
    background: #D32F2F !important;
    border: none !important;
    border-radius: 14px !important;
    font-weight: 900 !important;
    font-size: 18px !important;
    letter-spacing: 0.5em !important;
    font-style: italic !important;
    height: 300px !important;
    min-height: 300px !important;
    padding-top: 120px !important;
    padding-bottom: 120px !important;
    min-width: 260px !important;
    width: 100% !important;
    transition: all 0.3s !important;
}
[data-testid="stButton"]:has(button[data-testid="baseButton-primary"]) > button:hover,
button[data-testid="baseButton-primary"]:hover:not(:disabled) {
    background: #B71C1C !important;
    box-shadow: 0 20px 40px -10px rgba(211,47,47,0.35) !important;
    transform: translateY(-2px);
}
[data-testid="stButton"]:has(button[data-testid="baseButton-primary"]) > button:disabled,
button[data-testid="baseButton-primary"]:disabled {
    background: #F1F5F9 !important;
    color: #CBD5E1 !important;
}

/* ダウンロードボタン */
[data-testid="stDownloadButton"] button {
    border-radius: 12px !important;
    font-weight: 700 !important;
    font-size: 11px !important;
    letter-spacing: 0.3em !important;
    text-transform: uppercase !important;
}

/* divider */
hr { border-color: #F8F8F8 !important; margin: 0 0 0.5rem 0 !important; }


/* Streamlit の列・ブロック間ギャップを最小化 */
[data-testid="stVerticalBlock"] > div { gap: 0.15rem !important; }
[data-testid="stHorizontalBlock"] { gap: 0.3rem !important; }

/* number_input の高さを詰める */
[data-testid="stNumberInput"] input {
    height: 22px !important;
    padding: 0 3px !important;
    font-size: 14px !important;
}
/* number_input 全体のマージンも詰める */
[data-testid="stNumberInput"] {
    margin-bottom: 0 !important;
    padding-bottom: 0 !important;
}
[data-testid="stNumberInput"] > div {
    margin-bottom: 0 !important;
    gap: 0 !important;
}
/* numberInput のラベル部分（hidden でも余白が残る）を消す */
[data-testid="stNumberInput"] > label {
    display: none !important;
    height: 0 !important;
    margin: 0 !important;
    padding: 0 !important;
}

/* 条件ダイアログを大画面表示 */
[data-testid="stDialog"] > div {
    max-width: 860px !important;
    width: 90vw !important;
}
</style>
""", unsafe_allow_html=True)


# ── 使い方ダイアログ ─────────────────────────────────────────
@st.dialog('使い方 — How to Use', width='large')
def show_usage_modal():
    st.markdown(
        '<style>'
        '[data-testid="stDialog"] > div > div {'
        '  max-height: 88vh !important;'
        '  overflow-y: auto !important;'
        '  scrollbar-width: thin !important;'
        '}'
        '[data-testid="stDialog"] [data-testid="stVerticalBlock"],'
        '[data-testid="stDialog"] [data-testid="stVerticalBlockBorderWrapper"] > div {'
        '  overflow: visible !important;'
        '  max-height: none !important;'
        '}'
        '</style>',
        unsafe_allow_html=True,
    )
    steps = [
        ('1', 'Hudlでの準備', [
            'Hudlで分析対象の試合・プレー（※ <b>D# のみ</b>）を選び、1つのプレイリストにまとめる。',
            'このとき、大学ごとに並んでいる状態にしておく。',
        ]),
        ('2', 'Hudl → Excel', [
            '作成したプレイリストを、<b>Excel 形式</b>でエクスポートする。',
        ]),
        ('3', 'Excelの編集', [
            'エクスポートしたファイルを開き、先頭に新しく <b>1列追加</b> する。',
            'その列に、各大学の最初のプレーの行にのみ大学名を入力する。',
        ]),
    ]

    # ①②③ を縦に並べる。本文の各文は改行で1行ずつ表示
    for num, title, lines in steps:
        sentences_html = ''.join(
            f'<div style="white-space:nowrap;font-size:13px;color:#475569;'
            f'line-height:2;">{line}</div>'
            for line in lines
        )
        st.markdown(
            f'<div style="display:flex;align-items:flex-start;gap:16px;'
            f'padding:18px 20px;border-radius:14px;background:#FAFAFA;'
            f'border:1px solid #F0F0F0;margin-bottom:12px;">'
            f'  <div style="flex-shrink:0;width:28px;height:28px;border-radius:50%;'
            f'background:#1A1A1A;color:white;display:flex;align-items:center;'
            f'justify-content:center;font-weight:900;font-style:italic;font-size:13px;'
            f'margin-top:2px;">{num}</div>'
            f'  <div>'
            f'    <div style="font-size:13px;font-weight:800;color:#1A1A1A;'
            f'margin-bottom:6px;">{title}</div>'
            f'    {sentences_html}'
            f'  </div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    # 参考画像（static/usage_guide.png が存在する場合のみ表示）
    import os as _os
    _usage_img = _os.path.join(_os.path.dirname(__file__), 'static', 'usage_guide.png')
    if _os.path.exists(_usage_img):
        st.markdown(
            '<div style="height:4px;"></div>'
            '<div style="font-size:12px;font-weight:700;color:#475569;margin-bottom:4px;">参考画像</div>',
            unsafe_allow_html=True,
        )
        st.image(_usage_img, use_container_width=True)


# ── データ読み込み（キャッシュ） ─────────────────────────────
@st.cache_data
def load(file_bytes, file_name):
    import io
    from data_loader import load_data as _load_data
    return _load_data(io.BytesIO(file_bytes))


@st.cache_data
def validate(file_bytes, file_name):
    from data_loader import validate_excel as _validate
    return _validate(file_bytes)


# ── バリデーション エラーダイアログ ──────────────────────────────
@st.dialog('⚠️ ちょっと待って！', width='large')
def show_validation_error_modal(errors):
    for e in errors:
        st.error(e)
    st.markdown('<div style="height:8px;"></div>', unsafe_allow_html=True)
    if st.button('閉じる', use_container_width=True):
        st.rerun()


# ── 自動考察 条件ダイアログ ───────────────────────────────────
@st.dialog('自動考察 — 判定条件一覧', width='large')
def show_conditions_modal():
    # ダイアログ: 外側パネルのみスクロール、内側ブロックはスクロール無効
    st.markdown(
        '<style>'
        '[data-testid="stDialog"] > div > div {'
        '  max-height: 88vh !important;'
        '  overflow-y: auto !important;'
        '  scrollbar-width: thin !important;'
        '}'
        '[data-testid="stDialog"] [data-testid="stVerticalBlock"],'
        '[data-testid="stDialog"] [data-testid="stVerticalBlockBorderWrapper"] > div {'
        '  overflow: visible !important;'
        '  max-height: none !important;'
        '}'
        '</style>',
        unsafe_allow_html=True,
    )

    def _sh(text, size=15):
        st.markdown(
            f'<div style="font-size:{size}px;font-weight:800;color:#1A1A1A;margin:10px 0 6px 0;">{text}</div>',
            unsafe_allow_html=True,
        )

    def _sub(text):
        st.markdown(
            f'<div style="font-size:12px;color:#475569;margin-bottom:8px;line-height:1.5;">{text}</div>',
            unsafe_allow_html=True,
        )

    def _items_html(items):
        rows = ''.join(
            f'<div style="display:flex;gap:8px;padding:4px 0;border-bottom:1px solid #F8FAFC;">'
            f'<span style="color:#265F8E;font-weight:700;flex-shrink:0;">▶</span>'
            f'<span style="font-size:12px;color:#334155;line-height:1.5;">{item}</span>'
            f'</div>'
            for item in items
        )
        return rows

    def _tbl(headers, rows_data, col_styles=None):
        """汎用テーブルHTML生成"""
        ths = ''.join(
            f'<th style="padding:6px 10px;text-align:left;color:white;font-size:11px;white-space:nowrap;">{h}</th>'
            for h in headers
        )
        trs = ''
        for row in rows_data:
            tds = ''
            for i, cell in enumerate(row):
                cs = (col_styles[i] if col_styles and i < len(col_styles) else
                      'padding:6px 10px;font-size:12px;vertical-align:top;')
                tds += f'<td style="{cs}">{cell}</td>'
            trs += f'<tr style="border-bottom:1px solid #F1F5F9;">{tds}</tr>'
        return (
            '<table style="border-collapse:collapse;width:100%;background:#FAFAFA;'
            'border-radius:8px;overflow:hidden;border:1px solid #E2E8F0;">'
            f'<thead><tr style="background:#1A1A1A;">{ths}</tr></thead>'
            f'<tbody>{trs}</tbody></table>'
        )

    def _warn(text):
        st.markdown(
            f'<div style="background:#FFF7ED;border:1px solid #FED7AA;border-radius:8px;'
            f'padding:8px 12px;font-size:12px;color:#92400E;margin:6px 0;">{text}</div>',
            unsafe_allow_html=True,
        )

    def _info(text):
        st.markdown(
            f'<div style="background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;'
            f'padding:8px 12px;font-size:12px;color:#1D4ED8;margin:6px 0;">{text}</div>',
            unsafe_allow_html=True,
        )

    # ════════════════════════════════════════════
    # 1. 表現ルール
    # ════════════════════════════════════════════
    _sh('📝 自動考察の表現ルール')
    _sub('データから自動生成されるコメントは、以下の基準で表現を使い分けています。')
    st.markdown(_tbl(
        ['表現', '判定条件', '意味・例'],
        [
            ['<b style="color:#D32F2F;">必ず〜</b>',        '100%',
             '例外なく毎回そのパターンが出ている（サンプル内で100%）'],
            ['<b style="color:#D32F2F;">高確率で〜</b>',    '75% 以上',
             '4回に3回以上の頻度。強い傾向として信頼できるレベル'],
            ['<b style="color:#D32F2F;">最多は〜</b>',      '75% 未満でも最頻値',
             '最も多く出ているが支配的とまでは言えない。傾向の参考として必ず1件表示'],
            ['<b style="color:#D32F2F;">〜も多い</b>',      '全体平均より +20%pt 以上、かつ 50% 以上',
             '最頻値ではないが全体と比べて明らかに多いパターン'],
            ['<b style="color:#D32F2F;">主に〇〇大戦で</b>', '特定1大学が 75% 以上を占める',
             'ほぼ1大学のデータに偏っているため、汎用性に注意'],
        ],
        [
            'padding:6px 10px;font-size:12px;white-space:nowrap;',
            'padding:6px 10px;font-size:12px;white-space:nowrap;font-weight:700;',
            'padding:6px 10px;font-size:12px;color:#64748B;',
        ],
    ), unsafe_allow_html=True)
    _warn('⚠️ <b>サンプル数の下限：</b>3プレー未満の組み合わせは対象外。大学別の分析は 5プレー未満を対象外とします。')

    st.divider()

    # ════════════════════════════════════════════
    # 2. シチュエーション定義とフィルタ条件
    # ════════════════════════════════════════════
    _sh('🔍 シチュエーション定義（フィルタ条件）')
    _sub('各セクションのデータは以下の条件で絞り込まれます。')
    st.markdown(_tbl(
        ['シチュエーション', '適用条件'],
        [
            ['<b>Normal Situation</b>',
             'DN = 1, 2　／　2MIN 除外　／　YARD LN 1〜25（RedZone）除外'],
            ['<b>3rd Situation</b>',
             'DN = 3, 4　／　2MIN 除外　／　YARD LN 1〜25（RedZone）除外'],
            ['<b>Red Zone</b>',
             'YARD LN = 1〜25　／　2MIN 除外'],
            ['<b>2MIN</b>',
             '2MIN = Y　（DN / DIST / YARD LN 問わず）'],
        ],
        [
            'padding:6px 10px;font-size:12px;white-space:nowrap;font-weight:700;',
            'padding:6px 10px;font-size:12px;color:#334155;',
        ],
    ), unsafe_allow_html=True)

    st.divider()

    # ════════════════════════════════════════════
    # 3. データ前処理・正規化ルール
    # ════════════════════════════════════════════
    _sh('🗃️ データ前処理・正規化ルール')

    with st.expander('COVERAGE 正規化', expanded=False):
        st.markdown(_tbl(
            ['元の値', '→ 統一後'],
            [
                ['33 / ROTE3 / ROTE 3 / 3?', '→ <b>3</b>（Cover 3 に統合）'],
                ['1FREE',                    '→ <b>1 FREE</b>（表記ゆれ統一）'],
                ['3.0 / 2.0 など小数',       '→ <b>.0 を除去</b>（整数文字列に変換）'],
            ],
        ), unsafe_allow_html=True)

    with st.expander('Cover3 COMPONENT（内訳）の表示名', expanded=False):
        _sub('COMPONENT 列の値を以下の表示名に変換して集計します。')
        st.markdown(_tbl(
            ['元の値（Excel）', '→ 表示名'],
            [
                ['FS BUZZ', '→ FS Buzz'],
                ['FS SKY',  '→ FS Sky'],
                ['R BUZZ',  '→ R Buzz'],
                ['R SKY',   '→ R Sky'],
            ],
        ), unsafe_allow_html=True)

    with st.expander('SIGN(D) / BLITZ の言い換え一覧', expanded=False):
        _sub('SIGN(D) および BLITZ 列の値は以下のルールで表示名に変換されます。')
        st.markdown('<b style="font-size:12px;">完全一致（固定変換）</b>', unsafe_allow_html=True)
        st.markdown(_tbl(
            ['元の値', '→ 表示名'],
            [
                ['N',      '→ read'],
                ['FO',     '→ 広RIP'],
                ['MB',     '→ LB中ブリッツ'],
                ['MB<i>n</i>（例 MB2, MB3）', '→ LB中ブリッツ<i>n</i>枚'],
                ['CS',     '→ クロス'],
                ['OCS / OSC', '→ 外クロス'],
                ['ICS / ISC', '→ 中クロス'],
                ['TO',     '→ Tアウチャ'],
                ['EI',     '→ Eインチャ'],
                ['BOTH.A', '→ 両A-LB Blitz'],
            ],
        ), unsafe_allow_html=True)
        st.markdown('<b style="font-size:12px;margin-top:8px;display:block;">パターン変換</b>', unsafe_allow_html=True)
        st.markdown(_tbl(
            ['パターン', '→ 表示名', '例'],
            [
                ['F.A-M / F.A-W / F.A-FS / F.A-S など',
                 '→ 広<i>gap</i>-<i>LB</i> Blitz',
                 'F.A-M → 広A-Mike Blitz'],
                ['B.A-M / B.B-W / B.A-FS など',
                 '→ 狭<i>gap</i>-<i>LB</i> Blitz',
                 'B.B-W → 狭B-Willy Blitz'],
                ['BOTH.X',
                 '→ 両<i>X の変換結果</i>',
                 'BOTH.A → 両A-LB Blitz'],
                ['複合（+ 区切り）',
                 '→ 各トークンを変換して + で結合',
                 'MB+FO → LB中ブリッツ + 広RIP'],
            ],
        ), unsafe_allow_html=True)
        _info('💡 LB名称：M=Mike、W=Willy、FS=FS、S=Sam')

    with st.expander('PERSONNEL 正規化', expanded=False):
        _sub('PERSONNEL 列の値は数値の .0 を除去して文字列化します（例：10.0 → "10"）。')
        st.markdown(_tbl(
            ['グループ名', '対象PERSONNEL値'],
            [
                ['①10（10per）',        '10'],
                ['②11（11per）',        '11'],
                ['③12/21/22（12/21/22per）', '12, 21, 22'],
            ],
        ), unsafe_allow_html=True)

    with st.expander('OFF FORM 前処理・グルーピング', expanded=False):
        _sub('OFF FORM 列の先頭の "SG " プレフィックスは除去されます（例：SG ACE → ACE）。\n以下のグルーピングで集計されます。')
        st.markdown(_tbl(
            ['表示グループ名', '対象フォーメーション'],
            [
                ['ACE', 'ACE'],
                ['SUPER', 'SUPER'],
                ['WAC', 'WAC'],
                ['WAT', 'WAT'],
                ['SPREAD', 'SPREAD'],
                ['DISH', 'DISH'],
                ['SLOT（A/B/VEER合算）', 'SLOT A、SLOT B、SLOT VEER'],
                ['EMPTY', 'EMPTY を含む全て'],
                ['BUNCH合算', 'BUNCH、NEAR BUNCH、TIGHT BUNCH'],
                ['PRO合算（I/B/A）', 'PRO I、PRO B、PRO A'],
                ['PRO FAR', 'PRO FAR'],
                ['PRO NEAR', 'PRO NEAR'],
                ['TWIN合算（A/B/I）', 'TWIN A、TWIN B、TWIN I'],
                ['TWIN FAR', 'TWIN FAR'],
                ['TWIN NEAR', 'TWIN NEAR'],
                ['UMB系（合算）', 'UMB で始まる全て'],
                ['その他（個別表示）', '上記以外 → <b>3プレー以上</b>ある場合のみ表示'],
            ],
        ), unsafe_allow_html=True)

    with st.expander('RBサイド（RB_WIDE_SIDE）の算出方法', expanded=False):
        _sub('HASH / OFF STR / BACKFIELD 列から RB のワイド側を算出します。')
        st.markdown(_tbl(
            ['条件', '→ RBサイド'],
            [
                ['BACKFIELD が空', '→ <b>なし</b>'],
                ['BACKFIELD = PISTOL', '→ <b>PISTOL</b>'],
                ['HASH=L → 広い側=右（R）、BACKFIELD=ST かつ OFFSTR=R', '→ <b>広い</b>'],
                ['HASH=L → 広い側=右（R）、BACKFIELD=WE かつ OFFSTR=R', '→ <b>狭い</b>'],
                ['HASH=R → 広い側=左（L）、逆のロジック', '→ 同様に広い / 狭い'],
            ],
        ), unsafe_allow_html=True)
        _info('要素分析での表記：広い → <b>RB広</b>、狭い → <b>RB狭</b>、なし → <b>RBなし</b>')

    st.divider()

    # ════════════════════════════════════════════
    # 4. 表の備考列ルール
    # ════════════════════════════════════════════
    _sh('📋 各表の備考列ルール')
    _sub('各割合表の「備考」列には以下のルールでプレー情報が記載されます。')
    st.markdown(_tbl(
        ['表の種類', '備考列の記載条件'],
        [
            ['FRONT 表', '3プレー以下の場合、全プレーの 大学名 + PLAY# を記載'],
            ['SIGN(D) 表', '3プレー以下の場合、全プレーの 大学名 + PLAY# を記載（min_count=3 未満は行ごと非表示）'],
            ['COVERAGE 表（Cover3以外）', '3プレー以下の場合、全プレーの 大学名 + PLAY# を記載'],
            ['COVERAGE 表（Cover3）', '内訳（COMPONENT）側に備考を記載するため、Cover3 行自体は空欄'],
            ['パッケージ（3rd / RedZone）', 'プレー数に関わらず全プレーの 大学名 + PLAY# を常に記載'],
        ],
    ), unsafe_allow_html=True)

    st.divider()

    # ════════════════════════════════════════════
    # 5. 分析内容と記載場所
    # ════════════════════════════════════════════
    _sh('📍 何を分析して、どこに書かれるか')
    _sub('自動考察は Word レポート内の各箇所に出力されます。タイトルをクリックで詳細を確認できます。')

    sections = [
        (
            '📄 各 FRONT の特徴欄',
            'Word の「各Frontの特徴」欄に、そのFRONTが使われるときの傾向を書き込みます。',
            [
                'そのFRONTで使われるSIGNの最頻値を必ず表示（1位が N(read) なら2位も同一行に表示）',
                'どのOFF FORMに対してこのFRONTが多く使われるか（75%以上で記載）',
                'このFRONTが特定の1大学に75%以上集中していれば警告',
                '特定の大学でこのFRONTの使用率が全体より有意に高い場合に記載（50%以上かつ全体比+20pt以上、5プレー以上）',
                'このFRONTを使うとき、大学によってSIGNやCOVERAGEのパターンが異なれば記載（75%以上かつ全体比+20pt以上）',
                '【3rdのみ】Normalと3rdで出現頻度・最多SIGN・多く出るOFF FORMが変化するか',
            ],
        ),
        (
            '📄 各 SIGN(D) の特徴欄',
            'Word の「各Signの特徴」欄に、そのサインが出るときの傾向を書き込みます。',
            [
                'FRONTとの組み合わせ：最頻値を必ず1件 ＋ 全体比+20pt以上のFRONTも追記',
                'COVERAGEとの組み合わせ：最頻値を必ず1件 ＋ 全体比+20pt以上のCOVERAGEも追記（Cover3は内訳も括弧表示）',
                'このSIGNが特定の1大学に75%以上集中していれば警告',
                '特定の大学でこのSIGNの使用率が全体より有意に高い場合に記載（50%以上かつ全体比+20pt以上、5プレー以上）',
                'このSIGNを使うとき、大学によってFRONTやCOVERAGEのパターンが異なれば記載（75%以上かつ全体比+20pt以上）',
                'このSIGNが全体の75%以上を占めるOFF FORMやPERSONNELがあれば記載（B方向）',
                'このSIGNのプレーのうち75%以上が特定のOFF FORM / PERSONNEL / Down&Dist に集中している場合に記載（A方向）。出力例：「80%がACEで出現（8/10プレー）」',
                '【3rdのみ】出現頻度・最多FRONT・最多COVERAGEがNormalと変化するか',
            ],
        ),
        (
            '📄 各 COVERAGE の特徴欄',
            'Word の「各Coverの特徴」欄に、そのカバーが使われる状況の傾向を書き込みます。',
            [
                'FRONTとの組み合わせ：最頻値を必ず1件 ＋ 全体比+20pt以上のFRONTも追記',
                'SIGNとの組み合わせ：最頻値を必ず1件 ＋ 全体比+20pt以上のSIGNも追記',
                'このCOVERAGEが特定の1大学に75%以上集中していれば警告',
                '特定の大学でこのCOVERAGEの使用率が全体より有意に高い場合に記載（50%以上かつ全体比+20pt以上、5プレー以上）',
                'このCOVERAGEを使うとき、大学によってFRONTやSIGNのパターンが異なれば記載（75%以上かつ全体比+20pt以上）',
                '【3rdのみ】出現頻度・最多FRONT・最多SIGNがNormalと変化するか',
            ],
        ),
        (
            '📊 FRONT 表の上【自動考察ブロック】',
            'FRONT全体を横断して見えてくる傾向をまとめます。',
            [
                'FRONTとSIGNの「セット」として頻繁に出る組み合わせ上位3つ（全体の5%以上かつ3プレー以上）',
                '特定の大学でFRONTの使用率（全体比+20pt以上）や特定FRONT×SIGNの組み合わせが偏っていれば記載',
                'SIGNが5種以上に分散し最多が30%未満のFRONTを「分散」と記載',
                '【3rdのみ】Normalと比べてFRONTの構成が±20pt以上変化している場合に記載',
            ],
        ),
        (
            '📊 SIGN(D) 表の上【自動考察ブロック】',
            'ブリッツ全体の傾向と、場面ごとの変化をまとめます。',
            [
                '全体のブリッツ率（SIGN ≠ N の割合）',
                '①10per ②11per ③12/21/22per の3グループで各SIGNの使用率が全体比±20pt以上なら記載（3プレー以上）',
                '【3rdのみ】Normalと比べて3rdでSIGNの出現頻度が±20pt以上変化する場合に記載',
                'ブリッツ直後の再ブリッツ率が全体ブリッツ率と±20pt以上異なれば記載（3プレー以上）',
                '特定SIGNの直後に75%以上の確率で続くSIGNのパターンがあれば記載（3プレー以上）',
                '「SL2+MB」などの複合ブリッツが単体「SL2」と比べて特定PERSONNEL・フォームで出やすい場合（75%以上かつ+20pt差、3プレー以上）',
                '大きなゲインで1stDown更新直後（1st&10）のブリッツ率・ラッシュ人数の変化（7〜11ydsで最も差が大きい閾値を自動選択）',
                '①10per ②11per ③12/21/22per ④0per の4グループとEMPTYフォームについて5人以上ラッシュ率が全体比±20pt以上なら記載',
                'OFF FORM・PERSONNEL・Downのうち、ノーブリッツ（SIGN=N）が75%以上を占めるパターンを記載',
                '大学別にブリッツ率が±20pt以上異なれば記載。その大学で使われるSIGNの種類が変わっていれば合わせて記載（5プレー以上）',
            ],
        ),
        (
            '📊 COVERAGE 表の上【自動考察ブロック】',
            'カバー全体を横断した傾向をまとめます。',
            [
                '【3rdのみ】Normalと比べてカバーの構成が±20pt以上変化している場合に記載',
                '①10per ②11per ③12/21/22per の3グループで各COVERAGEの使用率が全体比±20pt以上なら記載',
                '特定の大学でCOVERAGEの使用率や特定COVERAGE×SIGNの組み合わせが偏っていれば記載（+20pt以上）',
                'SIGNとCOVERAGEの「セット」として頻繁に出る組み合わせ上位3つ（全体の5%以上）',
            ],
        ),
        (
            '📊 RedZone FRONT / COVERAGE 表の上【自動考察ブロック】',
            'RedZone と Normal・3rd を比較して変化しているFRONT・COVERAGEを記載します。',
            [
                'RedZone vs Normal / RedZone vs 3rd それぞれで各FRONT・COVERAGEの割合差が±20pt以上なら記載',
                '（COVERAGEのみ）①10per ②11per ③12/21/22per の3グループで割合が全体比±20pt以上なら記載',
                '特定大学でFRONT・COVERAGE使用率が全体比+20pt以上・50%以上の場合に記載（5プレー以上）',
                '特定大学でFRONT×SIGNやCOVERAGE×SIGNの組み合わせが全体比+20pt以上・50%以上の場合に記載',
            ],
        ),
        (
            '📊 RedZone パッケージ 各パッケージの特徴欄',
            'RedZoneの上位5パッケージについて個別の特徴を書き込みます。',
            [
                '大学分析：1大学のみ出現の場合はその大学名を記載',
                '大学分析：特定1大学が75%以上を占める場合は「主に〇〇大で出現」と記載',
                '大学分析：複数大学の場合は大学名と出現数の内訳を記載',
                'Down & Dist 集中：特定D&D帯（1st down / 2nd 1~3 / 3rd 7~10 など）に70%以上集中していれば記載（3プレー以上）',
            ],
        ),
        (
            '📊 RedZone ヤードゾーン別パッケージ表の上【自動考察ブロック】',
            '各ヤードゾーン内のパッケージに特定D&D集中があれば表の上に先に記載します。',
            [
                '各パッケージについてDown & Dist帯に70%以上集中していれば記載（3プレー以上）',
                '記載形式：「パッケージ名：〇〇に集中（XX%、N/Mプレー）」',
            ],
        ),
        (
            '📋 OFF FORM 別カバー表の「備考」列',
            '各フォーメーション × カバーの組み合わせについて追加の傾向を備考列に書き込みます。',
            [
                'ラッシュ人数（PRESSURE）の最頻値を常に記載。大学によってパターンが異なればその大学名も記載',
                'SIGNの最頻値を常に記載（1位がNなら2位も表示）。大学によって異なればその大学名も記載',
                '※ 組み合わせが3プレー未満の場合は記載されません',
            ],
        ),
    ]

    for title, lead, items in sections:
        with st.expander(title, expanded=False):
            st.markdown(
                f'<div style="font-size:12px;color:#64748B;padding:4px 0 8px 0;">{lead}</div>',
                unsafe_allow_html=True,
            )
            st.markdown(_items_html(items), unsafe_allow_html=True)

    st.divider()

    # ════════════════════════════════════════════
    # 6. SIGN 要素分析（シチュエーション分析）
    # ════════════════════════════════════════════
    _sh('🔬 SIGN 要素分析（シチュエーション分析）の条件')
    _sub('各SIGNがどのような状況で多く出るかを6要素の組み合わせで自動探索します。')

    with st.expander('6要素の定義と区分け', expanded=True):
        st.markdown(_tbl(
            ['要素', '説明', '区分け'],
            [
                ['<b>YD-LN</b>', 'ヤードライン',
                 '自陣深く（-1~-9yds）　自陣（-10~-49yds）　敵陣（50~26yds）　Redzone（1~25）<br>'
                 '<small>※ 敵陣のサンプルが3プレー以下の場合、自陣の区分は除外されます</small>'],
                ['<b>DN×DIST</b>', 'ダウン × 残り距離',
                 '1st down（DISTなし）　2nd 1~3 / 2nd 4~6 / 2nd 7~10 / 2nd 11+<br>'
                 '3rd 1~3 / 3rd 4~6 / 3rd 7~10 / 3rd 11+　など'],
                ['<b>パーソネル</b>', 'OFFのPERSONNEL',
                 '00per　10per　11per　12/21/22per'],
                ['<b>フォーメーション</b>', 'OFFのフォーメーション',
                 'グルーピング済みの隊形名（ACE / PRO合算 / TWIN合算 など）'],
                ['<b>前プレー</b>', '直前プレーのGN/LS',
                 '前プレーゲイン3yds以下　4~6yds　7~9yds　10+yds'],
                ['<b>RBサイド</b>', 'RBの広い側',
                 'RB広　RB狭　RBなし　PISTOL'],
            ],
        ), unsafe_allow_html=True)

    with st.expander('記載条件と冗長除去ルール', expanded=False):
        st.markdown(_items_html([
            '記載条件：ある状況（要素の組み合わせ）でそのSIGNが <b>60%以上</b> 出現し、かつ 3プレー以上の場合に記載',
            'カバレッジ条件：そのグループがSIGN全体の <b>20%以上</b> をカバーしている場合のみ（小さすぎる外れ値を除外）',
            '最大5要素の組み合わせまで探索（例：2nd 7~10 + 11per + ACE のような3要素組み合わせも検出）',
            '<b>冗長除去ルール：</b>「2nd 7~10のとき60%」と「2nd 7~10 + 11perのとき65%」が両方検出された場合 → 差が5pt以内ならシンプルな前者のみ残す。差が5pt超なら具体的な後者のみ残す',
            'すべてのSIGNで傾向小の場合：「※各SIGNは特定のシチュエーションに依存した出現傾向を示さない」と表示',
            '一部のSIGNで傾向小の場合：「その他Sign：出現シチュエーションの傾向小」とまとめて記載（個別のSIGN名は列挙しない）',
        ]), unsafe_allow_html=True)

    st.divider()

    # ════════════════════════════════════════════
    # 7. ヒートマップ（フィールド型）
    # ════════════════════════════════════════════
    _sh('🗺️ フィールド型ヒートマップの条件')

    with st.expander('セル表示ルール', expanded=False):
        st.markdown(_items_html([
            '各セル（YARD LNゾーン × フォーメーション帯）に 最頻値 と <b>(最頻値のプレー数 / 合計プレー数)</b> を表示',
            '合計プレー数が2かつ同数の最頻値が2つある場合 → 2つをセル内に並べて表示（"全Nプレー" 付き）',
            '合計プレー数が3以上かつ最頻値が同率で複数ある場合 → セルには代表1つのみ、残りは ※N 付きでヒートマップ下部に列挙',
            'カラー：出現率0〜20%=白、20〜100%=白→濃赤のグラデーション',
            'データなしのセル：グレー（#dddddd）で「—」表示',
        ]), unsafe_allow_html=True)

    with st.expander('4パネルの構成', expanded=False):
        st.markdown(_tbl(
            ['パネル', '集計列', '備考'],
            [
                ['COVERAGE', 'COVERAGE_NORM（正規化済み）', '最頻値 + 割合'],
                ['ラッシュ人数', 'PRESSURE_NUM', '最頻値 + 割合'],
                ['Blitz詳細（SIGN(D)）', 'SIGN_D', 'sign_display() 変換済みの表示名で表示'],
                ['Blitz種類（BLITZ）', 'BLITZ_CLEAN', 'sign_display() 変換済みの表示名で表示'],
            ],
        ), unsafe_allow_html=True)

    st.divider()

    # ════════════════════════════════════════════
    # 8. OFF FORM 表示条件
    # ════════════════════════════════════════════
    _sh('🏈 OFF FORM × SIGN(D) / COVERAGE 表に出る隊形の条件')

    _DEFINED_FORMS_LIST = [
        'ACE', 'SUPER', 'WAC', 'WAT', 'SPREAD', 'DISH',
        'SLOT（A/B/VEER合算）', 'EMPTY',
        'BUNCH合算', 'PRO合算（I/B/A）', 'PRO FAR', 'PRO NEAR',
        'TWIN合算（A/B/I）', 'TWIN FAR', 'TWIN NEAR', 'UMB系（合算）',
    ]
    _tags_html = ''.join(
        f'<span style="display:inline-block;background:#EFF6FF;border:1px solid #BFDBFE;'
        f'border-radius:5px;padding:2px 9px;font-size:12px;font-weight:700;'
        f'color:#1D4ED8;margin:3px 3px 3px 0;">{f}</span>'
        for f in _DEFINED_FORMS_LIST
    )
    st.markdown(
        '<div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:8px;padding:10px 14px;margin-bottom:8px;">'
        '<div style="font-size:12px;font-weight:800;color:#1A1A1A;margin-bottom:6px;">'
        '✅ 定義済み隊形 — <span style="color:#16A34A;">1プレー以上あれば必ず表示</span></div>'
        + _tags_html +
        '</div>',
        unsafe_allow_html=True,
    )
    _warn('⚠️ 上記以外（未分類）の隊形 — <b>3プレー以上で表示</b>。2プレー以下は集計対象外です。')


# ════════════════════════════════════════════════════════════
# トップバー（タイトル + 使い方 / 条件ボタンを1行に）
# ════════════════════════════════════════════════════════════
nav_l, nav_r = st.columns([6, 2])

with nav_l:
    st.markdown(
        f'<div class="rask-topbar">'
        f'  <div class="rask-nav-title">'
        f'    <span class="r" id="rask-r" style="cursor:pointer;">R</span><span class="b">avens</span>'
        f'    <span class="r sp" id="rask-a" style="cursor:pointer;">A</span><span class="b">uto</span>'
        f'    <span class="r sp" id="rask-s" style="cursor:pointer;">S</span><span class="b">cout</span>'
        f'    <span class="r sp" id="rask-k" style="cursor:pointer;">K</span><span class="b">it</span>'
        f'  </div>'
        f'</div>',
        unsafe_allow_html=True,
    )

components.html("""
<script>
(function() {
  var targets = {
    'rask-r': 'https://www.youtube.com/watch?v=FNuKjO-oUQQ',
    'rask-a': 'https://www.youtube.com/@DoABarrowRoll',
    'rask-s': 'https://youtube.com/playlist?list=PLOcWDkZNtlujYGzU33Z4aqi-t0j6RhHt-&si=wtUlOjbQ0477xXmJ',
    'rask-k': 'https://youtu.be/Kutb5vyzTls?si=z4N1lvIGximhyuw-'
  };
  var counts = { 'rask-r': 0, 'rask-a': 0, 'rask-s': 0, 'rask-k': 0 };

  function attach() {
    var doc = window.parent.document;
    var allFound = true;
    for (var id in targets) {
      if (doc.getElementById(id)) {
        (function(elemId, url) {
          var el = doc.getElementById(elemId);
          if (el && !el.dataset.raskBound) {
            el.dataset.raskBound = '1';
            el.addEventListener('click', function() {
              counts[elemId]++;
              if (counts[elemId] >= 3) {
                counts[elemId] = 0;
                window.open(url, '_blank');
              }
            });
          }
        })(id, targets[id]);
      } else {
        allFound = false;
      }
    }
    if (!allFound) {
      setTimeout(attach, 300);
    }
  }
  attach();

  // ブラウザ通知パーミッションを事前リクエスト
  if ('Notification' in window.parent && window.parent.Notification.permission === 'default') {
    window.parent.Notification.requestPermission();
  }
})();
</script>
""", height=0)

with nav_r:
    _nb1, _nb2 = st.columns(2)
    with _nb1:
        if st.button('使い方', key='usage_btn', use_container_width=True):
            show_usage_modal()
    with _nb2:
        if st.button('条件', key='cond_btn', use_container_width=True):
            show_conditions_modal()


# ════════════════════════════════════════════════════════════
# メイン 3 列レイアウト
# ════════════════════════════════════════════════════════════
def _sec(num: int, title: str, subtitle: str = ''):
    sub_html = (
        f'<span style="font-size:11px;font-weight:400;color:#64748B;margin-left:8px;">'
        f'{subtitle}</span>'
    ) if subtitle else ''
    st.markdown(
        f'<div style="display:flex;align-items:center;gap:8px;min-height:36px;'
        f'margin-bottom:4px;margin-top:0;">'
        f'<span style="width:20px;height:20px;border-radius:50%;background:#1A1A1A;color:white;'
        f'display:inline-flex;align-items:center;justify-content:center;font-size:11px;'
        f'font-weight:900;font-style:italic;flex-shrink:0;">{num}</span>'
        f'<span style="font-size:13px;font-weight:900;color:#D32F2F;'
        f'letter-spacing:0.15em;text-transform:uppercase;">{title}</span>'
        f'{sub_html}'
        f'</div>',
        unsafe_allow_html=True,
    )


col1, col2, col3 = st.columns([1, 1, 1], gap='medium')

# 左・右: 見出し1 + カード。中央: 見出し② + 大学 + 見出し③ + ヤード。
# 中央は見出しが1つ多いため、2枚の合計高さを (メインカード − 見出し相当分) にする。
_MAIN_CARD_H = 530
_RASK_EXTRA_SEC = 36
_MID_PAIR_H = _MAIN_CARD_H - _RASK_EXTRA_SEC
_MID_UNIV_H = (_MID_PAIR_H + 1) // 2
_MID_YARD_H = _MID_PAIR_H // 2


# ────────────────────────────────────────────────────────────
# ① データ入力
# ────────────────────────────────────────────────────────────
with col1:
    _sec(1, 'データ入力')
    with st.container(height=_MAIN_CARD_H, border=True):
        uploaded = st.file_uploader(
            'Excel (.xlsx)',
            type=['xlsx'],
            label_visibility='collapsed',
        )
        if not uploaded:
            # components.html は独立した iframe として描画される。
            # JS で window.parent.document の <input type="file"> を
            # .click() することでブラウザのファイル選択ダイアログを開く。
            components.html("""
<!DOCTYPE html>
<html>
<head>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: transparent; font-family: sans-serif; }

  .upload-btn {
    width: 100%;
    height: 110px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 10px;
    border: 2px dashed #E5E5E5;
    border-radius: 16px;
    background: white;
    cursor: pointer;
    transition: border-color 0.25s, background 0.25s, box-shadow 0.25s;
    user-select: none;
  }
  .upload-btn:hover {
    border-color: #D32F2F;
    background: #FFF5F5;
    box-shadow: 0 4px 20px rgba(211,47,47,0.08);
  }
  .upload-btn:active { transform: scale(0.98); }

  .icon-wrap {
    width: 44px; height: 44px;
    background: white;
    border: 1px solid #F0F0F0;
    border-radius: 11px;
    display: flex; align-items: center; justify-content: center;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    font-size: 20px;
    line-height: 1;
  }
  .label {
    font-size: 10px;
    font-weight: 800;
    color: #1A1A1A;
    letter-spacing: 0.2em;
    text-transform: uppercase;
  }
  .sub {
    font-size: 8px;
    font-weight: 700;
    color: #BDBDBD;
    letter-spacing: 0.3em;
    text-transform: uppercase;
    margin-top: -4px;
  }
</style>
</head>
<body>
  <div class="upload-btn" id="btn">
    <div class="icon-wrap">⬆</div>
    <div class="label">ゲームデータを選択</div>
    <div class="sub">EXCEL</div>
  </div>

  <script>
    document.getElementById('btn').addEventListener('click', function() {
      // 親ドキュメント（Streamlit本体）の file input を探してクリック
      try {
        var inp = window.parent.document.querySelector('input[type="file"]');
        if (inp) { inp.click(); }
      } catch(e) { console.warn('file input click failed:', e); }
    });
  </script>
</body>
</html>
""", height=118, scrolling=False)
        else:
            _val_errors = validate(uploaded.getvalue(), uploaded.name)
            # 新しいファイルでエラーがある場合はポップアップを1回だけ表示
            if _val_errors and st.session_state.get('_val_err_file') != uploaded.name:
                st.session_state['_val_err_file'] = uploaded.name
                show_validation_error_modal(_val_errors)
            # col1 のファイル表示（エラー有無でスタイルを変える）
            if _val_errors:
                st.markdown(
                    f'<div style="background:#FFF5F5;border-radius:12px;border:1px solid #FFCDD2;'
                    f'padding:16px;text-align:center;margin-top:8px;">'
                    f'<div style="font-size:20px;margin-bottom:8px;">⚠️</div>'
                    f'<div style="font-size:9px;font-weight:700;color:#C62828;'
                    f'word-break:break-all;padding:0 8px;">{uploaded.name}</div>'
                    f'<div style="font-size:7px;font-weight:700;color:#EF9A9A;'
                    f'letter-spacing:0.2em;text-transform:uppercase;margin-top:6px;">'
                    f'FORMAT ERROR</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f'<div style="background:#F8FAFC;border-radius:12px;border:1px solid #F0F0F0;'
                    f'padding:16px;text-align:center;margin-top:8px;">'
                    f'<div style="font-size:20px;margin-bottom:8px;">📊</div>'
                    f'<div style="font-size:9px;font-weight:700;color:#1A1A1A;'
                    f'word-break:break-all;padding:0 8px;">{uploaded.name}</div>'
                    f'<div style="font-size:7px;font-weight:700;color:#BDBDBD;'
                    f'letter-spacing:0.2em;text-transform:uppercase;margin-top:6px;">'
                    f'{uploaded.size // 1024} KB</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )


# ────────────────────────────────────────────────────────────
# ② 大学選択  ＋  ③ ヤード設定
# ────────────────────────────────────────────────────────────
with col2:
    # ── ② 大学選択 ──────────────────────────────────────────
    h2l, h2r = st.columns([3, 1], vertical_alignment='center')
    with h2l:
        _sec(2, '大学選択', subtitle='分析に使用する大学を選択します')
    with h2r:
        sel_all = st.button('全選択', key='btn_sel_all', use_container_width=True)

    # height 固定でアップロード後にチェックボックスが増えてもページが伸びない
    with st.container(height=_MID_UNIV_H, border=True):
        if uploaded:
            df = load(uploaded.getvalue(), uploaded.name)
            opponents = sorted(df['OPPONENT'].unique())

            _SIT_KEYS = [('Normal', 'normal'), ('3rd', '3rd'), ('Red', 'red'), ('2MIN', '2min')]

            if st.session_state.get('_opp_file') != uploaded.name:
                for _, sit in _SIT_KEYS:
                    for o in opponents:
                        st.session_state[f'opp_{sit}_{o}'] = True
                st.session_state['_opp_file'] = uploaded.name

            if sel_all:
                for _, sit in _SIT_KEYS:
                    for o in opponents:
                        st.session_state[f'opp_{sit}_{o}'] = True

            tabs = st.tabs([label for label, _ in _SIT_KEYS])
            for tab, (_, sit) in zip(tabs, _SIT_KEYS):
                with tab:
                    gcols = st.columns(2)
                    for i, opp in enumerate(opponents):
                        with gcols[i % 2]:
                            st.checkbox(opp, key=f'opp_{sit}_{opp}')
        else:
            st.markdown(
                '<div style="text-align:center;color:#BDBDBD;padding:16px 0;font-size:11px;line-height:1.7;">'
                'ファイルをアップロードすると<br>大学が表示されます'
                '</div>',
                unsafe_allow_html=True,
            )

    # ── ③ ヤード設定 ────────────────────────────────────────
    _sec(3, 'ヤード設定', subtitle='3rdのヤード区分け設定')
    with st.container(height=_MID_YARD_H, border=True):
        for zone, d_min, d_max in [
            ('SHORT',    1,   3),
            ('MIDDLE ①', 4,   6),
            ('MIDDLE ②', 7,  10),
            ('LONG',    11,  99),
        ]:
            za, zb, zc, zd, ze = st.columns([2.5, 1.2, 0.4, 1.2, 0.8])
            za.markdown(
                f'<div style="padding:0;font-weight:700;font-size:14px;'
                f'color:#475569;text-transform:uppercase;letter-spacing:0.05em;">'
                f'{zone}</div>',
                unsafe_allow_html=True,
            )
            zb.number_input('', value=d_min, min_value=0, max_value=999,
                            key=f'yd_{zone}_min', label_visibility='collapsed')
            zc.markdown(
                '<div style="text-align:center;padding:1px 0;color:#94A3B8;font-size:13px;">−</div>',
                unsafe_allow_html=True,
            )
            zd.number_input('', value=d_max, min_value=0, max_value=999,
                            key=f'yd_{zone}_max', label_visibility='collapsed')
            ze.markdown(
                '<div style="padding:1px 0;font-size:11px;color:#94A3B8;">yds</div>',
                unsafe_allow_html=True,
            )




# ────────────────────────────────────────────────────────────
# ④ 解析結果
# ────────────────────────────────────────────────────────────
with col3:
    _cur_file = uploaded.name if uploaded else None
    if st.session_state.get('_word_file') != _cur_file:
        st.session_state.pop('word_buf', None)
        st.session_state.pop('_generating', None)
        st.session_state['_word_file'] = _cur_file

    is_generated  = 'word_buf' in st.session_state
    is_generating = st.session_state.get('_generating', False)
    _has_val_error = bool(uploaded and validate(uploaded.getvalue(), uploaded.name))

    _sec(4, '解析結果')

    with st.container(height=_MAIN_CARD_H, border=True):
        # ── START ANALYSIS ボタン ────────────────────────────────
        if st.button(
            '▶  START  ANALYSIS',
            type='primary',
            use_container_width=True,
            disabled=(uploaded is None or _has_val_error),
            key='start_btn',
        ):
            st.session_state['_generating'] = True
            st.rerun()

        st.markdown('<div style="height:4px;"></div>', unsafe_allow_html=True)
        result_slot = st.empty()

        if is_generating:
            with result_slot.container():
                st.markdown(_loading_gif_html(), unsafe_allow_html=True)
            _gen_error = None
            try:
                with st.spinner(''):
                    df_report = load(uploaded.getvalue(), uploaded.name)
                    _all_opps = sorted(df_report['OPPONENT'].unique())
                    _sit_opps = {}
                    for _sit in ('normal', '3rd', 'red', '2min'):
                        _sel = [o for o in _all_opps if st.session_state.get(f'opp_{_sit}_{o}', True)]
                        _sit_opps[_sit] = _sel if _sel else _all_opps
                    _yd_zones = [
                        ('SHORT',    st.session_state.get('yd_SHORT_min',    1),
                                     st.session_state.get('yd_SHORT_max',    3)),
                        ('MIDDLE ①', st.session_state.get('yd_MIDDLE ①_min', 4),
                                     st.session_state.get('yd_MIDDLE ①_max', 6)),
                        ('MIDDLE ②', st.session_state.get('yd_MIDDLE ②_min', 7),
                                     st.session_state.get('yd_MIDDLE ②_max', 10)),
                        ('LONG',     st.session_state.get('yd_LONG_min',    11),
                                     st.session_state.get('yd_LONG_max',    99)),
                    ]
                    from report_generator import generate_word_report as _gen_word
                    st.session_state.word_buf = _gen_word(
                        df_report, situation_opps=_sit_opps, yd_zones=_yd_zones
                    )
                    _notify_pc('Ravens Auto Scout Kit', 'Wordファイルの生成が完了しました')
                    st.toast('✅ Wordファイルの生成が完了しました', icon='🏈')
                    st.session_state['_notify_done'] = True
            except Exception as _e:
                import traceback as _tb
                _gen_error = _tb.format_exc()
                st.session_state.pop('word_buf', None)
                st.session_state['_gen_error'] = _gen_error
            finally:
                st.session_state.pop('_generating', None)
            st.rerun()
        elif st.session_state.get('_gen_error'):
            with result_slot.container():
                st.error('⚠️ レポート生成中にエラーが発生しました。ターミナルのログも確認してください。')
                with st.expander('エラー詳細', expanded=True):
                    st.code(st.session_state['_gen_error'], language='python')
            if st.button('🔄 再試行', use_container_width=True):
                st.session_state.pop('_gen_error', None)
                st.rerun()
        elif is_generated:
            # ブラウザ通知（生成直後の1回のみ発火）
            if st.session_state.pop('_notify_done', False):
                components.html("""
<script>
(function() {
  function fire() {
    if (!('Notification' in window.parent)) return;
    if (window.parent.Notification.permission === 'granted') {
      new window.parent.Notification('Ravens Auto Scout Kit 🏈', {
        body: 'Wordファイルの生成が完了しました',
        icon: 'https://em-content.zobj.net/source/google/387/american-football_1f3c8.png'
      });
    }
  }
  if (!('Notification' in window.parent)) return;
  if (window.parent.Notification.permission === 'default') {
    window.parent.Notification.requestPermission().then(function(p) {
      if (p === 'granted') fire();
    });
  } else {
    fire();
  }
})();
</script>
""", height=0)
            with result_slot.container():
                st.markdown(
                    '<div style="text-align:center;padding:18px 0 14px;">'
                    '<div style="width:72px;height:72px;background:#FFF5F5;border-radius:16px;'
                    'border:2px solid #FFCDD2;display:inline-flex;align-items:center;'
                    'justify-content:center;box-shadow:0 4px 16px rgba(211,47,47,0.12);'
                    'margin-bottom:12px;font-size:36px;">📄</div>'
                    '<div style="font-size:15px;font-weight:900;color:#D32F2F;'
                    'text-transform:uppercase;letter-spacing:0.2em;margin-bottom:4px;">生成完了</div>'
                    '<div style="font-size:10px;color:#94A3B8;font-weight:600;'
                    'letter-spacing:0.1em;">レポートをダウンロードできます</div>'
                    '</div>',
                    unsafe_allow_html=True,
                )
                st.download_button(
                    label='ダウンロード  ⬇',
                    data=st.session_state.word_buf,
                    file_name='scouting_report.docx',
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    use_container_width=True,
                )
        else:
            with result_slot.container():
                st.markdown(
                    '<div style="text-align:center;padding:12px 0 10px;'
                    'opacity:0.35;filter:grayscale(1);">'
                    '<div style="width:40px;height:40px;background:white;border-radius:10px;'
                    'border:1px solid #F0F0F0;display:inline-flex;align-items:center;'
                    'justify-content:center;box-shadow:0 2px 8px rgba(0,0,0,0.04);'
                    'margin-bottom:8px;font-size:20px;">📋</div>'
                    '<div style="font-size:9px;font-weight:900;color:#1A1A1A;'
                    'text-transform:uppercase;letter-spacing:0.25em;">解析待ち</div>'
                    '</div>'
                    '<div style="text-align:center;margin-bottom:6px;">'
                    '<div style="display:inline-block;padding:8px 24px;background:#F8FAFC;'
                    'border:1px solid #F0F0F0;border-radius:10px;font-size:9px;font-weight:700;'
                    'color:#CBD5E1;letter-spacing:0.3em;text-transform:uppercase;">'
                    'ダウンロード ⬇</div>'
                    '</div>',
                    unsafe_allow_html=True,
                )


st.markdown(
    '<div style="text-align:center;padding:4px 0 2px;border-top:1px solid #F8F8F8;margin-top:6px;">'
    '<span style="font-size:7px;font-weight:700;color:#E0E0E0;letter-spacing:0.5em;text-transform:uppercase;">'
    'RAVENS PRECISION ANALYTICS SYSTEM &copy; 2024'
    '</span>'
    '</div>',
    unsafe_allow_html=True,
)
