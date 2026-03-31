import streamlit as st
import pandas as pd

from data_loader import load_data
from report_generator import generate_word_report
import analyzer as az

st.set_page_config(
    page_title='スカウティングレポート自動生成',
    page_icon='🏈',
    layout='wide',
)

st.title('🏈 スカウティングレポート 自動生成')
st.caption('Excelファイルをアップロードすると、集計表とWordレポートを自動生成します')

# ── サイドバー ──────────────────────────────────────────────
with st.sidebar:
    st.header('📂 ファイルアップロード')
    uploaded = st.file_uploader('Excelファイル (.xlsx)', type=['xlsx'])

    if uploaded:
        st.success('ファイル読み込み完了')

    st.divider()
    st.caption('© Ravens Scouting Tool')

# ── メイン ─────────────────────────────────────────────────
if not uploaded:
    st.info('👈 サイドバーからExcelファイルをアップロードしてください')
    st.stop()

# データ読み込み
@st.cache_data
def load(file_bytes, file_name):
    import io
    return load_data(io.BytesIO(file_bytes))

df = load(uploaded.getvalue(), uploaded.name)

st.success(f'読み込み完了：{len(df)} プレー ／ 対戦校：{", ".join(sorted(df["OPPONENT"].unique()))}')

# ── Word ダウンロード ──────────────────────────────────────
st.subheader('📄 Wordレポート生成')
with st.spinner('レポートを生成中...'):
    word_buf = generate_word_report(df)

st.download_button(
    label='⬇️ Wordレポートをダウンロード',
    data=word_buf,
    file_name='scouting_report.docx',
    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
)

st.divider()

# ── 集計プレビュー（タブ） ──────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    '1️⃣ Normal Situation',
    '2️⃣ 3rd Situation',
    '3️⃣ Red Zone',
    '4️⃣ 2MIN',
])

# ── Tab 1: Normal Situation ──────────────────────────────
with tab1:
    df_n = az.filter_normal(df)
    st.caption(f'該当プレー数：{len(df_n)}  （DN=1,2 ／ 2MIN除外 ／ RedZone除外）')

    col_r, col_p = st.columns(2)

    with col_r:
        st.subheader('1-1. Run Defense')

        st.markdown('**① DEF FRONT 割合**')
        front = az.analyze_front(df_n)
        st.dataframe(front, use_container_width=True, hide_index=True)

        st.markdown('**② SIGN(D) 割合**')
        sign = az.analyze_sign_d(df_n)
        st.dataframe(sign, use_container_width=True, hide_index=True)

    with col_p:
        st.subheader('1-2. Pass Defense')

        st.markdown('**① パスカバー割合（COVERAGE）**')
        cov, comp3 = az.analyze_coverage(df_n)
        st.dataframe(cov, use_container_width=True, hide_index=True)
        if len(comp3) > 0:
            st.markdown('　▼ **Cover 3 内訳（COMPONENT）**')
            st.dataframe(comp3, use_container_width=True, hide_index=True)

        st.markdown('**② OFF FORM ごとの割合**')
        form = az.analyze_off_form(df_n)
        st.dataframe(form, use_container_width=True, hide_index=True)

# ── Tab 2: 3rd Situation ─────────────────────────────────
with tab2:
    df_3 = az.filter_3rd(df)
    st.caption(f'該当プレー数：{len(df_3)}  （DN=3,4 ／ 2MIN除外 ／ RedZone除外）')

    zone_results = az.analyze_3rd_zones(df_3)

    run_tab, pass_tab = st.tabs(['2-1. Run Defense', '2-2. Pass Defense'])

    with run_tab:
        for zone_name, data in zone_results.items():
            with st.expander(f'{zone_name}  （n={data["n"]}）', expanded=data['n'] > 0):
                if data['n'] == 0:
                    st.write('該当プレーなし')
                    continue
                st.markdown('**① DEF FRONT 割合**')
                st.dataframe(data['front'], use_container_width=True, hide_index=True)

    with pass_tab:
        for zone_name, data in zone_results.items():
            with st.expander(f'{zone_name}  （n={data["n"]}）', expanded=data['n'] > 0):
                if data['n'] == 0:
                    st.write('該当プレーなし')
                    continue
                st.markdown('**② COVERAGE 割合**')
                st.dataframe(data['coverage'], use_container_width=True, hide_index=True)
                if len(data['comp3']) > 0:
                    st.markdown('　▼ **Cover 3 内訳（COMPONENT）**')
                    st.dataframe(data['comp3'], use_container_width=True, hide_index=True)
                pkg = data['packages']
                st.markdown('**③ よく出るパッケージ**')
                if len(pkg) > 0:
                    st.dataframe(pkg, use_container_width=True, hide_index=True)
                else:
                    st.write('3プレー以上の組み合わせなし')

# ── Tab 3: Red Zone ──────────────────────────────────────
with tab3:
    df_r = az.filter_redzone(df)
    st.caption(f'該当プレー数：{len(df_r)}  （YARD LN 1〜25 ／ 2MIN除外）')

    col_r2, col_p2 = st.columns(2)

    with col_r2:
        st.subheader('3-1. Run Defense')
        st.markdown('**① DEF FRONT 割合**')
        st.dataframe(az.analyze_front(df_r), use_container_width=True, hide_index=True)

    with col_p2:
        st.subheader('3-2. Pass Defense')
        st.markdown('**② COVERAGE 割合**')
        cov_r, comp3_r = az.analyze_coverage(df_r)
        st.dataframe(cov_r, use_container_width=True, hide_index=True)
        if len(comp3_r) > 0:
            st.markdown('　▼ **Cover 3 内訳（COMPONENT）**')
            st.dataframe(comp3_r, use_container_width=True, hide_index=True)

# ── Tab 4: 2MIN ─────────────────────────────────────────
with tab4:
    df_2 = az.filter_2min(df)
    st.caption(f'該当プレー数：{len(df_2)}  （2MIN = Y）')

    col_s, col_c = st.columns(2)

    with col_s:
        st.markdown('**① SIGN(D) 割合**')
        st.dataframe(az.analyze_sign_d(df_2), use_container_width=True, hide_index=True)

    with col_c:
        st.markdown('**② COVERAGE 割合**')
        cov_2, comp3_2 = az.analyze_coverage(df_2)
        st.dataframe(cov_2, use_container_width=True, hide_index=True)
        if len(comp3_2) > 0:
            st.markdown('　▼ **Cover 3 内訳（COMPONENT）**')
            st.dataframe(comp3_2, use_container_width=True, hide_index=True)
