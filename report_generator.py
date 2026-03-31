import io
from datetime import date

import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

import analyzer as az


# ── スタイルヘルパー ────────────────────────────────────────
def _set_row_bold(row):
    for cell in row.cells:
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True


def _shade_cell(cell, hex_color='D9E1F2'):
    from docx.oxml import OxmlElement
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def _add_table(doc, df, title=None, header_color='D9E1F2', indent_level=0):
    """DataFrameをWordテーブルとして追記する"""
    if df is None or len(df) == 0:
        if title:
            p = doc.add_paragraph(f'  {'　' * indent_level}{title}：該当データなし')
            p.runs[0].italic = True
        return

    if title:
        p = doc.add_paragraph()
        indent_pt = indent_level * 360  # 1レベル = 0.5インチ相当
        p.paragraph_format.left_indent = indent_pt
        run = p.add_run(title)
        run.bold = True

    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = 'Table Grid'

    # ヘッダー行
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df.columns):
        hdr_cells[i].text = str(col_name)
        _shade_cell(hdr_cells[i], header_color)
    _set_row_bold(table.rows[0])

    # データ行
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = '' if pd.isna(val) else str(val)

    doc.add_paragraph()  # 余白


# ── レポート生成 ─────────────────────────────────────────────
def generate_word_report(df):
    doc = Document()

    # ── 表紙 ──
    doc.add_heading('スカウティングレポート', 0)
    doc.add_paragraph(f'生成日：{date.today().strftime("%Y年%m月%d日")}')
    doc.add_paragraph(f'対象チーム：{", ".join(sorted(df["OPPONENT"].unique()))}')
    doc.add_paragraph(f'総プレー数：{len(df)}')
    doc.add_page_break()

    # ════════════════════════════════════════
    # 1. Normal Situation
    # ════════════════════════════════════════
    doc.add_heading('1. Normal Situation', 1)
    doc.add_paragraph('フィルタ条件：DN=1,2 ／ 2MIN除外 ／ RedZone（YARD LN 1〜25）除外')

    df_n = az.filter_normal(df)
    doc.add_paragraph(f'該当プレー数：{len(df_n)}')

    # 1-1 Run Defense
    doc.add_heading('1-1. Run Defense', 2)

    _add_table(doc, az.analyze_front(df_n), '① DEF FRONT 割合')
    _add_table(doc, az.analyze_sign_d(df_n), '② SIGN(D) 割合')

    # 1-2 Pass Defense
    doc.add_heading('1-2. Pass Defense', 2)

    cov, comp3 = az.analyze_coverage(df_n)
    _add_table(doc, cov, '① パスカバー割合（COVERAGE）')
    if len(comp3) > 0:
        _add_table(doc, comp3, '　▼ Cover 3 内訳（COMPONENT）', header_color='E2EFDA', indent_level=1)

    _add_table(doc, az.analyze_off_form(df_n), '② OFF FORM ごとの割合')

    doc.add_page_break()

    # ════════════════════════════════════════
    # 2. 3rd Situation
    # ════════════════════════════════════════
    doc.add_heading('2. 3rd Situation', 1)
    doc.add_paragraph('フィルタ条件：DN=3,4 ／ 2MIN除外 ／ RedZone除外')

    df_3 = az.filter_3rd(df)
    doc.add_paragraph(f'該当プレー数：{len(df_3)}')

    zone_results = az.analyze_3rd_zones(df_3)

    # Run Defense
    doc.add_heading('2-1. Run Defense', 2)
    for zone_name, data in zone_results.items():
        n = data['n']
        doc.add_heading(f'{zone_name}  （n={n}）', 3)
        if n == 0:
            doc.add_paragraph('  該当プレーなし')
            continue
        _add_table(doc, data['front'], '① DEF FRONT 割合')

    # Pass Defense
    doc.add_heading('2-2. Pass Defense', 2)
    for zone_name, data in zone_results.items():
        n = data['n']
        doc.add_heading(f'{zone_name}  （n={n}）', 3)
        if n == 0:
            doc.add_paragraph('  該当プレーなし')
            continue
        _add_table(doc, data['coverage'], '② COVERAGE 割合')
        if len(data['comp3']) > 0:
            _add_table(doc, data['comp3'], '　▼ Cover 3 内訳（COMPONENT）',
                       header_color='E2EFDA', indent_level=1)
        pkg = data['packages']
        if len(pkg) > 0:
            _add_table(doc, pkg, '③ よく出るパッケージ（3プレー以上）', header_color='FCE4D6')
        else:
            doc.add_paragraph('  ③ よく出るパッケージ：なし')

    doc.add_page_break()

    # ════════════════════════════════════════
    # 3. Red Zone
    # ════════════════════════════════════════
    doc.add_heading('3. Red Zone', 1)
    doc.add_paragraph('フィルタ条件：YARD LN 1〜25 ／ 2MIN除外')

    df_r = az.filter_redzone(df)
    doc.add_paragraph(f'該当プレー数：{len(df_r)}')

    # Run Defense
    doc.add_heading('3-1. Run Defense', 2)
    _add_table(doc, az.analyze_front(df_r), '① DEF FRONT 割合')

    # Pass Defense
    doc.add_heading('3-2. Pass Defense', 2)
    cov, comp3 = az.analyze_coverage(df_r)
    _add_table(doc, cov, '② COVERAGE 割合')
    if len(comp3) > 0:
        _add_table(doc, comp3, '　▼ Cover 3 内訳（COMPONENT）', header_color='E2EFDA', indent_level=1)

    doc.add_page_break()

    # ════════════════════════════════════════
    # 4. 2MIN
    # ════════════════════════════════════════
    doc.add_heading('4. 2MIN', 1)
    doc.add_paragraph('フィルタ条件：2MIN = Y（DN・DIST・YARD LN によらず）')

    df_2 = az.filter_2min(df)
    doc.add_paragraph(f'該当プレー数：{len(df_2)}')

    _add_table(doc, az.analyze_sign_d(df_2), '① SIGN(D) 割合')

    cov, comp3 = az.analyze_coverage(df_2)
    _add_table(doc, cov, '② COVERAGE 割合')
    if len(comp3) > 0:
        _add_table(doc, comp3, '　▼ Cover 3 内訳（COMPONENT）', header_color='E2EFDA', indent_level=1)

    # ── BytesIO に保存して返す ──
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf
