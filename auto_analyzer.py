"""
自動考察生成モジュール (auto_analyzer.py)
各FRONT / SIGN / COVERAGEの傾向を自動検出し、
Wordレポートに挿入するテキストリストを返す。
"""

import pandas as pd
from collections import defaultdict
from itertools import combinations as _combinations
from analyzer import (sign_display as _sign_display,
                      _get_form_group as _get_form_group_az,
                      build_redzone_pkg_map as _build_redzone_pkg_map)

# ─── 閾値 ────────────────────────────────────────────────────
MIN_N        = 3     # 基本最低サンプル数
MIN_N_OPP    = 5     # 大学別分析の最低サンプル数
THR_HIGH     = 0.75  # 高確率 / 大学集中（共通）
THR_ABS      = 0.50  # 傾向あり（絶対値下限）
THR_DIFF     = 0.20  # 傾向あり（全体との差）/ Normal vs 3rd 差（共通）
THR_PREDICT_SCATTER = 0.30  # 分散判定（最多SIGNがこれ未満かつ5種以上で「分散」）

# PERSONNELグループ定義
PERSONNEL_GROUPS = [
    ('10per', ['10']),
    ('11per', ['11']),
    ('12/21/22per', ['12', '21', '22']),
]
PERSONNEL_GROUPS_PRESSURE = [
    ('10per', ['10']),
    ('11per', ['11']),
    ('12/21/22per', ['12', '21', '22']),
    ('0per', ['0']),
]


# ══════════════════════════════════════════════════════════════
# 基本ヘルパー
# ══════════════════════════════════════════════════════════════

def _s(series):
    """文字列クリーニング: nan/空文字を '' に統一"""
    return series.astype(str).str.strip().replace({'nan': '', 'NaN': '', 'None': ''})


def _counts(df, col):
    """列のvalue_counts（空文字除外）"""
    return _s(df[col]).replace('', pd.NA).dropna().value_counts()


def _fmt(pct, cnt, n):
    """割合とプレー数を必ず両方示す"""
    return f'{pct:.0%}（{cnt}/{n}プレー）'


def _display_val(col, val):
    """列に応じた表示用の値に変換する"""
    v = str(val)
    if col in ('SIGN_D', 'BLITZ', 'BLITZ_CLEAN'):
        return _sign_display(v)
    if col == 'COVERAGE_NORM':
        try:
            int(v)
            return f'COVER{v}'
        except ValueError:
            return v
    if col == 'PERSONNEL_NORM':
        display_v = '00' if v == '0' else v
        return f'{display_v}per'
    if col == 'DN':
        return f'{v}Down'
    return v


def _col_prefix(col):
    """列名を日本語プレフィックスに変換する"""
    mapping = {
        'DEF_FRONT':      'FRONTは',
        'SIGN_D':         'SIGNは',
        'COVERAGE_NORM':  'カバーは',
        'OFF_FORM_NORM':  'フォームは',
        'PERSONNEL_NORM': 'PERSONNELは',
        'RB_WIDE_SIDE':   'RBは',
        'DN':             'ダウンは',
    }
    return mapping.get(col, '')


def _fmt_top(col, val, cnt, n, pct):
    """最頻値の表示文字列（必ず/高確率/最多は）"""
    dv = _display_val(col, val)
    # 末尾の「は」を除去して「〇〇は最多は」の二重「は」を防ぐ
    prefix = _col_prefix(col).rstrip('は')
    if pct == 1.0:
        return f'{prefix}必ず{dv}（100%、{cnt}/{n}プレー）'
    elif pct >= THR_HIGH:
        return f'{prefix}高確率で{dv}（{_fmt(pct, cnt, n)}）'
    else:
        return f'{prefix}最多は{dv}（{_fmt(pct, cnt, n)}）'


# ══════════════════════════════════════════════════════════════
# 分析ヘルパー
# ══════════════════════════════════════════════════════════════

def _sign_top(df_sub, min_n=MIN_N):
    """SIGNの最頻値を1行で常に表示。1位がNなら2位も同一行に追記（1項目として返す）"""
    n = len(df_sub)
    if n < min_n or 'SIGN_D' not in df_sub.columns:
        return []
    vc = _counts(df_sub, 'SIGN_D')
    if len(vc) == 0:
        return []
    top_val = vc.index[0]
    top_cnt = vc.iloc[0]
    top_pct = top_cnt / n
    top_dv = _display_val('SIGN_D', top_val)

    if top_pct == 1.0:
        label = f'必ず{top_dv}（100%、{top_cnt}/{n}プレー）'
    elif top_pct >= THR_HIGH:
        label = f'高確率で{top_dv}（{_fmt(top_pct, top_cnt, n)}）'
    else:
        label = f'最多は{top_dv}（{_fmt(top_pct, top_cnt, n)}）'

    if top_val == 'N' and len(vc) >= 2:
        sec_val = vc.index[1]
        sec_cnt = vc.iloc[1]
        sec_pct = sec_cnt / n
        sec_dv = _display_val('SIGN_D', sec_val)
        label += f'、2番目は{sec_dv}（{_fmt(sec_pct, sec_cnt, n)}）'

    return [f'SIGN{label}']


def _top_and_notable(df_sub, col, df_all, min_n=MIN_N):
    """
    最頻値を常に1件表示 + 全体比+20pt以上のものも追記（重複除く）
    SIGN以外の列（FRONT、COVERAGEなど）に使用
    """
    n = len(df_sub)
    N = len(df_all)
    if n < min_n or N == 0 or col not in df_sub.columns:
        return []
    sub_vc = _counts(df_sub, col)
    all_vc = _counts(df_all, col) if col in df_all.columns else pd.Series(dtype=int)
    if len(sub_vc) == 0:
        return []

    results = []
    shown = set()

    # 最頻値（常に1件）
    top_val, top_cnt = sub_vc.index[0], sub_vc.iloc[0]
    top_pct = top_cnt / n
    results.append(_fmt_top(col, top_val, top_cnt, n, top_pct))
    shown.add(top_val)

    # 全体比+20pt以上のもの（最頻値以外も含む）
    for val, cnt in sub_vc.items():
        if val in shown or cnt < min_n:
            continue
        cnt_all = all_vc.get(val, 0)
        pct = cnt / n
        pct_all = cnt_all / N
        if pct >= THR_ABS and (pct - pct_all) >= THR_DIFF:
            dv = _display_val(col, val)
            prefix = _col_prefix(col).rstrip('は')
            results.append(
                f'{prefix}{dv}も多い（{pct:.0%}（{cnt}/{n}プレー） vs 全体{pct_all:.0%}（{cnt_all}/{N}プレー））')
            shown.add(val)

    return results


def _scan_context(df_all, context_col, target_col, target_val,
                  threshold=THR_HIGH, min_n=MIN_N):
    """
    B方向スキャン: context_colの各値に対して target_col==target_val の出現率を調べ、
    threshold以上のものをまとめて1項目として返す（複数ヒット時は名称を列挙）
    """
    matches = []  # (ctx_dv, pct, cnt, n_ctx)
    target_str = str(target_val)
    is_read_mode = (target_col == 'SIGN_D' and target_str == 'N')

    for ctx_val in _counts(df_all, context_col).index:
        df_ctx = df_all[_s(df_all[context_col]) == str(ctx_val)]
        n_ctx = len(df_ctx)
        if n_ctx < min_n:
            continue
        cnt = (_s(df_ctx[target_col]) == target_str).sum()
        if cnt < min_n:
            continue
        pct = cnt / n_ctx
        if pct >= threshold:
            ctx_dv = _display_val(context_col, ctx_val)
            matches.append((ctx_dv, pct, cnt, n_ctx))

    if not matches:
        return []

    action = 'ではreadが多い' if is_read_mode else 'において使用頻度高い'

    if len(matches) == 1:
        ctx_dv, pct, cnt, n_ctx = matches[0]
        return [f'{ctx_dv}{action}（{ctx_dv}中{pct:.0%}、{cnt}/{n_ctx}プレー）']
    else:
        names = '・'.join(dv for dv, _, _, _ in matches)
        details = '、'.join(
            f'{dv} {pct:.0%}（{cnt}/{n_ctx}）' for dv, pct, cnt, n_ctx in matches)
        return [f'{names}{action}（{details}）']


def _check_opp(df_sub, min_n=MIN_N):
    """特定大学への集中を検出（基本）"""
    n = len(df_sub)
    if n < min_n or 'OPPONENT' not in df_sub.columns:
        return []
    vc = _counts(df_sub, 'OPPONENT')
    if len(vc) == 0:
        return []
    if len(vc) == 1:
        opp, cnt = vc.index[0], vc.iloc[0]
        return [f'{opp}大戦のみで出現（100%、{cnt}/{n}プレー）']
    top, cnt = vc.index[0], vc.iloc[0]
    pct = cnt / n
    if pct >= THR_HIGH:
        return [f'主に{top}大戦で出現（{_fmt(pct, cnt, n)}）']
    return []


def _check_opp_pattern_diff(df_sub, compare_col, min_n=MIN_N):
    """
    df_sub（このSIGN/FRONT/COVに絞ったデータ）において、
    compare_colのパターンが大学ごとに異なるか検出。
    「このSIGNのとき〇〇大だけFRONTが違う」の実装。
    """
    n_total = len(df_sub)
    if n_total < min_n or 'OPPONENT' not in df_sub.columns:
        return []
    if compare_col not in df_sub.columns:
        return []

    overall_vc = _counts(df_sub, compare_col)
    if len(overall_vc) == 0:
        return []

    results = []
    for opp in _counts(df_sub, 'OPPONENT').index:
        df_opp = df_sub[_s(df_sub['OPPONENT']) == str(opp)]
        n_opp = len(df_opp)
        if n_opp < min_n:
            continue
        opp_vc = _counts(df_opp, compare_col)
        for val, cnt in opp_vc.items():
            cnt_all = overall_vc.get(val, 0)
            pct_opp = cnt / n_opp
            pct_all = cnt_all / n_total
            if pct_opp >= THR_HIGH and (pct_opp - pct_all) >= THR_DIFF and cnt >= min_n:
                dv = _display_val(compare_col, val)
                prefix = _col_prefix(compare_col)
                results.append(
                    f'{opp}大では{prefix}{dv}が多い'
                    f'（{opp}大{pct_opp:.0%}（{cnt}/{n_opp}プレー） vs 全体{pct_all:.0%}（{cnt_all}/{n_total}プレー））')

    return results


def _check_opp_usage_rate(df_all, col, val, min_n_opp=MIN_N_OPP):
    """
    大学ごとの col=val 使用率が全体使用率より有意に高い大学を検出。
    例：「XX大ではこのSIGNが多い（XX大70% vs 全体40%、8/20プレー）」
    _check_opp（集中度）とは別視点——大学側からの使用傾向。
    """
    total = len(df_all)
    if total == 0 or col not in df_all.columns or 'OPPONENT' not in df_all.columns:
        return []

    val_str = str(val)
    cnt_all = (_s(df_all[col]) == val_str).sum()
    pct_all = cnt_all / total
    if pct_all == 0:
        return []

    results = []
    for opp in _counts(df_all, 'OPPONENT').index:
        df_opp = df_all[_s(df_all['OPPONENT']) == str(opp)]
        n_opp = len(df_opp)
        if n_opp < min_n_opp:
            continue
        cnt = (_s(df_opp[col]) == val_str).sum()
        if cnt < MIN_N:
            continue
        pct_opp = cnt / n_opp
        if pct_opp >= THR_ABS and (pct_opp - pct_all) >= THR_DIFF:
            results.append(
                f'{opp}大戦での使用率が高い'
                f'（{opp}大{pct_opp:.0%}（{cnt}/{n_opp}プレー） vs 全体{pct_all:.0%}（{cnt_all}/{total}プレー））')

    return results


def _check_opp_modal_diff(df, col, min_n_opp=MIN_N_OPP):
    """
    各大学の最頻値が全体の最頻値と異なれば記載。
    「〇〇大では最頻値が△△（XX%、N/Mプレー）、全体は□□」の形式で返す。
    """
    total = len(df)
    if total == 0 or col not in df.columns:
        return []
    overall_vc = _counts(df, col)
    if len(overall_vc) == 0:
        return []
    overall_modal = overall_vc.index[0]
    overall_dv    = _display_val(col, overall_modal)
    results = []
    for opp in _counts(df, 'OPPONENT').index:
        df_opp = df[_s(df['OPPONENT']) == str(opp)]
        n_opp = len(df_opp)
        if n_opp < min_n_opp:
            continue
        opp_vc = _counts(df_opp, col)
        if len(opp_vc) == 0:
            continue
        opp_modal = opp_vc.index[0]
        if opp_modal != overall_modal:
            opp_dv  = _display_val(col, opp_modal)
            opp_cnt = int(opp_vc.iloc[0])
            opp_pct = opp_cnt / n_opp
            results.append(
                f'{opp}大では最頻値が{opp_dv}（{opp_pct:.0%}、{opp_cnt}/{n_opp}プレー）'
                f'、全体は{overall_dv}')
    return results


def _check_n3(val, col, df_n, df_3, min_n=MIN_N):
    """Normal vs 3rd — valの出現頻度変化"""
    if df_n is None or df_3 is None:
        return []
    if col not in df_n.columns or col not in df_3.columns:
        return []
    n_n, n_3 = len(df_n), len(df_3)
    if n_n < min_n or n_3 < min_n:
        return []

    cnt_n = (_s(df_n[col]) == str(val)).sum()
    cnt_3 = (_s(df_3[col]) == str(val)).sum()
    pct_n = cnt_n / n_n
    pct_3 = cnt_3 / n_3
    diff = pct_3 - pct_n

    if abs(diff) < THR_DIFF:
        return []
    if diff > 0:
        if cnt_n == 0:
            return [f'Normalでは出現せず3rdで出現（{_fmt(pct_3, cnt_3, n_3)}）']
        return [f'Normalと比べて3rdで増加（Normal {pct_n:.0%} → 3rd {pct_3:.0%}、{cnt_3}/{n_3}プレー）']
    else:
        if cnt_3 == 0:
            return [f'3rdでは出現しない（Normal {pct_n:.0%}、{cnt_n}/{n_n}プレー）']
        return [f'Normalと比べて3rdで減少（Normal {pct_n:.0%} → 3rd {pct_3:.0%}、{cnt_3}/{n_3}プレー）']


def _check_n3_top(filter_val, filter_col, compare_col, df_normal, df_3rd, min_n=MIN_N):
    """
    Normal vs 3rd — filter_col==filter_val のサブセットで
    compare_col の最頻値・割合が変わるか確認
    """
    if df_normal is None or df_3rd is None:
        return []
    for df in [df_normal, df_3rd]:
        if filter_col not in df.columns or compare_col not in df.columns:
            return []

    df_n_sub = df_normal[_s(df_normal[filter_col]) == str(filter_val)]
    df_3_sub = df_3rd[_s(df_3rd[filter_col]) == str(filter_val)]
    if len(df_n_sub) < min_n or len(df_3_sub) < min_n:
        return []

    n_vc = _counts(df_n_sub, compare_col)
    t_vc = _counts(df_3_sub, compare_col)
    if len(n_vc) == 0 or len(t_vc) == 0:
        return []

    top_n = n_vc.index[0]
    top_3 = t_vc.index[0]
    pct_n = n_vc.iloc[0] / len(df_n_sub)
    pct_3 = t_vc.iloc[0] / len(df_3_sub)
    prefix = _col_prefix(compare_col)

    if top_n != top_3:
        dv_n = _display_val(compare_col, top_n)
        dv_3 = _display_val(compare_col, top_3)
        return [f'3rdでは{prefix}最多が{dv_n}（Normal）→ {dv_3}（3rd {pct_3:.0%}）に変化']
    elif abs(pct_3 - pct_n) >= THR_DIFF:
        dv = _display_val(compare_col, top_n)
        trend = '増加' if pct_3 > pct_n else '減少'
        return [f'3rdでは{prefix}{dv}が{trend}（Normal {pct_n:.0%} → 3rd {pct_3:.0%}）']

    return []


# ══════════════════════════════════════════════════════════════
# PART 1：各項目の特徴量
# ══════════════════════════════════════════════════════════════

def _cover3_comp_note(df_sub):
    """COVERAGE_NORM='3'の行のCOMPONENT_DISPLAY内訳を括弧書きで返す"""
    if 'COMPONENT_DISPLAY' not in df_sub.columns:
        return ''
    df_c3 = df_sub[_s(df_sub['COVERAGE_NORM']) == '3']
    n_c3 = len(df_c3)
    if n_c3 < MIN_N:
        return ''
    comp_vc = _counts(df_c3, 'COMPONENT_DISPLAY')
    if len(comp_vc) == 0:
        return ''
    parts = [f'{cv} {cnt/n_c3:.0%}' for cv, cnt in comp_vc.items()]
    return f'（内訳：{" / ".join(parts)}）'

def _check_sign_play_concentration(df_sub, min_n=MIN_N, threshold=THR_HIGH):
    """
    A方向チェック：このSIGNのプレーのうち threshold(75%) 以上が
    特定の OFF FORM / PERSONNEL / Down&Dist に集中していれば記載。
    """
    n = len(df_sub)
    if n < min_n:
        return []
    results = []

    # ── OFF FORM（グループ化して集計） ──────────────────────────
    if 'OFF_FORM_NORM' in df_sub.columns:
        form_grps = _s(df_sub['OFF_FORM_NORM']).apply(
            lambda v: _get_form_group_az(v) or '')
        form_vc = form_grps.replace('', pd.NA).dropna().value_counts()
        for fgrp, cnt in form_vc.items():
            if cnt < min_n:
                continue
            pct = cnt / n
            if pct >= threshold:
                results.append(f'{pct:.0%}が{fgrp}で出現（{cnt}/{n}プレー）')

    # ── PERSONNEL ────────────────────────────────────────────────
    if 'PERSONNEL_NORM' in df_sub.columns:
        def _pg(v):
            v = str(v).strip()
            if v == '0':             return '00per'
            if v == '10':            return '10per'
            if v == '11':            return '11per'
            if v in ('12','21','22'): return '12/21/22per'
            return ''
        pers_grps = _s(df_sub['PERSONNEL_NORM']).apply(_pg)
        pers_vc = pers_grps.replace('', pd.NA).dropna().value_counts()
        for pgrp, cnt in pers_vc.items():
            if cnt < min_n:
                continue
            pct = cnt / n
            if pct >= threshold:
                results.append(f'{pct:.0%}が{pgrp}で出現（{cnt}/{n}プレー）')

    # ── Down & Dist ───────────────────────────────────────────────
    if 'DN' in df_sub.columns and 'DIST' in df_sub.columns:
        dd_labels = df_sub.apply(_pkg_dn_dist_label, axis=1)
        dd_vc = dd_labels.replace('', pd.NA).dropna().value_counts()
        for ddval, cnt in dd_vc.items():
            if cnt < min_n:
                continue
            pct = cnt / n
            if pct >= threshold:
                results.append(f'{pct:.0%}が{ddval}で出現（{cnt}/{n}プレー）')

    return results


def consider_front_item(front_val, df_all, df_normal=None, df_3rd=None):
    """特定FRONT値の考察テキストリストを生成"""
    df_sub = df_all[_s(df_all['DEF_FRONT']) == str(front_val)]
    b = []

    # ① SIGNの最頻値（常に表示、Nなら2位も）
    b += _sign_top(df_sub)

    # ② このFRONTが多く使われるOFF FORM（B方向）
    b += _scan_context(df_all, 'OFF_FORM_NORM', 'DEF_FRONT', front_val)

    # ③ 大学集中（基本）
    b += _check_opp(df_sub)

    # ③ 大学別使用率比較（大学側の視点）
    b += _check_opp_usage_rate(df_all, 'DEF_FRONT', front_val)

    # ③ 大学別SIGN/COVERAGEパターン差（このFRONTを使う時に大学によって異なるか）
    b += _check_opp_pattern_diff(df_sub, 'SIGN_D')
    b += _check_opp_pattern_diff(df_sub, 'COVERAGE_NORM')

    # ⑤ 3rdのみ: 頻度変化 + SIGN傾向変化 + OFF FORM傾向変化
    if df_normal is not None and df_3rd is not None:
        b += _check_n3(front_val, 'DEF_FRONT', df_normal, df_3rd)
        b += _check_n3_top(front_val, 'DEF_FRONT', 'SIGN_D', df_normal, df_3rd)
        b += _check_n3_top(front_val, 'DEF_FRONT', 'OFF_FORM_NORM', df_normal, df_3rd)

    return b


def consider_sign_item(sign_val, df_all, df_normal=None, df_3rd=None):
    """特定SIGN値の考察テキストリストを生成"""
    sign_str = str(sign_val)
    # sign_val は表示名（sign_display済み）の可能性があるため逆引きを試みる
    mask = _s(df_all['SIGN_D']) == sign_str
    if mask.sum() == 0:
        for raw in _s(df_all['SIGN_D']).unique():
            if raw and _sign_display(raw) == sign_str:
                sign_str = raw
                mask = _s(df_all['SIGN_D']) == raw
                break
    df_sub = df_all[mask]
    b = []

    # ① FRONTとの組み合わせ（最頻値 + 全体比+20pt）
    b += _top_and_notable(df_sub, 'DEF_FRONT', df_all)

    # ② COVERAGEとの組み合わせ（最頻値 + 全体比+20pt）、Cover3は内訳も追記
    cov_bullets = _top_and_notable(df_sub, 'COVERAGE_NORM', df_all)
    comp_note = _cover3_comp_note(df_sub)
    if comp_note:
        cov_bullets = [
            bullet + comp_note if 'COVER3' in bullet else bullet
            for bullet in cov_bullets
        ]
    b += cov_bullets

    # ③ 大学集中（基本）
    b += _check_opp(df_sub)

    # ③ 大学別使用率比較（大学側の視点）※生値で比較
    b += _check_opp_usage_rate(df_all, 'SIGN_D', sign_str)

    # ③ 大学別FRONT/COVERAGEパターン差
    b += _check_opp_pattern_diff(df_sub, 'DEF_FRONT')
    b += _check_opp_pattern_diff(df_sub, 'COVERAGE_NORM')

    # ④ OFF FORM・PERSONNELで出現率が高い場合（B方向）※生値で比較
    b += _scan_context(df_all, 'OFF_FORM_NORM', 'SIGN_D', sign_str)
    b += _scan_context(df_all, 'PERSONNEL_NORM', 'SIGN_D', sign_str)

    # ④' このSIGNのプレー内でOFF FORM / PERSONNEL / Down&Dist が75%以上を占める（A方向）
    b += _check_sign_play_concentration(df_sub)

    # ⑤ 3rdのみ: 頻度変化 + FRONT傾向変化 + COVERAGE傾向変化 ※生値で比較
    if df_normal is not None and df_3rd is not None:
        b += _check_n3(sign_str, 'SIGN_D', df_normal, df_3rd)
        b += _check_n3_top(sign_str, 'SIGN_D', 'DEF_FRONT', df_normal, df_3rd)
        b += _check_n3_top(sign_str, 'SIGN_D', 'COVERAGE_NORM', df_normal, df_3rd)

    return b


def consider_cov_item(cov_val, df_all, df_normal=None, df_3rd=None):
    """特定COVERAGE値の考察テキストリストを生成"""
    df_sub = df_all[_s(df_all['COVERAGE_NORM']) == str(cov_val)]
    b = []

    # ① FRONTとの組み合わせ（最頻値 + 全体比+20pt）
    b += _top_and_notable(df_sub, 'DEF_FRONT', df_all)

    # ① SIGNとの組み合わせ（最頻値 + 全体比+20pt）
    b += _top_and_notable(df_sub, 'SIGN_D', df_all)

    # ② 大学集中（基本）
    b += _check_opp(df_sub)

    # ② 大学別使用率比較（大学側の視点）
    b += _check_opp_usage_rate(df_all, 'COVERAGE_NORM', cov_val)

    # ② 大学別FRONT/SIGNパターン差
    b += _check_opp_pattern_diff(df_sub, 'DEF_FRONT')
    b += _check_opp_pattern_diff(df_sub, 'SIGN_D')

    # ③ 3rdのみ: 頻度変化 + FRONT傾向変化 + SIGN傾向変化
    if df_normal is not None and df_3rd is not None:
        dv = _display_val('COVERAGE_NORM', cov_val)
        for obs in _check_n3(cov_val, 'COVERAGE_NORM', df_normal, df_3rd):
            b.append(f'{dv}：{obs}')
        b += _check_n3_top(cov_val, 'COVERAGE_NORM', 'DEF_FRONT', df_normal, df_3rd)
        b += _check_n3_top(cov_val, 'COVERAGE_NORM', 'SIGN_D', df_normal, df_3rd)

    return b


# ══════════════════════════════════════════════════════════════
# PART 2：表の上（横断的考察）
# ══════════════════════════════════════════════════════════════

def consider_front_header(df, df_normal=None, df_3rd=None):
    """FRONT表の上の横断考察"""
    bullets = []
    total = len(df)
    if total == 0:
        return bullets

    # ① FRONT×SIGN 頻出組み合わせ上位3（全体の5%以上）
    pairs = []
    for fv in _counts(df, 'DEF_FRONT').index:
        df_f = df[_s(df['DEF_FRONT']) == str(fv)]
        for sv, cnt in _counts(df_f, 'SIGN_D').items():
            pct = cnt / total
            if pct >= 0.05 and cnt >= MIN_N:
                pairs.append((pct, cnt, fv, sv))
    _RANK_LABEL = ['最多', '2番目に多い', '3番目に多い']
    pairs.sort(reverse=True)
    for rank, (pct, cnt, fv, sv) in enumerate(pairs[:3]):
        label = _RANK_LABEL[rank]
        bullets.append(
            f'{fv}+{_display_val("SIGN_D", sv)}の組み合わせが{label}（全体の{pct:.0%}、{cnt}/{total}プレー）')

    # ③ 大学別FRONT偏り + FRONT×SIGNの大学別偏り
    front_all_vc = _counts(df, 'DEF_FRONT')
    for opp in _counts(df, 'OPPONENT').index:
        df_opp = df[_s(df['OPPONENT']) == str(opp)]
        n_opp = len(df_opp)
        if n_opp < MIN_N_OPP:
            continue
        # FRONT単体の偏り
        for fv, cnt in _counts(df_opp, 'DEF_FRONT').items():
            cnt_fv_all = front_all_vc.get(fv, 0)
            pct = cnt / n_opp
            pct_all = cnt_fv_all / total
            if pct >= THR_ABS and (pct - pct_all) >= THR_DIFF and cnt >= MIN_N:
                bullets.append(
                    f'{opp}大では{fv}が多い'
                    f'（{pct:.0%}（{cnt}/{n_opp}プレー） vs 全体{pct_all:.0%}（{cnt_fv_all}/{total}プレー））')
        # FRONT×SIGNの組み合わせの偏り
        for fv in _counts(df_opp, 'DEF_FRONT').index:
            df_f_opp = df_opp[_s(df_opp['DEF_FRONT']) == str(fv)]
            for sv, cnt in _counts(df_f_opp, 'SIGN_D').items():
                pct_fs_opp = cnt / n_opp
                cnt_fs_all = ((_s(df['DEF_FRONT']) == fv) & (_s(df['SIGN_D']) == sv)).sum()
                pct_fs_all = cnt_fs_all / total
                if (pct_fs_opp >= THR_ABS and (pct_fs_opp - pct_fs_all) >= THR_DIFF
                        and cnt >= MIN_N):
                    bullets.append(
                        f'{opp}大では{fv}+{_display_val("SIGN_D", sv)}の組み合わせが多い'
                        f'（{pct_fs_opp:.0%}（{cnt}/{n_opp}プレー） vs 全体{pct_fs_all:.0%}（{cnt_fs_all}/{total}プレー））')

    # ③b 大学別最頻FRONT差
    bullets += _check_opp_modal_diff(df, 'DEF_FRONT')

    # ④ SIGN分散（5種以上かつ最多30%未満）
    for fv in _counts(df, 'DEF_FRONT').index:
        df_f = df[_s(df['DEF_FRONT']) == str(fv)]
        n_f = len(df_f)
        if n_f < MIN_N:
            continue
        sv_vc = _counts(df_f, 'SIGN_D')
        if len(sv_vc) == 0:
            continue
        top_pct = sv_vc.iloc[0] / n_f
        if top_pct < THR_PREDICT_SCATTER and len(sv_vc) >= 5:
            bullets.append(
                f'{fv}はSIGNが{len(sv_vc)}種に分散'
                f'（最多 {_display_val("SIGN_D", sv_vc.index[0])} {top_pct:.0%}、{n_f}プレー）')

    # ⑤ 3rdのみ: FRONTの構成変化
    if df_normal is not None and df_3rd is not None:
        for fv in _counts(df_normal, 'DEF_FRONT').index:
            for obs in _check_n3(fv, 'DEF_FRONT', df_normal, df_3rd):
                bullets.append(f'{fv}：{obs}')

    return bullets


def consider_sign_header(df, df_normal=None, df_3rd=None):
    """SIGN表の上の横断考察"""
    bullets = []
    total = len(df)
    if total == 0:
        return bullets

    # ① 全体ブリッツ率
    blitz_n = (_s(df['SIGN_D']) != 'N').sum()
    blitz_pct = blitz_n / total
    bullets.append(f'ブリッツ率：{blitz_pct:.0%}（{blitz_n}/{total}プレー）')

    # ② PERSONNELグループ別SIGN使用率変化
    for grp_label, grp_vals in PERSONNEL_GROUPS:
        df_p = df[_s(df['PERSONNEL_NORM']).isin(grp_vals)]
        n_p = len(df_p)
        if n_p < MIN_N:
            continue
        for sv in _counts(df, 'SIGN_D').index:
            cnt_p = (_s(df_p['SIGN_D']) == sv).sum()
            pct_p = cnt_p / n_p
            cnt_all = (_s(df['SIGN_D']) == sv).sum()
            pct_all = cnt_all / total
            if abs(pct_p - pct_all) >= THR_DIFF and cnt_p >= MIN_N:
                trend = '多い' if pct_p > pct_all else '少ない'
                bullets.append(
                    f'{grp_label}では{_display_val("SIGN_D", sv)}が{trend}'
                    f'（{pct_p:.0%}（{cnt_p}/{n_p}プレー） vs 全体{pct_all:.0%}（{cnt_all}/{total}プレー））')

    # ②b 大学別最頻SIGN差
    bullets += _check_opp_modal_diff(df, 'SIGN_D')

    # ③ 3rdのみ: SIGNの出現頻度変化
    if df_normal is not None and df_3rd is not None:
        for sv in _counts(df_normal, 'SIGN_D').index:
            for obs in _check_n3(sv, 'SIGN_D', df_normal, df_3rd):
                bullets.append(f'{_display_val("SIGN_D", sv)}：{obs}')

    # ④ 連続ブリッツ
    bullets += _consider_sign_consecutive(df, blitz_pct)

    # ⑤ SIGN連続パターン
    bullets += _consider_sign_sequence(df)

    # ⑥ 大学別ブリッツ率 + SIGNの種類変化
    for opp in _counts(df, 'OPPONENT').index:
        df_opp = df[_s(df['OPPONENT']) == str(opp)]
        n_opp = len(df_opp)
        if n_opp < MIN_N_OPP:
            continue
        b_opp = (_s(df_opp['SIGN_D']) != 'N').sum()
        b_total = (_s(df['SIGN_D']) != 'N').sum()
        pct_opp = b_opp / n_opp
        if abs(pct_opp - blitz_pct) >= THR_DIFF and b_opp >= MIN_N:
            trend = '高い' if pct_opp > blitz_pct else '低い'
            bullets.append(
                f'{opp}大戦はブリッツ率が{trend}'
                f'（{pct_opp:.0%}（{b_opp}/{n_opp}プレー） vs 全体{blitz_pct:.0%}（{b_total}/{total}プレー））')
        # SIGNの種類変化
        for sv in _counts(df, 'SIGN_D').index:
            if sv == 'N':
                continue
            cnt_sv_opp = (_s(df_opp['SIGN_D']) == sv).sum()
            cnt_sv_all = (_s(df['SIGN_D']) == sv).sum()
            pct_sv_opp = cnt_sv_opp / n_opp
            pct_sv_all = cnt_sv_all / total
            if (pct_sv_opp >= THR_ABS and (pct_sv_opp - pct_sv_all) >= THR_DIFF
                    and cnt_sv_opp >= MIN_N):
                bullets.append(
                    f'{opp}大では{_display_val("SIGN_D", sv)}が多い'
                    f'（{pct_sv_opp:.0%}（{cnt_sv_opp}/{n_opp}プレー） vs 全体{pct_sv_all:.0%}（{cnt_sv_all}/{total}プレー））')

    # ⑦ 複合ブリッツの出現条件
    bullets += _consider_complex_sign(df)

    # ⑧ ビッグゲイン後（1stDown更新後の1st&10）
    bullets += _consider_after_1st_down_gain(df)

    # ⑨ PERSONNEL・EMPTYでの5人以上ラッシュ
    bullets += _consider_pressure_by_personnel(df)

    # ⑩ ノーブリッツ条件（OFF FORM・PERSONNEL・Down）
    bullets += _consider_no_blitz(df)

    return bullets


def consider_coverage_header(df, df_normal=None, df_3rd=None):
    """COVERAGE表の上の横断考察"""
    bullets = []
    total = len(df)
    if total == 0:
        return bullets

    # ① 3rdのみ: COVERAGEの構成変化
    if df_normal is not None and df_3rd is not None:
        for cv in _counts(df_normal, 'COVERAGE_NORM').index:
            dv = _display_val('COVERAGE_NORM', cv)
            for obs in _check_n3(cv, 'COVERAGE_NORM', df_normal, df_3rd):
                bullets.append(f'{dv}：{obs}')

    # ② PERSONNELグループ別COVERAGE変化
    for grp_label, grp_vals in PERSONNEL_GROUPS:
        df_p = df[_s(df['PERSONNEL_NORM']).isin(grp_vals)]
        n_p = len(df_p)
        if n_p < MIN_N:
            continue
        for cv in _counts(df, 'COVERAGE_NORM').index:
            cnt_p = (_s(df_p['COVERAGE_NORM']) == cv).sum()
            pct_p = cnt_p / n_p
            cnt_all = (_s(df['COVERAGE_NORM']) == cv).sum()
            pct_all = cnt_all / total
            if abs(pct_p - pct_all) >= THR_DIFF and cnt_p >= MIN_N:
                dv = _display_val('COVERAGE_NORM', cv)
                trend = '多い' if pct_p > pct_all else '少ない'
                bullets.append(
                    f'{grp_label}では{dv}が{trend}'
                    f'（{pct_p:.0%}（{cnt_p}/{n_p}プレー） vs 全体{pct_all:.0%}（{cnt_all}/{total}プレー））')

    # ③ 大学別COVERAGE偏り + COVERAGE×SIGNの大学別偏り
    cov_all_vc = _counts(df, 'COVERAGE_NORM')
    for opp in _counts(df, 'OPPONENT').index:
        df_opp = df[_s(df['OPPONENT']) == str(opp)]
        n_opp = len(df_opp)
        if n_opp < MIN_N_OPP:
            continue
        for cv, cnt in _counts(df_opp, 'COVERAGE_NORM').items():
            cnt_cv_all = cov_all_vc.get(cv, 0)
            pct = cnt / n_opp
            pct_all = cnt_cv_all / total
            if pct >= THR_ABS and (pct - pct_all) >= THR_DIFF and cnt >= MIN_N:
                dv = _display_val('COVERAGE_NORM', cv)
                bullets.append(
                    f'{opp}大では{dv}が多い'
                    f'（{pct:.0%}（{cnt}/{n_opp}プレー） vs 全体{pct_all:.0%}（{cnt_cv_all}/{total}プレー））')
        # COVERAGE×SIGNの組み合わせ偏り
        for cv in _counts(df_opp, 'COVERAGE_NORM').index:
            df_c_opp = df_opp[_s(df_opp['COVERAGE_NORM']) == str(cv)]
            for sv, cnt in _counts(df_c_opp, 'SIGN_D').items():
                pct_cs_opp = cnt / n_opp
                cnt_cs_all = ((_s(df['COVERAGE_NORM']) == cv) & (_s(df['SIGN_D']) == sv)).sum()
                pct_cs_all = cnt_cs_all / total
                dv = _display_val('COVERAGE_NORM', cv)
                if (pct_cs_opp >= THR_ABS and (pct_cs_opp - pct_cs_all) >= THR_DIFF
                        and cnt >= MIN_N):
                    bullets.append(
                        f'{opp}大では{dv}+{_display_val("SIGN_D", sv)}の組み合わせが多い'
                        f'（{pct_cs_opp:.0%}（{cnt}/{n_opp}プレー） vs 全体{pct_cs_all:.0%}（{cnt_cs_all}/{total}プレー））')

    # ③b 大学別最頻COVERAGE差
    bullets += _check_opp_modal_diff(df, 'COVERAGE_NORM')

    # ④ SIGN×COVERAGE 頻出組み合わせ上位3（全体の5%以上）
    pairs = []
    for cv in _counts(df, 'COVERAGE_NORM').index:
        df_c = df[_s(df['COVERAGE_NORM']) == str(cv)]
        for sv, cnt in _counts(df_c, 'SIGN_D').items():
            pct = cnt / total
            if pct >= 0.05 and cnt >= MIN_N:
                dv = _display_val('COVERAGE_NORM', cv)
                pairs.append((pct, cnt, sv, dv))
    _RANK_LABEL = ['最多', '2番目に多い', '3番目に多い']
    pairs.sort(reverse=True)
    for rank, (pct, cnt, sv, dv) in enumerate(pairs[:3]):
        label = _RANK_LABEL[rank]
        bullets.append(
            f'{_display_val("SIGN_D", sv)}+{dv}の組み合わせが{label}（全体の{pct:.0%}、{cnt}/{total}プレー）')

    return bullets


# ══════════════════════════════════════════════════════════════
# PART 3 / PART 4：特殊分析（SIGNヘッダー用サブ関数）
# ══════════════════════════════════════════════════════════════

def _consider_sign_consecutive(df, blitz_pct):
    """ブリッツ直後の再ブリッツ率（全体ブリッツ率と比較）"""
    bullets = []
    if len(df) < MIN_N * 2:
        return bullets

    df_s = df.sort_values(['OPPONENT', 'PLAY #']).reset_index(drop=True)
    signs = _s(df_s['SIGN_D'])
    opps = _s(df_s['OPPONENT'])
    total = len(df_s)

    after_blitz = []
    for i in range(total - 1):
        if opps.iloc[i] != opps.iloc[i + 1]:
            continue
        if signs.iloc[i] != 'N':
            after_blitz.append(signs.iloc[i + 1] != 'N')

    if len(after_blitz) >= MIN_N:
        n_consec = sum(after_blitz)
        n_after = len(after_blitz)
        after_rate = n_consec / n_after
        diff = after_rate - blitz_pct
        if abs(diff) >= THR_DIFF:
            b_total = (_s(df_s['SIGN_D']) != 'N').sum()
            trend = '連続しやすい' if diff > 0 else '単発が多い'
            bullets.append(
                f'ブリッツ直後の再ブリッツ率：{after_rate:.0%}（{n_consec}/{n_after}プレー）'
                f' vs 全体{blitz_pct:.0%}（{b_total}/{total}プレー） → {trend}')

    return bullets


def _consider_sign_sequence(df):
    """特定SIGNの直後に続きやすいSIGNのパターン（75%以上）"""
    bullets = []
    if len(df) < MIN_N * 2:
        return bullets

    df_s = df.sort_values(['OPPONENT', 'PLAY #']).reset_index(drop=True)
    signs = _s(df_s['SIGN_D'])
    opps = _s(df_s['OPPONENT'])
    total = len(df_s)

    sign_after = defaultdict(list)
    for i in range(total - 1):
        if opps.iloc[i] != opps.iloc[i + 1]:
            continue
        sv = signs.iloc[i]
        if sv != 'N':
            sign_after[sv].append(signs.iloc[i + 1])

    for sv, nexts in sign_after.items():
        n = len(nexts)
        if n < MIN_N:
            continue
        nvc = pd.Series(nexts).value_counts()
        top_next = nvc.index[0]
        top_cnt = nvc.iloc[0]
        top_pct = top_cnt / n
        if top_pct >= THR_HIGH:
            sv_dv = _display_val('SIGN_D', sv)
            next_dv = _display_val('SIGN_D', top_next)
            bullets.append(
                f'SIGN {sv_dv}直後は高確率でSIGN {next_dv}（{_fmt(top_pct, top_cnt, n)}）')

    return bullets


def _consider_complex_sign(df):
    """複合SIGN（+）の出現条件（単体との比較）"""
    bullets = []
    signs = _s(df['SIGN_D'])
    complex_vc = signs[signs.str.contains(r'\+', na=False)].value_counts()

    for comp_sign, comp_cnt in complex_vc.items():
        if comp_cnt < MIN_N:
            continue
        base = comp_sign.split('+')[0].strip()
        df_comp = df[_s(df['SIGN_D']) == comp_sign]
        df_base = df[_s(df['SIGN_D']) == base]
        base_n = len(df_base)
        if base_n < MIN_N:
            continue

        for col, label in [('PERSONNEL_NORM', 'Personnel'),
                            ('OFF_FORM_NORM', 'フォーム')]:
            if col not in df_comp.columns:
                continue
            comp_vc = _counts(df_comp, col)
            base_vc = _counts(df_base, col)
            if len(comp_vc) == 0:
                continue
            top_val = comp_vc.index[0]
            top_cnt = comp_vc.iloc[0]
            top_pct = top_cnt / comp_cnt
            base_pct = base_vc.get(top_val, 0) / base_n
            dv = _display_val(col, top_val)
            if top_pct >= THR_HIGH and (top_pct - base_pct) >= THR_DIFF and top_cnt >= MIN_N:
                bullets.append(
                    f'{comp_sign}は{base}と比べて{label}:{dv}で出やすい'
                    f'（{_fmt(top_pct, top_cnt, comp_cnt)} vs {base}:{base_pct:.0%}）')

    return bullets


def _consider_after_1st_down_gain(df):
    """大きなゲインで1stDown更新後の1st&10プレーでのブリッツ/ラッシュ変化"""
    for col in ['GN/LS', 'DN', 'DIST']:
        if col not in df.columns:
            return []

    df_s = df.sort_values(['OPPONENT', 'PLAY #']).reset_index(drop=True)
    opps = _s(df_s['OPPONENT'])
    total = len(df_s)
    gn = pd.to_numeric(df_s['GN/LS'], errors='coerce')
    dn = pd.to_numeric(df_s['DN'], errors='coerce')
    dist = pd.to_numeric(df_s['DIST'], errors='coerce')

    if total < MIN_N * 2:
        return []

    b_all_total = (_s(df_s['SIGN_D']) != 'N').sum()
    overall_blitz = b_all_total / total
    overall_press = pd.to_numeric(df_s['PRESSURE_NUM'], errors='coerce').mean()
    best_blitz = None
    best_press = None

    for threshold in [7, 8, 9, 10, 11]:
        next_idx = []
        for i in range(total - 1):
            if opps.iloc[i] != opps.iloc[i + 1]:
                continue
            if pd.isna(gn.iloc[i]) or gn.iloc[i] < threshold:
                continue
            if dn.iloc[i + 1] == 1 and 9 <= dist.iloc[i + 1] <= 11:
                next_idx.append(i + 1)

        if len(next_idx) < MIN_N:
            continue
        df_after = df_s.iloc[next_idx]
        n_after = len(next_idx)

        after_blitz = (_s(df_after['SIGN_D']) != 'N').mean()
        bdiff = abs(after_blitz - overall_blitz)
        if bdiff >= THR_DIFF and (best_blitz is None or bdiff > best_blitz['diff']):
            b_cnt = (_s(df_after['SIGN_D']) != 'N').sum()
            best_blitz = {
                'threshold': threshold, 'n': n_after,
                'after': after_blitz, 'overall': overall_blitz,
                'diff': bdiff, 'b_cnt': b_cnt,
                'b_all': b_all_total, 'n_all': total,
                'trend': '増加' if after_blitz > overall_blitz else '減少',
            }

        after_press = pd.to_numeric(df_after['PRESSURE_NUM'], errors='coerce').mean()
        if pd.notna(after_press) and pd.notna(overall_press):
            pdiff = abs(after_press - overall_press)
            if pdiff >= 0.4 and (best_press is None or pdiff > best_press['diff']):
                best_press = {
                    'threshold': threshold, 'n': n_after,
                    'after': after_press, 'overall': overall_press, 'diff': pdiff,
                    'trend': '増加' if after_press > overall_press else '減少',
                }

    bullets = []
    if best_blitz:
        t = best_blitz
        bullets.append(
            f'{t["threshold"]}yds以上ゲインで1stDown更新後はブリッツ率が{t["trend"]}'
            f'（{t["after"]:.0%}（{t["b_cnt"]}/{t["n"]}プレー） vs 全体{t["overall"]:.0%}（{t["b_all"]}/{t["n_all"]}プレー））')
    if best_press:
        t = best_press
        bullets.append(
            f'{t["threshold"]}yds以上ゲインで1stDown更新後はラッシュ人数が{t["trend"]}'
            f'（直後平均{t["after"]:.1f}人 vs 全体{t["overall"]:.1f}人、{t["n"]}プレー）')
    return bullets


def _consider_pressure_by_personnel(df):
    """PERSONNELグループ別・EMPTYでの5人以上ラッシュ傾向"""
    bullets = []
    pressure = pd.to_numeric(df['PRESSURE_NUM'], errors='coerce').dropna()
    if len(pressure) < MIN_N:
        return bullets

    high_all = int((pressure >= 5).sum())
    n_all_press = len(pressure)
    overall_high_pct = high_all / n_all_press

    for grp_label, grp_vals in PERSONNEL_GROUPS_PRESSURE:
        df_p = df[_s(df['PERSONNEL_NORM']).isin(grp_vals)]
        n_p = len(df_p)
        if n_p < MIN_N:
            continue
        pres_p = pd.to_numeric(df_p['PRESSURE_NUM'], errors='coerce').dropna()
        if len(pres_p) < MIN_N:
            continue
        high_p = (pres_p >= 5).sum()
        high_pct = high_p / len(pres_p)
        if abs(high_pct - overall_high_pct) >= THR_DIFF and high_p >= MIN_N:
            trend = '高い' if high_pct > overall_high_pct else '低い'
            bullets.append(
                f'{grp_label}は5人以上ラッシュ率が{trend}'
                f'（{high_pct:.0%}（{high_p}/{len(pres_p)}プレー） vs 全体{overall_high_pct:.0%}（{high_all}/{n_all_press}プレー））')

    # EMPTYフォーメーション
    if 'OFF_FORM_NORM' in df.columns:
        df_empty = df[_s(df['OFF_FORM_NORM']).str.contains('EMPTY', na=False)]
        n_e = len(df_empty)
        if n_e >= MIN_N:
            pres_e = pd.to_numeric(df_empty['PRESSURE_NUM'], errors='coerce').dropna()
            high_e = (pres_e >= 5).sum()
            pct_e = high_e / n_e if n_e > 0 else 0
            if pct_e >= THR_ABS and (pct_e - overall_high_pct) >= THR_DIFF and high_e >= MIN_N:
                bullets.append(
                    f'EMPTYフォームでは5人以上ラッシュの傾向'
                    f'（{pct_e:.0%}（{high_e}/{n_e}プレー） vs 全体{overall_high_pct:.0%}（{high_all}/{n_all_press}プレー））')

    return bullets


def _consider_no_blitz(df):
    """ノーブリッツ（SIGN=N）の出現条件 — OFF FORM・PERSONNEL・Down"""
    df_n = df[_s(df['SIGN_D']) == 'N']
    n_n = len(df_n)
    N = len(df)
    if n_n < MIN_N or N == 0:
        return []

    bullets = []
    for col in ['OFF_FORM_NORM', 'PERSONNEL_NORM', 'DN']:
        if col not in df.columns:
            continue
        for val, cnt in _counts(df_n, col).items():
            df_val = df[_s(df[col]) == str(val)]
            n_val = len(df_val)
            if n_val < MIN_N:
                continue
            nb_val = (_s(df_val['SIGN_D']) == 'N').sum()
            pct_val = nb_val / n_val
            dv = _display_val(col, val)
            if pct_val >= THR_HIGH and cnt >= MIN_N:
                bullets.append(
                    f'{dv}において高確率でノーブリッツ（{_fmt(pct_val, nb_val, n_val)}）')

    return bullets


# ══════════════════════════════════════════════════════════════
# PART 5：2MIN 大学差分析
# ══════════════════════════════════════════════════════════════

def consider_2min_sign_header(df):
    """2MIN SIGN(D) の大学差分析（集中度 + 使用率比較のみ）"""
    bullets = []
    for val in _counts(df, 'SIGN_D').index:
        if not val or str(val) == 'nan':
            continue
        df_sub = df[_s(df['SIGN_D']) == str(val)]
        for b in _check_opp(df_sub):
            bullets.append(f'{val}：{b}')
        for b in _check_opp_usage_rate(df, 'SIGN_D', val):
            bullets.append(f'{val}：{b}')
    return bullets


def consider_2min_cov_header(df):
    """2MIN COVERAGE の大学差分析（集中度 + 使用率比較のみ）"""
    bullets = []
    for val in _counts(df, 'COVERAGE_NORM').index:
        if not val or str(val) == 'nan':
            continue
        df_sub = df[_s(df['COVERAGE_NORM']) == str(val)]
        for b in _check_opp(df_sub):
            bullets.append(f'{val}：{b}')
        for b in _check_opp_usage_rate(df, 'COVERAGE_NORM', val):
            bullets.append(f'{val}：{b}')
    return bullets


# ══════════════════════════════════════════════════════════════
# PART 5b：Red Zone 専用ヘッダー考察
# ══════════════════════════════════════════════════════════════

def consider_redzone_front_header(df, df_normal=None, df_3rd=None):
    """
    Red Zone FRONT表の上の考察。
    ・Normal / 3rd と比較して±20pt以上の FRONT → 記載
    ・特定大学でFRONT使用率が全体比+20pt以上 → 記載
    ・特定大学でFRONT×SIGNの組み合わせが全体比+20pt以上 → 記載
    """
    bullets = []
    total = len(df)
    if total == 0:
        return bullets

    # RedZone vs Normal / 3rd 比較
    for comp_df, comp_label in [(df_normal, 'Normal'), (df_3rd, '3rd')]:
        if comp_df is None or len(comp_df) == 0:
            continue
        n_c = len(comp_df)
        for fv in _counts(df, 'DEF_FRONT').index:
            cnt_rz = (_s(df['DEF_FRONT']) == str(fv)).sum()
            pct_rz = cnt_rz / total
            cnt_c  = (_s(comp_df['DEF_FRONT']) == str(fv)).sum()
            pct_c  = cnt_c / n_c
            diff   = pct_rz - pct_c
            if abs(diff) >= THR_DIFF:
                trend = '多い' if diff > 0 else '少ない'
                bullets.append(
                    f'RedZone vs {comp_label}：{fv}が{trend}'
                    f'（RZ {pct_rz:.0%}（{cnt_rz}/{total}） vs {comp_label} {pct_c:.0%}（{cnt_c}/{n_c}プレー））')

    front_all_vc = _counts(df, 'DEF_FRONT')
    for opp in _counts(df, 'OPPONENT').index:
        df_opp = df[_s(df['OPPONENT']) == str(opp)]
        n_opp = len(df_opp)
        if n_opp < MIN_N_OPP:
            continue
        # FRONT単体の偏り
        for fv, cnt in _counts(df_opp, 'DEF_FRONT').items():
            cnt_fv_all = front_all_vc.get(fv, 0)
            pct = cnt / n_opp
            pct_all = cnt_fv_all / total
            if pct >= THR_ABS and (pct - pct_all) >= THR_DIFF and cnt >= MIN_N:
                bullets.append(
                    f'{opp}大では{fv}が多い'
                    f'（{pct:.0%}（{cnt}/{n_opp}プレー） vs 全体{pct_all:.0%}（{cnt_fv_all}/{total}プレー））')
        # FRONT×SIGNの組み合わせ偏り
        for fv in _counts(df_opp, 'DEF_FRONT').index:
            df_f_opp = df_opp[_s(df_opp['DEF_FRONT']) == str(fv)]
            for sv, cnt in _counts(df_f_opp, 'SIGN_D').items():
                pct_fs_opp = cnt / n_opp
                cnt_fs_all = ((_s(df['DEF_FRONT']) == fv) & (_s(df['SIGN_D']) == sv)).sum()
                pct_fs_all = cnt_fs_all / total
                if (pct_fs_opp >= THR_ABS and (pct_fs_opp - pct_fs_all) >= THR_DIFF
                        and cnt >= MIN_N):
                    bullets.append(
                        f'{opp}大では{fv}+{_display_val("SIGN_D", sv)}の組み合わせが多い'
                        f'（{pct_fs_opp:.0%}（{cnt}/{n_opp}プレー）'
                        f' vs 全体{pct_fs_all:.0%}（{cnt_fs_all}/{total}プレー））')
    return bullets


def consider_redzone_cov_header(df, df_normal=None, df_3rd=None):
    """
    Red Zone COVERAGE表の上の考察。
    ・Normal / 3rd と比較して±20pt以上の COVERAGE → 記載
    ① PERSONNELグループ（10per / 11per / 12・21・22per）別COVERAGE使用率が全体比±20pt以上 → 記載
    ② 特定大学でCOVERAGE使用率が全体比+20pt以上 → 記載
    ③ 特定大学でCOVERAGE×SIGN組み合わせが全体比+20pt以上 → 記載
    """
    bullets = []
    total = len(df)
    if total == 0:
        return bullets

    # RedZone vs Normal / 3rd 比較
    for comp_df, comp_label in [(df_normal, 'Normal'), (df_3rd, '3rd')]:
        if comp_df is None or len(comp_df) == 0:
            continue
        n_c = len(comp_df)
        for cv in _counts(df, 'COVERAGE_NORM').index:
            cnt_rz = (_s(df['COVERAGE_NORM']) == str(cv)).sum()
            pct_rz = cnt_rz / total
            cnt_c  = (_s(comp_df['COVERAGE_NORM']) == str(cv)).sum()
            pct_c  = cnt_c / n_c
            diff   = pct_rz - pct_c
            if abs(diff) >= THR_DIFF:
                dv    = _display_val('COVERAGE_NORM', cv)
                trend = '多い' if diff > 0 else '少ない'
                bullets.append(
                    f'RedZone vs {comp_label}：{dv}が{trend}'
                    f'（RZ {pct_rz:.0%}（{cnt_rz}/{total}） vs {comp_label} {pct_c:.0%}（{cnt_c}/{n_c}プレー））')

    # ① PERSONNELグループ別COVERAGE使用率変化（±20pt）
    for grp_label, grp_vals in PERSONNEL_GROUPS:
        df_p = df[_s(df['PERSONNEL_NORM']).isin(grp_vals)]
        n_p = len(df_p)
        if n_p < MIN_N:
            continue
        for cv in _counts(df, 'COVERAGE_NORM').index:
            cnt_p = (_s(df_p['COVERAGE_NORM']) == cv).sum()
            pct_p = cnt_p / n_p
            cnt_all = (_s(df['COVERAGE_NORM']) == cv).sum()
            pct_all = cnt_all / total
            if abs(pct_p - pct_all) >= THR_DIFF and cnt_p >= MIN_N:
                dv = _display_val('COVERAGE_NORM', cv)
                trend = '多い' if pct_p > pct_all else '少ない'
                bullets.append(
                    f'{grp_label}では{dv}が{trend}'
                    f'（{pct_p:.0%}（{cnt_p}/{n_p}プレー） vs 全体{pct_all:.0%}（{cnt_all}/{total}プレー））')

    # ② 大学別COVERAGE偏り
    cov_all_vc = _counts(df, 'COVERAGE_NORM')
    for opp in _counts(df, 'OPPONENT').index:
        df_opp = df[_s(df['OPPONENT']) == str(opp)]
        n_opp = len(df_opp)
        if n_opp < MIN_N_OPP:
            continue
        for cv, cnt in _counts(df_opp, 'COVERAGE_NORM').items():
            cnt_cv_all = cov_all_vc.get(cv, 0)
            pct = cnt / n_opp
            pct_all = cnt_cv_all / total
            if pct >= THR_ABS and (pct - pct_all) >= THR_DIFF and cnt >= MIN_N:
                dv = _display_val('COVERAGE_NORM', cv)
                bullets.append(
                    f'{opp}大では{dv}が多い'
                    f'（{pct:.0%}（{cnt}/{n_opp}プレー） vs 全体{pct_all:.0%}（{cnt_cv_all}/{total}プレー））')
        # ③ COVERAGE×SIGNの組み合わせ偏り
        for cv in _counts(df_opp, 'COVERAGE_NORM').index:
            df_c_opp = df_opp[_s(df_opp['COVERAGE_NORM']) == str(cv)]
            for sv, cnt in _counts(df_c_opp, 'SIGN_D').items():
                pct_cs_opp = cnt / n_opp
                cnt_cs_all = ((_s(df['COVERAGE_NORM']) == cv) & (_s(df['SIGN_D']) == sv)).sum()
                pct_cs_all = cnt_cs_all / total
                dv = _display_val('COVERAGE_NORM', cv)
                if (pct_cs_opp >= THR_ABS and (pct_cs_opp - pct_cs_all) >= THR_DIFF
                        and cnt >= MIN_N):
                    bullets.append(
                        f'{opp}大では{dv}+{_display_val("SIGN_D", sv)}の組み合わせが多い'
                        f'（{pct_cs_opp:.0%}（{cnt}/{n_opp}プレー）'
                        f' vs 全体{pct_cs_all:.0%}（{cnt_cs_all}/{total}プレー））')
    return bullets


def consider_redzone_packages_header(df, df_overall=None):
    """
    Red Zone パッケージ割合表の上の考察。
    ① 各パッケージの使用率が全体比（df_overall）±20pt以上なら記載
       （df_overall=None のとき、①はスキップ）
    ② 特定大学でパッケージ使用率が偏っていれば記載
    """
    bullets = []
    total = len(df)
    if total == 0:
        return bullets

    # パッケージ列を構築（DEF_FRONT + SIGN_D + COVERAGE_NORM）
    tmp = df.copy()
    tmp['_pkg'] = (tmp['DEF_FRONT'].fillna('').astype(str).str.strip()
                   + ' + ' + tmp['SIGN_D'].fillna('').astype(str).str.strip()
                   + ' + ' + tmp['COVERAGE_NORM'].fillna('').astype(str).str.strip())

    pkg_vc = tmp['_pkg'].value_counts()

    # ① ヤード帯ごとの全体比較（df_overall が渡されたとき）
    if df_overall is not None and len(df_overall) > 0:
        total_all = len(df_overall)
        tmp_all = df_overall.copy()
        tmp_all['_pkg'] = (tmp_all['DEF_FRONT'].fillna('').astype(str).str.strip()
                           + ' + ' + tmp_all['SIGN_D'].fillna('').astype(str).str.strip()
                           + ' + ' + tmp_all['COVERAGE_NORM'].fillna('').astype(str).str.strip())
        pkg_vc_all = tmp_all['_pkg'].value_counts()

        for pkg, cnt in pkg_vc.items():
            if cnt < MIN_N:
                continue
            pct = cnt / total
            cnt_all = pkg_vc_all.get(pkg, 0)
            pct_all = cnt_all / total_all
            diff = pct - pct_all
            if abs(diff) >= THR_DIFF:
                trend = '多い' if diff > 0 else '少ない'
                bullets.append(
                    f'{pkg} が全体比より{trend}'
                    f'（{pct:.0%}（{cnt}/{total}プレー） vs 全体{pct_all:.0%}（{cnt_all}/{total_all}プレー））')

    # ② 大学別パッケージ偏り
    for opp in _counts(df, 'OPPONENT').index:
        df_opp = df[_s(df['OPPONENT']) == str(opp)]
        n_opp = len(df_opp)
        if n_opp < MIN_N_OPP:
            continue
        tmp_opp = df_opp.copy()
        tmp_opp['_pkg'] = (tmp_opp['DEF_FRONT'].fillna('').astype(str).str.strip()
                           + ' + ' + tmp_opp['SIGN_D'].fillna('').astype(str).str.strip()
                           + ' + ' + tmp_opp['COVERAGE_NORM'].fillna('').astype(str).str.strip())
        for pkg, cnt in tmp_opp['_pkg'].value_counts().items():
            if cnt < MIN_N:
                continue
            pct = cnt / n_opp
            cnt_all = pkg_vc.get(pkg, 0)
            pct_all = cnt_all / total
            if pct >= THR_ABS and (pct - pct_all) >= THR_DIFF:
                bullets.append(
                    f'{opp}大では{pkg} が多い'
                    f'（{pct:.0%}（{cnt}/{n_opp}プレー） vs 全体{pct_all:.0%}（{cnt_all}/{total}プレー））')

    return bullets


def _pkg_dn_dist_label(row):
    """パッケージ分析用 DN×DIST ラベル生成"""
    try:
        dn = int(float(row['DN']))
    except (ValueError, TypeError):
        return ''
    if dn == 1:
        return '1st down'
    suffix = {2: 'nd', 3: 'rd'}.get(dn, 'th')
    try:
        dv = int(float(row['DIST']))
    except (ValueError, TypeError):
        return f'{dn}{suffix}'
    if dv <= 3:    ds = '1~3'
    elif dv <= 6:  ds = '4~6'
    elif dv <= 10: ds = '7~10'
    else:          ds = '11+'
    return f'{dn}{suffix} {ds}'


def consider_redzone_package_item(pkg_label, df_pkg, df_rz):
    """
    RedZone パッケージ個別特徴分析。
    ・大学分析：1大学のみ出現 / 主要大学 / 複数大学内訳
    ・Down&Dist：70%以上で特定帯に集中していれば記載
    """
    if df_pkg is None or len(df_pkg) == 0:
        return []
    n_pkg = len(df_pkg)
    bullets = []

    # 大学分析
    opp_vc = _counts(df_pkg, 'OPPONENT')
    if len(opp_vc) == 1:
        opp = opp_vc.index[0]
        bullets.append(f'{opp}大のみで出現（{n_pkg}プレー）')
    elif len(opp_vc) > 0:
        top_opp = opp_vc.index[0]
        top_cnt = int(opp_vc.iloc[0])
        top_pct = top_cnt / n_pkg
        if top_pct >= THR_HIGH:
            bullets.append(f'主に{top_opp}大で出現（{top_pct:.0%}、{top_cnt}/{n_pkg}プレー）')
        else:
            parts = [f'{o}大 {c}プレー' for o, c in opp_vc.items()]
            bullets.append('内訳：' + '、'.join(parts))

    # Down & Dist 集中（70%以上）
    tmp = df_pkg.copy()
    tmp['_dd'] = tmp.apply(_pkg_dn_dist_label, axis=1)
    dd_vc = _counts(tmp, '_dd')
    for dd_val, cnt in dd_vc.items():
        if cnt < MIN_N:
            continue
        pct = cnt / n_pkg
        if pct >= 0.70:
            bullets.append(f'{dd_val}に集中（{pct:.0%}、{cnt}/{n_pkg}プレー）')

    return bullets


def consider_redzone_packages_downdist(df_zone):
    """
    ゾーン別パッケージ表の上に表示。
    各パッケージについて特定 Down&Dist 帯に70%以上集中していれば記載。
    """
    if len(df_zone) < MIN_N:
        return []
    pkg_map = _build_redzone_pkg_map(df_zone)
    bullets = []
    for pkg_label, df_pkg in pkg_map.items():
        n_pkg = len(df_pkg)
        if n_pkg < MIN_N:
            continue
        tmp = df_pkg.copy()
        tmp['_dd'] = tmp.apply(_pkg_dn_dist_label, axis=1)
        dd_vc = _counts(tmp, '_dd')
        for dd_val, cnt in dd_vc.items():
            if cnt < MIN_N:
                continue
            pct = cnt / n_pkg
            if pct >= 0.70:
                bullets.append(
                    f'{pkg_label}：{dd_val}に集中（{pct:.0%}、{cnt}/{n_pkg}プレー）')
    return bullets


# ══════════════════════════════════════════════════════════════
# PART 6：3変数複合パターン（OFF FORM × COVERAGE 備考列）
# ══════════════════════════════════════════════════════════════

def consider_3var_offform_cov(df):
    """
    OFF FORM × COVERAGE の組み合わせに対する備考を生成。
    戻り値: { (form_str, cov_str): [observation_string, ...] }
    """
    result = {}
    if len(df) == 0:
        return result

    for form in _counts(df, 'OFF_FORM_NORM').index:
        df_f = df[_s(df['OFF_FORM_NORM']) == str(form)]
        if len(df_f) < MIN_N:
            continue
        for cov in _counts(df_f, 'COVERAGE_NORM').index:
            df_fc = df_f[_s(df_f['COVERAGE_NORM']) == str(cov)]
            n_fc = len(df_fc)
            if n_fc < MIN_N:
                continue

            obs = []

            # ① ラッシュ人数（最頻値を常に記載）
            press = pd.to_numeric(df_fc['PRESSURE_NUM'], errors='coerce').dropna().astype(int)
            if len(press) >= MIN_N:
                pvc = press.value_counts()
                top_p, top_cnt = pvc.index[0], pvc.iloc[0]
                n_press = len(press)
                top_pct = top_cnt / n_press
                _label = f'{top_p}人ラッシュ'
                if top_pct == 1.0:
                    obs.append(f'必ず{_label}（100%、{top_cnt}/{n_press}プレー）')
                elif top_pct >= THR_HIGH:
                    obs.append(f'高確率で{_label}（{_fmt(top_pct, top_cnt, n_press)}）')
                else:
                    obs.append(f'最多は{_label}（{_fmt(top_pct, top_cnt, n_press)}）')

                # 大学別差異
                for opp in _counts(df_fc, 'OPPONENT').index:
                    df_opp_fc = df_fc[_s(df_fc['OPPONENT']) == str(opp)]
                    n_o = len(df_opp_fc)
                    if n_o < MIN_N:
                        continue
                    press_o = pd.to_numeric(df_opp_fc['PRESSURE_NUM'], errors='coerce').dropna().astype(int)
                    if len(press_o) < MIN_N:
                        continue
                    top_o = press_o.value_counts().index[0]
                    if top_o != top_p:
                        obs.append(f'※{opp}でのみ{top_o}人ラッシュが最も多い（{n_o}プレー）')

            # ② SIGN（最頻値を常に記載）
            sign_vc = _counts(df_fc, 'SIGN_D')
            if len(sign_vc) >= 1:
                top_s, top_s_cnt = sign_vc.index[0], sign_vc.iloc[0]
                top_s_pct = top_s_cnt / n_fc
                sign_obs = _fmt_top('SIGN_D', top_s, top_s_cnt, n_fc, top_s_pct)
                if top_s == 'N' and len(sign_vc) >= 2:
                    sec_s = sign_vc.index[1]
                    sec_cnt = sign_vc.iloc[1]
                    sec_pct = sec_cnt / n_fc
                    sec_dv = _display_val('SIGN_D', sec_s)
                    sign_obs += f'、2番目は{sec_dv}（{_fmt(sec_pct, sec_cnt, n_fc)}）'
                obs.append(sign_obs)

                # 大学別差異
                for opp in _counts(df_fc, 'OPPONENT').index:
                    df_opp_fc = df_fc[_s(df_fc['OPPONENT']) == str(opp)]
                    n_o = len(df_opp_fc)
                    if n_o < MIN_N:
                        continue
                    opp_sign_vc = _counts(df_opp_fc, 'SIGN_D')
                    if len(opp_sign_vc) == 0:
                        continue
                    top_so = opp_sign_vc.index[0]
                    if top_so != top_s:
                        obs.append(f'※{opp}でのみSIGN:{_display_val("SIGN_D", top_so)}が最も多い（{n_o}プレー）')

            if obs:
                result[(str(form), str(cov))] = obs

    return result


# ══════════════════════════════════════════════════════════════
# PART 7：SIGNシチュエーション分析（Req C）
# ══════════════════════════════════════════════════════════════

# 8要素の定義（列名, 表示ラベル）
_SIT_ELEMENTS = [
    ('_sit_yd',      'YD-LN'),
    ('_sit_dn_dist', 'DN×DIST'),
    ('_sit_pers',    'パーソネル'),
    ('_sit_form',    'フォーメーション'),
    ('_sit_prev',    '前プレー'),
    ('_sit_rb',      'RBサイド'),
]


def _build_situation_df(df):
    """SIGNシチュエーション分析用の派生列を追加したDataFrameを返す"""
    d = df.copy()

    # 1. YARD LN zone
    def _yd_zone(v):
        try:
            y = int(float(v))
        except (ValueError, TypeError):
            return ''
        if -9 <= y <= -1:   return '自陣深く（-1~-9yds）'
        if -49 <= y <= -10: return '自陣（-10~-49yds）'
        if 26 <= y <= 50:   return '敵陣（50~26yds）'
        if 1  <= y <= 25:   return 'Redzone'
        return ''
    d['_sit_yd'] = d['YARD LN'].apply(_yd_zone)

    # 自陣除外：敵陣プレー数が3以下の場合、自陣を空文字に変換
    _tekki_n = (d['_sit_yd'] == '敵陣').sum()
    if _tekki_n <= 3:
        d['_sit_yd'] = d['_sit_yd'].replace({'自陣深く（-1~-9yds）': '', '自陣（-10~-49yds）': ''})

    # 2. DN × DIST 統合（1st のみ DIST なし）
    def _dn_dist(row):
        try:
            dn = int(float(row['DN']))
        except (ValueError, TypeError):
            return ''
        if dn == 1:
            return '1st down'
        suffix = {1: 'nd', 2: 'nd', 3: 'rd'}.get(dn, 'th')
        try:
            dv = int(float(row['DIST']))
        except (ValueError, TypeError):
            return f'{dn}{suffix}'
        if dv <= 3:   dist_str = '1~3'
        elif dv <= 6:  dist_str = '4~6'
        elif dv <= 10: dist_str = '7~10'
        else:          dist_str = '11+'
        return f'{dn}{suffix} {dist_str}'
    d['_sit_dn_dist'] = d.apply(_dn_dist, axis=1)

    # 3. Personnel group
    def _pers_g(v):
        v = str(v).strip()
        if v == '0':  return '00per'
        if v == '10': return '10per'
        if v == '11': return '11per'
        if v in ('12', '21', '22'): return '12/21/22per'
        return ''
    d['_sit_pers'] = d['PERSONNEL_NORM'].apply(_pers_g)

    # 4. Formation group（analyzer._get_form_group を使用）
    d['_sit_form'] = d['OFF_FORM_NORM'].apply(
        lambda v: _get_form_group_az(v) or '')

    # 5. Previous GN/LS（同一試合の前プレーを参照）
    ds = d.sort_values(['OPPONENT', 'PLAY #'])
    gnls_raw  = pd.to_numeric(ds['GN/LS'], errors='coerce')
    opp_s     = _s(ds['OPPONENT'])
    prev_gnls = gnls_raw.shift(1)
    # 対戦校が変わったら前プレーを無効化
    prev_opps = opp_s.shift(1).fillna('')
    prev_gnls[prev_opps != opp_s] = float('nan')

    def _gnls_z(v):
        if pd.isna(v): return ''
        fv = float(v)
        if fv <= 3:  return '前プレーゲイン3yds以下'
        if fv <= 6:  return '前プレーゲイン4~6yds'
        if fv <= 9:  return '前プレーゲイン7~9yds'
        return '前プレーゲイン10+yds'
    ds = ds.copy()
    ds['_sit_prev'] = prev_gnls.apply(_gnls_z)
    d['_sit_prev'] = ds['_sit_prev']

    # 6. RB Wide Side（なし・PISTOL も有効値として含める）
    _rb_map = {'広い': 'RB広', '狭い': 'RB狭', 'なし': 'RBなし'}
    if 'RB_WIDE_SIDE' in d.columns:
        d['_sit_rb'] = d['RB_WIDE_SIDE'].apply(
            lambda v: _rb_map.get(str(v).strip(), str(v).strip())
            if str(v).strip() not in ('', 'nan') else '')
    else:
        d['_sit_rb'] = ''

    return d


def consider_sign_situation(df_all, sign_val, min_n=MIN_N, threshold=0.60, max_combo=3,
                            min_coverage=0.20, _df_sit=None):
    """
    特定SIGN値がどんな状況（YD-LN/DN/DIST/パーソネル/フォーメーション/前プレー/QTR/RBサイド）
    で多く出るかを多変量探索で分析し、箇条書きリストを返す。

    ・各要素の単独〜max_combo要素組み合わせを試し、以下を同時に満たすものを収集：
      1. 集中度 (sign_in / combo_total) >= threshold
      2. そのグループ内の sign_in が n_sign * min_coverage 以上（外れ値的小集団を除外）
    ・冗長除去：より具体的な条件（要素が多い）が同等以上の集中度を持つなら汎用側を除去。
    ・条件なし → 空リストを返す（呼び出し元で「傾向なし」と表示）。
    """
    sign_str = str(sign_val)
    n_total  = len(df_all)
    if n_total == 0:
        return []
    n_sign = int((_s(df_all['SIGN_D']) == sign_str).sum())
    if n_sign < min_n:
        return []

    # min_sign_in: そのグループがSIGN全体の何割以上をカバーするか
    min_sign_in = max(min_n, int(n_sign * min_coverage))

    d = _df_sit if _df_sit is not None else _build_situation_df(df_all)

    # 使える要素（MIN_N 以上の非空値がある列のみ）
    valid_cols = [
        (col, label) for col, label in _SIT_ELEMENTS
        if (_s(d[col]) != '').sum() >= min_n * 2
    ]
    if not valid_cols:
        return []

    hits = []  # (frozenset{(col,val),...}, concentration, n_sign_in, n_combo)

    for r in range(1, min(max_combo, len(valid_cols)) + 1):
        for cols_combo in _combinations(valid_cols, r):
            cols = [c for c, _ in cols_combo]

            # 全列が非空の行だけ使う
            mask = pd.Series(True, index=d.index)
            for c in cols:
                mask &= (_s(d[c]) != '')
            d_valid = d[mask]
            if len(d_valid) < min_n:
                continue

            for gv, grp in d_valid.groupby(cols, observed=True):
                n_combo = len(grp)
                if n_combo < min_n:
                    continue
                if isinstance(gv, str):
                    gv = (gv,)
                n_sign_in = int((_s(grp['SIGN_D']) == sign_str).sum())
                # 集中度条件 + カバレッジ条件
                if n_sign_in < min_sign_in:
                    continue
                conc = n_sign_in / n_combo
                if conc >= threshold:
                    key = frozenset(zip(cols, [str(v) for v in gv]))
                    hits.append((key, float(conc), n_sign_in, int(n_combo)))

    if not hits:
        return []

    # 重複キー除去
    # 同一キーで複数ヒット：
    #   min_conc ≤ 50% かつ (max - min) > 10pt → 最大集中度の1件のみ残す
    #   それ以外（差が小さい or min が50%超）    → 全件残す
    key_groups: dict = {}
    for k, c, ns, nc in hits:
        key_groups.setdefault(k, []).append((c, ns, nc))
    hits = []
    for k, entries in key_groups.items():
        if len(entries) == 1:
            hits.append((k, *entries[0]))
            continue
        concs   = [c for c, _, _ in entries]
        max_c   = max(concs)
        min_c   = min(concs)
        if min_c <= 0.50 and (max_c - min_c) > 0.10:
            # 最大集中度の1件だけ残す
            best = max(entries, key=lambda x: x[0])
            hits.append((k, *best))
        else:
            # 全件残す
            for c, ns, nc in entries:
                hits.append((k, c, ns, nc))

    # 冗長除去：サブセット(i) vs スーパーセット(j) の比較
    #   差 5pt 以内 (cj - ci <= 0.05) → 追加条件の効果が小さい → シンプルな i を残し j を削除
    #   差 5pt 超  (cj - ci >  0.05) → 追加条件が有意 → 具体的な j を残し i を削除
    n_hits = len(hits)
    redundant = set()
    for i in range(n_hits):
        if i in redundant:
            continue
        ki, ci = hits[i][0], hits[i][1]
        for j in range(n_hits):
            if i == j or j in redundant:
                continue
            kj, cj = hits[j][0], hits[j][1]
            if ki < kj:          # i は j のサブセット
                if cj - ci <= 0.05:
                    redundant.add(j)   # スーパーセット(j)を削除、シンプル(i)を残す
                else:
                    redundant.add(i)   # サブセット(i)を削除、具体的(j)を残す
                    break

    kept = [(k, c, ns, nc)
            for idx, (k, c, ns, nc) in enumerate(hits)
            if idx not in redundant]

    # 集中度降順 → 要素数降順でソート、上位10件に制限
    kept.sort(key=lambda x: (-x[1], -len(x[0])))
    kept = kept[:10]

    # 出力文字列生成（値のみ・ラベルなし）
    elem_order = [c for c, _ in _SIT_ELEMENTS]

    bullets = []
    for key, conc, n_s, n_c in kept:
        # 要素を定義順に並べ、値のみ表示
        sorted_items = sorted(key, key=lambda x: elem_order.index(x[0]) if x[0] in elem_order else 99)
        cond_parts = [val for _, val in sorted_items]
        cond_str = ' + '.join(cond_parts)
        bullets.append(f'【状況】{cond_str} のとき {conc:.0%}（{n_s}/{n_c}プレー）')

    return bullets


def consider_sign_situation_header(df, top_n=5, min_n=4, threshold=0.60, max_items_per_sign=5):
    """
    3rd Situation 用：SIGNヘッダー前に、上位N個のSIGNのシチュエーション分析をまとめて返す。
    横断データ（DISTゾーンを問わない全3rdプレー）で分析する。
    ・min_n=4、threshold=0.60（Normal の per-SIGN 分析より厳しめに設定）
    ・1SIGNあたり最大 max_items_per_sign 件に制限
    """
    bullets = []
    n_total = len(df)
    if n_total < min_n:
        return bullets

    vc = _counts(df, 'SIGN_D')
    if len(vc) == 0:
        return bullets

    # situation df を1回だけ構築してすべてのSIGNで使い回す
    _df_sit = _build_situation_df(df)

    no_trend_signs = []
    for sign_val in vc.index[:top_n]:
        sign_dv = _display_val('SIGN_D', sign_val)
        n_sign  = int(vc[sign_val])
        pct     = n_sign / n_total

        sit_bullets = consider_sign_situation(
            df, sign_val, min_n=min_n, threshold=threshold, max_combo=3,
            min_coverage=0.20, _df_sit=_df_sit)
        sit_bullets = sit_bullets[:max_items_per_sign]

        if sit_bullets:
            bullets.append(f'▶ {sign_dv}（{pct:.0%}、{n_sign}/{n_total}プレー）の出現状況')
            bullets.extend([f'  {b}' for b in sit_bullets])
        else:
            no_trend_signs.append(sign_dv)

    if no_trend_signs:
        if len(no_trend_signs) == len(vc.index[:top_n]):
            # 全SIGNが傾向なし
            bullets.append('※各SIGNは特定のシチュエーションに依存した出現傾向を示さない')
        else:
            bullets.append('その他Sign：出現シチュエーションの傾向小')

    return bullets


def consider_n3_comparison(df_n, df_3):
    """Normal vs 3rdの横断比較（ヤード非依存）— 3rdセクション冒頭に表示"""
    bullets = []
    if df_n is None or df_3 is None or len(df_n) == 0 or len(df_3) == 0:
        return bullets

    # FRONTの構成変化
    for fv in _counts(df_n, 'DEF_FRONT').index:
        for obs in _check_n3(fv, 'DEF_FRONT', df_n, df_3):
            bullets.append(f'FRONT {fv}：{obs}')

    # SIGNの出現頻度変化
    for sv in _counts(df_n, 'SIGN_D').index:
        for obs in _check_n3(sv, 'SIGN_D', df_n, df_3):
            bullets.append(f'SIGN {_display_val("SIGN_D", sv)}：{obs}')

    # COVERAGEの構成変化
    for cv in _counts(df_n, 'COVERAGE_NORM').index:
        dv = _display_val('COVERAGE_NORM', cv)
        for obs in _check_n3(cv, 'COVERAGE_NORM', df_n, df_3):
            bullets.append(f'{dv}：{obs}')

    return bullets
