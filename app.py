import os
import re

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ─────────────────────────────────────────────────────────
# 🟢 1. 기본 설정
# ─────────────────────────────────────────────────────────
st.set_page_config(page_title="2027년 사업계획 at a glance", layout="wide")

BASE_YEAR = 2026   # 기준연도 (계획 vs 실적)
PLAN_YEAR = 2027   # 계획연도 (당초 vs 실천)

# 4개 비교 계열: 어느 시트의 몇 년도 데이터를 쓸지 여기서만 바꾸면 됩니다.
#   ① 2026 계획 : 공급량_사업계획      (연=2026)
#   ② 2026 실적 : 공급량_실천사업계획  (연=2026, 1~6월 실적 + 7~12월 예상)
#   ③ 2027 당초 : 공급량_사업계획      (연=2027)
#   ④ 2027 실천 : 공급량_실천사업계획  (연=2027)
SERIES_SOURCE = {
    "①": {"sheet": "공급량_사업계획", "year": BASE_YEAR},
    "②": {"sheet": "공급량_실천사업계획", "year": BASE_YEAR},
    "③": {"sheet": "공급량_사업계획", "year": PLAN_YEAR},
    "④": {"sheet": "공급량_실천사업계획", "year": PLAN_YEAR},
}
SERIES_HEADER = {"①": "① 계획", "②": "② 실적", "③": "③ 당초", "④": "④ 실천"}
SERIES_FULL = {
    "①": f"{BASE_YEAR} 계획",
    "②": f"{BASE_YEAR} 실적(예상)",
    "③": f"{PLAN_YEAR} 당초",
    "④": f"{PLAN_YEAR} 실천",
}
# 증감 / 대비 열: (대상, 기준)
COMPARISONS = [("②", "①"), ("③", "②"), ("④", "②"), ("④", "③")]

# ─────────────────────────────────────────────────────────
# 🟢 2. 용도 매핑 (수송용은 표에서 CNG / BIO 로 분리)
# ─────────────────────────────────────────────────────────
MAPPING_SUPPLY = {
    "취사용": "가정용", "개별난방용": "가정용", "중앙난방용": "가정용",
    "개별난방": "가정용", "중앙난방": "가정용",
    "영업용": "영업/업무용",
    "일반용(1)": "영업/업무용", "일반용1": "영업/업무용",
    "일반용1(영업)": "영업/업무용", "일반용1(업무)": "영업/업무용",
    "일반용(2)": "영업/업무용", "일반용2": "영업/업무용",
    "업무난방용": "영업/업무용", "냉난방용": "영업/업무용", "냉방용": "영업/업무용",
    "주한미군": "영업/업무용", "업무용": "영업/업무용", "영업/업무용": "영업/업무용",
    "산업용": "산업용",
    "수송용(CNG)": "CNG", "CNG": "CNG",
    "수송용(BIO)": "BIO", "BIO": "BIO",
    "열병합용": "열병합용", "열병합용1": "열병합용",
    "연료전지": "연료전지", "연료전지용": "연료전지",
    "자가열전용": "자가열전용",
    "열전용설비용": "열전용설비용(주택외)", "열전용설비용(주택외)": "열전용설비용(주택외)",
}

TRANSPORT_GROUPS = ["CNG", "BIO"]
MAIN_ROWS = ["가정용", "산업용", "영업/업무용", "열병합용", "연료전지", "자가열전용", "열전용설비용(주택외)"]
DETAIL_ORDER = MAIN_ROWS + TRANSPORT_GROUPS
CHART_ORDER = MAIN_ROWS + ["수송용"]
DISPLAY_NAME = {"열병합용": "열병합", "열전용설비용(주택외)": "열전용설비용"}

EXCLUDE_COLS = ['연', '월', '날짜', '평균기온', '총공급량', '총합계', '비교(V-W)', '소 계', '소계']


# ─────────────────────────────────────────────────────────
# 🟢 3. 전처리 함수
# ─────────────────────────────────────────────────────────
@st.cache_data(ttl=600)
def load_all_sheets(uploaded_file):
    if uploaded_file is None:
        return {}
    data_dict = {}
    try:
        excel = pd.ExcelFile(uploaded_file, engine='openpyxl')
        for sheet in excel.sheet_names:
            data_dict[sheet] = excel.parse(sheet)
    except Exception:
        if hasattr(uploaded_file, 'seek'):
            uploaded_file.seek(0)
        data_dict["default"] = pd.read_csv(uploaded_file, encoding='utf-8-sig')
    return data_dict


def clean_df(df):
    if df is None:
        return pd.DataFrame()
    df = df.copy()
    if len(df.columns) > 0 and isinstance(df.columns[0], str) and "데이터 학습기간" in df.columns[0]:
        new_header = df.iloc[0]
        df = df[1:]
        df.columns = new_header
    df.columns = df.columns.astype(str).str.strip()
    cols = [c for c in df.columns if not ("Unnamed" in c or re.search(r'^열\s*\d+', c) or c == '0')]
    df = df[cols]
    if '날짜' in df.columns:
        df['날짜'] = pd.to_datetime(df['날짜'], errors='coerce')
        if '연' not in df.columns:
            df['연'] = df['날짜'].dt.year
        if '월' not in df.columns:
            df['월'] = df['날짜'].dt.month
    return df


def make_long_data(df):
    df = clean_df(df)
    if df.empty or '연' not in df.columns:
        return pd.DataFrame()
    if '월' not in df.columns:
        df['월'] = 1

    df['연'] = pd.to_numeric(df['연'], errors='coerce')
    df['월'] = pd.to_numeric(df['월'], errors='coerce')
    df = df.dropna(subset=['연'])   # 시트 하단의 연도 없는 잔여 행 제거

    records = []
    for col in df.columns:
        if col in EXCLUDE_COLS:
            continue
        val_series = pd.to_numeric(df[col], errors='coerce').fillna(0)
        if val_series.sum() == 0:
            continue
        sub = df[['연', '월']].copy()
        sub['그룹'] = MAPPING_SUPPLY.get(col, col)
        sub['값'] = val_series
        records.append(sub[sub['값'] != 0])

    if not records:
        return pd.DataFrame()
    return pd.concat(records, ignore_index=True)


def find_sheet(data_dict, name):
    """정확한 시트명이 없으면 키워드로 대체 탐색"""
    if name in data_dict:
        return data_dict[name]
    if "실천" in name:
        return next((d for n, d in data_dict.items() if "실천" in n), None)
    if "사업계획" in name:
        return next((d for n, d in data_dict.items() if "사업계획" in n and "실천" not in n), None)
    return None


def build_summary(data_dict, factor):
    """용도(세부) × ①~④ 연간 합계 표 + 계열별 데이터 존재 여부"""
    long_cache, series, available = {}, {}, {}
    for key, src in SERIES_SOURCE.items():
        sheet = src["sheet"]
        if sheet not in long_cache:
            raw = find_sheet(data_dict, sheet)
            long_cache[sheet] = make_long_data(raw) if raw is not None else pd.DataFrame()
        long_df = long_cache[sheet]
        if long_df.empty:
            s = pd.Series(dtype=float)
        else:
            s = long_df[long_df['연'] == src["year"]].groupby('그룹')['값'].sum()
        series[key] = s * factor
        available[key] = (not s.empty) and s.sum() != 0

    df = pd.DataFrame(series).fillna(0)
    extras = sorted(g for g in df.index if g not in DETAIL_ORDER)
    df = df.reindex(DETAIL_ORDER + extras).fillna(0)
    return df, available


def to_chart_df(df_detail):
    """CNG/BIO → 수송용으로 묶은 차트용 표"""
    df = df_detail.copy()
    df.index = ["수송용" if g in TRANSPORT_GROUPS else g for g in df.index]
    df = df.groupby(level=0).sum()
    extras = sorted(g for g in df.index if g not in CHART_ORDER)
    return df.reindex([g for g in CHART_ORDER + extras if g in df.index])


# ─────────────────────────────────────────────────────────
# 🟢 4. 비교 요약표 (HTML, 엑셀 양식과 동일 구조)
# ─────────────────────────────────────────────────────────
TABLE_CSS = """
<style>
.glance-wrap { overflow-x: auto; margin-bottom: 1rem; }
.glance { width: 100%; min-width: 1150px; border-collapse: collapse; font-family: sans-serif; font-size: 14px; }
.glance th, .glance td { border: 1px solid #d0d4da; padding: 7px 9px; color: #31333F; }
.glance thead th { background: #FFF2CC; text-align: center; font-weight: 600; }
.glance tbody td { text-align: right; background: #ffffff; }
.glance td.label { text-align: center; background: #f8f9fa; font-weight: 600; }
.glance .blk { border-left: 3px solid #333333 !important; }
.glance .strong { font-weight: 700; }
.glance tr.subtotal td { background: #DDEBF7; }
.glance tr.total td { background: #FCE4D6; font-weight: 700; }
.glance tr.spacer td { border: none; background: transparent; height: 12px; padding: 0; }
.glance td.pos { color: #0055a4; }
.glance td.neg { color: #cc0000; }
.glance-unit { text-align: right; color: gray; font-size: 12px; }
</style>
"""


def value_cells(vals, available):
    """①~④ 값 4칸 + 증감 4칸 + 대비 4칸"""
    cells = []
    for i, key in enumerate(["①", "②", "③", "④"]):
        cls = []
        if i == 0:
            cls.append("blk")
        if key == "④":
            cls.append("strong")
        txt = f"{vals[key]:,.0f}" if available[key] else "-"
        cells.append(f'<td class="{" ".join(cls)}">{txt}</td>')

    for i, (tgt, base) in enumerate(COMPARISONS):
        cls = ["blk"] if i == 0 else []
        if available[tgt] and available[base]:
            d = vals[tgt] - vals[base]
            if round(d) > 0:
                cls.append("pos")
            elif round(d) < 0:
                cls.append("neg")
            txt = f"{d:,.0f}"
        else:
            txt = "-"
        cells.append(f'<td class="{" ".join(cls)}">{txt}</td>')

    for i, (tgt, base) in enumerate(COMPARISONS):
        cls = ["blk"] if i == 0 else []
        if (tgt, base) == ("④", "③"):
            cls.append("strong")
        if available[tgt] and available[base] and vals[base] != 0:
            txt = f"{vals[tgt] / vals[base] * 100:,.1f}%"
        else:
            txt = "-"
        cells.append(f'<td class="{" ".join(cls)}">{txt}</td>')
    return "".join(cells)


def render_glance_table(df_detail, available, short_unit):
    head = (
        '<thead>'
        '<tr>'
        '<th rowspan="2" colspan="2">구분</th>'
        f'<th colspan="2" class="blk">{BASE_YEAR}년</th>'
        f'<th colspan="2">{PLAN_YEAR}년</th>'
        '<th colspan="4" class="blk">증감</th>'
        '<th colspan="4" class="blk">대비</th>'
        '</tr><tr>'
        + "".join(
            f'<th class="{"blk" if k == "①" else ""}">'
            f'{"<b>" + SERIES_HEADER[k] + "</b>" if k == "④" else SERIES_HEADER[k]}</th>'
            for k in ["①", "②", "③", "④"]
        )
        + "".join(f'<th class="{"blk" if i == 0 else ""}">{t}-{b}</th>' for i, (t, b) in enumerate(COMPARISONS))
        + "".join(
            f'<th class="{"blk" if i == 0 else ""}">{"<b>" if (t, b) == ("④", "③") else ""}{t}/{b}'
            f'{"</b>" if (t, b) == ("④", "③") else ""}</th>'
            for i, (t, b) in enumerate(COMPARISONS)
        )
        + '</tr></thead>'
    )

    rows = []
    # 일반 용도
    for g in [g for g in df_detail.index if g not in TRANSPORT_GROUPS]:
        rows.append(
            f'<tr><td class="label" colspan="2">{DISPLAY_NAME.get(g, g)}</td>'
            f'{value_cells(df_detail.loc[g], available)}</tr>'
        )
    # 수송용 (CNG / BIO / 소계)
    transport_sum = df_detail.loc[TRANSPORT_GROUPS].sum()
    rows.append(f'<tr><td class="label" rowspan="3">수송용</td><td class="label">CNG</td>'
                f'{value_cells(df_detail.loc["CNG"], available)}</tr>')
    rows.append(f'<tr><td class="label">BIO</td>{value_cells(df_detail.loc["BIO"], available)}</tr>')
    rows.append(f'<tr class="subtotal"><td class="label">소계</td>{value_cells(transport_sum, available)}</tr>')

    # 합계
    total = df_detail.sum()
    rows.append(f'<tr class="total"><td class="label" colspan="2">합계</td>{value_cells(total, available)}</tr>')

    # 분류 요약 (가정용 / 산업용 / 기타)
    rows.append('<tr class="spacer"><td colspan="14"></td></tr>')
    home, ind = df_detail.loc["가정용"], df_detail.loc["산업용"]
    for name, vals in [("가정용", home), ("산업용", ind), ("기타", total - home - ind)]:
        rows.append(f'<tr class="total"><td class="label" colspan="2">{name}</td>{value_cells(vals, available)}</tr>')

    html = (
        TABLE_CSS
        + f'<div class="glance-unit">(단위 : {short_unit})</div>'
        + f'<div class="glance-wrap"><table class="glance">{head}<tbody>{"".join(rows)}</tbody></table></div>'
    )
    st.markdown(html, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────
# 🟢 5. 차트
# ─────────────────────────────────────────────────────────
def unit_annotation(fig, short_unit):
    fig.add_annotation(x=1.0, y=1.05, xref="paper", yref="paper", text=f"단위: {short_unit}",
                       showarrow=False, font=dict(size=12, color="gray"), xanchor="right", yanchor="bottom")


def draw_waterfall(df_chart, base, target, short_unit):
    base_tot, tgt_tot = df_chart[base].sum(), df_chart[target].sum()
    diffs = (df_chart[target] - df_chart[base]).tolist()

    labels = [SERIES_FULL[base]] + [DISPLAY_NAME.get(g, g) for g in df_chart.index] + [SERIES_FULL[target]]
    fig = go.Figure(go.Waterfall(
        orientation="v",
        measure=["absolute"] + ["relative"] * len(diffs) + ["total"],
        x=labels,
        y=[base_tot] + diffs + [tgt_tot],
        text=[f"{base_tot:,.0f}"] + [f"{v:,.0f}" if round(v) != 0 else "" for v in diffs] + [f"{tgt_tot:,.0f}"],
        textposition="outside",
        decreasing={"marker": {"color": "#ff7f0e"}},
        increasing={"marker": {"color": "#1f77b4"}},
        totals={"marker": {"color": "#2ca02c"}},
        connector={"line": {"color": "gray", "width": 1.5}},
    ))

    running, cur = [base_tot], base_tot
    for v in diffs:
        cur += v
        running.append(cur)
    y_min, y_max = min(running + [tgt_tot]), max(running + [tgt_tot])
    pad = (y_max - y_min) if y_max != y_min else y_max * 0.1
    fig.update_yaxes(range=[max(0, y_min - pad), y_max + pad])
    fig.update_layout(title=f"{SERIES_FULL[base]} → {SERIES_FULL[target]} 용도별 증감 브릿지",
                      margin=dict(t=60, b=40))
    unit_annotation(fig, short_unit)
    return fig


def draw_bullet_chart(df_data, base, target, short_unit, show_legend=False):
    fig = go.Figure()
    fig.add_trace(go.Bar(y=df_data.index, x=df_data['기준'], name=SERIES_FULL[base], orientation='h',
                         marker_color='#e2e6ea', width=0.8, hoverinfo='x+name'))
    fig.add_trace(go.Bar(y=df_data.index, x=df_data['대상'], name=SERIES_FULL[target], orientation='h',
                         marker_color='#1f77b4', width=0.5,
                         text=[f"<b>{v:,.0f}</b>" for v in df_data['대상']], textposition='outside',
                         textfont=dict(size=14, color='black'), hoverinfo='x+name'))

    for idx, row in df_data.iterrows():
        d = row['차이']
        diff_text = f"▲ {d:,.0f}" if round(d) > 0 else (f"▼ {abs(d):,.0f}" if round(d) < 0 else "-")
        diff_color = "#d62728" if d < 0 else "#2ca02c"
        fig.add_annotation(
            y=idx, x=1.02, xref="paper", yref="y", showarrow=False, xanchor="left", align="left",
            text=f"<span style='color:{diff_color}; font-size:14px'><b>{diff_text}</b></span> "
                 f"<span style='color:gray; font-size:13px'>({row['대비']:.1f}%)</span>",
        )

    max_val = max(df_data['기준'].max(), df_data['대상'].max()) or 1
    unit_annotation(fig, short_unit)
    fig.update_layout(
        barmode='overlay',
        height=len(df_data) * 70 + 80,
        xaxis=dict(showgrid=True, gridcolor='#f0f0f0', title="", range=[0, max_val * 1.15]),
        yaxis=dict(title="", tickfont=dict(size=14), automargin=False),
        margin=dict(l=180, r=150, t=60, b=20),
        showlegend=show_legend,
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="left", x=0) if show_legend else None,
    )
    return fig


def comparison_selector(label, options, key):
    fmt = {f"{t}vs{b}": f"{SERIES_FULL[b]} → {SERIES_FULL[t]}  ({t}-{b})" for t, b in options}
    default = next((i for i, o in enumerate(options) if o == ("④", "②")), 0)
    choice = st.radio(label, list(fmt.keys()), index=default, horizontal=True,
                      format_func=lambda k: fmt[k], key=key)
    t, b = choice.split("vs")
    return t, b


# ─────────────────────────────────────────────────────────
# 🟢 6. One Page Review
# ─────────────────────────────────────────────────────────
def render_one_page_review(data_dict, unit, heating_value):
    short_unit = "GJ" if "GJ" in unit else "천m³"
    factor = 1 / 1000 if "GJ" in unit else 1 / heating_value / 1000   # 원자료 MJ 기준

    df_detail, available = build_summary(data_dict, factor)

    if not any(available.values()):
        st.warning("데이터가 부족합니다. 공급량_사업계획 및 공급량_실천사업계획 시트를 확인해주세요.")
        return

    missing = [SERIES_FULL[k] for k, ok in available.items() if not ok]
    if missing:
        st.info(f"💡 아직 데이터가 없는 항목: **{', '.join(missing)}** — 표에는 '-'로 표시되고 관련 비교는 생략됩니다.")

    df_chart = to_chart_df(df_detail)
    totals = df_detail.sum()
    valid_comps = [(t, b) for t, b in COMPARISONS if available[t] and available[b]]

    # ==========================================
    # 💡 1. 핵심 지표 & 폭포수 차트
    # ==========================================
    st.markdown("#### 🌊 핵심 지표 및 용도별 증감 요인 폭포수 차트")

    def metric(col, key, base=None):
        if not available[key]:
            col.metric(f"{key} {SERIES_FULL[key]}", "-")
            return
        delta = None
        if base and available[base] and totals[base] != 0:
            d = totals[key] - totals[base]
            delta = f"{d:,.0f} ({totals[key] / totals[base] * 100:.1f}%) vs {base}"
        col.metric(f"{key} {SERIES_FULL[key]}", f"{totals[key]:,.0f} {short_unit}", delta=delta)

    c1, c2, c3, c4 = st.columns(4)
    metric(c1, "①")
    metric(c2, "②", "①")
    metric(c3, "③", "②")
    metric(c4, "④", "②")

    if valid_comps:
        t, b = comparison_selector("비교 기준 선택 (폭포수 차트)", valid_comps, "wf_comp")
        st.plotly_chart(draw_waterfall(df_chart, b, t, short_unit), width="stretch")
    st.markdown("---")

    # ==========================================
    # 💡 2. 비교 요약표
    # ==========================================
    st.markdown(f"#### 🚥 {BASE_YEAR}년 계획·실적 vs {PLAN_YEAR}년 사업계획 요약표")
    render_glance_table(df_detail, available, short_unit)
    st.markdown("---")

    # ==========================================
    # 💡 3. 불릿 차트
    # ==========================================
    if not valid_comps:
        return
    st.markdown("#### 🎯 용도별 세부 비교")
    t, b = comparison_selector("비교 기준 선택 (세부 비교)", valid_comps, "bullet_comp")

    df_perf = df_chart[[b, t]].copy()
    df_perf.columns = ['기준', '대상']
    df_perf.loc['총계'] = df_perf.sum()
    df_perf['차이'] = df_perf['대상'] - df_perf['기준']
    df_perf['대비'] = (df_perf['대상'] / df_perf['기준'].replace(0, pd.NA) * 100).fillna(0).astype(float)

    others_mask = ~df_perf.index.isin(['총계', '가정용', '산업용'])
    others = df_perf.loc[others_mask, ['기준', '대상']].sum()
    others_row = pd.DataFrame([others], index=['기타'])
    others_row['차이'] = others_row['대상'] - others_row['기준']
    others_row['대비'] = others_row['대상'] / others_row['기준'] * 100 if others['기준'] else 0.0

    part1 = pd.concat([others_row, df_perf.loc[['산업용', '가정용', '총계']]])
    part2 = df_perf.loc[others_mask].sort_values('대상', ascending=True)

    st.markdown("##### 📌 [요약 (분류 변경)]")
    st.plotly_chart(draw_bullet_chart(part1, b, t, short_unit, show_legend=True), width="stretch")
    if not part2.empty:
        st.markdown("##### 📌 [세부용도 (기타 용도 나타냄)]")
        st.plotly_chart(draw_bullet_chart(part2, b, t, short_unit), width="stretch")


# ─────────────────────────────────────────────────────────
# 🟢 7. 메인 실행
# ─────────────────────────────────────────────────────────
def main():
    st.title(f"📈 {PLAN_YEAR}년 사업계획 at a glance")
    st.caption(f"{BASE_YEAR}년 계획 대비 실적 · {BASE_YEAR}년 실적 대비 {PLAN_YEAR}년 계획량 비교")

    with st.sidebar:
        st.header("⚙️ 기본 설정")
        unit = st.radio("단위 선택", ["열량 (GJ)", "부피 (천m³)"], index=0)
        heating_value = 42.563
        if "부피" in unit:
            heating_value = st.number_input("(기준열량 MJ/Nm3 : 42.563 )", value=42.563, format="%.3f")
        st.markdown("---")
        st.subheader("📂 데이터 업로드")
        up_supply = st.file_uploader("공급량 데이터 업로드 (새 파일이 있으면 우선 반영됩니다)", type=["xlsx", "csv"])

    default_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "공급량실적_계획_실적_MJ.xlsx")
    target_file = up_supply if up_supply is not None else (default_file if os.path.exists(default_file) else None)

    if target_file is None:
        st.info("👈 좌측 사이드바에서 공급량 파일을 업로드하거나, 프로젝트 폴더에 공급량실적_계획_실적_MJ.xlsx 파일을 배치해 주세요.")
        return

    render_one_page_review(load_all_sheets(target_file), unit, heating_value)


if __name__ == "__main__":
    main()
