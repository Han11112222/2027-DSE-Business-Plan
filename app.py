import os
import re

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ─────────────────────────────────────────────────────────
# 🟢 1. 기본 설정
# ─────────────────────────────────────────────────────────
st.set_page_config(page_title="2027년 사업계획 at a glance", layout="wide")

BASE_YEAR = 2026   # 기준연도 (계획 vs 실적)
PLAN_YEAR = 2027   # 계획연도 (당초 vs 실천)

# 실적 데이터: 구글 스프레드시트 [new ver] 상품별 분배(MJ) 블록 (C48 헤더, C49:ZZ65 값)
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1gIhArPlLBJ9fwlaqXtZWxiKlSK9hbRuz6HcDw_Yf7Is/edit?gid=0#gid=0"
GSHEET_BLOCK_LABEL = "[new ver]"   # 이 제목이 있는 블록을 찾아서 읽음
GSHEET_HEADER_ROW = 48             # 제목을 못 찾을 때 사용할 헤더 행 (엑셀 행번호)
GSHEET_LAST_ROW = 65
GSHEET_LAST_COL = 702              # ZZ열

EXCEL_FILE = "공급량실적_계획_실적_MJ.xlsx"

# 계획 데이터: 구글 스프레드시트 (3. 공급량 분석 탭 전용)
#   블록 1개 = 제목행 + 헤더행('구분','1월'~'12월','합계') + 용도별 행 + '합계' 행
#   같은 연도에 블록이 2개면 1번째 = V1(제출 버전), 2번째 = V2(마케팅 버전)
#   2027~2029년 계획은 시트 하단에 같은 형식으로 블록을 추가하면 자동 인식
PLAN_GSHEET_URL = "https://docs.google.com/spreadsheets/d/1PSzKts5lL_zNNi_vfKlW1CdNZasTA-jBj_UCNEtu-x0/edit?gid=0#gid=0"
PLAN_SHEET_UNIT_TO_MJ = 1000   # 계획 시트 단위(GJ) → 앱 내부 단위(MJ)
# 제목/헤더를 못 찾을 때 사용하는 고정 위치 (연도, 버전, 시작행, 끝행 / 엑셀 행번호, 헤더 포함)
PLAN_FIXED_BLOCKS = [(2026, "V1", 3, 23), (2026, "V2", 27, 47)]

SERIES_FULL = {
    "①": f"{BASE_YEAR} 계획",
    "②": f"{BASE_YEAR} 실적(예상)",
    "③": f"{PLAN_YEAR} 당초",
    "④": f"{PLAN_YEAR} 실천",
}
SERIES_HEADER = {"①": "① 계획", "②": "② 실적", "③": "③ 당초", "④": "④ 실천"}
COMPARISONS = [("②", "①"), ("③", "②"), ("④", "②"), ("④", "③")]   # (대상, 기준)

# ─────────────────────────────────────────────────────────
# 🟢 2. 용도 매핑 (수송용은 표에서 CNG / BIO 로 분리)
# ─────────────────────────────────────────────────────────
MAPPING_SUPPLY = {
    "취사용": "가정용", "개별난방용": "가정용", "중앙난방용": "가정용",
    "개별난방": "가정용", "중앙난방": "가정용",
    "영업용": "영업/업무용", "일반용": "영업/업무용",
    "일반용(1)": "영업/업무용", "일반용1": "영업/업무용",
    "일반용1(영업)": "영업/업무용", "일반용1(업무)": "영업/업무용",
    "일반용(2)": "영업/업무용", "일반용2": "영업/업무용",
    "업무난방용": "영업/업무용", "냉난방용": "영업/업무용", "냉방용": "영업/업무용",
    "냉난방공조용": "영업/업무용",
    "주한미군": "영업/업무용", "업무용": "영업/업무용", "영업/업무용": "영업/업무용",
    "산업용": "산업용",
    "수송용(CNG)": "CNG", "CNG": "CNG", "수송용": "CNG",
    "수송용(BIO)": "BIO", "BIO": "BIO", "BIO가스": "BIO",
    "열병합용": "열병합용", "열병합용1": "열병합용",
    "연료전지": "연료전지", "연료전지용": "연료전지",
    "자가열전용": "자가열전용",
    "냉난방": "영업/업무용", "업무난방": "영업/업무용", "열병합": "열병합용",   # 계획 시트 표기
    "열전용설비용": "열전용설비용(주택외)", "열전용설비용(주택외)": "열전용설비용(주택외)",
}

TRANSPORT_GROUPS = ["CNG", "BIO"]
MAIN_ROWS = ["가정용", "산업용", "영업/업무용", "열병합용", "연료전지", "자가열전용", "열전용설비용(주택외)"]
DETAIL_ORDER = MAIN_ROWS + TRANSPORT_GROUPS
CHART_ORDER = MAIN_ROWS + ["수송용"]
DISPLAY_NAME = {"열병합용": "열병합", "열전용설비용(주택외)": "열전용설비용"}

EXCLUDE_COLS = ['연', '월', '날짜', '평균기온', '총공급량', '총합계', '비교(V-W)', '소 계', '소계']

# 세련된 블루 팔레트
C_NAVY = "#1E3A5F"      # 합계(시작/끝) 막대
C_UP = "#2F6FB3"        # 증가
C_DOWN = "#9CC3E6"      # 감소
C_BASE_BAR = "#DCE6F2"  # 불릿 차트 기준 막대
C_GRID = "#EEF2F6"


def to_chart_group(g):
    return "수송용" if g in TRANSPORT_GROUPS else g


# ─────────────────────────────────────────────────────────
# 🟢 3. 데이터 로드
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


def to_csv_url(url):
    """구글시트 편집 URL → CSV 내보내기 URL (구글시트가 아니면 그대로 사용)"""
    m = re.search(r"/spreadsheets/d/([\w-]+)", url)
    if not m:
        return url
    g = re.search(r"gid=(\d+)", url)
    return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=csv&gid={g.group(1) if g else 0}"


def _norm(v):
    return re.sub(r"\s+", "", str(v)) if v is not None else ""


@st.cache_data(ttl=600, show_spinner="구글시트 실적 불러오는 중...")
def fetch_gsheet_actual(url):
    raw = pd.read_csv(to_csv_url(url), header=None, dtype=str, keep_default_na=False)
    raw = raw.iloc[:, :GSHEET_LAST_COL]

    # 1) 헤더 행 찾기: '[new ver]' 제목 아래의 '정산항목' 행
    hdr = None
    title_rows = [i for i, v in raw.iloc[:, 1].items() if GSHEET_BLOCK_LABEL in str(v)]
    if title_rows:
        start = title_rows[0]
        hdr = next((i for i in range(start, min(start + 5, len(raw))) if _norm(raw.iat[i, 2]) == "정산항목"), None)
    if hdr is None:
        hdr = GSHEET_HEADER_ROW - 1

    # 2) 월 헤더 파싱 (D열~)
    month_cols = {}
    for j in range(3, raw.shape[1]):
        m = re.search(r"(\d{4})\D+(\d{1,2})", str(raw.iat[hdr, j]))
        if m:
            month_cols[j] = (int(m.group(1)), int(m.group(2)))

    # 3) 항목 행 파싱 ('합계' 행까지, 소계·합계 제외)
    records = []
    last = min(len(raw), hdr + 1 + (GSHEET_LAST_ROW - GSHEET_HEADER_ROW) + 10)
    for i in range(hdr + 1, last):
        b, c = _norm(raw.iat[i, 1]), _norm(raw.iat[i, 2])
        if b == "합계" or c == "합계":
            break
        if not c or c in ("소계",):
            continue
        group = MAPPING_SUPPLY.get(c, c)
        for j, (y, mth) in month_cols.items():
            val = pd.to_numeric(str(raw.iat[i, j]).replace(",", "").strip(), errors="coerce")
            if pd.notna(val) and val != 0:
                records.append((y, mth, group, float(val)))

    return pd.DataFrame(records, columns=['연', '월', '그룹', '값'])


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
    empty = pd.DataFrame(columns=['연', '월', '그룹', '값'])
    df = clean_df(df)
    if df.empty or '연' not in df.columns:
        return empty
    if '월' not in df.columns:
        df['월'] = 1
    df['연'] = pd.to_numeric(df['연'], errors='coerce')
    df['월'] = pd.to_numeric(df['월'], errors='coerce')
    df = df.dropna(subset=['연', '월'])

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
        return empty
    out = pd.concat(records, ignore_index=True)
    out[['연', '월']] = out[['연', '월']].astype(int)
    return out


def find_sheet(data_dict, name):
    if name in data_dict:
        return data_dict[name]
    if "실천" in name:
        return next((d for n, d in data_dict.items() if "실천" in n), None)
    if "사업계획" in name:
        return next((d for n, d in data_dict.items() if "사업계획" in n and "실천" not in n), None)
    if "실적" in name:
        return next((d for n, d in data_dict.items() if "실적" in n and "계획" not in n), None)
    return None


def assemble_data(data_dict, gs_long):
    """①~④ 월별 계열과 세부내용용 연도별 이력 데이터 생성 (단위: MJ)"""
    plan = make_long_data(find_sheet(data_dict, "공급량_사업계획"))
    action = make_long_data(find_sheet(data_dict, "공급량_실천사업계획"))
    act_excel = make_long_data(find_sheet(data_dict, "공급량_실적"))

    # 실적 = 구글시트 우선, 구글시트에 없는 연·월만 엑셀 공급량_실적 사용
    if gs_long is not None and not gs_long.empty:
        keys = set(zip(gs_long['연'], gs_long['월']))
        keep = [(y, m) not in keys for y, m in zip(act_excel['연'], act_excel['월'])]
        actual = pd.concat([gs_long, act_excel[keep]], ignore_index=True)
    else:
        actual = act_excel

    # ② 기준연도 실적(예상) = 실적 있는 달 + 나머지 달은 실천사업계획
    act_base = actual[actual['연'] == BASE_YEAR]
    tot = act_base.groupby('월')['값'].sum()
    act_months = sorted(int(m) for m in tot[tot > 0].index)
    blend = pd.concat([act_base, action[(action['연'] == BASE_YEAR) & (~action['월'].isin(act_months))]],
                      ignore_index=True)

    series = {
        "①": plan[plan['연'] == BASE_YEAR],
        "②": blend,
        "③": plan[plan['연'] == PLAN_YEAR],
        "④": action[action['연'] == PLAN_YEAR],
    }
    series = {k: v[['월', '그룹', '값']].reset_index(drop=True) for k, v in series.items()}

    # 세부내용용 이력: 실적 / 계획 / 예상실적 / 당초 / 실천
    p = plan.copy()
    p['구분'] = p['연'].apply(lambda y: "계획" if y <= BASE_YEAR else "당초")
    a = action[action['연'] != BASE_YEAR].copy()
    a['구분'] = a['연'].apply(lambda y: "예상실적" if y < BASE_YEAR else "실천")
    hist = pd.concat([actual.assign(구분="실적"), p, a, blend.assign(연=BASE_YEAR, 구분="예상실적")],
                     ignore_index=True)
    hist = hist[hist['연'] <= PLAN_YEAR]
    hist['그룹'] = hist['그룹'].map(to_chart_group)

    return {"series": series, "hist": hist, "act_months": act_months, "actual": actual}


def scale(df, factor):
    df = df.copy()
    df['값'] = df['값'] * factor
    return df


# ─────────────────────────────────────────────────────────
# 🟢 4. 공통 차트 요소
# ─────────────────────────────────────────────────────────
def unit_annotation(fig, short_unit):
    fig.add_annotation(x=1.0, y=1.05, xref="paper", yref="paper", text=f"단위: {short_unit}",
                       showarrow=False, font=dict(size=12, color="gray"), xanchor="right", yanchor="bottom")


def style_fig(fig):
    fig.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                      font=dict(color="#31333F"), hoverlabel=dict(bgcolor="white"))
    fig.update_yaxes(gridcolor=C_GRID, zerolinecolor=C_GRID)
    return fig


def draw_waterfall(df_chart, base, target, short_unit, title=None):
    base_tot, tgt_tot = df_chart[base].sum(), df_chart[target].sum()
    diffs = (df_chart[target] - df_chart[base]).tolist()
    labels = [SERIES_FULL[base]] + [DISPLAY_NAME.get(g, g) for g in df_chart.index] + [SERIES_FULL[target]]

    fig = go.Figure(go.Waterfall(
        orientation="v",
        measure=["absolute"] + ["relative"] * len(diffs) + ["total"],
        x=labels,
        y=[base_tot] + diffs + [tgt_tot],
        text=[f"<b>{base_tot:,.0f}</b>"] + [f"{v:+,.0f}" if round(v) != 0 else "" for v in diffs]
             + [f"<b>{tgt_tot:,.0f}</b>"],
        textposition="outside",
        textfont=dict(size=12, color="#1E3A5F"),
        increasing={"marker": {"color": C_UP}},
        decreasing={"marker": {"color": C_DOWN}},
        totals={"marker": {"color": C_NAVY}},
        connector={"line": {"color": "#B8C4D1", "width": 1, "dash": "dot"}},
        hovertemplate="%{x}<br>%{y:,.0f}<extra></extra>",
    ))

    running, cur = [base_tot], base_tot
    for v in diffs:
        cur += v
        running.append(cur)
    y_min, y_max = min(running + [tgt_tot]), max(running + [tgt_tot])
    pad = (y_max - y_min) if y_max != y_min else y_max * 0.1
    fig.update_yaxes(range=[max(0, y_min - pad), y_max + pad], tickformat=",.0f")

    fig.update_layout(
        title=title or f"{SERIES_FULL[base]} → {SERIES_FULL[target]} 용도별 증감 브릿지",
        margin=dict(t=70, b=40), height=440, showlegend=False,
    )

    # 마지막(결과) 막대 위에 [증감량 증가/감소] 표시
    diff_tot = tgt_tot - base_tot
    rate = tgt_tot / base_tot * 100 if base_tot else 0
    if round(diff_tot) > 0:
        badge, badge_color = f"[▲ {diff_tot:,.0f} 증가]", "#1F5FA8"
    elif round(diff_tot) < 0:
        badge, badge_color = f"[▼ {abs(diff_tot):,.0f} 감소]", "#C0392B"
    else:
        badge, badge_color = "[변동 없음]", "#555555"
    fig.add_annotation(
        x=labels[-1], y=tgt_tot, xref="x", yref="y", yanchor="bottom", yshift=30, showarrow=False,
        text=f"<b>{badge}</b><br><span style='font-size:13px; color:gray'>{SERIES_FULL[base]} 대비 {rate:.1f}%</span>",
        font=dict(size=17, color=badge_color), align="center",
    )
    fig.update_layout(margin=dict(t=70, b=40, r=60))
    unit_annotation(fig, short_unit)
    return style_fig(fig)


def draw_bullet_chart(df_data, base, target, short_unit, show_legend=False):
    fig = go.Figure()
    fig.add_trace(go.Bar(y=df_data.index, x=df_data['기준'], name=SERIES_FULL[base], orientation='h',
                         marker_color=C_BASE_BAR, width=0.8, hoverinfo='x+name'))
    fig.add_trace(go.Bar(y=df_data.index, x=df_data['대상'], name=SERIES_FULL[target], orientation='h',
                         marker_color=C_UP, width=0.5,
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


def comparison_selector(label, options, key, default=("④", "②")):
    fmt = {f"{t}vs{b}": f"{SERIES_FULL[b]} → {SERIES_FULL[t]}  ({t}-{b})" for t, b in options}
    idx = next((i for i, o in enumerate(options) if o == default), 0)
    choice = st.radio(label, list(fmt.keys()), index=idx, horizontal=True,
                      format_func=lambda k: fmt[k], key=key)
    t, b = choice.split("vs")
    return t, b


# ─────────────────────────────────────────────────────────
# 🟢 5. 비교 요약표 (엑셀 양식과 동일 구조)
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
    cells = []
    for i, key in enumerate(["①", "②", "③", "④"]):
        cls = (["blk"] if i == 0 else []) + (["strong"] if key == "④" else [])
        txt = f"{vals[key]:,.0f}" if available[key] else "-"
        cells.append(f'<td class="{" ".join(cls)}">{txt}</td>')

    for i, (tgt, base) in enumerate(COMPARISONS):
        cls = ["blk"] if i == 0 else []
        if available[tgt] and available[base]:
            d = vals[tgt] - vals[base]
            cls += ["pos"] if round(d) > 0 else (["neg"] if round(d) < 0 else [])
            txt = f"{d:,.0f}"
        else:
            txt = "-"
        cells.append(f'<td class="{" ".join(cls)}">{txt}</td>')

    for i, (tgt, base) in enumerate(COMPARISONS):
        cls = (["blk"] if i == 0 else []) + (["strong"] if (tgt, base) == ("④", "③") else [])
        ok = available[tgt] and available[base] and vals[base] != 0
        txt = f"{vals[tgt] / vals[base] * 100:,.1f}%" if ok else "-"
        cells.append(f'<td class="{" ".join(cls)}">{txt}</td>')
    return "".join(cells)


def render_glance_table(df_detail, available, short_unit):
    bold = lambda s, on: f"<b>{s}</b>" if on else s
    head = (
        '<thead><tr>'
        '<th rowspan="2" colspan="2">구분</th>'
        f'<th colspan="2" class="blk">{BASE_YEAR}년</th>'
        f'<th colspan="2">{PLAN_YEAR}년</th>'
        '<th colspan="4" class="blk">증감</th>'
        '<th colspan="4" class="blk">대비</th>'
        '</tr><tr>'
        + "".join(f'<th class="{"blk" if k == "①" else ""}">{bold(SERIES_HEADER[k], k == "④")}</th>'
                  for k in ["①", "②", "③", "④"])
        + "".join(f'<th class="{"blk" if i == 0 else ""}">{t}-{b}</th>' for i, (t, b) in enumerate(COMPARISONS))
        + "".join(f'<th class="{"blk" if i == 0 else ""}">{bold(f"{t}/{b}", (t, b) == ("④", "③"))}</th>'
                  for i, (t, b) in enumerate(COMPARISONS))
        + '</tr></thead>'
    )

    rows = []
    for g in [g for g in df_detail.index if g not in TRANSPORT_GROUPS]:
        rows.append(f'<tr><td class="label" colspan="2">{DISPLAY_NAME.get(g, g)}</td>'
                    f'{value_cells(df_detail.loc[g], available)}</tr>')
    rows.append(f'<tr><td class="label" rowspan="3">수송용</td><td class="label">CNG</td>'
                f'{value_cells(df_detail.loc["CNG"], available)}</tr>')
    rows.append(f'<tr><td class="label">BIO</td>{value_cells(df_detail.loc["BIO"], available)}</tr>')
    rows.append(f'<tr class="subtotal"><td class="label">소계</td>'
                f'{value_cells(df_detail.loc[TRANSPORT_GROUPS].sum(), available)}</tr>')

    total = df_detail.sum()
    rows.append(f'<tr class="total"><td class="label" colspan="2">합계</td>{value_cells(total, available)}</tr>')
    rows.append('<tr class="spacer"><td colspan="14"></td></tr>')
    home, ind = df_detail.loc["가정용"], df_detail.loc["산업용"]
    for name, vals in [("가정용", home), ("산업용", ind), ("기타", total - home - ind)]:
        rows.append(f'<tr class="total"><td class="label" colspan="2">{name}</td>{value_cells(vals, available)}</tr>')

    st.markdown(
        TABLE_CSS
        + f'<div class="glance-unit">(단위 : {short_unit})</div>'
        + f'<div class="glance-wrap"><table class="glance">{head}<tbody>{"".join(rows)}</tbody></table></div>',
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────
# 🟢 6. 탭 1 — One Page Review
# ─────────────────────────────────────────────────────────
def render_one_page_review(data, factor, short_unit):
    series = {k: scale(v, factor) for k, v in data["series"].items()}
    available = {k: not v.empty and v['값'].sum() != 0 for k, v in series.items()}

    if not any(available.values()):
        st.warning("데이터가 부족합니다. 사업계획 / 실천사업계획 / 실적 데이터를 확인해주세요.")
        return

    missing = [SERIES_FULL[k] for k, ok in available.items() if not ok]
    if missing:
        st.info(f"💡 아직 데이터가 없는 항목: **{', '.join(missing)}** — 표에는 '-'로 표시되고 관련 비교는 생략됩니다.")

    df_detail = pd.DataFrame({k: v.groupby('그룹')['값'].sum() for k, v in series.items()}).fillna(0)
    extras = sorted(g for g in df_detail.index if g not in DETAIL_ORDER)
    df_detail = df_detail.reindex(DETAIL_ORDER + extras).fillna(0)

    df_chart = df_detail.copy()
    df_chart.index = [to_chart_group(g) for g in df_chart.index]
    df_chart = df_chart.groupby(level=0).sum()
    df_chart = df_chart.reindex([g for g in CHART_ORDER + sorted(set(df_chart.index) - set(CHART_ORDER))
                                 if g in df_chart.index])

    totals = df_detail.sum()
    valid_comps = [(t, b) for t, b in COMPARISONS if available[t] and available[b]]

    # ── 1. 핵심 지표 ──
    st.markdown("#### 📌 핵심 지표")

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
    st.markdown("---")

    # ── 2. 폭포수 차트 2종 ──
    st.markdown("#### 🌊 용도별 증감 요인 폭포수 차트")

    st.markdown(f"##### ① {BASE_YEAR}년 계획 대비 실적")
    if available["①"] and available["②"]:
        st.plotly_chart(draw_waterfall(df_chart, "①", "②", short_unit), width="stretch")
    else:
        st.info(f"{BASE_YEAR}년 계획 또는 실적 데이터가 없습니다.")

    st.markdown(f"##### ② {BASE_YEAR}년 실적 대비 {PLAN_YEAR}년 계획")
    plan_keys = [k for k in ["④", "③"] if available[k]]
    if available["②"] and plan_keys:
        if len(plan_keys) > 1:
            pk = st.radio(f"{PLAN_YEAR}년 계획 기준", plan_keys, horizontal=True, key="wf_plan27",
                          format_func=lambda k: SERIES_FULL[k])
        else:
            pk = plan_keys[0]
        st.plotly_chart(draw_waterfall(df_chart, "②", pk, short_unit), width="stretch")
    else:
        st.info(f"💡 {PLAN_YEAR}년 계획 데이터(사업계획 / 실천사업계획 시트의 {PLAN_YEAR}년 행)가 들어오면 표시됩니다.")
    st.markdown("---")

    # ── 3. 비교 요약표 ──
    st.markdown(f"#### 🚥 {BASE_YEAR}년 계획·실적 vs {PLAN_YEAR}년 사업계획 요약표")
    render_glance_table(df_detail, available, short_unit)
    st.markdown("---")

    # ── 4. 불릿 차트 ──
    if not valid_comps:
        return
    st.markdown("#### 🎯 용도별 세부 비교")
    t, b = comparison_selector("비교 기준 선택 (세부 비교)", valid_comps, "bullet_comp")

    df_perf = df_chart[[b, t]].copy()
    df_perf.columns = ['기준', '대상']
    df_perf.loc['총계'] = df_perf.sum()
    df_perf['차이'] = df_perf['대상'] - df_perf['기준']
    df_perf['대비'] = [(r['대상'] / r['기준'] * 100) if r['기준'] else 0.0 for _, r in df_perf.iterrows()]

    others_mask = ~df_perf.index.isin(['총계', '가정용', '산업용'])
    others = df_perf.loc[others_mask, ['기준', '대상']].sum()
    others_row = pd.DataFrame([others], index=['기타'])
    others_row['차이'] = others_row['대상'] - others_row['기준']
    others_row['대비'] = others['대상'] / others['기준'] * 100 if others['기준'] else 0.0

    part1 = pd.concat([others_row, df_perf.loc[['산업용', '가정용', '총계']]])
    part2 = df_perf.loc[others_mask].sort_values('대상', ascending=True)

    st.markdown("##### 📌 [요약 (분류 변경)]")
    st.plotly_chart(draw_bullet_chart(part1, b, t, short_unit, show_legend=True), width="stretch")
    if not part2.empty:
        st.markdown("##### 📌 [세부용도 (기타 용도 나타냄)]")
        st.plotly_chart(draw_bullet_chart(part2, b, t, short_unit), width="stretch")


# ─────────────────────────────────────────────────────────
# 🟢 7. 탭 2 — 세부내용
# ─────────────────────────────────────────────────────────
SMALL_TABLE_CSS = """
<style>
.custom-table table { width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 14px; margin-bottom: 1rem; }
.custom-table th, .custom-table td { border: 1px solid #e2e6ea; padding: 8px; color: #31333F; text-align: right; }
.custom-table th { background-color: #ffffff; text-align: center; }
.custom-table th:first-child, .custom-table td:first-child { background-color: #f8f9fa !important; text-align: center !important; font-weight: bold !important; }
.custom-table th:last-child, .custom-table td:last-child { background-color: #e2e6ea !important; font-weight: bold !important; }
</style>
"""
TYPE_ORDER = ["실적", "계획", "예상실적", "당초", "실천"]
LINE_PALETTE = px.colors.qualitative.Plotly + px.colors.qualitative.D3


def render_detail(data, factor, short_unit):
    series = {k: scale(v, factor) for k, v in data["series"].items()}
    available = {k: not v.empty and v['값'].sum() != 0 for k, v in series.items()}
    valid_comps = [(t, b) for t, b in COMPARISONS if available[t] and available[b]]
    act_months = data["act_months"]

    st.subheader(f"📊 공급량 실적 및 계획 통합 분석 ({short_unit})")

    # ── 1. 월별 비교 ──
    st.markdown("### 1️⃣ 계획 vs 실적 월별 비교")
    if not valid_comps:
        st.info("비교할 수 있는 계획/실적 데이터가 없습니다.")
    else:
        c_sel1, c_sel2 = st.columns([3, 2])
        with c_sel1:
            t, b = comparison_selector("비교 기준 선택", valid_comps, "detail_comp", default=("②", "①"))
        with c_sel2:
            option = st.selectbox("📂 조회할 항목 선택", ["전체"] + CHART_ORDER, index=0, key="sb_detail_grp")

        name_b, name_t = SERIES_FULL[b], SERIES_FULL[t]
        parts = []
        for key, nm in [(b, name_b), (t, name_t)]:
            d = series[key].copy()
            d['그룹'] = d['그룹'].map(to_chart_group)
            d['구분_비교'] = nm
            parts.append(d)
        df_comp = pd.concat(parts, ignore_index=True)
        if option != "전체":
            df_comp = df_comp[df_comp['그룹'] == option]
        title_suffix = "전체량" if option == "전체" else option
        cmap = {name_b: C_DOWN, name_t: C_NAVY}

        col_bar, col_line = st.columns([3, 7])
        with col_bar:
            df_tot = df_comp.groupby('구분_비교')['값'].sum().reindex([name_b, name_t]).fillna(0).reset_index()
            base_val = df_tot.loc[0, '값']
            df_tot['텍스트'] = [f"{v:,.0f}" if i == 0 or not base_val else f"{v:,.0f}<br>({v / base_val * 100:.1f}%)"
                              for i, v in enumerate(df_tot['값'])]
            fig_bar = px.bar(df_tot, x='구분_비교', y='값', text='텍스트', color='구분_비교', color_discrete_map=cmap)
            fig_bar.update_traces(textfont_size=16)
            fig_bar.update_layout(title=f"{title_suffix} 연간 총합 비교", showlegend=False, xaxis_title="", yaxis_title="")
            unit_annotation(fig_bar, short_unit)
            st.plotly_chart(style_fig(fig_bar), width="stretch")

        with col_line:
            df_mon = df_comp.groupby(['구분_비교', '월'])['값'].sum().reset_index().sort_values('월')
            fig_line = px.line(df_mon, x='월', y='값', color='구분_비교', markers=True, color_discrete_map=cmap,
                               category_orders={'구분_비교': [name_b, name_t]})
            fig_line.update_xaxes(tickvals=list(range(1, 13)), ticktext=[f"{i}월" for i in range(1, 13)])
            if "②" in (t, b) and act_months and max(act_months) < 12:
                fig_line.add_vrect(x0=max(act_months) + 0.5, x1=12.5, fillcolor="#9CC3E6", opacity=0.12,
                                   line_width=0, annotation_text="예상 구간", annotation_position="top left")
            fig_line.update_layout(title=f"{title_suffix} 월별 추이", xaxis_title="", yaxis_title="", legend_title="")
            unit_annotation(fig_line, short_unit)
            st.plotly_chart(style_fig(fig_line), width="stretch")

        st.markdown("##### 📋 세부 수치")
        tbl = df_comp.pivot_table(index='구분_비교', columns='월', values='값', aggfunc='sum')
        tbl = tbl.reindex(index=[name_b, name_t], columns=range(1, 13)).fillna(0)
        tbl.columns = [f"{m}월" for m in tbl.columns]
        tbl['연간 총합'] = tbl.sum(axis=1)
        tbl.loc['증감'] = tbl.loc[name_t] - tbl.loc[name_b]
        ratio = (tbl.loc[name_t] / tbl.loc[name_b].where(tbl.loc[name_b] != 0) * 100).fillna(0)

        disp = tbl.map(lambda v: f"{v:,.0f}").astype(object)
        disp.loc['대비'] = [f"{v:,.1f}%" for v in ratio]
        disp.index.name = "구분"
        html = disp.reset_index().to_html(index=False, border=0)
        st.markdown(SMALL_TABLE_CSS + f'<div class="custom-table">{html}</div>', unsafe_allow_html=True)

    st.markdown("---")

    # ── 이력 데이터 준비 ──
    df = scale(data["hist"], factor)
    if df.empty:
        return
    df['연'] = df['연'].astype(int)
    df['sort_key'] = df['연'] * 10 + df['구분'].map({k: i for i, k in enumerate(TYPE_ORDER)})
    df = df.sort_values(['sort_key', '월'])
    df['연_구분'] = df['연'].astype(str) + " (" + df['구분'] + ")"
    label_info = df.drop_duplicates('연_구분').set_index('연_구분')[['연', '구분']].to_dict('index')
    all_labels = list(label_info.keys())
    color_map = {lbl: LINE_PALETTE[i % len(LINE_PALETTE)] for i, lbl in enumerate(all_labels)}

    groups = df['그룹'].unique()
    final_group_order = [g for g in CHART_ORDER if g in groups] + sorted(g for g in groups if g not in CHART_ORDER)
    years = sorted(df['연'].unique().tolist())
    types = [t for t in TYPE_ORDER if t in df['구분'].unique()]

    # 12개월 미만 실적(진행 중인 연도)은 연간 구성비 막대에서 제외
    month_cnt = df.groupby('연_구분')['월'].nunique()
    partial = {lbl for lbl, n in month_cnt.items() if n < 12 and label_info[lbl]['구분'] == "실적"}

    def pick_labels(sel_years, sel_types):
        return [l for l in all_labels if label_info[l]['연'] in sel_years and label_info[l]['구분'] in sel_types]

    def detail_pivot(dff, labels):
        piv = dff.pivot_table(index='연_구분', columns='그룹', values='값', aggfunc='sum').fillna(0)
        piv = piv.reindex(index=labels, columns=[c for c in final_group_order if c in piv.columns]).fillna(0)
        piv['총계'] = piv.sum(axis=1)
        return piv

    # ── 2. 전체량 분석 ──
    st.markdown("### 2️⃣ 전체량 분석")
    c1, c2 = st.columns(2)
    with c1:
        sel_years_1 = st.multiselect("📅 [전체량] 조회할 연도 선택", years, default=years[-5:], key="sec1")
    with c2:
        sel_types_1 = st.multiselect("📊 [전체량] 조회할 구분 선택", types, default=types, key="type_sec1")

    if sel_years_1 and sel_types_1:
        labels_1 = pick_labels(sel_years_1, sel_types_1)
        dff1 = df[df['연_구분'].isin(labels_1)]

        st.markdown("#### 📈 전체 월별 공급량 추이")
        mon = dff1.groupby(['연_구분', '월'])['값'].sum().reset_index()
        fig1 = px.line(mon, x='월', y='값', color='연_구분', markers=True,
                       category_orders={"연_구분": labels_1}, color_discrete_map=color_map)
        fig1.update_xaxes(tickvals=list(range(1, 13)), ticktext=[f"{i}월" for i in range(1, 13)])
        fig1.update_layout(xaxis_title="", yaxis_title="", legend_title="")
        unit_annotation(fig1, short_unit)
        st.plotly_chart(style_fig(fig1), width="stretch")

        st.markdown("#### 🧱 연도/구분별 용도 구성비")
        labels_bar = [l for l in labels_1 if l not in partial]
        yr = dff1[dff1['연_구분'].isin(labels_bar)].groupby(['연_구분', '그룹'])['값'].sum().reset_index()
        fig2 = px.bar(yr, x='연_구분', y='값', color='그룹', text_auto=',.0f',
                      category_orders={"연_구분": labels_bar, "그룹": final_group_order})
        for _, row in yr.groupby('연_구분')['값'].sum().reset_index().iterrows():
            fig2.add_annotation(x=row['연_구분'], y=row['값'], text=f"{row['값']:,.0f}", showarrow=False,
                                yanchor="bottom", yshift=15, font=dict(size=14, color=C_NAVY))
        fig2.update_layout(xaxis_title="", yaxis_title="", legend_title="")
        unit_annotation(fig2, short_unit)
        st.plotly_chart(style_fig(fig2), width="stretch")
        if partial & set(labels_1):
            st.caption(f"※ 12개월 미만 실적({', '.join(sorted(partial & set(labels_1)))})은 구성비 막대에서 제외했습니다.")

        st.markdown("##### 📋 전체량 상세 수치")
        st.dataframe(detail_pivot(dff1, labels_1).style.format("{:,.0f}"), width="stretch")

    st.markdown("---")

    # ── 3. 용도별 구성 분석 ──
    st.markdown("### 3️⃣ 용도별 구성 분석")
    default_2 = [y for y in years if y in (BASE_YEAR, PLAN_YEAR)] or years[-2:]
    c1, c2 = st.columns(2)
    with c1:
        sel_years_2 = st.multiselect("📅 [용도별] 조회할 연도 선택", years, default=default_2, key="sec2")
    with c2:
        sel_types_2 = st.multiselect("📊 [용도별] 조회할 구분 선택", types, default=types, key="type_sec2")

    if sel_years_2 and sel_types_2:
        labels_2 = pick_labels(sel_years_2, sel_types_2)
        dff2 = df[df['연_구분'].isin(labels_2)]
        sel_group = st.radio("📂 조회할 용도 선택", final_group_order, index=0, horizontal=True, key="rb_group_sec3")

        st.markdown("#### 📈 연도/구분별 용도 꺾은선 추이 (월별 비교)")
        mon2 = dff2[dff2['그룹'] == sel_group].groupby(['연_구분', '월'])['값'].sum().reset_index()
        fig3 = px.line(mon2, x='월', y='값', color='연_구분', markers=True,
                       category_orders={"연_구분": labels_2}, color_discrete_map=color_map)
        fig3.update_xaxes(tickvals=list(range(1, 13)), ticktext=[f"{i}월" for i in range(1, 13)])
        fig3.update_layout(xaxis_title="", yaxis_title="", legend_title="")
        unit_annotation(fig3, short_unit)
        st.plotly_chart(style_fig(fig3), width="stretch")

        st.markdown("##### 📋 용도별 상세 수치 (비교 테이블)")
        st.dataframe(detail_pivot(dff2, labels_2).style.format("{:,.0f}"), width="stretch")


# ─────────────────────────────────────────────────────────
# 🟢 7-2. 탭 3 — 공급량 분석 (연도별 실적 vs 당초 계획 V1/V2)
# ─────────────────────────────────────────────────────────
KIND_LABEL = {"실적": "실적", "V1": "계획 V1(제출)", "V2": "계획 V2(마케팅)"}
KIND_COLOR = {"실적": C_NAVY, "V1": "#9CC3E6", "V2": "#F2A65A"}
KIND_DASH = {"실적": "solid", "V1": "dash", "V2": "dot"}
KIND_SYMBOL = {"실적": "circle", "V1": "square", "V2": "diamond"}
SA_ROWS = ["가정용", "영업/업무용", "산업용", "열병합용", "연료전지", "자가열전용", "열전용설비용(주택외)"]
SA_AGG = {"가정용", "영업/업무용"}   # 시트상 '소계' 성격의 행 (하위 품목 합산) → 소계 배경색

SA_CSS = """
<style>
.sa-wrap { overflow-x: auto; margin-bottom: 1rem; }
.sa { border-collapse: collapse; font-family: sans-serif; font-size: 13.5px; min-width: 100%; }
.sa th, .sa td { border: 1px solid #d0d4da; padding: 6px 9px; color: #31333F; white-space: nowrap; }
.sa thead th { background: #FFF2CC; text-align: center; font-weight: 600; }
.sa tbody td { text-align: right; background: #ffffff; }
.sa td.label { text-align: center; background: #f8f9fa; font-weight: 600; }
.sa td.tcol { background: #DDEBF7; font-weight: 700; }
.sa .blk { border-left: 3px solid #333333 !important; }
.sa tr.subtotal td { background: #DDEBF7; }
.sa tr.total td { background: #FCE4D6; font-weight: 700; }
.sa tr.ratio td { background: #F3F6FA; color: #1F5FA8; }
.sa-unit { text-align: right; color: gray; font-size: 12px; }
</style>
"""


def kind_label(k):
    return KIND_LABEL.get(k, f"계획 {k}")


# ── 계획 시트 파서 ────────────────────────────────────────
_MONTH_RE = re.compile(r"^(\d{1,2})월$")
_YEAR_RE = re.compile(r"(20\d{2})\s*년")


def _num(v):
    s = str(v).replace(",", "").replace(" ", "").strip()
    if s in ("", "-", "nan"):
        return float("nan")
    return pd.to_numeric(s, errors="coerce")


def parse_plan_sheet(raw):
    """계획 시트(CSV, header=None) → (long DataFrame[연,월,그룹,값(MJ),버전], 인식 결과 info)"""
    ncols = raw.shape[1]
    n = len(raw)

    def cell(i, j):
        return _norm(raw.iat[i, j]) if j < ncols else ""

    # 1) 헤더 행 찾기: A/B열에 '구분' + '1월'~'12월'
    heads = []
    for i in range(n):
        if "구분" in (cell(i, 0), cell(i, 1)):
            mc = {}
            for j in range(ncols):
                m = _MONTH_RE.match(cell(i, j))
                if m and 1 <= int(m.group(1)) <= 12:
                    mc[j] = int(m.group(1))
            if len(mc) >= 12:
                heads.append({"hdr": i, "mc": mc, "year": None, "ver": None})

    # 2) 못 찾으면 고정 위치 사용 (A3:O23 / A27:O47, 월은 C~N열)
    if not heads:
        for y, v, r0, _r1 in PLAN_FIXED_BLOCKS:
            heads.append({"hdr": r0 - 1, "mc": {2 + k: k + 1 for k in range(12)}, "year": y, "ver": v})

    # 3) 블록별 연도/버전 결정 (제목 → 없으면 등장 순서)
    count, used, prev_year, blocks, records, unmapped = {}, set(), None, [], [], set()
    for b in heads:
        title = ""
        for k in range(b["hdr"] - 1, max(b["hdr"] - 4, -1), -1):
            t = " ".join(str(raw.iat[k, j]) for j in range(min(4, ncols)) if str(raw.iat[k, j]).strip())
            if _YEAR_RE.search(t):
                title = t
                break
        m = _YEAR_RE.search(title)
        year = b["year"] or (int(m.group(1)) if m else
                             (BASE_YEAR if prev_year is None else
                              (prev_year if count.get(prev_year, 0) < 2 else prev_year + 1)))
        tl = title.lower()
        ver = b["ver"] or ("V2" if ("마케팅" in tl or "marketing" in tl)
                           else "V1" if ("normal" in tl or "제출" in tl) else f"V{count.get(year, 0) + 1}")
        if (year, ver) in used:
            ver = f"V{count.get(year, 0) + 1}"
        used.add((year, ver))
        count[year] = count.get(year, 0) + 1
        prev_year = year

        # 4) 용도 행 파싱 ('합계' 행 또는 다음 블록 시작 전까지, 소계·합계 제외)
        next_hdr = min([h["hdr"] for h in heads if h["hdr"] > b["hdr"]], default=n)
        n_rows = 0
        for i in range(b["hdr"] + 1, min(next_hdr, n)):
            a, bb = cell(i, 0), cell(i, 1)
            if a == "합계" or bb == "합계":
                break
            if a == "소계" or bb == "소계":
                continue
            item = bb or a
            if not item:
                continue
            if _YEAR_RE.search(item):   # 다음 블록 제목
                break
            group = MAPPING_SUPPLY.get(item) or MAPPING_SUPPLY.get(a)
            if group is None:
                group = item
                unmapped.add(item)
            hit = False
            for j, mth in b["mc"].items():
                val = _num(raw.iat[i, j])
                if pd.notna(val) and val != 0:
                    records.append((year, mth, group, float(val) * PLAN_SHEET_UNIT_TO_MJ, ver))
                    hit = True
            n_rows += hit
        blocks.append({"연도": year, "버전": ver, "헤더 행": b["hdr"] + 1, "용도 행 수": n_rows, "제목": title or "(제목 없음)"})

    df = pd.DataFrame(records, columns=['연', '월', '그룹', '값', '버전'])
    return df, {"blocks": blocks, "unmapped": sorted(unmapped)}


@st.cache_data(ttl=600, show_spinner="구글시트 계획 불러오는 중...")
def fetch_gsheet_plan(url):
    raw = pd.read_csv(to_csv_url(url), header=None, dtype=str, keep_default_na=False)
    return parse_plan_sheet(raw)


# ── 집계 도우미 ───────────────────────────────────────────
def kinds_in(series):
    plan_kinds = sorted({k for (k, _y) in series if k != "실적"})
    return (["실적"] if any(k == "실적" for (k, _y) in series) else []) + plan_kinds


def sa_avail_months(df):
    t = df.groupby('월')['값'].sum()
    return sorted(int(m) for m in t[t != 0].index)


def sa_month_vec(df, group=None):
    """월(1~12) 벡터. 데이터가 없는 달은 NaN (진행 중인 연도의 미래 월 등)"""
    av = sa_avail_months(df)
    d = df if group is None else df[df['그룹'].map(to_chart_group) == group]
    v = d.groupby('월')['값'].sum().reindex(range(1, 13)).fillna(0.0)
    return v.where(v.index.isin(av))


def sa_cumulative(v):
    return v.fillna(0).cumsum().where(v.notna())


def sa_build_rows(cols, tcols=()):
    """cols: [Series(그룹→값) | None]. 시트 양식과 같은 행 구성 (소계=파랑, 합계=주황)"""
    names = set()
    for c in cols:
        if c is not None:
            names |= set(c.index)
    extras = sorted(g for g in names if g not in SA_ROWS and g not in TRANSPORT_GROUPS)

    def cells(fn):
        out = []
        for k, c in enumerate(cols):
            cls = "tcol" if k in tcols else ""
            out.append(("-" if c is None else f"{fn(c):,.0f}", cls))
        return out

    rows = []
    for g in SA_ROWS + extras:
        rows.append(dict(label=DISPLAY_NAME.get(g, g) + (" (소계)" if g in SA_AGG else ""),
                         cls="subtotal" if g in SA_AGG else "", cells=cells(lambda c, g=g: c.get(g, 0.0))))
    for g in TRANSPORT_GROUPS:
        rows.append(dict(label=f"수송용 · {g}", cls="", cells=cells(lambda c, g=g: c.get(g, 0.0))))
    rows.append(dict(label="수송용 (소계)", cls="subtotal",
                     cells=cells(lambda c: sum(c.get(g, 0.0) for g in TRANSPORT_GROUPS))))
    rows.append(dict(label="합계", cls="total", cells=cells(lambda c: c.sum())))
    return rows


def sa_render_table(head_html, rows, short_unit):
    body = ""
    for r in rows:
        tds = "".join(f'<td class="{c}">{t}</td>' for t, c in r["cells"])
        body += f'<tr class="{r["cls"]}"><td class="label">{r["label"]}</td>{tds}</tr>'
    st.markdown(SA_CSS + f'<div class="sa-unit">(단위 : {short_unit})</div>'
                f'<div class="sa-wrap"><table class="sa"><thead>{head_html}</thead><tbody>{body}</tbody></table></div>',
                unsafe_allow_html=True)


def build_sa_series(actual, plan_long, factor):
    """{(구분, 연도): DataFrame[월,그룹,값]} — 구분 = 실적 / V1 / V2 ...  (단위 변환 완료)"""
    out = {}

    def add(kind, df):
        if df is None or df.empty:
            return
        for y, g in df.groupby('연'):
            out[(kind, int(y))] = scale(g[['월', '그룹', '값']], factor).reset_index(drop=True)

    add("실적", actual)
    if plan_long is not None and not plan_long.empty:
        for v, g in plan_long.groupby('버전'):
            add(v, g)
    return out


# ── 화면 ─────────────────────────────────────────────────
def render_supply_analysis(series, short_unit):
    st.subheader(f"📊 공급량 분석 — 연도별 실적 vs 당초 계획 ({short_unit})")
    if not series:
        st.warning("표시할 실적/계획 데이터가 없습니다.")
        return

    years = sorted({y for (_k, y) in series})
    kinds = kinds_in(series)
    act_years = [y for (k, y) in series if k == "실적"]
    latest_act = max(act_years) if act_years else None
    default_n = 12
    if latest_act is not None:
        av = sa_avail_months(series[("실적", latest_act)])
        if av and max(av) < 12:
            default_n = max(av)

    all_groups = {to_chart_group(g) for df in series.values() for g in df['그룹'].unique()}
    prod_opts = ["전체"] + [g for g in CHART_ORDER if g in all_groups] + sorted(all_groups - set(CHART_ORDER))
    color_by_year = {y: LINE_PALETTE[i % len(LINE_PALETTE)] for i, y in enumerate(years)}

    # ── 1. 연도별 누적 공급량 (전체량) ──
    st.markdown("### 1️⃣ 연도별 누적 공급량 비교 (전체량)")
    c1, c2 = st.columns([2, 3])
    with c1:
        n = st.select_slider("📆 누적 기준월 (1월 ~ N월)", options=list(range(1, 13)), value=default_n,
                             format_func=lambda m: f"{m}월", key="sa_n")
    with c2:
        sel_years = st.multiselect("📅 조회 연도 (실적 + 계획)", years, default=years[-8:], key="sa_years1")

    def complete(kind, y):
        df = series.get((kind, y))
        return df is not None and set(range(1, n + 1)) <= set(sa_avail_months(df))

    def cum_total(kind, y):
        if not complete(kind, y):
            return None
        df = series[(kind, y)]
        return float(df.loc[df['월'] <= n, '값'].sum())

    if sel_years:
        period = "1~12월(연간)" if n == 12 else f"1~{n}월 누적"
        fig = go.Figure()
        for k in kinds:
            ys = [cum_total(k, y) for y in sel_years]
            fig.add_trace(go.Bar(name=kind_label(k), x=[str(y) for y in sel_years], y=ys,
                                 marker_color=KIND_COLOR.get(k, "#B0B0B0"),
                                 text=[f"{v:,.0f}" if v is not None else "" for v in ys],
                                 textposition="outside", textfont=dict(size=12), cliponaxis=False,
                                 hovertemplate="%{x} " + kind_label(k) + "<br>%{y:,.0f}<extra></extra>"))
        fig.update_layout(barmode="group", title=f"연도별 공급량 — {period}", xaxis=dict(type="category"),
                          xaxis_title="", yaxis_title="", legend_title="", height=460, margin=dict(t=70))
        fig.update_yaxes(tickformat=",.0f")
        unit_annotation(fig, short_unit)
        st.plotly_chart(style_fig(fig), width="stretch")

        skipped = [f"{y} {kind_label(k)}" for y in sel_years for k in kinds
                   if (k, y) in series and not complete(k, y)]
        if skipped:
            st.caption(f"※ 1~{n}월 데이터가 모두 없는 항목은 그래프에서 제외했습니다: {', '.join(skipped)}")

        st.markdown(f"##### 📋 연도별 용도별 수치 ({period})")
        cols, head1, head2 = [], "", ""
        for y in sel_years:
            ks = [k for k in kinds if (k, y) in series]
            if not ks:
                continue
            head1 += f'<th colspan="{len(ks)}" class="blk">{y}년</th>'
            for i, k in enumerate(ks):
                head2 += f'<th class="{"blk" if i == 0 else ""}">{kind_label(k)}</th>'
                df = series[(k, y)]
                cols.append(df[df['월'] <= n].groupby('그룹')['값'].sum() if complete(k, y) else None)
        rows = sa_build_rows(cols)
        # 실적 ÷ 계획 (%) 행: 실적 열에만 표시
        col_meta = [(k, y) for y in sel_years for k in kinds if (k, y) in series]
        for pk in [k for k in kinds if k != "실적"]:
            cells = []
            for (k, y), c in zip(col_meta, cols):
                pc = cols[col_meta.index((pk, y))] if (pk, y) in col_meta else None
                ok = k == "실적" and c is not None and pc is not None and pc.sum() != 0
                cells.append((f"{c.sum() / pc.sum() * 100:,.1f}%" if ok else "-", ""))
            rows.append(dict(label=f"실적 ÷ {kind_label(pk)}", cls="ratio", cells=cells))
        head = f'<tr><th rowspan="2">구분</th>{head1}</tr><tr>{head2}</tr>'
        sa_render_table(head, rows, short_unit)
    st.markdown("---")

    # ── 2. 상품별 연도별 꺾은선 ──
    st.markdown("### 2️⃣ 상품별 연도별 추이 (꺾은선)")
    product = st.radio("📂 상품 선택", prod_opts, horizontal=True, key="sa_prod")
    c1, c2, c3 = st.columns([2, 3, 3])
    with c1:
        mode = st.radio("표시 방식", ["누적", "월별"], horizontal=True, key="sa_mode")
    with c2:
        kinds_sel = st.multiselect("📊 구분", kinds, default=kinds, format_func=kind_label, key="sa_kinds")
    with c3:
        base = latest_act if latest_act is not None else years[-1]
        years2 = st.multiselect("📅 연도", years, default=[y for y in years if y >= base - 3], key="sa_years2")

    fig2, tbl_rows = go.Figure(), []
    for y in years2:
        for k in kinds_sel:
            if (k, y) not in series:
                continue
            v = sa_month_vec(series[(k, y)], None if product == "전체" else product)
            vv = sa_cumulative(v) if mode == "누적" else v
            name = f"{y} {kind_label(k)}"
            fig2.add_trace(go.Scatter(
                x=list(range(1, 13)), y=vv.values, name=name, mode="lines+markers",
                line=dict(color=color_by_year[y], dash=KIND_DASH.get(k, "dash"), width=3 if k == "실적" else 2),
                marker=dict(size=7 if k == "실적" else 6, symbol=KIND_SYMBOL.get(k, "triangle-up")),
                connectgaps=False, hovertemplate=name + " %{x}월<br>%{y:,.0f}<extra></extra>"))
            tot = vv.dropna().iloc[-1] if mode == "누적" and vv.notna().any() else vv.sum()
            tbl_rows.append(dict(label=name, cls="", cells=[("-" if pd.isna(x) else f"{x:,.0f}", "") for x in vv]
                                 + [(f"{tot:,.0f}", "tcol")]))
    if not tbl_rows:
        st.info("선택한 연도/구분에 해당하는 데이터가 없습니다.")
    else:
        fig2.update_xaxes(tickvals=list(range(1, 13)), ticktext=[f"{i}월" for i in range(1, 13)])
        fig2.update_yaxes(tickformat=",.0f")
        fig2.update_layout(title=f"{product} — 연도별 {'누적' if mode == '누적' else '월별'} 공급량",
                           xaxis_title="", yaxis_title="", legend_title="", height=520, hovermode="closest")
        unit_annotation(fig2, short_unit)
        st.plotly_chart(style_fig(fig2), width="stretch")

        st.markdown(f"##### 📋 {product} — 연도별 {'누적' if mode == '누적' else '월별'} 수치")
        head = "<tr><th>구분</th>" + "".join(f"<th>{m}월</th>" for m in range(1, 13)) + "<th>합계</th></tr>"
        sa_render_table(head, tbl_rows, short_unit)
        st.caption("※ 합계 = " + ("해당 연도·구분의 마지막 누적값 (진행 중인 연도는 실적이 있는 달까지)" if mode == "누적"
                                  else "월별 값의 합 (진행 중인 연도는 실적이 있는 달까지)"))
    st.markdown("---")

    # ── 3. 선택 연도·구분의 용도 × 월 상세 ──
    st.markdown("### 3️⃣ 용도별 월별 상세 (연도·구분 선택)")
    c1, c2 = st.columns([1, 3])
    with c1:
        y3 = st.selectbox("📅 연도", years, index=years.index(latest_act) if latest_act in years else len(years) - 1,
                          key="sa_y3")
    ks3 = [k for k in kinds if (k, y3) in series]
    with c2:
        k3 = st.radio("📊 구분", ks3, horizontal=True, format_func=kind_label, key="sa_k3") if ks3 else None
    if k3:
        df3 = series[(k3, y3)]
        av3 = set(sa_avail_months(df3))
        fig3 = go.Figure()
        d3 = df3.assign(그룹=df3['그룹'].map(to_chart_group))
        gorder = [g for g in CHART_ORDER if g in set(d3['그룹'])] + sorted(set(d3['그룹']) - set(CHART_ORDER))
        for i, g in enumerate(gorder):
            s = d3[d3['그룹'] == g].groupby('월')['값'].sum().reindex(range(1, 13)).where(lambda x: x.index.isin(av3))
            fig3.add_trace(go.Bar(name=DISPLAY_NAME.get(g, g), x=[f"{m}월" for m in range(1, 13)], y=s.values,
                                  marker_color=LINE_PALETTE[i % len(LINE_PALETTE)],
                                  hovertemplate=DISPLAY_NAME.get(g, g) + " %{x}<br>%{y:,.0f}<extra></extra>"))
        fig3.update_layout(barmode="stack", title=f"{y3}년 {kind_label(k3)} — 월별 용도 구성",
                           xaxis_title="", yaxis_title="", legend_title="", height=460)
        fig3.update_yaxes(tickformat=",.0f")
        unit_annotation(fig3, short_unit)
        st.plotly_chart(style_fig(fig3), width="stretch")

        st.markdown(f"##### 📋 {y3}년 {kind_label(k3)} — 용도 × 월")
        cols3 = [df3[df3['월'] == m].groupby('그룹')['값'].sum() if m in av3 else None for m in range(1, 13)]
        cols3.append(df3.groupby('그룹')['값'].sum())
        head = "<tr><th>구분</th>" + "".join(f"<th>{m}월</th>" for m in range(1, 13)) + "<th>합계</th></tr>"
        sa_render_table(head, sa_build_rows(cols3, tcols={12}), short_unit)


def render_supply_page(data_dict, gs_long, plan_url, plan_status, factor, short_unit):
    """3번 탭: 실적(구글시트) + 계획(구글시트)만으로 구성. 엑셀 업로드는 필요 없음"""
    # 실적 = '2. 세부내용'과 동일 기준 (구글시트 우선, 엑셀이 있으면 구글시트에 없는 연·월만 보충)
    if data_dict:
        actual = assemble_data(data_dict, gs_long)["actual"]
    else:
        actual = gs_long
    if actual is None or actual.empty:
        st.error("실적 데이터를 불러오지 못했습니다. 사이드바의 구글시트 주소/공유 설정을 확인해주세요.")
        return

    plan_long, info = pd.DataFrame(columns=['연', '월', '그룹', '값', '버전']), None
    try:
        plan_long, info = fetch_gsheet_plan(plan_url.strip())
        if plan_long.empty:
            plan_status.warning("계획 시트에서 값을 찾지 못했습니다. 실적만 표시합니다.")
        else:
            desc = ", ".join(f"{b['연도']} {b['버전']}" for b in info["blocks"])
            plan_status.success(f"✅ 계획 반영: {desc}")
    except Exception as e:
        plan_status.warning(f"계획 구글시트 연결 실패 → 실적만 표시합니다.\n\n({type(e).__name__}: {e})")

    render_supply_analysis(build_sa_series(actual, plan_long, factor), short_unit)

    if info:
        with st.expander("🔍 계획 시트 인식 결과 (확인용)"):
            st.dataframe(pd.DataFrame(info["blocks"]), hide_index=True)
            if info["unmapped"]:
                st.warning("용도 매핑에 없는 항목(그대로 별도 용도로 표시됨): " + ", ".join(info["unmapped"]))


# ─────────────────────────────────────────────────────────
# 🟢 8. 메인 실행
# ─────────────────────────────────────────────────────────
def main():
    st.title(f"📈 {PLAN_YEAR}년 사업계획 at a glance")
    st.caption(f"{BASE_YEAR}년 계획 대비 실적 · {BASE_YEAR}년 실적 대비 {PLAN_YEAR}년 계획량 비교")

    with st.sidebar:
        st.header("⚙️ 메뉴 및 기본 설정")
        menu = st.radio("📋 보고서 탭 선택", ["1. One page review", "2. 세부내용", "3. 공급량 분석"])
        st.markdown("---")
        unit = st.radio("단위 선택", ["열량 (GJ)", "부피 (천m³)"], index=0)
        heating_value = 42.563
        if "부피" in unit:
            heating_value = st.number_input("(기준열량 MJ/Nm3 : 42.563 )", value=42.563, format="%.3f")

        st.markdown("---")
        st.subheader("🔗 실적 데이터 (구글시트)")
        gs_url = st.text_input("스프레드시트 주소", value=GSHEET_URL)
        if st.button("🔄 실적 새로고침"):
            fetch_gsheet_actual.clear()
        gs_status = st.empty()

        st.markdown("---")
        st.subheader("🔗 계획 데이터 (구글시트) · 3번 탭")
        plan_url = st.text_input("계획 스프레드시트 주소", value=PLAN_GSHEET_URL)
        if st.button("🔄 계획 새로고침"):
            fetch_gsheet_plan.clear()
        plan_status = st.empty()

        st.markdown("---")
        st.subheader("📂 계획 데이터 업로드 (1·2번 탭용)")
        up_supply = st.file_uploader("공급량 데이터 업로드 (새 파일이 있으면 우선 반영됩니다)", type=["xlsx", "csv"])
        st.caption("3. 공급량 분석 탭은 업로드 없이 구글시트만 사용합니다.")

    is_tab3 = menu.startswith("3.")

    # 계획 데이터 (엑셀) — 1·2번 탭 전용 (3번 탭은 없어도 동작)
    default_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), EXCEL_FILE)
    target_file = up_supply if up_supply is not None else (default_file if os.path.exists(default_file) else None)
    if target_file is None and not is_tab3:
        st.info(f"👈 좌측 사이드바에서 공급량 파일을 업로드하거나, 프로젝트 폴더에 {EXCEL_FILE} 파일을 배치해 주세요.")
        return
    data_dict = load_all_sheets(target_file) if target_file is not None else {}

    # 실적 데이터 (구글시트)
    gs_long = pd.DataFrame(columns=['연', '월', '그룹', '값'])
    try:
        gs_long = fetch_gsheet_actual(gs_url.strip())
        if gs_long.empty:
            gs_status.warning("구글시트에서 실적 값을 찾지 못했습니다. 엑셀 실적으로 대체합니다.")
        else:
            last = gs_long.loc[gs_long['연'].idxmax(), '연']
            last_m = gs_long[gs_long['연'] == last]['월'].max()
            gs_status.success(f"✅ 실적 반영: {gs_long['연'].min()}-01 ~ {last}-{last_m:02d}")
    except Exception as e:
        gs_status.warning(f"구글시트 연결 실패 → 엑셀 실적으로 대체합니다.\n\n({type(e).__name__})")

    short_unit = "GJ" if "GJ" in unit else "천m³"
    factor = 1 / 1000 if "GJ" in unit else 1 / heating_value / 1000   # 원자료 MJ 기준

    if is_tab3:
        render_supply_page(data_dict, gs_long, plan_url, plan_status, factor, short_unit)
        return

    data = assemble_data(data_dict, gs_long)

    act_months = data["act_months"]
    if act_months and max(act_months) < 12:
        st.caption(f"ℹ️ ② {BASE_YEAR} 실적(예상) = 1~{max(act_months)}월 실적 + "
                   f"{max(act_months) + 1}~12월 실천사업계획")

    if menu == "1. One page review":
        render_one_page_review(data, factor, short_unit)
    else:
        render_detail(data, factor, short_unit)


if __name__ == "__main__":
    main()
