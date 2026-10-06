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

# 2026년 실적 추정: 상품별공급량계획 시트의 '4. 2026년 추정실적' 섹션
#   양식: B열 상품명, C~N열 1~12월 (단위 GJ → ×1000 → MJ 변환)
ACTUAL_EST_GSHEET_URL = "https://docs.google.com/spreadsheets/d/1PSzKts5lL_zNNi_vfKlW1CdNZasTA-jBj_UCNEtu-x0/edit?gid=0#gid=0"

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
    return parse_gsheet_actual(raw)


@st.cache_data(ttl=600, show_spinner="2026년 추정실적 불러오는 중...")
def fetch_actual_est_from_plan_sheet(url):
    """상품별공급량계획 시트 '4. 2026년 추정실적' 섹션 파싱
    구조: B열=상품명, C~N열=1~12월 (단위 GJ → ×1000 → MJ)
    반환: DataFrame(columns=['연', '월', '그룹', '값'])  — 값은 MJ
    """
    raw = pd.read_csv(to_csv_url(url), header=None, dtype=str, keep_default_na=False)

    # '2026년 추정실적' 섹션 찾기
    start_idx = None
    for i in range(len(raw)):
        for j in range(min(5, raw.shape[1])):
            cell = str(raw.iat[i, j])
            if "2026" in cell and "추정실적" in cell:
                start_idx = i
                break
        if start_idx is not None:
            break

    if start_idx is None:
        return pd.DataFrame(columns=['연', '월', '그룹', '값'])

    # '구분' 헤더 행 찾기 (start_idx 이후 5행 이내)
    header_idx = start_idx
    for i in range(start_idx, min(start_idx + 5, len(raw))):
        if _norm(raw.iat[i, 1]) == "구분":
            header_idx = i
            break

    # 데이터 행: header_idx+1 부터 '합계' 행까지
    # B열(col1)=상품명, C~N열(col2~col13)=1~12월, 단위 GJ → ×1000 → MJ
    records = []
    for i in range(header_idx + 1, len(raw)):
        prod_name = _norm(raw.iat[i, 1])
        if not prod_name:
            continue
        if prod_name in ("합계", "합 계"):
            break
        if prod_name in ("소계", "소 계"):
            continue
        group = MAPPING_SUPPLY.get(prod_name, prod_name)
        for m in range(1, 13):
            col_idx = m + 1  # col2=C=1월, col3=D=2월, ..., col13=N=12월
            if col_idx < raw.shape[1]:
                val = pd.to_numeric(str(raw.iat[i, col_idx]).replace(",", "").strip(), errors="coerce")
                if pd.notna(val) and val != 0:
                    records.append((2026, m, group, float(val) * 1000))  # GJ → MJ

    return pd.DataFrame(records, columns=['연', '월', '그룹', '값'])


def parse_gsheet_actual(raw):
    raw = raw.iloc[:, :GSHEET_LAST_COL]

    # 1) 헤더 행 찾기: '[new ver]' 제목 아래의 '정산항목' 행
    hdr = None
    title_rows = [i for i, v in raw.iloc[:, 1].items() if GSHEET_BLOCK_LABEL in str(v)]
    if title_rows:
        start = title_rows[0]
        hdr = next((i for i in range(start, min(start + 5, len(raw))) if _norm(raw.iat[i, 2]) == "정산항목"), None)
    if hdr is None:   # 제목이 없는 간단한 양식(업로드용): '정산항목'/'상품' 헤더 + 'YYYY-MM' 월 헤더가 있는 첫 행
        month_re = re.compile(r"\d{4}\D+\d{1,2}")
        for i in range(len(raw)):
            n_month = sum(1 for j in range(3, raw.shape[1]) if month_re.search(str(raw.iat[i, j])))
            if n_month >= 1 and _norm(raw.iat[i, 2]) in ("정산항목", "상품", "상품명", "용도", "항목"):
                hdr = i
                break
        if hdr is None:
            hdr = next((i for i in range(len(raw))
                        if sum(1 for j in range(3, raw.shape[1]) if month_re.search(str(raw.iat[i, j]))) >= 3), None)
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


def draw_waterfall(df_chart, base, target, short_unit, title=None, names=None, height=440):
    names = names or SERIES_FULL
    base_tot, tgt_tot = df_chart[base].sum(), df_chart[target].sum()
    diffs = (df_chart[target] - df_chart[base]).tolist()
    labels = [names[base]] + [DISPLAY_NAME.get(g, g) for g in df_chart.index] + [names[target]]

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
        title=title or f"{names[base]} → {names[target]} 용도별 증감 브릿지",
        margin=dict(t=70, b=40), height=height, showlegend=False,
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
        text=f"<b>{badge}</b><br><span style='font-size:13px; color:gray'>{names[base]} 대비 {rate:.1f}%</span>",
        font=dict(size=17, color=badge_color), align="center",
    )
    fig.update_layout(margin=dict(t=70, b=40, r=60))
    unit_annotation(fig, short_unit)
    return style_fig(fig)


def draw_bullet_chart(df_data, base, target, short_unit, show_legend=False, names=None):
    names = names or SERIES_FULL
    fig = go.Figure()
    fig.add_trace(go.Bar(y=df_data.index, x=df_data['기준'], name=names[base], orientation='h',
                         marker_color=C_BASE_BAR, width=0.8, hoverinfo='x+name'))
    fig.add_trace(go.Bar(y=df_data.index, x=df_data['대상'], name=names[target], orientation='h',
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
# 🟢 5. 비교 요약표 (엑셀 양식과 동일 구조 · 1번/2번 탭 공용)
# ─────────────────────────────────────────────────────────
TABLE_CSS = """
<style>
.glance-wrap { overflow-x: auto; margin-bottom: 1rem; }
.glance { width: 100%; min-width: 760px; border-collapse: collapse; font-family: sans-serif; font-size: 14px; }
.glance th, .glance td { border: 1px solid #d0d4da; padding: 7px 9px; color: #31333F; }
.glance thead th { background: #E4E7EB; text-align: center; font-weight: 600; }
.glance tbody td { text-align: right; background: #ffffff; }
.glance td.label { text-align: center; background: #f8f9fa; font-weight: 600; }
.glance .blk { border-left: 3px solid #333333 !important; }
.glance .strong { font-weight: 700; }
.glance tr.subtotal td { background: #F3F5F7; }
.glance tr.total td { background: #E4E7EB; font-weight: 700; }
.glance tr.plain td { background: #ffffff; font-weight: 700; }
.glance tr.plain td.label { background: #f8f9fa; }
.glance tr.spacer td { border: none; background: transparent; height: 12px; padding: 0; }
.glance td.pos { color: #0055a4; }
.glance td.neg { color: #cc0000; }
.glance-unit { text-align: right; color: gray; font-size: 12px; }
</style>
"""

PLAN_VER_NAME = {"V1": "제출", "V2": "마케팅팀"}


def value_cells(vals, available, cfg):
    keys, comps = cfg["keys"], cfg["comps"]
    cells = []
    for i, key in enumerate(keys):
        cls = (["blk"] if i == 0 else []) + (["strong"] if key == keys[-1] else [])
        txt = f"{vals[key]:,.0f}" if available[key] else "-"
        cells.append(f'<td class="{" ".join(cls)}">{txt}</td>')

    for i, (tgt, base) in enumerate(comps):
        cls = ["blk"] if i == 0 else []
        if available[tgt] and available[base]:
            d = vals[tgt] - vals[base]
            cls += ["pos"] if round(d) > 0 else (["neg"] if round(d) < 0 else [])
            txt = f"{d:,.0f}"
        else:
            txt = "-"
        cells.append(f'<td class="{" ".join(cls)}">{txt}</td>')

    for i, (tgt, base) in enumerate(comps):
        cls = (["blk"] if i == 0 else []) + ["strong"]
        ok = available[tgt] and available[base] and vals[base] != 0
        txt = f"{vals[tgt] / vals[base] * 100:,.1f}%" if ok else "-"
        cells.append(f'<td class="{" ".join(cls)}">{txt}</td>')
    return "".join(cells)


def render_glance_table(df_detail, available, short_unit, cfg):
    keys, comps = cfg["keys"], cfg["comps"]
    bold = lambda s, on: f"<b>{s}</b>" if on else s
    head = (
        '<thead><tr>'
        '<th rowspan="2" colspan="2">구분</th>'
        + "".join(f'<th colspan="{len(ks)}" class="{"blk" if gi == 0 else ""}">{t}</th>'
                  for gi, (t, ks) in enumerate(cfg["groups"]))
        + f'<th colspan="{len(comps)}" class="blk">증감</th>'
        + f'<th colspan="{len(comps)}" class="blk">대비</th>'
        '</tr><tr>'
        + "".join(f'<th class="{"blk" if i == 0 else ""}">{bold(cfg["headers"][k], k == keys[-1])}</th>'
                  for i, k in enumerate(keys))
        + "".join(f'<th class="{"blk" if i == 0 else ""}">{t}-{b}</th>' for i, (t, b) in enumerate(comps))
        + "".join(f'<th class="{"blk" if i == 0 else ""}">{bold(f"{t}/{b}", True)}</th>'
                  for i, (t, b) in enumerate(comps))
        + '</tr></thead>'
    )
    vc = lambda vals: value_cells(vals, available, cfg)

    rows = []
    for g in [g for g in df_detail.index if g not in TRANSPORT_GROUPS]:
        rows.append(f'<tr><td class="label" colspan="2">{DISPLAY_NAME.get(g, g)}</td>{vc(df_detail.loc[g])}</tr>')
    rows.append(f'<tr><td class="label" rowspan="3">수송용</td><td class="label">CNG</td>{vc(df_detail.loc["CNG"])}</tr>')
    rows.append(f'<tr><td class="label">BIO</td>{vc(df_detail.loc["BIO"])}</tr>')
    rows.append(f'<tr class="subtotal"><td class="label">소계</td>{vc(df_detail.loc[TRANSPORT_GROUPS].sum())}</tr>')

    total = df_detail.sum()
    rows.append(f'<tr class="total"><td class="label" colspan="2">합계</td>{vc(total)}</tr>')
    rows.append(f'<tr class="spacer"><td colspan="{2 + len(keys) + 2 * len(comps)}"></td></tr>')
    home, ind = df_detail.loc["가정용"], df_detail.loc["산업용"]
    for name, vals in [("가정용", home), ("산업용", ind), ("기타", total - home - ind)]:
        rows.append(f'<tr class="plain"><td class="label" colspan="2">{name}</td>{vc(vals)}</tr>')

    st.markdown(
        TABLE_CSS
        + f'<div class="glance-unit">(단위 : {short_unit})</div>'
        + f'<div class="glance-wrap"><table class="glance">{head}<tbody>{"".join(rows)}</tbody></table></div>',
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────
# 🟢 6. 탭 1·2 — One Page Review (구글시트 실적·계획 사용)
#     page 1: 2026 계획 vs 2026 실적 / page 2: 2026 실적 vs 2027 계획
# ─────────────────────────────────────────────────────────
def op_config(page, ver):
    vn = PLAN_VER_NAME.get(ver, ver)
    act_name = f"{BASE_YEAR} 실적(예상)"
    if page == 1:
        return dict(
            keys=["①", "②"], comps=[("②", "①")],
            names={"①": f"{BASE_YEAR} 계획({vn})", "②": act_name},
            headers={"①": "① 계획", "②": "② 실적"},
            groups=[(f"{BASE_YEAR}년", ["①", "②"])],
            wf_title=f"① {BASE_YEAR}년 계획({vn}) 대비 실적",
            table_title=f"🚥 {BASE_YEAR}년 계획·실적 요약표",
            plan_year=BASE_YEAR,
        )
    return dict(
        keys=["①", "②"], comps=[("②", "①")],
        names={"①": act_name, "②": f"{PLAN_YEAR} 계획"},
        headers={"①": "① 실적", "②": "② 계획"},
        groups=[(f"{BASE_YEAR}년", ["①"]), (f"{PLAN_YEAR}년", ["②"])],
        wf_title=f"① {BASE_YEAR}년 실적 대비 {PLAN_YEAR}년 계획",
        table_title=f"🚥 {BASE_YEAR}년 실적 vs {PLAN_YEAR}년 사업계획 요약표",
        plan_year=PLAN_YEAR,
    )


def op_series(page, actual, plan_long, ver):
    empty = pd.DataFrame(columns=['월', '그룹', '값'])
    cfg = op_config(page, ver)

    def plan_of(year):
        if plan_long is None or plan_long.empty:
            return empty
        d = plan_long[(plan_long['연'] == year) & (plan_long['버전'] == ver)]
        return d[['월', '그룹', '값']]

    act = actual[actual['연'] == BASE_YEAR][['월', '그룹', '값']] if actual is not None and not actual.empty else empty
    return {"①": plan_of(BASE_YEAR), "②": act} if page == 1 else {"①": act, "②": plan_of(PLAN_YEAR)}


def render_one_page_review(page, actual, plan_long, factor, short_unit):
    pk = f"op{page}"
    py = BASE_YEAR if page == 1 else PLAN_YEAR
    vers = set(plan_long[plan_long['연'] == py]['버전']) if plan_long is not None and not plan_long.empty else set()
    notice = st.container()
    h_col, t_col = st.columns([3, 1])
    h_col.markdown("#### 📌 핵심 지표")
    if page == 1:   # 2026년만 제출(V1)/마케팅팀(V2) 계획이 있음
        base_ver = "V1" if "V1" in vers or not vers else sorted(vers)[0]
        use_mkt = t_col.toggle("마케팅팀 계획", value=False, key=f"{pk}_mkt", disabled="V2" not in vers)
        ver = "V2" if (use_mkt and "V2" in vers) else base_ver
    else:           # 2027년 계획은 1개 버전
        ver = "V1" if "V1" in vers or not vers else sorted(vers)[0]
    cfg = op_config(page, ver)
    keys, names = cfg["keys"], cfg["names"]

    series = {k: scale(v, factor) for k, v in op_series(page, actual, plan_long, ver).items()}
    available = {k: not v.empty and v['값'].sum() != 0 for k, v in series.items()}

    with notice:
        if not any(available.values()):
            st.warning("데이터가 부족합니다. 사이드바의 실적/계획 구글시트 주소와 공유 설정을 확인해주세요.")
        else:
            missing = [names[k] for k, ok in available.items() if not ok]
            if missing:
                st.info(f"💡 아직 데이터가 없는 항목: **{', '.join(missing)}** — 표에는 '-'로 표시되고 관련 비교는 생략됩니다.")
        act_df = actual[actual['연'] == BASE_YEAR] if actual is not None and not actual.empty else None
        if act_df is not None and not act_df.empty:
            av = sa_avail_months(act_df)
            if av and max(av) < 12:
                st.caption(f"ℹ️ {BASE_YEAR}년 실적은 1월–{max(av)}월 데이터만 반영되어 있습니다 (시트에 입력된 만큼).")
    if not any(available.values()):
        return

    df_detail = pd.DataFrame({k: v.groupby('그룹')['값'].sum() for k, v in series.items()}).reindex(columns=keys).fillna(0)
    extras = sorted(g for g in df_detail.index if g not in DETAIL_ORDER)
    df_detail = df_detail.reindex(DETAIL_ORDER + extras).fillna(0)

    df_chart = df_detail.copy()
    df_chart.index = [to_chart_group(g) for g in df_chart.index]
    df_chart = df_chart.groupby(level=0).sum()
    df_chart = df_chart.reindex([g for g in CHART_ORDER + sorted(set(df_chart.index) - set(CHART_ORDER))
                                 if g in df_chart.index])

    totals = df_detail.sum()
    valid_comps = [(t, b) for t, b in cfg["comps"] if available[t] and available[b]]

    # ── 1. 핵심 지표 ──
    def metric(col, key, base=None):
        label = f"{key} {names[key]}"
        if not available[key]:
            col.metric(label, "-")
            return
        delta = None
        if base and available[base] and totals[base] != 0:
            d = totals[key] - totals[base]
            delta = f"{d:,.0f} ({totals[key] / totals[base] * 100:.1f}%)"
        col.metric(label, f"{totals[key]:,.0f} {short_unit}", delta=delta)

    st.markdown("""<style>
    div[data-testid="stMetricDelta"] { padding: 4px 12px; border-radius: 16px; }
    div[data-testid="stMetricDelta"] * { font-size: 1.25rem !important; font-weight: 700 !important; }
    div[data-testid="stMetricDelta"] svg { width: 1.3rem; height: 1.3rem; }
    </style>""", unsafe_allow_html=True)
    mcols = st.columns(4)
    metric(mcols[0], "①")
    metric(mcols[1], "②", "①")
    st.markdown("---")

    # ── 2. 폭포수 차트 ──
    st.markdown("#### 🌊 용도별 증감 요인 폭포수 차트")
    wf_h, wf_t = st.columns([3, 1])
    wf_h.markdown(f"##### {cfg['wf_title']}")
    simple = wf_t.toggle("심플버전", value=False, key=f"{pk}_simple", disabled=not valid_comps)
    if valid_comps:
        if simple:   # 가정용 · 산업용 · 연료전지 · 기타로 묶어서 표시
            home, ind, fc = df_detail.loc["가정용"], df_detail.loc["산업용"], df_detail.loc["연료전지"]
            df3 = pd.DataFrame([home, ind, fc, df_detail.sum() - home - ind - fc],
                               index=["가정용", "산업용", "연료전지", "기타"])
            st.plotly_chart(draw_waterfall(df3, "①", "②", short_unit, names=names, height=380), width="stretch")
        else:
            st.plotly_chart(draw_waterfall(df_chart, "①", "②", short_unit, names=names), width="stretch")
    elif page == 2:
        st.info(f"💡 {PLAN_YEAR}년 계획 데이터(계획 시트 하단에 '{PLAN_YEAR}년' 블록)가 들어오면 표시됩니다.")
    else:
        st.info(f"{BASE_YEAR}년 계획 또는 실적 데이터가 없습니다.")
    st.markdown("---")

    # ── 3. 비교 요약표 ──
    st.markdown(f"#### {cfg['table_title']}")
    render_glance_table(df_detail, available, short_unit, cfg)
    st.markdown("---")

    # ── 4. 불릿 차트 ──
    if not valid_comps:
        return
    st.markdown("#### 🎯 용도별 세부 비교")
    t, b = valid_comps[0]

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
    st.plotly_chart(draw_bullet_chart(part1, b, t, short_unit, show_legend=True, names=names), width="stretch")
    if not part2.empty:
        st.markdown("##### 📌 [세부용도 (기타 용도 나타냄)]")
        st.plotly_chart(draw_bullet_chart(part2, b, t, short_unit, names=names), width="stretch")


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
KIND_COLOR = {"실적": C_NAVY, "V1": "#4A90D9", "V2": "#8FB8E8"}
KIND_DASH = {"실적": "solid", "V1": "dash", "V2": "dot"}
KIND_SYMBOL = {"실적": "circle", "V1": "square", "V2": "diamond"}
SA_ROWS = ["가정용", "영업/업무용", "산업용", "열병합용", "연료전지", "자가열전용", "열전용설비용(주택외)"]
SA_AGG = {"가정용", "영업/업무용"}   # 시트상 '소계' 성격의 행 (하위 품목 합산) → 소계 배경색

SA_PLOT_CONFIG = {"scrollZoom": True, "displaylogo": False, "doubleClick": "reset"}

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
    per_year = {}
    for bl in blocks:
        per_year.setdefault(bl["연도"], []).append(bl)
    for y, bls in per_year.items():   # 블록이 1개뿐인 연도(예: 2027)는 버전 구분 없이 V1
        if len(bls) == 1:
            df.loc[df['연'] == y, '버전'] = "V1"
            bls[0]["버전"] = "V1"
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
    # 푸른색 계열: 오래된 연도 = 연한 파랑 → 최근 연도 = 진한 남색
    def _blue(t):
        lo, hi = (150, 190, 230), (11, 42, 85)
        return "#%02X%02X%02X" % tuple(round(l + (h - l) * t) for l, h in zip(lo, hi))
    color_by_year = {y: _blue(i / max(len(years) - 1, 1)) for i, y in enumerate(years)}

    # ── 1. 상품별 연도별 꺾은선 ──
    st.markdown("### 1️⃣ 상품별 연도별 추이 (꺾은선)")
    product = st.radio("📂 상품 선택", prod_opts, horizontal=True, key="sa_prod")
    st.markdown("📊 구분 (체크한 항목을 표시)")
    chk1 = st.columns(max(len(kinds), 1))
    kinds_sel = [k for i, k in enumerate(kinds)
                 if chk1[i].checkbox(kind_label(k), value=True, key=f"sa_line_chk_{k}")]
    base = latest_act if latest_act is not None else years[-1]
    years2 = st.multiselect("📅 연도", years, default=[y for y in years if y >= base - 3], key="sa_years2")

    sy = sorted(years2)

    def _shade(lo, hi, t):
        return "#%02X%02X%02X" % tuple(round(l + (h - l) * t) for l, h in zip(lo, hi))

    def line_color(kind, y):
        # 기준연도(BASE_YEAR) 실적 = 붉은색(메인), 과거 실적 = 회색 계열, V1 = 파랑, V2 = 초록빛 청록
        t = sy.index(y) / max(len(sy) - 1, 1)
        if kind == "실적":
            if y == BASE_YEAR:
                return "#D32F2F"
            return _shade((185, 192, 202), (85, 95, 110), t)   # 오래된 연도 = 연한 회색 → 최근 = 진한 회색
        # V1 = 파랑, V2 = 초록빛 청록 (색상 차이를 크게) — 단일 연도여도 너무 연하지 않도록 0.35~1.0 구간 사용
        ramp = {"V1": ((80, 135, 225), (15, 55, 150)),
                "V2": ((60, 200, 170), (5, 120, 95))}.get(kind, ((150, 190, 230), (11, 42, 85)))
        return _shade(ramp[0], ramp[1], 0.35 + 0.65 * t)

    fig2, tbl_rows = go.Figure(), []
    for y in years2:
        for k in kinds_sel:
            if (k, y) not in series:
                continue
            v = sa_month_vec(series[(k, y)], None if product == "전체" else product)
            vv = v   # 월별 공급량
            name = f"{y} {kind_label(k)}"
            fig2.add_trace(go.Scatter(
                x=list(range(1, 13)), y=vv.values, name=name, mode="lines+markers",
                line=dict(color=line_color(k, y), dash=KIND_DASH.get(k, "dash"),
                          width=4 if (k == "실적" and y == BASE_YEAR) else 2.8),
                marker=dict(size=7 if k == "실적" else 6, symbol=KIND_SYMBOL.get(k, "triangle-up")),
                connectgaps=False, hovertemplate=name + " %{x}월<br>%{y:,.0f}<extra></extra>"))
            tot = vv.sum()
            half = lambda h: "-" if h.isna().any() else f"{h.sum():,.0f}"
            tbl_rows.append(dict(label=name, cls="", cells=[("-" if pd.isna(x) else f"{x:,.0f}", "") for x in vv]
                                 + [(half(vv.iloc[:6]), "tcol"), (half(vv.iloc[6:]), "tcol"), (f"{tot:,.0f}", "tcol")]))
    if not tbl_rows:
        st.info("선택한 연도/구분에 해당하는 데이터가 없습니다.")
    else:
        fig2.update_xaxes(tickvals=list(range(1, 13)), ticktext=[f"{i}월" for i in range(1, 13)])
        fig2.update_yaxes(tickformat=",.0f")
        fig2.update_layout(title=f"{product} — 연도별 월별 공급량",
                           xaxis_title="", yaxis_title="", legend_title="", height=520, hovermode="closest")
        unit_annotation(fig2, short_unit)
        fig2.update_layout(dragmode="pan")   # 드래그 = 이동, 마우스 휠 = 확대/축소 (더블클릭 = 원래대로)
        st.plotly_chart(style_fig(fig2), width="stretch", config=SA_PLOT_CONFIG)

        st.markdown(f"##### 📋 {product} — 연도별 월별 수치")
        head = ("<tr><th>구분</th>" + "".join(f"<th>{m}월</th>" for m in range(1, 13))
                + "<th>상반기</th><th>하반기</th><th>합계</th></tr>")
        sa_render_table(head, tbl_rows, short_unit)
        st.caption("※ 상반기 = 1월–6월, 하반기 = 7월–12월, 합계 = 월별 값의 합. 해당 기간 데이터가 다 없으면 '-' (진행 중인 연도는 합계만 실적이 있는 달까지)")
    st.markdown("---")

    # ── 2. 선택 연도·구분의 용도 × 월 상세 ──
    st.markdown("### 2️⃣ 용도별 월별 상세 (연도·구분 선택)")
    c1, c2 = st.columns([1, 3])
    with c1:
        y3 = st.selectbox("📅 연도", years, index=years.index(latest_act) if latest_act in years else len(years) - 1,
                          key="sa_y3")
    ks3 = [k for k in kinds if (k, y3) in series]
    with c2:
        st.markdown("📊 구분 (체크한 항목을 월별로 나란히 비교)")
        chk = st.columns(max(len(ks3), 1))
        sel3 = [k for i, k in enumerate(ks3)
                if chk[i].checkbox(kind_label(k), value=(k == "실적"), key=f"sa_chk_{k}")]
    if not sel3:
        st.info("비교할 구분(실적, 계획 V1, 계획 V2)을 하나 이상 체크해주세요.")
    else:
        SHORT = {"실적": "실적", "V1": "V1", "V2": "V2"}
        # 구분별 색상 계열 (진함=가정용 / 중간=산업용 / 연함=기타): 실적=기존 파랑, V1=회색, V2=청록빛 파랑
        STACK_COLORS = {"실적": [C_NAVY, C_UP, C_DOWN],               # 기존 푸른색 (진함/중간/연함)
                        "V1": ["#4B5563", "#8A94A3", "#C3CAD4"],       # 회색 계열
                        "V2": ["#0A6987", "#32A0B9", "#8CCDDC"]}       # 청록빛 파랑 계열
        simple3 = lambda g: g if g in ("가정용", "산업용") else "기타"
        mlabels = [f"{m}월" for m in range(1, 13) for _ in sel3]
        klabels = [SHORT.get(k, k) for _m in range(1, 13) for k in sel3]

        fig3 = go.Figure()
        for k in sel3:
            df3 = series[(k, y3)]
            av_k = set(sa_avail_months(df3))
            pal = STACK_COLORS.get(k, STACK_COLORS["V1"])
            for gi, g in enumerate(["가정용", "산업용", "기타"]):
                ys = []
                for m in range(1, 13):
                    for kk in sel3:
                        if kk != k or m not in av_k:
                            ys.append(None)
                            continue
                        d = df3[df3['월'] == m]
                        ys.append(float(d[d['그룹'].map(simple3) == g]['값'].sum()))
                nm = f"{SHORT.get(k, k)} · {g}"
                fig3.add_trace(go.Bar(name=nm, legendgroup=k, x=[mlabels, klabels], y=ys, marker_color=pal[gi],
                                      hovertemplate=nm + " %{x}<br>%{y:,.0f}<extra></extra>"))
        names = " · ".join(kind_label(k) for k in sel3)
        fig3.update_layout(barmode="stack", title=f"{y3}년 월별 공급량 — {names} (가정용·산업용·기타)",
                           xaxis_title="", yaxis_title="", legend_title="", height=500, bargap=0.15)
        fig3.update_yaxes(tickformat=",.0f")
        unit_annotation(fig3, short_unit)
        fig3.update_layout(dragmode="pan")

        def ach(pk, g=None):
            """달성률(%) = 실적 ÷ 계획. 실적이 있는 달만 같은 기간으로 비교 (g: 가정용/산업용/기타, None=전체)"""
            a_df, p_df = series.get(("실적", y3)), series.get((pk, y3))
            if a_df is None or p_df is None:
                return None
            av = set(sa_avail_months(a_df))
            if not av:
                return None

            def tot(df):
                d = df[df['월'].isin(av)]
                if g:
                    d = d[d['그룹'].map(simple3) == g]
                return float(d['값'].sum())
            p = tot(p_df)
            return tot(a_df) / p * 100 if p else None

        # 좌측: 전체 누계 (선택한 구분별 합계, 가정용·산업용·기타 스택)
        fig_t = go.Figure()
        totals, partial = {}, []
        for k in sel3:
            df_k = series[(k, y3)]
            av_k = sa_avail_months(df_k)
            if len(av_k) < 12:
                partial.append(f"{SHORT.get(k, k)} {max(av_k)}월까지" if av_k else SHORT.get(k, k))
            pal = STACK_COLORS.get(k, STACK_COLORS["V1"])
            sums = df_k.groupby(df_k['그룹'].map(simple3))['값'].sum()
            totals[k] = float(sums.sum())
            for gi, g in enumerate(["가정용", "산업용", "기타"]):
                fig_t.add_trace(go.Bar(name=f"{SHORT.get(k, k)} · {g}", x=[SHORT.get(k, k)], y=[float(sums.get(g, 0.0))],
                                       marker_color=pal[gi], showlegend=False,
                                       hovertemplate=f"{SHORT.get(k, k)} · {g}<br>%{{y:,.0f}}<extra></extra>"))
        # 막대 위: 합계 숫자(15px) + 그 위에 달성률(실적/계획 = %, 큰 글씨)
        for k in sel3:
            fig_t.add_annotation(x=SHORT.get(k, k), y=totals[k], text=f"<b>{totals[k]:,.0f}</b>", showarrow=False,
                                 yanchor="bottom", yshift=4, font=dict(size=15, color="#31333F"))
            if k != "실적" and "실적" in sel3:
                r = ach(k)
                if r is not None:
                    fig_t.add_annotation(
                        x=SHORT.get(k, k), y=totals[k], showarrow=False, yanchor="bottom", yshift=28,
                        text=f"<span style='font-size:12px'>실적/{SHORT.get(k, k)} =</span><br><b>{r:,.1f}%</b>",
                        font=dict(size=20, color="#C0392B"), align="center")
        fig_t.update_layout(barmode="stack", title=f"{y3}년 전체 누계", xaxis_title="", yaxis_title="",
                            height=500, bargap=0.25, dragmode="pan", margin=dict(t=70))
        fig_t.update_yaxes(tickformat=",.0f", range=[0, max(totals.values()) * 1.32 if totals else 1])
        fig_t.update_xaxes(type="category")
        unit_annotation(fig_t, short_unit)

        col_tot, col_mon = st.columns([4, 7])
        with col_tot:
            st.plotly_chart(style_fig(fig_t), width="stretch", config=SA_PLOT_CONFIG)
        with col_mon:
            st.plotly_chart(style_fig(fig3), width="stretch", config=SA_PLOT_CONFIG)
        if partial:
            st.caption("※ 전체 누계는 데이터가 있는 달까지의 합계입니다: " + ", ".join(partial))

        # 달성률 (실적 ÷ 계획 V1 / V2): 월별 + 합계(실적이 있는 달 기준)
        plan_ks = [k for k in ks3 if k != "실적"]
        if "실적" in ks3 and plan_ks:
            st.markdown(f"##### 🎯 {y3}년 달성률 (실적 ÷ 계획)")
            a_df = series[("실적", y3)]
            av_a = set(sa_avail_months(a_df))
            rows_a = []
            for pk in plan_ks:
                p_df = series[(pk, y3)]
                for gname in ["전체", "가정용", "산업용", "기타"]:
                    gg = None if gname == "전체" else gname
                    cs = []
                    for m in range(1, 13):
                        da, dp = a_df[a_df['월'] == m], p_df[p_df['월'] == m]
                        if gg:
                            da, dp = da[da['그룹'].map(simple3) == gg], dp[dp['그룹'].map(simple3) == gg]
                        pv = float(dp['값'].sum())
                        cs.append((f"{float(da['값'].sum()) / pv * 100:,.1f}%" if (m in av_a and pv) else "-", ""))
                    r = ach(pk, gg)
                    cs.append((f"{r:,.1f}%" if r is not None else "-", "tcol"))
                    rows_a.append(dict(label=f"실적/{SHORT.get(pk, pk)} · {gname}", cls="total" if gname == "전체" else "",
                                       cells=cs))
            head_a = ("<tr><th>구분</th>" + "".join(f"<th>{m}월</th>" for m in range(1, 13))
                      + "<th>합계</th></tr>")
            sa_render_table(head_a, rows_a, "%")
            if len(av_a) < 12:
                st.caption(f"※ 합계 달성률은 실적이 있는 달({min(av_a)}월~{max(av_a)}월) 기준으로 계획과 같은 기간끼리 비교한 값입니다.")

        for k in sel3:
            df3 = series[(k, y3)]
            av3 = set(sa_avail_months(df3))
            st.markdown(f"##### 📋 {y3}년 {kind_label(k)} — 용도 × 월")
            cols3 = [df3[df3['월'] == m].groupby('그룹')['값'].sum() if m in av3 else None for m in range(1, 13)]
            cols3.append(df3.groupby('그룹')['값'].sum())
            head = "<tr><th>구분</th>" + "".join(f"<th>{m}월</th>" for m in range(1, 13)) + "<th>합계</th></tr>"
            sa_render_table(head, sa_build_rows(cols3, tcols={12}), short_unit)


def load_plan(plan_url, plan_status):
    """계획 구글시트 → (long DataFrame, 인식 결과 info)  (1·2·4번 탭 공용)"""
    empty = pd.DataFrame(columns=['연', '월', '그룹', '값', '버전'])
    try:
        plan_long, info = fetch_gsheet_plan(plan_url.strip())
        if plan_long.empty:
            plan_status.warning("계획 시트에서 값을 찾지 못했습니다.")
        else:
            desc = ", ".join(f"{b['연도']} {b['버전']}" for b in info["blocks"])
            plan_status.success(f"✅ 계획 반영: {desc}")
        return plan_long, info
    except Exception as e:
        plan_status.warning(f"계획 구글시트 연결 실패\n\n({type(e).__name__}: {e})")
        return empty, None


def render_supply_page(actual, plan_long, info, factor, short_unit):
    """4번 탭: 실적(구글시트) + 계획(구글시트)만으로 구성"""
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
    sub_ph = st.empty()   # 탭별 소제목 (메뉴 선택 후 채움)

    TAB1 = f"1. One page review({BASE_YEAR}년 실적)"
    TAB2 = f"2. One page review({PLAN_YEAR}년 계획)"
    TAB3 = "3. 세부내용"
    TAB4 = "4. 공급량 분석"

    with st.sidebar:
        st.header("⚙️ 메뉴 및 기본 설정")
        menu = st.radio("📋 보고서 탭 선택", [TAB1, TAB2, TAB3, TAB4])
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
        st.subheader("🔗 계획 데이터 (구글시트) · 1·2·4번 탭")
        plan_url = st.text_input("계획 스프레드시트 주소", value=PLAN_GSHEET_URL)
        if st.button("🔄 계획 새로고침"):
            fetch_gsheet_plan.clear()
        plan_status = st.empty()

        st.markdown("---")
        st.subheader("🔗 2026년 실적 추정 (구글시트)")
        est_url = st.text_input("실적 추정 스프레드시트 주소", value=ACTUAL_EST_GSHEET_URL)
        st.caption("이 시트에 있는 연·월은 위 '실적' 시트 대신 사용되고, 나머지 기간은 '실적' 시트 값을 씁니다. "
                   "주소를 비우면 사용하지 않습니다.")
        if st.button("🔄 실적 추정 새로고침"):
            fetch_gsheet_actual.clear()
        est_status = st.empty()

        st.markdown("---")
        st.subheader("📂 계획 데이터 업로드 (3. 세부내용 탭용)")
        up_supply = st.file_uploader("공급량 데이터 업로드 (새 파일이 있으면 우선 반영됩니다)", type=["xlsx", "csv"])
        st.caption("1·2·4번 탭은 업로드 없이 구글시트만 사용합니다.")

    SUBTITLES = {
        TAB1: (f"{BASE_YEAR}년 계획 대비 실적", f"{BASE_YEAR}년 계획(제출·마케팅팀)과 {BASE_YEAR}년 실적(예상)을 비교합니다."),
        TAB2: (f"{BASE_YEAR}년 실적 대비 {PLAN_YEAR}년 계획", f"{BASE_YEAR}년 실적(예상)과 {PLAN_YEAR}년 사업계획을 비교합니다."),
        TAB3: ("연도별 용도별 세부내용", "실적·계획·예상실적·당초·실천 값을 연도별/용도별로 확인합니다."),
        TAB4: ("공급량 분석", "과거 실적과 계획(V1·V2)을 연도·월별로 비교합니다."),
    }
    sub_t, sub_c = SUBTITLES[menu]
    with sub_ph.container():
        st.subheader(sub_t)
        st.caption(sub_c)

    need_excel = menu == TAB3

    # 엑셀 (3. 세부내용 전용)
    default_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), EXCEL_FILE)
    target_file = up_supply if up_supply is not None else (default_file if os.path.exists(default_file) else None)
    if target_file is None and need_excel:
        st.info(f"👈 좌측 사이드바에서 공급량 파일을 업로드하거나, 프로젝트 폴더에 {EXCEL_FILE} 파일을 배치해 주세요.")
        return
    data_dict = load_all_sheets(target_file) if target_file is not None else {}

    # 실적 데이터 (구글시트)
    gs_long = pd.DataFrame(columns=['연', '월', '그룹', '값'])
    try:
        gs_long = fetch_gsheet_actual(gs_url.strip())
        if gs_long.empty:
            gs_status.warning("구글시트에서 실적 값을 찾지 못했습니다.")
        else:
            last = gs_long.loc[gs_long['연'].idxmax(), '연']
            last_m = gs_long[gs_long['연'] == last]['월'].max()
            gs_status.success(f"✅ 실적 반영: {gs_long['연'].min()}-01 ~ {last}-{last_m:02d}")
    except Exception as e:
        gs_status.warning(f"구글시트 연결 실패\n\n({type(e).__name__})")

    # 2026년 실적 추정 시트: 있는 연·월은 이 값으로 덮어씀 (모든 탭 공통)
    if est_url.strip():
        try:
            est_long = fetch_actual_est_from_plan_sheet(est_url.strip())
            if est_long.empty:
                est_status.warning("실적 추정 시트에서 '4. 2026년 추정실적' 섹션을 찾지 못했습니다. '실적' 시트 값만 사용합니다.")
            else:
                keys = set(zip(est_long['연'], est_long['월']))
                keep = [(y, m) not in keys for y, m in zip(gs_long['연'], gs_long['월'])]
                gs_long = pd.concat([gs_long[keep], est_long], ignore_index=True)
                f_y, l_y = est_long['연'].min(), est_long['연'].max()
                f_m = est_long.loc[est_long['연'] == f_y, '월'].min()
                l_m = est_long.loc[est_long['연'] == l_y, '월'].max()
                est_status.success(f"✅ 실적 추정 반영: {f_y}-{f_m:02d} ~ {l_y}-{l_m:02d}")
        except Exception as e:
            est_status.warning(f"실적 추정 구글시트 연결 실패 → '실적' 시트 값만 사용합니다.\n\n({type(e).__name__})")

    short_unit = "GJ" if "GJ" in unit else "천m³"
    factor = 1 / 1000 if "GJ" in unit else 1 / heating_value / 1000   # 원자료 MJ 기준

    if menu == TAB3:
        data = assemble_data(data_dict, gs_long)
        act_months = data["act_months"]
        if act_months and max(act_months) < 12:
            st.caption(f"ℹ️ ② {BASE_YEAR} 실적(예상) = 1~{max(act_months)}월 실적 + "
                       f"{max(act_months) + 1}~12월 실천사업계획")
        render_detail(data, factor, short_unit)
        return

    # 1·2·4번 탭: 실적·계획 모두 구글시트
    actual = assemble_data(data_dict, gs_long)["actual"] if data_dict else gs_long
    if actual is None or actual.empty:
        st.error("실적 데이터를 불러오지 못했습니다. 사이드바의 구글시트 주소/공유 설정을 확인해주세요.")
        return
    plan_long, info = load_plan(plan_url, plan_status)

    if menu == TAB1:
        render_one_page_review(1, actual, plan_long, factor, short_unit)
    elif menu == TAB2:
        render_one_page_review(2, actual, plan_long, factor, short_unit)
    else:
        render_supply_page(actual, plan_long, info, factor, short_unit)


if __name__ == "__main__":
    main()
