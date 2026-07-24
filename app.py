"""
STCO 온라인팀 광고/마케팅 성과 대시보드
==================================
매주 "STCO_주간보고서_...xlsx" 파일을 업로드하면 아래 시트들을 자동으로 인식해서
누적 저장하고, ROAS/KPI를 웹에서 바로 볼 수 있는 대시보드.

자동으로 읽는 시트:
  - "매체통합" 시트의 1) 월별 통합데이터 / 2) 매체별 현황(당월 GA비교) / 통합 주간별
  - "(SA)/(DA)/(SSP)/(브검) ○○" 형태의 매체별 요약 시트 (약 17개)
  - "GA-RAW" 시트 (소스/매체별 유입 스냅샷)

실행:
    streamlit run app.py
배포:
    README.md 참고 (GitHub + Supabase + Streamlit Community Cloud)
"""

import io
import re
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="STCO 광고성과 대시보드", page_icon="📊", layout="wide")

# ── 디자인 톤 (색상/버튼/표는 Toss(TDS Mobile) 스타일, 폰트는 당근마켓 SEED 시스템폰트 스택) ──
THEME_FONT_STACK = (
    "-apple-system, BlinkMacSystemFont, \"Pretendard Variable\", Pretendard, "
    "\"Apple SD Gothic Neo\", \"Malgun Gothic\", system-ui, sans-serif"
)
THEME_COLORS = {
    "primary": "#3182f6",
    "primary_hover": "#2272eb",
    "canvas": "#ffffff",
    "surface": "#f2f4f6",
    "foreground": "#191f28",
    "body": "#4e5968",
    "muted": "#8b95a1",
    "border": "#e5e8eb",
    "on_primary": "#ffffff",
    "weak_bg": "#e8f3ff",
    "weak_fg": "#1b64da",
    "danger": "#e42939",
}
px.defaults.color_discrete_sequence = ["#3182f6", "#191f28", "#8b95a1", "#1b64da"]


def theme_chart(fig):
    """Plotly 차트에 테마 톤(시스템 폰트, 화이트 배경, 옅은 그리드)을 적용."""
    fig.update_layout(
        font_family=THEME_FONT_STACK,
        font_color=THEME_COLORS["foreground"],
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        title_font_size=17,
        title_font_color=THEME_COLORS["foreground"],
        legend_title_font_color=THEME_COLORS["muted"],
        margin=dict(l=10, r=10, t=50, b=10),
    )
    fig.update_xaxes(gridcolor=THEME_COLORS["surface"], zerolinecolor=THEME_COLORS["border"], linecolor=THEME_COLORS["border"])
    fig.update_yaxes(gridcolor=THEME_COLORS["surface"], zerolinecolor=THEME_COLORS["border"], linecolor=THEME_COLORS["border"])
    return fig


def inject_theme():
    st.markdown(
        f"""
        <style>
        html, body, [class*="css"], .stApp, .stMarkdown, .stText {{
            font-family: {THEME_FONT_STACK} !important;
        }}
        .stApp, [data-testid="stAppViewContainer"] {{
            background-color: {THEME_COLORS["canvas"]};
        }}
        [data-testid="stHeader"] {{ background-color: transparent; }}
        .block-container {{ padding-top: 2rem; padding-left: 3rem; padding-right: 3rem; max-width: 100%; }}

        [data-testid="stSidebar"] {{
            background-color: {THEME_COLORS["surface"]};
            border-right: 1px solid {THEME_COLORS["border"]};
        }}

        h1, h2, h3, h4 {{
            color: {THEME_COLORS["foreground"]} !important;
            font-weight: 700 !important;
            letter-spacing: -0.01em;
        }}
        h1 {{ font-size: 32px !important; font-weight: 700 !important; }}
        h2 {{ font-size: 24px !important; font-weight: 600 !important; }}
        h3 {{ font-size: 19px !important; font-weight: 600 !important; }}
        h4 {{ font-size: 17px !important; font-weight: 600 !important; }}
        p, span, label, div {{ color: {THEME_COLORS["body"]}; }}
        [data-testid="stCaptionContainer"], .stCaption, small {{
            color: {THEME_COLORS["muted"]} !important;
        }}

        [data-testid="stMetric"] {{
            background: {THEME_COLORS["canvas"]};
            border: 1px solid {THEME_COLORS["border"]};
            border-radius: 14px;
            padding: 16px 20px;
        }}
        [data-testid="stMetricLabel"] {{ color: {THEME_COLORS["muted"]} !important; font-weight: 400 !important; }}
        [data-testid="stMetricValue"] {{ color: {THEME_COLORS["foreground"]} !important; font-weight: 700 !important; }}

        .stButton > button, .stDownloadButton > button, .stLinkButton > a {{
            background-color: {THEME_COLORS["primary"]} !important;
            color: {THEME_COLORS["on_primary"]} !important;
            border: none !important;
            border-radius: 10px !important;
            padding: 8px 20px !important;
            font-weight: 600 !important;
            font-size: 15px !important;
            box-shadow: none !important;
            transition: background-color .15s ease;
        }}
        .stButton > button:hover, .stDownloadButton > button:hover, .stLinkButton > a:hover {{
            background-color: {THEME_COLORS["primary_hover"]} !important;
            color: {THEME_COLORS["on_primary"]} !important;
        }}
        .stButton > button[kind="secondary"] {{
            background-color: {THEME_COLORS["weak_bg"]} !important;
            color: {THEME_COLORS["weak_fg"]} !important;
            border: none !important;
        }}

        [data-testid="stTabs"] button {{ color: {THEME_COLORS["muted"]}; font-weight: 600; }}
        [data-testid="stTabs"] button[aria-selected="true"] {{
            color: {THEME_COLORS["primary"]} !important;
            border-bottom-color: {THEME_COLORS["primary"]} !important;
        }}

        [data-testid="stPopover"] > div > button {{
            background-color: {THEME_COLORS["canvas"]} !important;
            color: {THEME_COLORS["foreground"]} !important;
            border: 1px solid {THEME_COLORS["border"]} !important;
            border-radius: 8px !important;
            font-weight: 500 !important;
            padding: 4px 10px !important;
            font-size: 12.5px !important;
            white-space: nowrap !important;
        }}
        [data-testid="stPopover"] > div > button:hover {{
            background-color: {THEME_COLORS["surface"]} !important;
            color: {THEME_COLORS["foreground"]} !important;
            border-color: {THEME_COLORS["primary"]} !important;
        }}
        [data-testid="stPopoverBody"] {{
            border-radius: 12px;
            border: 1px solid {THEME_COLORS["border"]};
        }}
        div[data-baseweb="popover"] {{
            z-index: 999999 !important;
        }}

        [data-baseweb="select"] > div {{
            border-radius: 8px !important;
            border-color: {THEME_COLORS["border"]} !important;
        }}
        .stTextInput > div > div, .stDateInput > div > div {{ border-radius: 8px !important; }}

        [data-testid="stFileUploader"] section {{
            border-radius: 10px;
            border: 1px dashed {THEME_COLORS["border"]};
            background: {THEME_COLORS["surface"]};
        }}

        [data-testid="stDataFrame"] {{
            border-radius: 10px;
            overflow: hidden;
            border: 1px solid {THEME_COLORS["border"]};
        }}

        hr {{ border-color: {THEME_COLORS["border"]}; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


inject_theme()

TABLES = {
    "weekly_overview": "weekly_overview",
    "monthly_overview": "monthly_overview",
    "daily_overview": "daily_overview",
    "channel_monthly": "channel_monthly",
    "channel_snapshot": "channel_snapshot",
    "ga_source": "ga_source",
}

# 채널 요약 시트로 취급하지 않을 시트들
SHEET_SKIP_EXACT = {"매체통합", "GA-RAW", "RD_네이버"}
SHEET_SKIP_SUBSTR = ["_data", "_date", "확인용", "소재"]


# ──────────────────────────────────────────────────────────────
# Supabase 연결 (secrets.toml 에 SUPABASE_URL / SUPABASE_KEY 필요)
# 설정이 없으면 로컬 세션 메모리로 동작 (테스트용, 새로고침 시 초기화됨)
# ──────────────────────────────────────────────────────────────
@st.cache_resource
def get_supabase_client():
    try:
        from supabase import create_client

        return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
    except Exception:
        return None


def _local_store():
    if "local_store" not in st.session_state:
        st.session_state["local_store"] = {k: pd.DataFrame() for k in TABLES}
    return st.session_state["local_store"]


@st.cache_data(ttl=60, show_spinner=False)
def load_table(name: str) -> pd.DataFrame:
    client = get_supabase_client()
    if client is None:
        return _local_store().get(name, pd.DataFrame()).copy()

    rows, page, page_size = [], 0, 1000
    while True:
        resp = (
            client.table(TABLES[name])
            .select("*")
            .range(page * page_size, page * page_size + page_size - 1)
            .execute()
        )
        chunk = resp.data or []
        rows.extend(chunk)
        if len(chunk) < page_size:
            break
        page += 1
    return pd.DataFrame(rows)


def save_table(name: str, df: pd.DataFrame, on_conflict: str, source_file: str):
    if df is None or df.empty:
        return 0
    df = df.copy()
    # 같은 업로드 안에 동일 키(예: 같은 월+매체) 행이 중복되면 upsert 한 번의 요청 안에서
    # 같은 행을 두 번 건드리게 되어 Postgres가 에러를 내므로, 저장 전에 미리 정리한다.
    key_cols = [c.strip() for c in on_conflict.split(",")]
    df = df.drop_duplicates(subset=key_cols, keep="last")
    df["source_file"] = source_file
    df["uploaded_at"] = datetime.utcnow().isoformat()

    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].astype(str)
        elif df[col].dtype == object:
            df[col] = df[col].apply(lambda v: v.isoformat() if isinstance(v, (date, datetime)) else v)

    client = get_supabase_client()
    if client is None:
        store = _local_store()
        prev = store.get(name, pd.DataFrame())
        merged = pd.concat([prev, df], ignore_index=True)
        keys = on_conflict.split(",")
        merged = merged.drop_duplicates(subset=keys, keep="last")
        store[name] = merged
        return len(df)

    records = df.to_dict(orient="records")
    for i in range(0, len(records), 500):
        client.table(TABLES[name]).upsert(records[i : i + 500], on_conflict=on_conflict).execute()
    return len(df)


# ──────────────────────────────────────────────────────────────
# 파싱 유틸
# ──────────────────────────────────────────────────────────────
def clean_col(c) -> str:
    if c is None:
        return ""
    return str(c).replace("\n", "").replace(" ", "").strip()


def match_col(columns, include_all=None, include_any=None, exclude=None):
    include_all, include_any, exclude = include_all or [], include_any or [], exclude or []
    for c in columns:
        lc = clean_col(c).lower()
        if not lc:
            continue
        if any(ex in lc for ex in exclude):
            continue
        if include_all and not all(tok in lc for tok in include_all):
            continue
        if include_any and not any(tok in lc for tok in include_any):
            continue
        return c
    return None


def metric_cols(columns):
    return dict(
        impr=match_col(columns, include_any=["노출"]),
        clicks=match_col(columns, include_any=["클릭"]),
        cost_ex=match_col(columns, include_all=["제외"], include_any=["광고비", "비용"], exclude=["마크업", "최종"]),
        cost_in=match_col(columns, include_all=["포함"], include_any=["광고비", "비용"], exclude=["마크업"]),
        signup=match_col(columns, include_any=["가입"]),
        conv=match_col(columns, include_all=["전환"], exclude=["금액", "ga", "율"]),
        rev=match_col(columns, include_any=["매출", "전환금액"], exclude=["ga", "객단가"]),
        ga_conv=match_col(columns, include_all=["ga"], include_any=["전환"]),
        ga_rev=match_col(columns, include_all=["ga"], include_any=["매출"]),
    )


def numcol(data: pd.DataFrame, c):
    if not c or c not in data.columns:
        return np.zeros(len(data))
    return pd.to_numeric(data[c], errors="coerce").fillna(0).values


def find_header_row(raw: pd.DataFrame, required=("노출수", "클릭수"), scan=10):
    for i in range(min(scan, len(raw))):
        row_text = " ".join(str(x) for x in raw.iloc[i].tolist())
        if all(tok in row_text for tok in required):
            return i
    return None


SECTION_MARKERS = {
    "monthly": ["월별 통합데이터", "월간 데이터"],
    "channel_snap": ["매체별 현황"],
    "weekly": ["통합 주간별", "주간 데이터"],
    "daily": ["통합 일자별", "일일 데이터", "일별 데이터"],
}


def find_sections(raw: pd.DataFrame, scan_cols=(0, 1, 2)):
    """시트 안에 세로로 쌓인 여러 표(월간/주간/일별/매체별 현황 등)의 경계를 찾는다.
    각 표는 '■ 월간 데이터' 같은 제목 행 다음 줄이 헤더, 그 다음부터 다음 제목 전까지가 데이터."""
    hits = []
    for i in range(len(raw)):
        for col in scan_cols:
            if col >= raw.shape[1]:
                continue
            v = raw.iat[i, col]
            if isinstance(v, str):
                for key, tokens in SECTION_MARKERS.items():
                    if any(tok in v for tok in tokens):
                        hits.append((i, key))
                        break
    hits.sort()
    bounds = {}
    for idx, (row_i, key) in enumerate(hits):
        end = hits[idx + 1][0] if idx + 1 < len(hits) else len(raw)
        bounds.setdefault(key, (row_i, end))
    return bounds


def section_dataframe(raw: pd.DataFrame, start_row: int, end_row: int, date_tokens=("기간", "월별")):
    """섹션 제목(start_row) 바로 다음 줄을 헤더로 보고 데이터프레임 구성."""
    header_row = start_row + 1
    headers = raw.iloc[header_row].tolist()
    date_idx = None
    for i, h in enumerate(headers):
        if clean_col(h) in date_tokens:
            date_idx = i
            break
    if date_idx is None:
        return None, None
    data = raw.iloc[header_row + 1 : end_row].copy()
    data.columns = headers
    data = data[data.iloc[:, date_idx].notna()]
    return data, date_idx


def parse_monthly(raw: pd.DataFrame, bounds, today: date):
    if "monthly" not in bounds:
        return pd.DataFrame()
    data, date_idx = section_dataframe(raw, *bounds["monthly"], date_tokens=("월별", "기간"))
    if data is None or data.empty:
        return pd.DataFrame()
    m = metric_cols(list(data.columns))
    out = pd.DataFrame()
    out["report_month"] = pd.to_datetime(data.iloc[:, date_idx], errors="coerce")
    out = out[out["report_month"].notna()]
    data = data.loc[out.index]
    out["impressions"] = numcol(data, m["impr"])
    out["clicks"] = numcol(data, m["clicks"])
    out["cost_excl_vat"] = numcol(data, m["cost_ex"])
    out["cost_incl_vat"] = numcol(data, m["cost_in"])
    out["signups"] = numcol(data, m["signup"])
    out["conversions"] = numcol(data, m["conv"])
    out["revenue"] = numcol(data, m["rev"])
    out["ga_conversions"] = numcol(data, m["ga_conv"])
    out["ga_revenue"] = numcol(data, m["ga_rev"])
    out["report_month"] = out["report_month"].dt.date
    cutoff = today.replace(day=1)
    out = out[out["report_month"] <= cutoff]
    return out.reset_index(drop=True)


def parse_weekly(raw: pd.DataFrame, bounds, today: date):
    if "weekly" not in bounds:
        return pd.DataFrame()
    data, date_idx = section_dataframe(raw, *bounds["weekly"], date_tokens=("기간", "월별"))
    if data is None or data.empty:
        return pd.DataFrame()
    m = metric_cols(list(data.columns))

    year_state = {"year": None, "prev_month": None}
    rows = []
    for _, r in data.iterrows():
        label = r.iloc[date_idx]
        mm = re.search(r"\((\d{1,2})/(\d{1,2})\s*~\s*(\d{1,2})/(\d{1,2})\)", str(label))
        lead = re.search(r"^(\d{1,2})월", str(label))
        if not mm or not lead:
            continue
        month_lead = int(lead.group(1))
        if year_state["year"] is None:
            # 통합 주간별 섹션은 월별 섹션의 첫 달과 같은 해에서 시작
            first_month_row = raw.iloc[bounds["monthly"][0] + 2] if "monthly" in bounds else None
            year_state["year"] = pd.to_datetime(first_month_row.iloc[1]).year if first_month_row is not None else today.year
        elif year_state["prev_month"] is not None and month_lead < year_state["prev_month"] - 6:
            year_state["year"] += 1
        year_state["prev_month"] = month_lead

        sm, sd, em, ed = map(int, mm.groups())
        s_year = year_state["year"]
        e_year = s_year if em >= sm else s_year + 1
        try:
            week_start = date(s_year, sm, sd)
            week_end = date(e_year, em, ed)
        except ValueError:
            continue

        rows.append(
            {
                "week_start": week_start,
                "week_end": week_end,
                "label": str(label).strip(),
                "impressions": float(pd.to_numeric(r.get(m["impr"]), errors="coerce") or 0) if m["impr"] else 0,
                "clicks": float(pd.to_numeric(r.get(m["clicks"]), errors="coerce") or 0) if m["clicks"] else 0,
                "cost_excl_vat": float(pd.to_numeric(r.get(m["cost_ex"]), errors="coerce") or 0) if m["cost_ex"] else 0,
                "cost_incl_vat": float(pd.to_numeric(r.get(m["cost_in"]), errors="coerce") or 0) if m["cost_in"] else 0,
                "signups": float(pd.to_numeric(r.get(m["signup"]), errors="coerce") or 0) if m["signup"] else 0,
                "conversions": float(pd.to_numeric(r.get(m["conv"]), errors="coerce") or 0) if m["conv"] else 0,
                "revenue": float(pd.to_numeric(r.get(m["rev"]), errors="coerce") or 0) if m["rev"] else 0,
            }
        )
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    out = out[out["week_end"] <= today]
    return out.reset_index(drop=True)


def parse_daily(raw: pd.DataFrame, bounds, today: date):
    """'3) 통합 일자별' 표를 파싱한다. 날짜 컬럼명은 '일자', 바로 옆에 '요일' 컬럼이 있다."""
    if "daily" not in bounds:
        return pd.DataFrame()
    data, date_idx = section_dataframe(raw, *bounds["daily"], date_tokens=("일자", "기간", "월별"))
    if data is None or data.empty:
        return pd.DataFrame()
    m = metric_cols(list(data.columns))
    out = pd.DataFrame()
    out["report_date"] = pd.to_datetime(data.iloc[:, date_idx], errors="coerce")
    out = out[out["report_date"].notna()]
    data = data.loc[out.index]
    out["impressions"] = numcol(data, m["impr"])
    out["clicks"] = numcol(data, m["clicks"])
    out["cost_excl_vat"] = numcol(data, m["cost_ex"])
    out["cost_incl_vat"] = numcol(data, m["cost_in"])
    out["signups"] = numcol(data, m["signup"])
    out["conversions"] = numcol(data, m["conv"])
    out["revenue"] = numcol(data, m["rev"])
    out["report_date"] = out["report_date"].dt.date
    out = out[out["report_date"] <= today]
    out = out.sort_values("report_date").reset_index(drop=True)
    # 아직 보고되지 않은(전부 0인) 말미 날짜는 잘라낸다 (리포트 템플릿의 미래 placeholder 행)
    metric_sum = out[["impressions", "clicks", "cost_excl_vat", "cost_incl_vat", "conversions", "revenue"]].sum(axis=1)
    nonzero_idx = metric_sum[metric_sum > 0].index
    if len(nonzero_idx):
        out = out.loc[: nonzero_idx.max()]
    return out.reset_index(drop=True)


def parse_channel_snapshot(raw: pd.DataFrame, bounds, monthly_df: pd.DataFrame):
    if "channel_snap" not in bounds:
        return pd.DataFrame()
    data, date_idx_unused = section_dataframe(raw, *bounds["channel_snap"], date_tokens=("매체",))
    if data is None or data.empty:
        return pd.DataFrame()
    m = metric_cols(list(data.columns))
    channel_col = data.columns[1] if clean_col(data.columns[1]) == "매체" else data.columns[0]
    out = pd.DataFrame()
    out["channel"] = data[channel_col].astype(str)
    out = out[~out["channel"].str.contains("TOTAL", case=False, na=False)]
    data = data.loc[out.index]
    out["impressions"] = numcol(data, m["impr"])
    out["clicks"] = numcol(data, m["clicks"])
    out["cost_excl_vat"] = numcol(data, m["cost_ex"])
    out["cost_incl_vat"] = numcol(data, m["cost_in"])
    out["signups"] = numcol(data, m["signup"])
    out["conversions"] = numcol(data, m["conv"])
    out["revenue"] = numcol(data, m["rev"])
    out["ga_conversions"] = numcol(data, m["ga_conv"])
    out["ga_revenue"] = numcol(data, m["ga_rev"])
    as_of = monthly_df["report_month"].max() if len(monthly_df) else date.today().replace(day=1)
    out["as_of_month"] = as_of
    return out.reset_index(drop=True)


def discover_channel_sheets(xls: pd.ExcelFile):
    names = []
    for s in xls.sheet_names:
        if s in SHEET_SKIP_EXACT:
            continue
        low = s.lower()
        if any(p in low for p in SHEET_SKIP_SUBSTR):
            continue
        names.append(s)
    return names


def parse_channel_sheet(xls: pd.ExcelFile, sheet: str, today: date):
    raw = pd.read_excel(xls, sheet_name=sheet, header=None)
    bounds = find_sections(raw)
    if "monthly" in bounds:
        data, date_idx = section_dataframe(raw, *bounds["monthly"], date_tokens=("기간", "월별"))
    else:
        # '■ 월간 데이터' 같은 섹션 제목이 없는 단순 시트는 기존 방식으로 폴백
        hdr = find_header_row(raw)
        if hdr is None:
            return None
        headers = raw.iloc[hdr].tolist()
        date_idx = next((i for i, h in enumerate(headers) if clean_col(h) in ("기간", "월별")), None)
        if date_idx is None:
            return None
        data = raw.iloc[hdr + 1 :].copy()
        data.columns = headers
        data = data[data.iloc[:, date_idx].notna()]
    if data is None or data.empty:
        return None
    m = metric_cols(list(data.columns))
    out = pd.DataFrame()
    out["report_month"] = pd.to_datetime(data.iloc[:, date_idx], errors="coerce")
    out = out[out["report_month"].notna()]
    data = data.loc[out.index]
    out["impressions"] = numcol(data, m["impr"])
    out["clicks"] = numcol(data, m["clicks"])
    out["cost_excl_vat"] = numcol(data, m["cost_ex"])
    out["cost_incl_vat"] = numcol(data, m["cost_in"])
    out["signups"] = numcol(data, m["signup"])
    out["conversions"] = numcol(data, m["conv"])
    out["revenue"] = numcol(data, m["rev"])
    out["report_month"] = out["report_month"].dt.date
    out = out[out["report_month"] <= today.replace(day=1)]
    out["channel"] = sheet
    return out.reset_index(drop=True)


def parse_ga_raw(xls: pd.ExcelFile, today: date):
    if "GA-RAW" not in xls.sheet_names:
        return pd.DataFrame()
    raw = pd.read_excel(xls, sheet_name="GA-RAW")
    raw.columns = [clean_col(c) for c in raw.columns]
    rename = {
        "매체": "source_medium",
        "사용자": "users",
        "신규방문자": "new_users",
        "세션": "sessions",
        "이탈률": "bounce_rate",
        "세션당페이지수": "pages_per_session",
        "평균세션시간": "avg_session_duration",
        "전자상거래전환율": "ecommerce_cvr",
        "거래수": "transactions",
        "수익": "revenue",
    }
    raw = raw.rename(columns=rename)
    keep = [c for c in rename.values() if c in raw.columns]
    out = raw[keep].dropna(subset=["source_medium"]).copy()
    for c in keep:
        if c != "source_medium":
            out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0)
    out["as_of_date"] = today
    return out.reset_index(drop=True)


def parse_workbook(file, today: date):
    xls = pd.ExcelFile(file)
    result = {
        "weekly": pd.DataFrame(),
        "monthly": pd.DataFrame(),
        "daily": pd.DataFrame(),
        "channel_snapshot": pd.DataFrame(),
        "channels": pd.DataFrame(),
        "ga": pd.DataFrame(),
        "channel_sheets_found": [],
        "channel_sheets_parsed": [],
    }
    if "매체통합" in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name="매체통합", header=None)
        bounds = find_sections(raw)
        result["monthly"] = parse_monthly(raw, bounds, today)
        result["weekly"] = parse_weekly(raw, bounds, today)
        result["daily"] = parse_daily(raw, bounds, today)
        result["channel_snapshot"] = parse_channel_snapshot(raw, bounds, result["monthly"])

    chan_frames = []
    for s in discover_channel_sheets(xls):
        result["channel_sheets_found"].append(s)
        df = parse_channel_sheet(xls, s, today)
        if df is not None and len(df):
            chan_frames.append(df)
            result["channel_sheets_parsed"].append(s)
    if chan_frames:
        result["channels"] = pd.concat(chan_frames, ignore_index=True)

    result["ga"] = parse_ga_raw(xls, today)
    return result


# ──────────────────────────────────────────────────────────────
# KPI
# ──────────────────────────────────────────────────────────────
def add_kpis(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    cost_in = df["cost_incl_vat"] if "cost_incl_vat" in df else 0
    cost_ex = df["cost_excl_vat"] if "cost_excl_vat" in df else 0
    df["ctr"] = np.where(df.get("impressions", 0) > 0, df["clicks"] / df["impressions"] * 100, 0)
    df["cpc"] = np.where(df.get("clicks", 0) > 0, cost_in / df["clicks"], 0)
    df["cpa"] = np.where(df.get("conversions", 0) > 0, cost_ex / df["conversions"], 0)
    df["cvr"] = np.where(df.get("clicks", 0) > 0, df["conversions"] / df["clicks"] * 100, 0)
    df["roas"] = np.where(cost_in > 0, df["revenue"] / cost_in * 100, 0)
    df["aov"] = np.where(df.get("conversions", 0) > 0, df["revenue"] / df["conversions"], 0)
    if "ga_revenue" in df.columns:
        df["ga_roas"] = np.where(cost_in > 0, df["ga_revenue"] / cost_in * 100, 0)
    return df


def kpi_cards(df: pd.DataFrame):
    cost = df["cost_incl_vat"].sum()
    revenue = df["revenue"].sum()
    conv = df["conversions"].sum()
    roas = (revenue / cost * 100) if cost > 0 else 0
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("총 광고비 (VAT포함)", f"{cost:,.0f} 원")
    c2.metric("총 매출", f"{revenue:,.0f} 원")
    c3.metric("총 전환수", f"{conv:,.0f} 건")
    c4.metric("ROAS", f"{roas:,.1f} %")


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    return buf.getvalue()


# 화면/엑셀에 표시할 때 쓰는 한글 컬럼명
KOR_COLS = {
    "channel": "매체",
    "impressions": "노출수",
    "clicks": "클릭수",
    "cost_excl_vat": "광고비(VAT제외)",
    "cost_incl_vat": "광고비(VAT포함)",
    "signups": "회원가입",
    "conversions": "전환수",
    "revenue": "매출",
    "ctr": "CTR(%)",
    "cpc": "CPC",
    "cpa": "CPA",
    "cvr": "CVR(%)",
    "roas": "ROAS(%)",
    "aov": "객단가",
    "ga_conversions": "GA-전환수",
    "ga_revenue": "GA-매출",
    "ga_roas": "GA-ROAS(%)",
    "report_month": "월",
    "report_date": "일자",
    "weekday": "요일",
    "week_start": "주 시작일",
    "week_end": "주 종료일",
    "label": "기간",
    "week_no": "주차",
    "week_range": "기간(월~일)",
    "as_of_month": "기준월",
    "as_of_date": "기준일",
    "source_medium": "소스/매체",
    "users": "사용자",
    "new_users": "신규방문자",
    "sessions": "세션",
    "bounce_rate": "이탈률(%)",
    "pages_per_session": "세션당 페이지수",
    "avg_session_duration": "평균 세션시간(초)",
    "ecommerce_cvr": "전자상거래 전환율(%)",
    "transactions": "거래수",
}


def korify(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns=KOR_COLS)


DATE_PRESETS = [
    "오늘", "어제", "이번주", "지난주",
    "최근 7일(오늘 포함)", "최근 7일(오늘 제외)",
    "이번달", "지난달",
    "최근 30일(오늘 포함)", "최근 30일(오늘 제외)",
]


def _preset_to_range(name: str, min_d: date, max_d: date):
    """프리셋 이름 → (start, end). '오늘' 기준일은 실제 오늘 날짜이되, 데이터 범위 밖이면 잘라낸다."""
    today = date.today()
    if name == "오늘":
        s, e = today, today
    elif name == "어제":
        s = e = today - timedelta(days=1)
    elif name == "이번주":
        s, e = today - timedelta(days=today.weekday()), today
    elif name == "지난주":
        this_mon = today - timedelta(days=today.weekday())
        s = this_mon - timedelta(days=7)
        e = s + timedelta(days=6)
    elif name == "최근 7일(오늘 포함)":
        s, e = today - timedelta(days=6), today
    elif name == "최근 7일(오늘 제외)":
        s, e = today - timedelta(days=7), today - timedelta(days=1)
    elif name == "이번달":
        s, e = today.replace(day=1), today
    elif name == "지난달":
        last_prev = today.replace(day=1) - timedelta(days=1)
        s, e = last_prev.replace(day=1), last_prev
    elif name == "최근 30일(오늘 포함)":
        s, e = today - timedelta(days=29), today
    elif name == "최근 30일(오늘 제외)":
        s, e = today - timedelta(days=30), today - timedelta(days=1)
    else:
        s, e = min_d, max_d
    s = max(min_d, min(s, max_d))
    e = min(max_d, max(e, min_d))
    if s > e:
        s = e
    return s, e


def period_filter(min_d: date, max_d: date, key: str, default_preset: str = "이번달"):
    """날짜범위 버튼(달력 아이콘 + 시작~종료일 + ◀/▶) 클릭 시 프리셋/직접선택 패널이 열리는
    기간 선택 UI. 반환값은 (start, end)."""
    start_key, end_key = f"{key}_drp_start", f"{key}_drp_end"
    cal_key = f"{key}_drp_calendar"

    if start_key not in st.session_state:
        s, e = _preset_to_range(default_preset, min_d, max_d)
        st.session_state[start_key], st.session_state[end_key] = s, e

    cur_start = max(min_d, min(st.session_state[start_key], max_d))
    cur_end = min(max_d, max(st.session_state[end_key], min_d))

    # 버튼 폭을 좁게 고정 (전체 폭으로 늘어나지 않도록 좁은 컬럼 안에만 배치하고 나머지는 빈 컬럼으로 남긴다)
    col_prev, col_main, col_next, _spacer = st.columns([1, 2, 1, 12])
    with col_prev:
        if st.button("◀", key=f"{key}_drp_prev", use_container_width=True):
            span = (cur_end - cur_start).days + 1
            new_end = cur_start - timedelta(days=1)
            new_start = new_end - timedelta(days=span - 1)
            if new_end >= min_d:
                st.session_state[start_key] = max(min_d, new_start)
                st.session_state[end_key] = new_end
                if cal_key in st.session_state:
                    del st.session_state[cal_key]
                st.rerun()
    with col_next:
        if st.button("▶", key=f"{key}_drp_next", use_container_width=True):
            span = (cur_end - cur_start).days + 1
            new_start = cur_end + timedelta(days=1)
            new_end = new_start + timedelta(days=span - 1)
            if new_start <= max_d:
                st.session_state[start_key] = new_start
                st.session_state[end_key] = min(max_d, new_end)
                if cal_key in st.session_state:
                    del st.session_state[cal_key]
                st.rerun()
    with col_main:
        with st.popover(f"📅 {cur_start:%Y.%m.%d} → {cur_end:%Y.%m.%d}", use_container_width=True):
            # date_input은 key로 지정된 세션 상태를 직접 소유한다. 프리셋 버튼에서 값을 바꾸려면
            # date_input을 호출하기 "전"에 같은 key(cal_key)로 세션 상태를 직접 써야 화면에 반영된다
            # (value= 인자는 위젯이 이미 존재하면 무시되기 때문에 별도 변수로는 반영되지 않았던 버그 수정).
            if cal_key not in st.session_state:
                st.session_state[cal_key] = (cur_start, cur_end)

            # 왼쪽: 프리셋 목록(세로 1열) / 오른쪽: 달력(직접 선택) + 취소·확인
            preset_col, calendar_col = st.columns([1, 2])
            with preset_col:
                for i, p in enumerate(DATE_PRESETS):
                    if st.button(p, key=f"{key}_drp_preset_{i}", use_container_width=True):
                        s, e = _preset_to_range(p, min_d, max_d)
                        st.session_state[cal_key] = (s, e)
                        st.rerun()

            with calendar_col:
                dr = st.date_input(
                    "직접 선택", min_value=min_d, max_value=max_d, key=cal_key,
                )
                if isinstance(dr, tuple) and len(dr) == 2:
                    pend_s, pend_e = dr
                else:
                    pend_s, pend_e = cur_start, cur_end

                bc1, bc2 = st.columns(2)
                if bc1.button("취소", key=f"{key}_drp_cancel", use_container_width=True):
                    st.session_state[cal_key] = (cur_start, cur_end)
                    st.rerun()
                if bc2.button("확인", key=f"{key}_drp_confirm", type="primary", use_container_width=True):
                    st.session_state[start_key] = pend_s
                    st.session_state[end_key] = pend_e
                    st.rerun()

    return st.session_state[start_key], st.session_state[end_key]


def preset_button_picker(options: list, key: str, default: str, label_prefix: str = "📅"):
    """연도/기간 프리셋을 달력 버튼 + 팝오버 패널 방식으로 고르는 UI (누적 표용).
    st.selectbox 대신 목업과 같은 '버튼 → 패널' 스타일을 쓰되, 옵션 내용은 그대로(연도 단위) 유지."""
    sel_key = f"{key}_preset_sel"
    if sel_key not in st.session_state or st.session_state[sel_key] not in options:
        st.session_state[sel_key] = default if default in options else options[0]

    current = st.session_state[sel_key]
    with st.popover(f"{label_prefix} {current}", use_container_width=False):
        cols = st.columns(2)
        for i, opt in enumerate(options):
            c = cols[i % 2]
            btn_type = "primary" if opt == current else "secondary"
            if c.button(opt, key=f"{key}_preset_opt_{i}", use_container_width=True, type=btn_type):
                st.session_state[sel_key] = opt
                st.rerun()
    return st.session_state[sel_key]


PAGE_SIZE_OPTIONS = [20, 50, 100, 200]

# 표시 포맷: 정수+콤마(돈/카운트), 소수점 2자리 %(CTR·CVR류), 소수점 0자리 %(ROAS류)
MONEY_COLS = {
    "impressions", "clicks", "signups", "conversions", "cost_excl_vat", "cost_incl_vat",
    "cpc", "cpa", "revenue", "aov", "ga_conversions", "ga_revenue",
    "users", "new_users", "sessions", "transactions",
}
PCT2_COLS = {"ctr", "cvr", "bounce_rate", "ecommerce_cvr"}
PCT0_COLS = {"roas", "ga_roas"}


def format_display(df: pd.DataFrame) -> pd.DataFrame:
    """화면/엑셀에 보여줄 때 쓰는 최종 포맷팅 (콤마, 소수점 자리수, 날짜 형식)."""
    df = df.copy()
    for c in df.columns:
        if c == "report_month":
            df[c] = pd.to_datetime(df[c]).dt.strftime("%Y-%m")
        elif c in ("week_start", "week_end", "report_date", "as_of_month", "as_of_date"):
            df[c] = pd.to_datetime(df[c]).dt.strftime("%Y-%m-%d")
        elif c in MONEY_COLS:
            df[c] = pd.to_numeric(df[c], errors="coerce").map(lambda v: f"{v:,.0f}" if pd.notna(v) else "")
        elif c in PCT2_COLS:
            df[c] = pd.to_numeric(df[c], errors="coerce").map(lambda v: f"{v:.2f}%" if pd.notna(v) else "")
        elif c in PCT0_COLS:
            df[c] = pd.to_numeric(df[c], errors="coerce").map(lambda v: f"{v:.0f}%" if pd.notna(v) else "")
    return df


def render_html_table(table: pd.DataFrame):
    """pandas Styler(jinja2 의존) 없이 순수 HTML로 표를 그린다.
    ▲(상승)는 빨간색, ▼(하락)는 파란색 글씨로 표시하고, 인덱스는 표시하지 않는다."""
    if table.empty:
        st.caption("데이터가 아직 없습니다.")
        return

    cols = list(table.columns)
    thead = "".join(f"<th>{c}</th>" for c in cols)

    row_htmls = []
    for _, row in table.iterrows():
        first_text = str(row[cols[0]]).strip()
        is_total = first_text == "TOTAL"
        is_label_row = is_total or first_text.endswith("대비")
        # TOTAL/증감 행에서 앞의 두 컬럼이 라벨용(둘 다 텍스트)이고 두 번째 칸이 비어있으면
        # 두 칸을 하나로 합쳐서(colspan) 가운데 정렬로 보여준다 (예: 주차 + 기간 컬럼).
        merge_first_two = False
        if is_label_row and len(cols) > 1:
            second_val = row[cols[1]]
            second_text = "" if pd.isna(second_val) else str(second_val).strip()
            merge_first_two = second_text == ""

        cells = []
        skip_next = False
        for i, c in enumerate(cols):
            if skip_next:
                skip_next = False
                continue
            val = row[c]
            text = "" if pd.isna(val) else str(val)
            style = ""
            colspan = ""
            if merge_first_two and i == 0:
                colspan = ' colspan="2"'
                style = "text-align:center;"
                skip_next = True
            if text.startswith("▲"):
                style += "color:#d93025;"
            elif text.startswith("▼"):
                style += "color:#1a73e8;"
            cells.append(f'<td{colspan} style="{style}">{text}</td>')
        row_class = ' class="stco-total-row"' if is_total else ""
        row_htmls.append(f"<tr{row_class}>{''.join(cells)}</tr>")

    html = f"""
    <style>
    .stco-table-wrap {{
        overflow-x:auto; border:1px solid {THEME_COLORS["border"]}; border-radius:10px; background:{THEME_COLORS["canvas"]};
    }}
    .stco-table {{
        width:100%; border-collapse:collapse; font-size:14px;
        font-family: {THEME_FONT_STACK};
    }}
    .stco-table th {{
        background:{THEME_COLORS["surface"]}; color:{THEME_COLORS["muted"]}; font-weight:600; padding:8px 14px;
        text-align:right; border-bottom:1px solid {THEME_COLORS["border"]}; white-space:nowrap;
    }}
    .stco-table th:first-child {{ text-align:left; border-top-left-radius:10px; }}
    .stco-table th:last-child {{ border-top-right-radius:10px; }}
    .stco-table td {{
        padding:8px 14px; text-align:right; color:{THEME_COLORS["foreground"]};
        border-bottom:1px solid {THEME_COLORS["border"]}; white-space:nowrap;
    }}
    .stco-table td:first-child {{ text-align:left; }}
    .stco-table tr:last-child td {{ border-bottom:none; }}
    .stco-table tr:hover td {{ background:{THEME_COLORS["surface"]}; }}
    .stco-table tr.stco-total-row td {{
        background:{THEME_COLORS["surface"]}; color:{THEME_COLORS["foreground"]};
        font-weight:700; border-top:2px solid {THEME_COLORS["border"]};
    }}
    .stco-table tr.stco-total-row:hover td {{ background:{THEME_COLORS["surface"]}; }}
    </style>
    <div class="stco-table-wrap">
    <table class="stco-table">
      <thead><tr>{thead}</tr></thead>
      <tbody>{''.join(row_htmls)}</tbody>
    </table>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)


def pct_change_row(d_full: pd.DataFrame, latest_pos: int, numeric_cols: list, label_col: str, label_text: str = "전기간 대비"):
    """d_full(전체 정렬 데이터)에서 latest_pos 위치의 행과 바로 이전 행을 비교해 증감율 행을 만든다."""
    if latest_pos <= 0 or latest_pos >= len(d_full):
        return None
    latest, prev = d_full.iloc[latest_pos], d_full.iloc[latest_pos - 1]
    row = {label_col: label_text}
    for c in numeric_cols:
        if c not in d_full.columns:
            continue
        pv, lv = prev.get(c), latest.get(c)
        if pd.isna(pv) or pv in (0, None):
            row[c] = "-"
            continue
        change = (lv - pv) / abs(pv) * 100
        arrow = "▲" if change >= 0 else "▼"
        row[c] = f"{arrow}{change:+.1f}%"
    return row


def build_total_row(view_raw: pd.DataFrame, display_cols: list, label_col: str, label_text: str = "TOTAL"):
    """현재 화면에 표시 중인 원본(raw) 행들을 합산해 TOTAL 행을 만든다.
    노출/클릭/비용/전환/매출 등은 단순 합산하고, CTR·CPC·CPA·CVR·ROAS·객단가·GA-ROAS 같은
    비율/단가 지표는 합산된 값 기준으로 다시 계산한다 (개별 행 비율의 평균이 아님)."""
    if view_raw is None or view_raw.empty:
        return None

    def s(col):
        return pd.to_numeric(view_raw[col], errors="coerce").sum() if col in view_raw.columns else 0

    imp, clk = s("impressions"), s("clicks")
    cost_ex, cost_in = s("cost_excl_vat"), s("cost_incl_vat")
    signups, conv, revenue = s("signups"), s("conversions"), s("revenue")
    ga_conv, ga_rev = s("ga_conversions"), s("ga_revenue")

    raw = {
        "impressions": imp, "clicks": clk, "cost_excl_vat": cost_ex, "cost_incl_vat": cost_in,
        "signups": signups, "conversions": conv, "revenue": revenue,
        "ga_conversions": ga_conv, "ga_revenue": ga_rev,
        "ctr": (clk / imp * 100) if imp else 0,
        "cpc": (cost_in / clk) if clk else 0,
        "cpa": (cost_ex / conv) if conv else 0,
        "cvr": (conv / clk * 100) if clk else 0,
        "roas": (revenue / cost_in * 100) if cost_in else 0,
        "aov": (revenue / conv) if conv else 0,
        "ga_roas": (ga_rev / cost_in * 100) if cost_in else 0,
    }
    row = {label_col: label_text}
    for c in display_cols:
        if c == label_col:
            continue
        if c not in raw:
            row[c] = ""
            continue
        v = raw[c]
        if c in MONEY_COLS:
            row[c] = f"{v:,.0f}"
        elif c in PCT2_COLS:
            row[c] = f"{v:.2f}%"
        elif c in PCT0_COLS:
            row[c] = f"{v:.0f}%"
        else:
            row[c] = v
    return row


def build_year_options(date_series: pd.Series):
    years = sorted({d.year for d in pd.to_datetime(date_series).dropna()})
    return years


def render_pager(total_pages: int, key: str) -> int:
    """« 1 2 3 » 형태의 페이지 버튼. 처음 볼 땐 최신 데이터가 있는 마지막 페이지부터 보여준다."""
    state_key = f"{key}_pagenum"
    if state_key not in st.session_state:
        st.session_state[state_key] = total_pages
    cur = min(max(st.session_state[state_key], 1), total_pages)

    window = 5
    start_p = max(1, cur - window // 2)
    end_p = min(total_pages, start_p + window - 1)
    start_p = max(1, end_p - window + 1)

    spacer, pager_area = st.columns([3, 4])
    with pager_area:
        n_buttons = end_p - start_p + 3
        btn_cols = st.columns(n_buttons)
        if btn_cols[0].button("«", key=f"{key}_prev", disabled=cur <= 1, use_container_width=True):
            cur = max(1, cur - 1)
        for i, p in enumerate(range(start_p, end_p + 1)):
            if btn_cols[i + 1].button(
                str(p), key=f"{key}_p{p}", type=("primary" if p == cur else "secondary"), use_container_width=True
            ):
                cur = p
        if btn_cols[-1].button("»", key=f"{key}_next", disabled=cur >= total_pages, use_container_width=True):
            cur = min(total_pages, cur + 1)
    st.session_state[state_key] = cur
    return cur


def render_cumulative_table(df: pd.DataFrame, date_col: str, show_cols: list, numeric_cols: list,
                             title: str, key: str, mode: str):
    """월별/주간/일자별 누적 표.
    mode: 'month' → 기본값 당해년도(1월~최신월) / 'week' → 기본값 최근 5주 / 'day' → 기본값 이번달
    프리셋: (week·day는) 기본, 연도별, 전체, 직접선택
    """
    st.markdown(f"#### {title}")
    if df is None or df.empty:
        st.caption("데이터가 아직 없습니다.")
        return

    d = df.sort_values(date_col).reset_index(drop=True)
    d[date_col] = pd.to_datetime(d[date_col])
    years = build_year_options(d[date_col])
    year_labels = [f"{y}년" for y in years]

    if mode == "month":
        options = year_labels + ["전체", "직접선택"]
        this_year = date.today().year
        default_label = f"{this_year}년" if this_year in years else (year_labels[-1] if year_labels else "전체")
    else:
        options = ["기본"] + year_labels + ["전체", "직접선택"]
        default_label = "기본"

    preset = preset_button_picker(options, key=f"{key}_periodpicker", default=default_label)

    need_pagination = True
    if preset == "기본":
        if mode == "week":
            view_all = d.tail(5)
        else:  # day
            cur_month = pd.Timestamp(date.today().replace(day=1))
            view_all = d[d[date_col] >= cur_month]
            if view_all.empty:
                view_all = d.tail(31)
        need_pagination = False
    elif preset == "전체":
        view_all = d
    elif preset == "직접선택":
        min_d, max_d = d[date_col].min().date(), d[date_col].max().date()
        date_range = st.date_input("기간 직접 선택", value=(min_d, max_d), min_value=min_d, max_value=max_d, key=f"{key}_manual")
        start, end = date_range if isinstance(date_range, tuple) and len(date_range) == 2 else (min_d, max_d)
        view_all = d[(d[date_col].dt.date >= start) & (d[date_col].dt.date <= end)]
    else:  # "YYYY년"
        y = int(preset.replace("년", ""))
        view_all = d[d[date_col].dt.year == y]

    total = len(view_all)
    if need_pagination and total > PAGE_SIZE_OPTIONS[0]:
        page_size = st.selectbox("페이지당 표시", PAGE_SIZE_OPTIONS, index=1, key=f"{key}_{preset}_pagesize")
        total_pages = max(1, -(-total // page_size))
        page = render_pager(total_pages, key=f"{key}_{preset}") if total_pages > 1 else 1
        start_i, end_i = (page - 1) * page_size, page * page_size
        view = view_all.iloc[start_i:end_i]
        show_change_row = page == total_pages
    else:
        view = view_all
        show_change_row = True

    display_cols = [c for c in show_cols if c in view.columns]
    table = format_display(view[display_cols])  # 먼저 숫자/날짜 포맷 적용 (증감율 행은 이미 문자열이라 따로 붙임)

    change_label = {"month": "전월 대비", "week": "전주 대비", "day": "전일 대비"}.get(mode, "전기간 대비")
    has_change_row = False
    if show_change_row and len(view):
        latest_pos = view.index[-1]
        change_row = pct_change_row(d, latest_pos, numeric_cols, display_cols[0], label_text=change_label)
        if change_row:
            table = pd.concat([table, pd.DataFrame([change_row])], ignore_index=True)
            has_change_row = True

    total_row = build_total_row(view[display_cols], display_cols, display_cols[0], label_text="TOTAL")
    if total_row:
        table = pd.concat([table, pd.DataFrame([total_row])], ignore_index=True)

    render_html_table(korify(table))
    st.download_button(
        f"⬇️ 엑셀 다운로드 ({title})",
        data=to_excel_bytes(korify(format_display(view[display_cols]))),
        file_name=f"{key}.xlsx",
        key=f"{key}_{preset}_dl",
    )


# ──────────────────────────────────────────────────────────────
# 업로드 패널
# ──────────────────────────────────────────────────────────────
def render_upload_panel():
    st.sidebar.header("⚙️ 데이터 관리")
    client = get_supabase_client()
    st.sidebar.caption(f"저장소: {'Supabase (Postgres)' if client else '로컬 세션 (테스트용, 새로고침 시 초기화)'}")

    file = st.sidebar.file_uploader("① 주간 리포트 업로드 (STCO_주간보고서_...xlsx)", type=["xlsx", "xls"])

    if file is not None:
        today = date.today()
        with st.sidebar.status("파일 분석 중...", expanded=True) as status:
            result = parse_workbook(file, today)
            st.write(f"📅 월별 통합데이터: {len(result['monthly'])}개월")
            st.write(f"📆 통합 주간별: {len(result['weekly'])}주")
            st.write(f"🗓️ 통합 일자별: {len(result['daily'])}일")
            st.write(f"🏷️ 당월 매체별 스냅샷: {len(result['channel_snapshot'])}개 매체")
            st.write(f"📊 매체별 시트 인식: {len(result['channel_sheets_parsed'])}/{len(result['channel_sheets_found'])}개")
            st.write(f"🔎 GA 유입경로: {len(result['ga'])}건")
            missing = set(result["channel_sheets_found"]) - set(result["channel_sheets_parsed"])
            if missing:
                st.warning(f"인식 실패한 매체 시트: {', '.join(missing)}")
            status.update(label="분석 완료", state="complete")

        if st.sidebar.button("💾 전체 저장하기", type="primary"):
            n1 = save_table("weekly_overview", result["weekly"], "week_start", file.name)
            n2 = save_table("monthly_overview", result["monthly"], "report_month", file.name)
            n3 = save_table("channel_monthly", result["channels"], "report_month,channel", file.name)
            n4 = save_table("channel_snapshot", result["channel_snapshot"], "as_of_month,channel", file.name)
            n5 = save_table("ga_source", result["ga"], "as_of_date,source_medium", file.name)
            n6 = save_table("daily_overview", result["daily"], "report_date", file.name)
            st.cache_data.clear()
            st.sidebar.success(f"저장 완료! 주간 {n1} · 월별 {n2} · 일자별 {n6} · 매체(월) {n3} · 매체(당월) {n4} · GA {n5}건")
            st.rerun()

    st.sidebar.markdown("---")
    wk = load_table("weekly_overview")
    st.sidebar.metric("누적 주간 데이터", f"{len(wk):,} 주")
    if st.sidebar.button("🔄 새로고침 (캐시 비우기)"):
        st.cache_data.clear()
        st.rerun()


# ──────────────────────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────────────────────
def main():
    st.title("📊 STCO 온라인팀 광고/마케팅 성과 대시보드")
    render_upload_panel()

    weekly = load_table("weekly_overview")
    monthly = load_table("monthly_overview")
    daily = load_table("daily_overview")
    channels = load_table("channel_monthly")
    snapshot = load_table("channel_snapshot")
    ga = load_table("ga_source")

    if weekly.empty and monthly.empty:
        st.info("아직 저장된 데이터가 없습니다. 왼쪽 사이드바에서 주간 리포트 파일을 업로드하고 '전체 저장하기'를 눌러주세요.")
        return

    for df, col in [(weekly, "week_start"), (weekly, "week_end"), (monthly, "report_month"), (daily, "report_date")]:
        if not df.empty and col in df.columns:
            df[col] = pd.to_datetime(df[col]).dt.date

    tab1, tab2, tab3, tab4 = st.tabs(["종합 대시보드", "매체별 성과", "GA 유입경로", "GA4 라이브 리포트"])

    # ── 종합 대시보드 ──────────────────────────────
    with tab1:
        if not weekly.empty:
            st.subheader("🔎 기간 필터 (주간 기준)")
            min_d, max_d = weekly["week_start"].min(), weekly["week_end"].max()
            start, end = period_filter(min_d, max_d, key="weekly")
            fw = weekly[(weekly["week_start"] >= start) & (weekly["week_start"] <= end)]
            fw = add_kpis(fw).sort_values("week_start")

            kpi_cards(fw)
            st.markdown("### 주간 추이")
            c1, c2 = st.columns(2)
            with c1:
                chart_df = fw.rename(columns={"cost_incl_vat": "광고비(VAT포함)", "revenue": "매출"})
                fig = px.bar(
                    chart_df, x="week_start", y=["광고비(VAT포함)", "매출"], barmode="group",
                    title="주간 비용(VAT포함) vs 매출",
                    labels={"week_start": "주 시작일", "value": "금액(원)", "variable": "구분"},
                )
                fig.update_yaxes(tickformat=",.0f")
                fig.for_each_trace(
                    lambda t: t.update(
                        hovertemplate=f"구분={t.name}<br>주 시작일=%{{x}}<br>금액(원)=%{{y:,.0f}}원<extra></extra>"
                    )
                )
                st.plotly_chart(theme_chart(fig), use_container_width=True)
            with c2:
                fig2 = px.line(
                    fw, x="week_start", y="roas", markers=True, title="주간 ROAS 추이 (%)",
                    labels={"week_start": "주 시작일", "roas": "ROAS(%)"},
                )
                st.plotly_chart(theme_chart(fig2), use_container_width=True)

        else:
            st.info("주간 데이터가 아직 없습니다.")

        if not monthly.empty:
            st.markdown("---")
            st.markdown("### 월별 GA-ROAS vs 플랫폼 ROAS")
            fm_chart = add_kpis(monthly).sort_values("report_month").rename(
                columns={"roas": "플랫폼 ROAS", "ga_roas": "GA ROAS"}
            )
            fig3 = px.line(
                fm_chart, x="report_month", y=["플랫폼 ROAS", "GA ROAS"], markers=True,
                labels={"report_month": "월", "value": "ROAS(%)", "variable": "기준"},
                title="플랫폼 리포팅 ROAS vs GA 기준 ROAS",
            )
            st.plotly_chart(theme_chart(fig3), use_container_width=True)
            st.caption("* GA-매출/GA-ROAS는 쇼핑검색 및 GFA 외부몰 데이터가 미집계될 수 있습니다 (원본 시트 주석 기준).")

        # ── 누적 데이터 (월별 / 주간별 / 일자별) ──────────────────
        st.markdown("---")
        st.markdown("## 📚 누적 데이터")
        st.caption("기본은 최근 데이터만 보여주고, '이전 데이터 더 보기'를 켜면 10/30/50/100/200개 단위로 넘겨볼 수 있어요.")

        month_show_cols = ["report_month", "impressions", "clicks", "ctr", "cpc", "cost_excl_vat", "cost_incl_vat",
                            "signups", "cpa", "conversions", "cvr", "revenue", "roas", "aov",
                            "ga_conversions", "ga_revenue", "ga_roas"]
        month_numeric_cols = [c for c in month_show_cols if c != "report_month"]
        render_cumulative_table(
            add_kpis(monthly) if not monthly.empty else monthly,
            date_col="report_month", show_cols=month_show_cols, numeric_cols=month_numeric_cols,
            title="1) 월별 누적", key="monthly_cum", mode="month",
        )

        wk = weekly.copy()
        if not wk.empty:
            wk["week_no"] = wk["label"].astype(str).str.replace(r"\s*\(.*\)\s*$", "", regex=True).str.strip()
            wk["week_range"] = wk.apply(lambda r: f"{r['week_start']:%Y-%m-%d}~{r['week_end']:%Y-%m-%d}", axis=1)
        week_show_cols = ["week_range", "week_no", "impressions", "clicks", "ctr", "cpc",
                           "cost_excl_vat", "cost_incl_vat", "signups", "cpa", "conversions", "cvr", "revenue", "roas", "aov"]
        week_numeric_cols = [c for c in week_show_cols if c not in ("week_no", "week_range")]
        render_cumulative_table(
            add_kpis(wk) if not wk.empty else wk,
            date_col="week_start", show_cols=week_show_cols, numeric_cols=week_numeric_cols,
            title="2) 주간별 누적", key="weekly_cum", mode="week",
        )

        day_show_cols = ["report_date", "impressions", "clicks", "ctr", "cpc", "cost_excl_vat", "cost_incl_vat",
                          "signups", "cpa", "conversions", "cvr", "revenue", "roas", "aov"]
        day_numeric_cols = [c for c in day_show_cols if c != "report_date"]
        render_cumulative_table(
            add_kpis(daily) if not daily.empty else daily,
            date_col="report_date", show_cols=day_show_cols, numeric_cols=day_numeric_cols,
            title="3) 일자별 누적", key="daily_cum", mode="day",
        )

    # ── 매체별 성과 ──────────────────────────────
    with tab2:
        if not channels.empty:
            channels["report_month"] = pd.to_datetime(channels["report_month"]).dt.date
            st.subheader("🔎 기간 필터 (월별 기준)")
            min_m = channels["report_month"].min()
            max_m = (pd.Timestamp(channels["report_month"].max()) + pd.offsets.MonthEnd(0)).date()
            mstart, mend = period_filter(min_m, max_m, key="channel")
            fc = channels[(channels["report_month"] >= mstart) & (channels["report_month"] <= mend)]

            by_channel = (
                fc.groupby("channel", as_index=False)
                .agg(impressions=("impressions", "sum"), clicks=("clicks", "sum"),
                     cost_excl_vat=("cost_excl_vat", "sum"), cost_incl_vat=("cost_incl_vat", "sum"),
                     conversions=("conversions", "sum"), revenue=("revenue", "sum"))
            )
            by_channel = add_kpis(by_channel).sort_values("cost_incl_vat", ascending=False)

            fig = px.bar(
                by_channel, x="channel", y="roas", title="매체별 ROAS (%, 선택 기간 합산)", text_auto=".1f",
                labels={"channel": "매체", "roas": "ROAS(%)"},
            )
            st.plotly_chart(theme_chart(fig), use_container_width=True)

            bc_cols = list(by_channel.columns)
            bc_table = format_display(by_channel[bc_cols])
            bc_total = build_total_row(by_channel[bc_cols], bc_cols, "channel", label_text="TOTAL")
            if bc_total:
                bc_table = pd.concat([bc_table, pd.DataFrame([bc_total])], ignore_index=True)
            render_html_table(korify(bc_table))
            st.download_button("⬇️ 엑셀 다운로드 (매체별·월별)", data=to_excel_bytes(korify(format_display(by_channel))), file_name="channel_performance.xlsx")
        else:
            st.info("매체별 데이터가 아직 없습니다.")

        if not snapshot.empty:
            st.markdown("---")
            st.markdown("### 당월 매체별 GA 비교 (최신 스냅샷)")
            latest_month = snapshot["as_of_month"].max()
            snap_latest = add_kpis(snapshot[snapshot["as_of_month"] == latest_month])
            st.caption(f"기준월: {latest_month}")
            cols = ["channel", "impressions", "clicks", "cost_incl_vat", "conversions", "revenue", "roas", "ga_conversions", "ga_revenue", "ga_roas"]
            cols = [c for c in cols if c in snap_latest.columns]
            st.dataframe(korify(format_display(snap_latest[cols].sort_values("cost_incl_vat", ascending=False))), use_container_width=True, hide_index=True)

    # ── GA 유입경로 ──────────────────────────────
    with tab3:
        if not ga.empty:
            ga["as_of_date"] = pd.to_datetime(ga["as_of_date"]).dt.date
            st.subheader("🔎 기간 필터")
            min_g, max_g = ga["as_of_date"].min(), ga["as_of_date"].max()
            gstart, gend = period_filter(min_g, max_g, key="ga")
            g_in_range = ga[(ga["as_of_date"] >= gstart) & (ga["as_of_date"] <= gend)]
            if g_in_range.empty:
                st.info("선택한 기간에 해당하는 GA 스냅샷이 없습니다.")
            else:
                latest = g_in_range["as_of_date"].max()
                g = g_in_range[g_in_range["as_of_date"] == latest].sort_values("revenue", ascending=False)
                st.caption(f"기준일: {latest} (선택 기간 내 가장 최신 업로드 스냅샷)")
                st.dataframe(korify(format_display(g)), use_container_width=True, hide_index=True)
                st.download_button("⬇️ 엑셀 다운로드 (GA 유입경로)", data=to_excel_bytes(korify(format_display(g))), file_name="ga_source.xlsx")
        else:
            st.info("GA 유입경로 데이터가 아직 없습니다.")

    # ── GA4 라이브 리포트 (Looker Studio) ──────────────────
    with tab4:
        looker_view_url = (
            "https://lookerstudio.google.com/u/0/reporting/"
            "7177b0a5-7d7e-4f07-af76-17f2436b317e/page/p_bbwwb7lo4c"
        )
        looker_embed_url = (
            "https://lookerstudio.google.com/embed/reporting/"
            "7177b0a5-7d7e-4f07-af76-17f2436b317e/page/p_bbwwb7lo4c"
        )
        st.markdown("### 구글 애널리틱스(GA4) 라이브 리포트")
        st.caption(
            "대행사가 만든 리포트라 일반 공개(링크가 있는 모든 사용자)로 바꾸기 어려우면, "
            "대시보드 안에 그대로 넣는(임베드) 대신 아래 버튼으로 본인 구글 계정 권한으로 새 창에서 열어 보세요."
        )
        st.link_button("📊 GA4 리포트 새 창에서 열기", looker_view_url, use_container_width=True)

        with st.expander("대시보드 안에 직접 띄워보기 (권한 있으면 아래에 표시됨)"):
            st.caption(
                "대행사에게 Looker Studio에서 파일 > 삽입 보고서(Embed report)를 켜달라고 요청하면 "
                "이 안에 화면이 그대로 뜹니다. 권한이 없으면 로그인 요청이나 빈 화면이 보일 수 있어요."
            )
            st.markdown(
                f'<iframe src="{looker_embed_url}" width="100%" height="900" '
                f'style="border:0" allowfullscreen></iframe>',
                unsafe_allow_html=True,
            )


if __name__ == "__main__":
    main()
