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


PERIOD_PRESETS = ["이번달", "전체", "최근 1주", "최근 2주", "최근 1개월", "최근 3개월", "최근 6개월", "최근 1년", "직접 선택"]
PERIOD_DAYS = {
    "최근 1주": 7,
    "최근 2주": 14,
    "최근 1개월": 30,
    "최근 3개월": 90,
    "최근 6개월": 180,
    "최근 1년": 365,
}


def period_filter(min_d: date, max_d: date, key: str):
    """Streamlit 기본 날짜범위 위젯의 영어 프리셋("Past Week" 등) 대신
    한글 프리셋 선택 UI로 기간을 고른다. 기본값은 '이번달' (1일 ~ 최신 데이터)."""
    preset = st.selectbox("기간 선택", PERIOD_PRESETS, index=0, key=f"{key}_preset")
    if preset == "이번달":
        month_start = date.today().replace(day=1)
        start = max(min_d, month_start)
        if start > max_d:
            start = min_d
        return start, max_d
    if preset == "전체":
        return min_d, max_d
    if preset in PERIOD_DAYS:
        return max(min_d, max_d - timedelta(days=PERIOD_DAYS[preset])), max_d
    # 직접 선택
    date_range = st.date_input("기간 직접 선택", value=(min_d, max_d), min_value=min_d, max_value=max_d, key=f"{key}_manual")
    if isinstance(date_range, tuple) and len(date_range) == 2:
        return date_range
    return min_d, max_d


PAGE_SIZE_OPTIONS = [10, 30, 50, 100, 200]


def pct_change_row(df_sorted: pd.DataFrame, numeric_cols: list, label_col: str, label_text: str = "전기간 대비"):
    """가장 최근 값 대비 그 이전 값의 증감율을 계산해 표 맨 아래에 붙일 한 줄을 만든다."""
    if len(df_sorted) < 2:
        return None
    latest, prev = df_sorted.iloc[-1], df_sorted.iloc[-2]
    row = {label_col: label_text}
    for c in numeric_cols:
        if c not in df_sorted.columns:
            continue
        pv, lv = prev.get(c), latest.get(c)
        if pd.isna(pv) or pv in (0, None):
            row[c] = "-"
            continue
        change = (lv - pv) / abs(pv) * 100
        arrow = "▲" if change >= 0 else "▼"
        row[c] = f"{arrow}{change:+.1f}%"
    return row


def render_cumulative_table(df: pd.DataFrame, date_col: str, show_cols: list, numeric_cols: list,
                             default_n: int, title: str, key: str, month_default: bool = False):
    """월별/주간/일자별 누적 표. 기본은 최근 N개(또는 이번 달)만 보여주고,
    '이전 데이터 더 보기'를 켜면 10/30/50/100/200개 단위 페이지네이션으로 전체를 볼 수 있다."""
    st.markdown(f"#### {title}")
    if df is None or df.empty:
        st.caption("데이터가 아직 없습니다.")
        return

    d = df.sort_values(date_col).reset_index(drop=True)
    total = len(d)
    more = st.checkbox(f"이전 데이터 더 보기 (전체 {total}건)", key=f"{key}_more")

    if not more:
        if month_default:
            cur_month = date.today().replace(day=1)
            view = d[pd.to_datetime(d[date_col]) >= pd.Timestamp(cur_month)]
            if view.empty:
                view = d.tail(default_n)
        else:
            view = d.tail(default_n)
        show_change_row = True
    else:
        page_size = st.selectbox("페이지당 표시", PAGE_SIZE_OPTIONS, index=2, key=f"{key}_pagesize")
        max_page = max(1, -(-total // page_size))
        page = st.number_input(
            f"페이지 (1~{max_page}, 클수록 최신)", min_value=1, max_value=max_page, value=max_page, step=1, key=f"{key}_page"
        )
        start_i, end_i = (page - 1) * page_size, page * page_size
        view = d.iloc[start_i:end_i]
        show_change_row = page == max_page

    display_cols = [c for c in show_cols if c in view.columns]
    table = view[display_cols].copy()

    if show_change_row:
        change_row = pct_change_row(d, numeric_cols, display_cols[0])
        if change_row:
            table = pd.concat([table, pd.DataFrame([change_row])], ignore_index=True)

    st.dataframe(korify(table), use_container_width=True)
    st.download_button(
        f"⬇️ 엑셀 다운로드 ({title})",
        data=to_excel_bytes(korify(view[display_cols])),
        file_name=f"{key}.xlsx",
        key=f"{key}_dl",
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
                fig = px.bar(fw, x="week_start", y=["cost_incl_vat", "revenue"], barmode="group", title="주간 비용(VAT포함) vs 매출")
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig2 = px.line(fw, x="week_start", y="roas", markers=True, title="주간 ROAS 추이 (%)")
                st.plotly_chart(fig2, use_container_width=True)

        else:
            st.info("주간 데이터가 아직 없습니다.")

        if not monthly.empty:
            st.markdown("---")
            st.markdown("### 월별 GA-ROAS vs 플랫폼 ROAS")
            fm_chart = add_kpis(monthly).sort_values("report_month")
            fig3 = px.line(fm_chart, x="report_month", y=["roas", "ga_roas"], markers=True,
                            labels={"value": "%", "variable": "기준"}, title="플랫폼 리포팅 ROAS vs GA 기준 ROAS")
            st.plotly_chart(fig3, use_container_width=True)
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
            default_n=12, title="1) 월별 누적 (최근 12개월 기본)", key="monthly_cum",
        )

        week_show_cols = ["label", "week_start", "week_end", "impressions", "clicks", "ctr", "cpc",
                           "cost_excl_vat", "cost_incl_vat", "signups", "cpa", "conversions", "cvr", "revenue", "roas", "aov"]
        week_numeric_cols = [c for c in week_show_cols if c not in ("label", "week_start", "week_end")]
        render_cumulative_table(
            add_kpis(weekly) if not weekly.empty else weekly,
            date_col="week_start", show_cols=week_show_cols, numeric_cols=week_numeric_cols,
            default_n=5, title="2) 주간별 누적 (최근 5주 기본)", key="weekly_cum",
        )

        day_show_cols = ["report_date", "impressions", "clicks", "ctr", "cpc", "cost_excl_vat", "cost_incl_vat",
                          "signups", "cpa", "conversions", "cvr", "revenue", "roas", "aov"]
        day_numeric_cols = [c for c in day_show_cols if c != "report_date"]
        render_cumulative_table(
            add_kpis(daily) if not daily.empty else daily,
            date_col="report_date", show_cols=day_show_cols, numeric_cols=day_numeric_cols,
            default_n=31, title="3) 일자별 누적 (이번 달 기본)", key="daily_cum", month_default=True,
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

            fig = px.bar(by_channel, x="channel", y="roas", title="매체별 ROAS (%, 선택 기간 합산)", text_auto=".1f")
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(korify(by_channel), use_container_width=True)
            st.download_button("⬇️ 엑셀 다운로드 (매체별·월별)", data=to_excel_bytes(korify(by_channel)), file_name="channel_performance.xlsx")
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
            st.dataframe(korify(snap_latest[cols].sort_values("cost_incl_vat", ascending=False)), use_container_width=True)

    # ── GA 유입경로 ──────────────────────────────
    with tab3:
        if not ga.empty:
            ga["as_of_date"] = pd.to_datetime(ga["as_of_date"]).dt.date
            latest = ga["as_of_date"].max()
            g = ga[ga["as_of_date"] == latest].sort_values("revenue", ascending=False)
            st.caption(f"기준일: {latest} (마지막 업로드 시점 스냅샷)")
            st.dataframe(korify(g), use_container_width=True)
            st.download_button("⬇️ 엑셀 다운로드 (GA 유입경로)", data=to_excel_bytes(korify(g)), file_name="ga_source.xlsx")
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
