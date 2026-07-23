"""
온라인팀 광고/마케팅 성과 대시보드
==================================
매주 광고 매체(구글/메타/네이버/카카오 등) 성과 리포트를 엑셀/CSV로 업로드하면
Supabase(Postgres)에 누적 저장하고, ROAS/KPI를 한눈에 볼 수 있는 웹 대시보드.

실행:
    streamlit run app.py

배포:
    README.md 참고 (GitHub + Supabase + Streamlit Community Cloud)
"""

import io
from datetime import date, datetime, timedelta

import pandas as pd
import numpy as np
import plotly.express as px
import streamlit as st

# ──────────────────────────────────────────────────────────────
# 기본 설정
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="온라인팀 광고성과 대시보드",
    page_icon="📊",
    layout="wide",
)

STANDARD_COLUMNS = {
    "report_date": "날짜",
    "channel": "매체",
    "campaign": "캠페인",
    "ad_group": "소재/광고그룹",
    "impressions": "노출수",
    "clicks": "클릭수",
    "cost": "비용",
    "conversions": "전환수",
    "revenue": "전환매출",
}
NUMERIC_COLS = ["impressions", "clicks", "cost", "conversions", "revenue"]
TABLE_NAME = "ad_performance"


# ──────────────────────────────────────────────────────────────
# Supabase 연결 (secrets.toml 에 SUPABASE_URL / SUPABASE_KEY 필요)
# 설정이 없으면 로컬 세션 메모리로 동작 (테스트용, 새로고침 시 초기화됨)
# ──────────────────────────────────────────────────────────────
@st.cache_resource
def get_supabase_client():
    try:
        from supabase import create_client

        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except Exception:
        return None


def _empty_df():
    return pd.DataFrame(columns=list(STANDARD_COLUMNS.keys()) + ["source_file", "uploaded_at"])


@st.cache_data(ttl=60, show_spinner=False)
def load_db() -> pd.DataFrame:
    client = get_supabase_client()
    if client is None:
        return st.session_state.get("local_db", _empty_df())

    rows, page, page_size = [], 0, 1000
    while True:
        resp = (
            client.table(TABLE_NAME)
            .select("*")
            .range(page * page_size, page * page_size + page_size - 1)
            .execute()
        )
        chunk = resp.data or []
        rows.extend(chunk)
        if len(chunk) < page_size:
            break
        page += 1

    if not rows:
        return _empty_df()
    df = pd.DataFrame(rows)
    df["report_date"] = pd.to_datetime(df["report_date"]).dt.date
    return df


def save_rows(df: pd.DataFrame, source_file: str):
    """업로드된 데이터를 DB에 저장 (동일 날짜/매체/캠페인/소재는 덮어쓰기)."""
    df = df.copy()
    df["source_file"] = source_file
    df["uploaded_at"] = datetime.utcnow().isoformat()
    df["report_date"] = pd.to_datetime(df["report_date"]).dt.date.astype(str)

    client = get_supabase_client()
    if client is None:
        local = st.session_state.get("local_db", _empty_df())
        local = pd.concat([local, df], ignore_index=True)
        local["report_date"] = pd.to_datetime(local["report_date"]).dt.date
        local = local.drop_duplicates(
            subset=["report_date", "channel", "campaign", "ad_group"], keep="last"
        )
        st.session_state["local_db"] = local
        return len(df)

    records = df.to_dict(orient="records")
    client.table(TABLE_NAME).upsert(
        records, on_conflict="report_date,channel,campaign,ad_group"
    ).execute()
    return len(df)


# ──────────────────────────────────────────────────────────────
# 업로드 파일 파싱 + 컬럼 매핑
# ──────────────────────────────────────────────────────────────
def read_uploaded_file(file) -> pd.DataFrame:
    if file.name.lower().endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)


def guess_column(columns, keywords):
    for col in columns:
        low = str(col).replace(" ", "").lower()
        for kw in keywords:
            if kw in low:
                return col
    return None


GUESS_KEYWORDS = {
    "report_date": ["날짜", "일자", "date", "기간"],
    "channel": ["매체", "채널", "channel", "media", "플랫폼"],
    "campaign": ["캠페인", "campaign"],
    "ad_group": ["소재", "광고그룹", "adgroup", "그룹", "adset"],
    "impressions": ["노출", "impression"],
    "clicks": ["클릭", "click"],
    "cost": ["비용", "지출", "spend", "cost", "광고비"],
    "conversions": ["전환수", "구매수", "conversion", "conv"],
    "revenue": ["전환매출", "매출", "revenue", "sales"],
}


def render_upload_panel():
    st.sidebar.header("⚙️ 데이터 관리")
    client = get_supabase_client()
    st.sidebar.caption(f"저장소: {'Supabase (Postgres)' if client else '로컬 세션 (테스트용)'}")

    files = st.sidebar.file_uploader(
        "① 주간 리포트 업로드 (여러 개 한 번에 가능)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
    )

    if files:
        for f in files:
            with st.sidebar.expander(f"📄 {f.name} — 컬럼 매핑", expanded=True):
                raw = read_uploaded_file(f)
                st.dataframe(raw.head(3), use_container_width=True, height=120)

                mapping = {}
                cols = list(raw.columns)
                for key, label in STANDARD_COLUMNS.items():
                    default = guess_column(cols, GUESS_KEYWORDS[key])
                    default_idx = (cols.index(default) + 1) if default in cols else 0
                    mapping[key] = st.selectbox(
                        label,
                        options=["(없음)"] + cols,
                        index=default_idx,
                        key=f"{f.name}_{key}",
                    )

                if st.button(f"'{f.name}' 저장하기", key=f"save_{f.name}"):
                    clean = pd.DataFrame(index=raw.index)
                    for key, chosen in mapping.items():
                        if chosen == "(없음)":
                            clean[key] = np.nan if key in NUMERIC_COLS else ""
                        else:
                            clean[key] = raw[chosen]

                    for c in NUMERIC_COLS:
                        clean[c] = pd.to_numeric(clean[c], errors="coerce").fillna(0)

                    clean["report_date"] = pd.to_datetime(
                        clean["report_date"], errors="coerce"
                    ).dt.date
                    clean = clean.dropna(subset=["report_date"])

                    n = save_rows(clean, f.name)
                    st.cache_data.clear()
                    st.success(f"{n:,}건 저장 완료!")
                    st.rerun()

    st.sidebar.markdown("---")
    db = load_db()
    st.sidebar.metric("현재 DB 누적", f"{len(db):,} 건")
    if st.sidebar.button("🔄 새로고침 (캐시 비우기)"):
        st.cache_data.clear()
        st.rerun()


# ──────────────────────────────────────────────────────────────
# KPI 계산
# ──────────────────────────────────────────────────────────────
def add_kpis(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["ctr"] = np.where(df["impressions"] > 0, df["clicks"] / df["impressions"] * 100, 0)
    df["cpc"] = np.where(df["clicks"] > 0, df["cost"] / df["clicks"], 0)
    df["cvr"] = np.where(df["clicks"] > 0, df["conversions"] / df["clicks"] * 100, 0)
    df["cpa"] = np.where(df["conversions"] > 0, df["cost"] / df["conversions"], 0)
    df["roas"] = np.where(df["cost"] > 0, df["revenue"] / df["cost"] * 100, 0)
    return df


def kpi_cards(df: pd.DataFrame):
    total_cost = df["cost"].sum()
    total_revenue = df["revenue"].sum()
    total_conv = df["conversions"].sum()
    roas = (total_revenue / total_cost * 100) if total_cost > 0 else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("총 광고비", f"{total_cost:,.0f} 원")
    c2.metric("총 전환매출", f"{total_revenue:,.0f} 원")
    c3.metric("총 전환수", f"{total_conv:,.0f} 건")
    c4.metric("평균 ROAS", f"{roas:,.1f} %")


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────
# 필터
# ──────────────────────────────────────────────────────────────
def render_filters(db: pd.DataFrame):
    st.subheader("🔎 필터")
    f1, f2, f3 = st.columns([2, 2, 3])

    min_d = db["report_date"].min() if len(db) else date.today() - timedelta(days=30)
    max_d = db["report_date"].max() if len(db) else date.today()

    with f1:
        date_range = st.date_input("기간", value=(min_d, max_d), min_value=min_d, max_value=max_d)
    with f2:
        channels = sorted(db["channel"].dropna().unique().tolist())
        sel_channels = st.multiselect("매체", options=channels, default=channels)
    with f3:
        campaigns = sorted(db["campaign"].dropna().unique().tolist())
        sel_campaigns = st.multiselect("캠페인 (미선택 시 전체)", options=campaigns, default=[])

    if isinstance(date_range, tuple) and len(date_range) == 2:
        start, end = date_range
    else:
        start, end = min_d, max_d

    filtered = db[(db["report_date"] >= start) & (db["report_date"] <= end)]
    if sel_channels:
        filtered = filtered[filtered["channel"].isin(sel_channels)]
    if sel_campaigns:
        filtered = filtered[filtered["campaign"].isin(sel_campaigns)]

    return filtered


# ──────────────────────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────────────────────
def main():
    st.title("📊 온라인팀 광고/마케팅 성과 대시보드")

    render_upload_panel()
    db = load_db()

    if db.empty:
        st.info("아직 업로드된 데이터가 없습니다. 왼쪽 사이드바에서 주간 리포트 파일을 업로드해주세요.")
        return

    filtered = add_kpis(render_filters(db))

    tab1, tab2, tab3 = st.tabs(["종합 대시보드", "매체별 성과", "캠페인별 상세"])

    # ── 종합 대시보드 ──────────────────────────────
    with tab1:
        kpi_cards(filtered)
        st.markdown("### 주간 추이")
        weekly = (
            filtered.assign(week=pd.to_datetime(filtered["report_date"]).dt.to_period("W").dt.start_time)
            .groupby("week", as_index=False)
            .agg(cost=("cost", "sum"), revenue=("revenue", "sum"), conversions=("conversions", "sum"))
        )
        weekly["roas"] = np.where(weekly["cost"] > 0, weekly["revenue"] / weekly["cost"] * 100, 0)

        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(weekly, x="week", y=["cost", "revenue"], barmode="group", title="주간 비용 vs 매출")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig2 = px.line(weekly, x="week", y="roas", markers=True, title="주간 ROAS 추이 (%)")
            st.plotly_chart(fig2, use_container_width=True)

    # ── 매체별 성과 ──────────────────────────────
    with tab2:
        by_channel = (
            filtered.groupby("channel", as_index=False)
            .agg(
                cost=("cost", "sum"),
                revenue=("revenue", "sum"),
                impressions=("impressions", "sum"),
                clicks=("clicks", "sum"),
                conversions=("conversions", "sum"),
            )
        )
        by_channel = add_kpis(by_channel).sort_values("cost", ascending=False)

        fig3 = px.bar(by_channel, x="channel", y="roas", title="매체별 ROAS (%)", text_auto=".1f")
        st.plotly_chart(fig3, use_container_width=True)

        st.dataframe(
            by_channel.rename(columns={**STANDARD_COLUMNS, "ctr": "CTR(%)", "cpc": "CPC", "cvr": "CVR(%)", "cpa": "CPA", "roas": "ROAS(%)"}),
            use_container_width=True,
        )
        st.download_button(
            "⬇️ 엑셀 다운로드 (매체별)",
            data=to_excel_bytes(by_channel),
            file_name="channel_performance.xlsx",
        )

    # ── 캠페인별 상세 ──────────────────────────────
    with tab3:
        by_campaign = (
            filtered.groupby(["channel", "campaign"], as_index=False)
            .agg(
                cost=("cost", "sum"),
                revenue=("revenue", "sum"),
                impressions=("impressions", "sum"),
                clicks=("clicks", "sum"),
                conversions=("conversions", "sum"),
            )
        )
        by_campaign = add_kpis(by_campaign).sort_values("cost", ascending=False)
        st.dataframe(
            by_campaign.rename(columns={**STANDARD_COLUMNS, "ctr": "CTR(%)", "cpc": "CPC", "cvr": "CVR(%)", "cpa": "CPA", "roas": "ROAS(%)"}),
            use_container_width=True,
        )
        st.download_button(
            "⬇️ 엑셀 다운로드 (캠페인별)",
            data=to_excel_bytes(by_campaign),
            file_name="campaign_performance.xlsx",
        )


if __name__ == "__main__":
    main()
