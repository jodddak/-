"""
STCO 성과 메일 — 전일(일별) / 전주(주간) 성과와 운영 제안을 메일로 보낸다.

  · 화~금 아침 10:00  →  어제 하루 성과 (전주 같은 요일과 비교)
  · 월요일   오후 13:00 →  지난주 월~일 성과 (그 전주와 비교)

Streamlit Community Cloud는 정해진 시각에 코드를 돌리는 기능(크론)이 없어서
누가 페이지를 열어야만 동작한다. 그래서 이 스크립트는 대시보드와 별개로
GitHub Actions에서 돈다(.github/workflows/daily_report.yml).

읽는 곳은 대시보드와 같은 Supabase다. 판정 기준(KPI 200~300%, 표본 문턱,
광고비 출처 우선순위)도 대시보드와 똑같이 맞춰서, 메일에서 '증액 검토'라고 하면
화면에서도 같은 판정이 나오게 했다.

로컬에서 시험 발송:
    python daily_report.py --dry-run                     # 오늘 요일에 맞춰 자동 판단
    python daily_report.py --mode daily  --date 2026-09-15
    python daily_report.py --mode weekly --date 2026-09-14   # 그 날이 속한 주(월~일)
"""

from __future__ import annotations

import argparse
import os
import smtplib
import sys
from datetime import date, datetime, timedelta
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr

import pandas as pd

# ──────────────────────────────────────────────────────────────
# 설정 — 전부 환경변수(GitHub Secrets)에서 읽는다. 코드에 비밀값을 적지 않는다.
# ──────────────────────────────────────────────────────────────
def env(name: str, default: str = "") -> str:
    """환경변수를 읽되 '빈 값'도 없는 것으로 친다.

    GitHub Actions는 Secret을 안 넣어도 변수 자체는 **빈 문자열로** 만들어 넘긴다.
    그래서 os.environ.get(name, 기본값) 을 쓰면 기본값이 안 먹고 빈 문자열이 온다
    (int("") 에서 터졌던 원인). 앞뒤 공백과 따옴표도 같이 털어낸다 —
    Streamlit Secrets에서 복사하면 따옴표가 딸려오는 일이 흔하다.
    """
    v = os.environ.get(name, "")
    v = str(v).strip().strip('"').strip("'").strip()
    return v if v else default


def env_int(name: str, default: int) -> int:
    v = env(name)
    if not v:                 # 안 넣은 값 — 기본값을 쓰는 게 정상이라 조용히 넘어간다
        return default
    try:
        return int(v)
    except ValueError:
        print(f"[경고] {name} 값이 숫자가 아닙니다({v!r}). {default} 로 진행합니다.", file=sys.stderr)
        return default


SUPABASE_URL = env("SUPABASE_URL")
SUPABASE_KEY = env("SUPABASE_KEY")

# 보내는 계정은 Gmail을 기본으로 둔다.
# 사내 SMTP는 보통 외부(GitHub) 서버에서의 발송을 막아둬서 못 쓴다.
# '어디서 보내느냐'와 '어디로 받느냐'는 별개다 — Gmail로 보내고 회사 주소로 받으면 된다.
# 사내 서버를 쓸 수 있게 되면 SMTP_HOST만 채우면 그대로 동작한다.
SMTP_HOST = env("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = env_int("SMTP_PORT", 587)
SMTP_USER = env("SMTP_USER")
SMTP_PASS = env("SMTP_PASS").replace(" ", "")   # 앱 비밀번호는 4자리씩 띄어서 보여준다
SMTP_SECURITY = env("SMTP_SECURITY", "starttls").lower()  # starttls | ssl | none

MAIL_FROM = env("MAIL_FROM", SMTP_USER)
MAIL_FROM_NAME = env("MAIL_FROM_NAME", "STCO 성과 대시보드")
MAIL_TO = [a.strip() for a in env("MAIL_TO").split(",") if a.strip()]
MAIL_CC = [a.strip() for a in env("MAIL_CC").split(",") if a.strip()]

DASHBOARD_URL = env("DASHBOARD_URL")

# ── 판정 기준 (app.py와 같은 값) ───────────────────────────────
KPI_ROAS_LOW = 200        # 목표 하단(%)
KPI_ROAS_HIGH = 300       # 목표 상단(%)
MIN_SPEND_PER_DAY = 100_000   # 하루 기준 표본 문턱. 주간은 일수만큼 곱해서 쓴다.
STEP = 0.10               # 증액·감액은 한 번에 10%씩 (급격한 이동 금지)

# 광고비 출처 우선순위 — 같은 날 여러 출처가 있으면 낮은 숫자가 이긴다.
SPEND_SOURCE_PRIO = {
    "meta_api": 0, "google_ads_api": 0, "naver_api": 0, "kakao_api": 0,
    "criteo_api": 0, "naver_gfa_api": 0, "kakao_msg": 0, "kakao_cash": 0,
    "contract": 1, "manual": 2, "agency_weekly": 3, "budget_prorate": 4,
}

PAGE = 1000               # Supabase(PostgREST)는 응답당 1000행으로 자른다
WD = "월화수목금토일"


# ──────────────────────────────────────────────────────────────
# Supabase 읽기
# ──────────────────────────────────────────────────────────────
def get_client():
    missing = [n for n, v in (("SUPABASE_URL", SUPABASE_URL), ("SUPABASE_KEY", SUPABASE_KEY)) if not v]
    if missing:
        raise SystemExit(
            f"{' / '.join(missing)} 가 비어 있습니다.\n"
            "GitHub → Settings → Secrets and variables → Actions 에서 값이 들어 있는지 확인해주세요.")
    from supabase import create_client
    return create_client(SUPABASE_URL, SUPABASE_KEY)


def load_table(client, name: str, since: date | None = None,
               order_cols: list[str] | None = None) -> pd.DataFrame:
    """표 하나를 통째로 읽는다. 1000행씩 나눠 받되 기본키로 정렬해 경계를 고정한다
    (정렬을 안 하면 페이지마다 순서가 달라져 같은 행이 두 번 오거나 빠질 수 있다)."""
    rows, page = [], 0
    while True:
        q = client.table(name).select("*")
        if since is not None:
            q = q.gte("report_date", since.isoformat())
        for c in (order_cols or []):
            q = q.order(c)
        try:
            resp = q.range(page * PAGE, page * PAGE + PAGE - 1).execute()
        except Exception as e:
            print(f"[경고] '{name}' 조회 실패: {e}", file=sys.stderr)
            break
        chunk = resp.data or []
        rows.extend(chunk)
        if len(chunk) < PAGE:
            break
        page += 1
    return pd.DataFrame(rows)


# ──────────────────────────────────────────────────────────────
# 집계 — 대시보드 '채널 성과' 탭과 같은 방식
# ──────────────────────────────────────────────────────────────
def spend_by_channel(ad_spend: pd.DataFrame, start: date, end: date) -> pd.DataFrame:
    """기간 매체별 노출·클릭·광고비.

    **날짜마다** 출처를 따로 고른 뒤에 기간으로 합친다. 이 순서가 중요하다.
    기간 전체에서 출처를 한 번만 고르면, 예컨대 8월치만 있는 메타 API 행이
    1~7월 예산 행을 이겨버려서 앞쪽 광고비가 통째로 날아간다.

    지표마다도 출처를 따로 고른다. 브랜드검색이 대표적인 이유인데, 정액 상품이라
    API가 노출·클릭은 주지만 집행액은 안 준다. 행 단위로 출처를 하나만 고르면
    API 행이 뽑히면서 계약 광고비가 사라진다.
    """
    cols = ["channel", "impressions", "clicks", "cost_incl_vat"]
    if ad_spend is None or ad_spend.empty:
        return pd.DataFrame(columns=cols)
    a = ad_spend.copy()
    a["report_date"] = pd.to_datetime(a["report_date"], errors="coerce").dt.date
    a = a[(a["report_date"] >= start) & (a["report_date"] <= end)]
    if a.empty:
        return pd.DataFrame(columns=cols)
    for c in ("impressions", "clicks", "cost_incl_vat"):
        a[c] = pd.to_numeric(a.get(c), errors="coerce").fillna(0.0)
    a["source"] = a["source"].astype(str) if "source" in a.columns else ""
    a["_p"] = a["source"].map(SPEND_SOURCE_PRIO).fillna(9)
    a = a.groupby(["report_date", "channel", "source", "_p"], as_index=False).agg(
        impressions=("impressions", "sum"), clicks=("clicks", "sum"),
        cost_incl_vat=("cost_incl_vat", "sum"))
    a = a.sort_values("_p")

    picked = []
    for (d, ch), sub in a.groupby(["report_date", "channel"]):
        rec = {"report_date": d, "channel": ch}
        for col in ("impressions", "clicks", "cost_incl_vat"):
            hit = sub[sub[col] > 0]
            rec[col] = float(hit.iloc[0][col]) if not hit.empty else 0.0
        picked.append(rec)
    if not picked:
        return pd.DataFrame(columns=cols)
    p = pd.DataFrame(picked)
    return p.groupby("channel", as_index=False)[
        ["impressions", "clicks", "cost_incl_vat"]].sum()


def ga_by_media(ga_daily: pd.DataFrame, master: pd.DataFrame,
                start: date, end: date) -> dict:
    """GA4 유입(소스/매체 단위)을 매체 정의의 utm_match로 묶는다.

    완전일치를 먼저 보고 안 걸리면 부분일치로 간다. 순서가 중요한데,
    'naver/gfa_애드부스트'는 'naver/gfa'를 문자열로 포함하기 때문이다.
    """
    out = {m: {"conv": 0.0, "rev": 0.0} for m in master["media"]} if not master.empty else {}
    out["_미매칭"] = {"conv": 0.0, "rev": 0.0}
    if ga_daily is None or ga_daily.empty or "source_medium" not in ga_daily.columns:
        return out
    g = ga_daily.copy()
    g["report_date"] = pd.to_datetime(g["report_date"], errors="coerce").dt.date
    g = g[(g["report_date"] >= start) & (g["report_date"] <= end)]
    if g.empty:
        return out

    def norm(s):
        return "".join(str(s or "").lower().split())

    exact, partial = {}, []
    if not master.empty:
        for _, m in master.sort_values("sort_order").iterrows():
            for kw in str(m.get("utm_match") or "").split(","):
                kw = norm(kw)
                if kw:
                    exact.setdefault(kw, m["media"])
                    partial.append((len(kw), kw, m["media"]))
    partial.sort(key=lambda x: -x[0])

    for _, row in g.iterrows():
        sm = norm(row.get("source_medium"))
        conv = float(pd.to_numeric(row.get("conversions"), errors="coerce") or 0)
        rev = float(pd.to_numeric(row.get("revenue"), errors="coerce") or 0)
        media = exact.get(sm)
        if media is None:
            for _, kw, mm in partial:
                if kw in sm:
                    media = mm
                    break
        key = media if media else "_미매칭"
        out.setdefault(key, {"conv": 0.0, "rev": 0.0})
        out[key]["conv"] += conv
        out[key]["rev"] += rev
    return out


def build_rows(ad_spend, ga_daily, master, start: date, end: date) -> pd.DataFrame:
    """매체 한 줄씩 — 노출·클릭·광고비·GA구매·GA매출·ROAS."""
    sp = spend_by_channel(ad_spend, start, end)
    sp_map = {r["channel"]: r for _, r in sp.iterrows()} if not sp.empty else {}
    ga_map = ga_by_media(ga_daily, master, start, end)

    recs = []
    for _, r in master.iterrows():
        sc = str(r.get("spend_channel") or "").strip()
        s = sp_map.get(sc, {})
        g = ga_map.get(r["media"], {"conv": 0.0, "rev": 0.0})
        cost = float(s.get("cost_incl_vat", 0) or 0)
        rev = float(g["rev"])
        recs.append({
            "scope": r.get("scope", "자사몰"),
            "media": r["media"],
            "impressions": float(s.get("impressions", 0) or 0),
            "clicks": float(s.get("clicks", 0) or 0),
            "cost": cost,
            "conv": float(g["conv"]),
            "rev": rev,
            "roas": (rev / cost * 100) if cost > 0 else None,
            "order": int(pd.to_numeric(r.get("sort_order"), errors="coerce") or 100),
        })
    df = pd.DataFrame(recs).sort_values("order")
    return df[(df["cost"] > 0) | (df["rev"] > 0) | (df["impressions"] > 0)]


def totals(df: pd.DataFrame) -> dict:
    if df is None or df.empty:
        return {"cost": 0.0, "rev": 0.0, "conv": 0.0, "imp": 0.0, "clk": 0.0, "roas": 0.0}
    cost = float(df["cost"].sum())
    rev = float(df["rev"].sum())
    return {"cost": cost, "rev": rev, "conv": float(df["conv"].sum()),
            "imp": float(df["impressions"].sum()), "clk": float(df["clicks"].sum()),
            "roas": (rev / cost * 100) if cost > 0 else 0.0}


def daily_series(ad_spend, ga_daily, master, start: date, end: date) -> list[dict]:
    """주간 메일에 넣을 일자별 추이. 어느 날이 끌어올렸고 어느 날이 깎았는지 보려는 것."""
    out = []
    d = start
    while d <= end:
        t = totals(build_rows(ad_spend, ga_daily, master, d, d))
        out.append({"date": d, **t})
        d += timedelta(days=1)
    return out


# ──────────────────────────────────────────────────────────────
# 코멘트 · 운영 제안 — 대시보드 'NEXT BEST ACTION'과 같은 규칙
# ──────────────────────────────────────────────────────────────
def build_actions(df: pd.DataFrame, n_days: int = 1) -> tuple[str, list[dict], float]:
    """(헤드라인, 제안 목록, 적용한 표본 문턱). 판단에서 빼야 할 경우를 먼저 걸러낸다.

    · 외부몰 — 매출이 GA4에 안 잡혀 ROAS가 0으로 보인다. 그대로 읽으면 '줄여라'가
      되는데 실제로는 판단 근거 자체가 없는 것이다.
    · 광고비가 안 들어온 매체 — 분모가 없으니 ROAS를 만들 수 없다.
    · 광고비가 너무 적은 매체 — 표본이 작아 우연이 성과처럼 보인다.
      기간이 길면 문턱도 같이 올린다(주간은 하루 문턱 × 7일). 안 그러면 주간에서는
      거의 모든 매체가 문턱을 넘어버려서 표본 검사가 무력해진다.
    """
    floor = MIN_SPEND_PER_DAY * max(1, n_days)
    floor_txt = (f"{n_days}일 합계 {floor:,.0f}원" if n_days > 1 else f"{floor:,.0f}원")

    up, keep, down, hold, skip = [], [], [], [], []
    for _, r in df.iterrows():
        m, cost, roas, rev = r["media"], r["cost"], r["roas"], r["rev"]
        if r["scope"] == "외부몰":
            skip.append((m, "외부몰 — 매출이 GA에 안 잡혀 판단 불가 (스마트스토어 연동 필요)"))
        elif cost <= 0 and rev > 0:
            skip.append((m, f"광고비 미연동 — 매출 {rev:,.0f}원은 잡히는데 비용이 없어 ROAS 계산 불가"))
        elif cost <= 0:
            continue
        elif cost < floor:
            hold.append((m, cost, roas))
        elif roas is None or pd.isna(roas):
            skip.append((m, "매출 데이터 없음"))
        elif roas >= KPI_ROAS_HIGH:
            up.append((m, cost, roas))
        elif roas >= KPI_ROAS_LOW:
            keep.append((m, cost, roas))
        else:
            down.append((m, cost, roas))

    up.sort(key=lambda x: -x[2])
    down.sort(key=lambda x: x[2])

    if up and down:
        head = f"{up[0][0]}를 늘리고 {down[0][0]}를 줄이세요."
    elif up:
        head = f"{up[0][0]}에 예산을 더 태울 여지가 있습니다."
    elif down:
        head = f"{down[0][0]}를 손봐야 합니다."
    elif keep:
        head = "지금은 손댈 매체가 없습니다."
    else:
        head = "아직 판단할 만한 데이터가 없습니다."

    def move(c):
        """증감 제안 금액은 항상 '하루치'로 말한다. 주간 합계에 10%를 붙여
        '+70만원'이라고 하면 하루 예산을 그만큼 올리라는 말로 읽혀서 위험하다."""
        return c * STEP / max(1, n_days)

    items = []
    for m, c, ro in up:
        items.append({"tag": "증액 검토", "cls": "up", "media": m,
                      "text": f"ROAS {ro:,.0f}% · 집행 {c:,.0f}원 → 예산 +{STEP*100:.0f}% "
                              f"(일 +{move(c):,.0f}원) 테스트 후 48시간 관찰"})
    for m, c, ro in down:
        items.append({"tag": "감액·점검", "cls": "down", "media": m,
                      "text": f"ROAS {ro:,.0f}% · 집행 {c:,.0f}원 → 예산 −{STEP*100:.0f}% "
                              f"(일 −{move(c):,.0f}원) 또는 소재·타겟 점검 먼저"})
    for m, c, ro in keep:
        items.append({"tag": "유지", "cls": "keep", "media": m,
                      "text": f"ROAS {ro:,.0f}% · 집행 {c:,.0f}원 → 현 수준 유지"})
    for m, c, ro in hold:
        rr = "—" if (ro is None or pd.isna(ro)) else f"{ro:,.0f}%"
        items.append({"tag": "판단 보류", "cls": "hold", "media": m,
                      "text": f"집행 {c:,.0f}원으로 표본이 작습니다 (ROAS {rr}). "
                              f"{floor_txt} 넘을 때까지 판단 유보"})
    for m, why in skip:
        items.append({"tag": "판단 제외", "cls": "skip", "media": m, "text": why})
    return head, items, floor


def spend_gap_note(ga_daily, ad_spend, start: date, end: date) -> str:
    """GA는 있는데 매체 API 광고비가 없는 날이면 알려준다.
    이걸 모르고 보면 ROAS가 실제보다 좋게 나온 걸 성과로 오해한다."""
    def in_range(s):
        d = pd.to_datetime(s, errors="coerce").dt.date
        return (d >= start) & (d <= end)

    ga_days = set()
    if ga_daily is not None and not ga_daily.empty:
        m = in_range(ga_daily["report_date"])
        ga_days = set(pd.to_datetime(ga_daily.loc[m, "report_date"], errors="coerce").dt.date)

    spend_days = set()
    if ad_spend is not None and not ad_spend.empty and "source" in ad_spend.columns:
        a = ad_spend.copy()
        a["_d"] = pd.to_datetime(a["report_date"], errors="coerce").dt.date
        a["_c"] = pd.to_numeric(a.get("cost_incl_vat"), errors="coerce").fillna(0)
        ok = in_range(a["report_date"]) & a["source"].astype(str).str.endswith("_api") & (a["_c"] > 0)
        spend_days = set(a.loc[ok, "_d"])

    missing = sorted(ga_days - spend_days)
    if not missing:
        return ""
    if len(missing) == 1 and start == end:
        return ("이 날짜에 매체 API 광고비가 한 건도 없습니다 — ROAS가 실제보다 높게 나옵니다. "
                "대시보드 '채널 성과 → 광고비 다시 받기'로 채워주세요.")
    days = " · ".join(f"{d:%m/%d}" for d in missing)
    return (f"매체 API 광고비가 없는 날이 있습니다 ({days}) — 그만큼 ROAS가 실제보다 높게 나옵니다. "
            "대시보드 '채널 성과 → 광고비 다시 받기'로 채워주세요.")


# ──────────────────────────────────────────────────────────────
# 메일 본문
# ──────────────────────────────────────────────────────────────
TAG_COLOR = {
    "up": ("#E4F6DC", "#2C7A3C"), "down": ("#FDE7E9", "#C0273A"),
    "keep": ("#EAF1FB", "#2C5AA0"), "hold": ("#EFEEE8", "#767668"),
    "skip": ("#F3F1EC", "#8A8A7C"),
}


def won(v):
    return f"₩{v:,.0f}" if v else "—"


def num(v):
    return f"{v:,.0f}" if v else "—"


def delta(cur: float, prev: float, good_up: bool | None = True) -> str:
    """전 기간 대비 증감을 한 줄로.

    good_up=None 은 '좋고 나쁨을 안 따진다'는 뜻이다. 광고비가 그렇다 —
    늘어난 게 잘한 건지 잘못한 건지는 매출을 같이 봐야 알 수 있는데,
    초록색을 칠해두면 '증액=좋음'으로 읽혀서 판단을 흐린다.
    """
    if not prev:
        return '<div style="color:#B9BEC5;font-size:11px;margin-top:4px">비교 기간 없음</div>'
    pct = (cur - prev) / prev * 100
    if abs(pct) < 0.5:
        return '<div style="color:#8a8a7c;font-size:11px;margin-top:4px">전 기간과 같음</div>'
    up = pct > 0
    if good_up is None:
        col = "#5A6070"
    else:
        col = "#2C7A3C" if up == good_up else "#C0273A"
    return (f'<div style="color:{col};font-size:11px;font-weight:700;margin-top:4px">'
            f'{"▲" if up else "▼"} {abs(pct):,.0f}%</div>')


def render_html(weekly: bool, start: date, end: date, df: pd.DataFrame,
                prev: dict, head: str, items: list[dict], gap: str,
                floor: float, series: list[dict] | None) -> str:
    t = totals(df)
    tot_roas = t["roas"]
    status = ("목표 초과 달성" if tot_roas >= KPI_ROAS_HIGH
              else "목표 구간 내" if tot_roas >= KPI_ROAS_LOW else "목표 미달")
    status_color = ("#2C7A3C" if tot_roas >= KPI_ROAS_HIGH
                    else "#8A6714" if tot_roas >= KPI_ROAS_LOW else "#C0273A")

    if weekly:
        kicker = "WEEKLY PERFORMANCE"
        title = (f"{start:%Y년 %m월 %d일}({WD[start.weekday()]}) ~ "
                 f"{end:%m월 %d일}({WD[end.weekday()]}) 성과")
        lead = ("지난주 한 주 집행 결과입니다. 비교 값은 그 전주 같은 기간입니다. "
                "광고비는 매체 실집행값, 구매·매출은 GA4 기준입니다.")
        cmp_label = "전주 대비"
        footer = "STCO 온라인팀 · 내부 의사결정용 · 매주 월요일 자동 발송"
    else:
        kicker = "DAILY PERFORMANCE"
        title = f"{end:%Y년 %m월 %d일} ({WD[end.weekday()]}) 성과"
        lead = ("어제 하루 집행 결과입니다. 비교 값은 전주 같은 요일입니다. "
                "광고비는 매체 실집행값, 구매·매출은 GA4 기준입니다.")
        cmp_label = "전주 동요일 대비"
        footer = "STCO 온라인팀 · 내부 의사결정용 · 평일 아침 자동 발송"

    card = ('background:#fff;border:1px solid #E8E6DC;border-radius:12px;padding:15px 17px')
    kpi = f"""
    <div style="color:#8a8a7c;font-size:11px;margin:0 0 7px">증감은 {cmp_label}입니다</div>
    <table width="100%" cellpadding="0" cellspacing="8" style="margin:0 0 20px">
      <tr>
        <td width="25%" style="{card};border-top:3px solid #3D5AFE">
          <div style="color:#8a8a7c;font-size:12px;margin-bottom:6px">광고비</div>
          <div style="color:#17170f;font-size:20px;font-weight:800">{won(t['cost'])}</div>
          {delta(t['cost'], prev.get('cost', 0), good_up=None)}
        </td>
        <td width="25%" style="{card};border-top:3px solid #7C4DFF">
          <div style="color:#8a8a7c;font-size:12px;margin-bottom:6px">GA 매출</div>
          <div style="color:#17170f;font-size:20px;font-weight:800">{won(t['rev'])}</div>
          {delta(t['rev'], prev.get('rev', 0), good_up=True)}
        </td>
        <td width="25%" style="{card};border-top:3px solid #63C132">
          <div style="color:#8a8a7c;font-size:12px;margin-bottom:6px">GA 구매</div>
          <div style="color:#17170f;font-size:20px;font-weight:800">{num(t['conv'])}건</div>
          {delta(t['conv'], prev.get('conv', 0), good_up=True)}
        </td>
        <td width="25%" style="{card};border-top:3px solid #3D5AFE">
          <div style="color:#8a8a7c;font-size:12px;margin-bottom:6px">GA ROAS</div>
          <div style="color:{status_color};font-size:20px;font-weight:800">{tot_roas:,.0f}%</div>
          {delta(tot_roas, prev.get('roas', 0), good_up=True)}
          <div style="color:#8a8a7c;font-size:11px;margin-top:3px">{status}</div>
        </td>
      </tr>
    </table>"""

    body_rows = []
    for _, r in df.iterrows():
        ro = r["roas"]
        if ro is None or pd.isna(ro):
            ro_txt, ro_col = "—", "#B9BEC5"
        else:
            ro_txt = f"{ro:,.0f}%"
            ro_col = ("#2C7A3C" if ro >= KPI_ROAS_HIGH
                      else "#8A6714" if ro >= KPI_ROAS_LOW else "#C0273A")
        badge = "#4B3FA8" if r["scope"] == "외부몰" else "#14181F"
        body_rows.append(f"""
        <tr>
          <td style="padding:11px 10px;border-bottom:1px solid #F0EFE7">
            <span style="background:{badge};color:#fff;font-size:10.5px;font-weight:700;
                  padding:2px 6px;border-radius:4px">{r['scope']}</span>
            <span style="color:#14181F;font-weight:700;margin-left:6px">{r['media']}</span>
          </td>
          <td align="right" style="padding:11px 10px;border-bottom:1px solid #F0EFE7">{num(r['impressions'])}</td>
          <td align="right" style="padding:11px 10px;border-bottom:1px solid #F0EFE7">{num(r['clicks'])}</td>
          <td align="right" style="padding:11px 10px;border-bottom:1px solid #F0EFE7">{won(r['cost'])}</td>
          <td align="right" style="padding:11px 10px;border-bottom:1px solid #F0EFE7">{num(r['conv'])}</td>
          <td align="right" style="padding:11px 10px;border-bottom:1px solid #F0EFE7">{won(r['rev'])}</td>
          <td align="right" style="padding:11px 10px;border-bottom:1px solid #F0EFE7;
              color:{ro_col};font-weight:700">{ro_txt}</td>
        </tr>""")

    th = ('style="padding:10px;background:#F7F6EF;color:#8a8a7c;font-size:11.5px;'
          'font-weight:700;border-bottom:1px solid #E3E1DC"')
    tdt = 'style="padding:12px 10px;background:#14181F;color:#F3F1EC;font-weight:800"'
    table = f"""
    <table width="100%" cellpadding="0" cellspacing="0"
           style="border-collapse:collapse;font-size:13px;border:1px solid #E8E6DC;border-radius:12px">
      <tr>
        <th align="left" {th}>매체</th><th align="right" {th}>노출</th>
        <th align="right" {th}>클릭</th><th align="right" {th}>광고비</th>
        <th align="right" {th}>GA 구매</th><th align="right" {th}>GA 매출</th>
        <th align="right" {th}>ROAS</th>
      </tr>
      {''.join(body_rows)}
      <tr>
        <td style="padding:12px 10px;background:#14181F;color:#D9F27E;font-weight:800">TOTAL</td>
        <td align="right" {tdt}>{num(t['imp'])}</td>
        <td align="right" {tdt}>{num(t['clk'])}</td>
        <td align="right" {tdt}>{won(t['cost'])}</td>
        <td align="right" {tdt}>{num(t['conv'])}</td>
        <td align="right" {tdt}>{won(t['rev'])}</td>
        <td align="right" style="padding:12px 10px;background:#14181F;color:#D9F27E;font-weight:800">{tot_roas:,.0f}%</td>
      </tr>
    </table>"""

    trend = ""
    if weekly and series:
        best = max((s for s in series if s["cost"] > 0), key=lambda s: s["roas"], default=None)
        worst = min((s for s in series if s["cost"] > 0), key=lambda s: s["roas"], default=None)
        tr = []
        for s in series:
            ro = s["roas"]
            col = ("#2C7A3C" if ro >= KPI_ROAS_HIGH
                   else "#8A6714" if ro >= KPI_ROAS_LOW else "#C0273A") if s["cost"] else "#B9BEC5"
            mark = ""
            if best and s["date"] == best["date"]:
                mark = ' <span style="color:#2C7A3C;font-size:10.5px;font-weight:700">최고</span>'
            elif worst and s["date"] == worst["date"]:
                mark = ' <span style="color:#C0273A;font-size:10.5px;font-weight:700">최저</span>'
            tr.append(f"""
            <tr>
              <td style="padding:9px 10px;border-bottom:1px solid #F0EFE7;color:#14181F;font-weight:700">
                {s['date']:%m/%d}({WD[s['date'].weekday()]}){mark}</td>
              <td align="right" style="padding:9px 10px;border-bottom:1px solid #F0EFE7">{won(s['cost'])}</td>
              <td align="right" style="padding:9px 10px;border-bottom:1px solid #F0EFE7">{num(s['conv'])}</td>
              <td align="right" style="padding:9px 10px;border-bottom:1px solid #F0EFE7">{won(s['rev'])}</td>
              <td align="right" style="padding:9px 10px;border-bottom:1px solid #F0EFE7;
                  color:{col};font-weight:700">{s['roas']:,.0f}%</td>
            </tr>""")
        trend = f"""
        <div style="font-size:14px;font-weight:800;color:#17170f;margin:26px 0 10px">일자별 추이</div>
        <table width="100%" cellpadding="0" cellspacing="0"
               style="border-collapse:collapse;font-size:13px;border:1px solid #E8E6DC;border-radius:12px">
          <tr>
            <th align="left" {th}>날짜</th><th align="right" {th}>광고비</th>
            <th align="right" {th}>GA 구매</th><th align="right" {th}>GA 매출</th>
            <th align="right" {th}>ROAS</th>
          </tr>
          {''.join(tr)}
        </table>"""

    li = []
    for it in items:
        bg, fg = TAG_COLOR.get(it["cls"], TAG_COLOR["skip"])
        li.append(f"""
        <tr><td style="padding:10px 0;border-top:1px solid rgba(0,0,0,.08)">
          <span style="background:{bg};color:{fg};font-size:11.5px;font-weight:700;
                padding:3px 9px;border-radius:999px">{it['tag']}</span>
          <b style="color:#1E2530;margin-left:7px">{it['media']}</b>
          <span style="color:#3F4652"> — {it['text']}</span>
        </td></tr>""")

    action = f"""
    <div style="background:#D9F27E;border-radius:12px;padding:20px 22px;margin:22px 0">
      <div style="font-size:11px;font-weight:700;letter-spacing:.12em;color:#5E6B2A">NEXT BEST ACTION</div>
      <div style="font-size:18px;font-weight:800;color:#1E2530;margin:7px 0 4px">{head}</div>
      <table width="100%" cellpadding="0" cellspacing="0" style="font-size:13px">{''.join(li)}</table>
      <div style="color:#5E6B2A;font-size:11.5px;margin-top:13px;line-height:1.6">
        기준: GA-ROAS 목표 {KPI_ROAS_LOW}~{KPI_ROAS_HIGH}% ·
        집행 {floor:,.0f}원 미만은 판단 보류 ·
        예산 이동은 한 번에 {STEP*100:.0f}%씩만 제안합니다
      </div>
    </div>"""

    warn = ""
    if gap:
        warn = (f'<div style="background:#FBF0CE;border:1px solid #E8D89A;border-radius:10px;'
                f'padding:13px 16px;margin-bottom:18px;color:#8A6714;font-size:13px">⚠️ {gap}</div>')

    link = ""
    if DASHBOARD_URL:
        link = (f'<a href="{DASHBOARD_URL}" style="display:inline-block;background:#7C5CFF;'
                f'color:#fff;text-decoration:none;font-weight:700;font-size:13px;'
                f'padding:11px 20px;border-radius:9px">대시보드에서 자세히 보기</a>')

    return f"""<!doctype html><html><body style="margin:0;padding:24px;background:#F4F5F7;
      font-family:-apple-system,BlinkMacSystemFont,'Malgun Gothic',sans-serif">
      <div style="max-width:860px;margin:0 auto;background:#fff;border-radius:16px;padding:28px 30px">
        <div style="font-size:11px;font-weight:700;letter-spacing:.12em;color:#8a8a7c">{kicker}</div>
        <div style="font-size:23px;font-weight:800;color:#17170f;margin:5px 0 3px">{title}</div>
        <div style="color:#8a8a7c;font-size:13px;margin-bottom:20px">{lead}</div>
        {warn}{kpi}{action}
        <div style="font-size:14px;font-weight:800;color:#17170f;margin:24px 0 10px">매체별 상세</div>
        {table}
        {trend}
        <div style="margin-top:24px">{link}</div>
        <div style="color:#A9A79A;font-size:11px;margin-top:22px;border-top:1px solid #EFEEE7;padding-top:14px">
          {footer}</div>
      </div></body></html>"""


def render_text(weekly: bool, start: date, end: date, df: pd.DataFrame,
                prev: dict, head: str, items: list[dict],
                series: list[dict] | None) -> str:
    """HTML을 못 보는 메일 앱을 위한 순수 텍스트 버전."""
    t = totals(df)
    if weekly:
        hdr = f"[{start:%Y-%m-%d} ~ {end:%m-%d}] 주간 성과 (전주 대비)"
    else:
        hdr = f"[{end:%Y-%m-%d}] 일별 성과 (전주 동요일 대비)"

    def d(cur, pv):
        if not pv:
            return ""
        return f" ({'+' if cur >= pv else '-'}{abs(cur-pv)/pv*100:,.0f}%)"

    lines = [hdr,
             f"광고비 {t['cost']:,.0f}원{d(t['cost'], prev.get('cost', 0))} · "
             f"GA 매출 {t['rev']:,.0f}원{d(t['rev'], prev.get('rev', 0))} · "
             f"ROAS {t['roas']:,.0f}%{d(t['roas'], prev.get('roas', 0))}",
             "", f"■ {head}", ""]
    for it in items:
        lines.append(f"  [{it['tag']}] {it['media']} — {it['text']}")
    lines += ["", "■ 매체별"]
    for _, r in df.iterrows():
        ro = "—" if (r["roas"] is None or pd.isna(r["roas"])) else f"{r['roas']:,.0f}%"
        lines.append(f"  {r['media']}: 광고비 {r['cost']:,.0f} / 매출 {r['rev']:,.0f} / ROAS {ro}")
    if weekly and series:
        lines += ["", "■ 일자별"]
        for s in series:
            lines.append(f"  {s['date']:%m/%d}({WD[s['date'].weekday()]}): "
                         f"광고비 {s['cost']:,.0f} / 매출 {s['rev']:,.0f} / ROAS {s['roas']:,.0f}%")
    if DASHBOARD_URL:
        lines += ["", DASHBOARD_URL]
    return "\n".join(lines)


# ──────────────────────────────────────────────────────────────
# 발송
# ──────────────────────────────────────────────────────────────
def send_mail(subject: str, html: str, text: str):
    if not MAIL_TO:
        raise SystemExit("MAIL_TO 가 비어 있습니다. 받는 사람을 지정해주세요.")
    if not SMTP_HOST:
        raise SystemExit("SMTP_HOST 가 없습니다. GitHub Secrets를 확인해주세요.")
    if not SMTP_USER or not SMTP_PASS:
        raise SystemExit(
            "SMTP_USER / SMTP_PASS 가 없습니다.\n"
            "Gmail을 쓰신다면 구글 계정 → 보안 → 2단계 인증 → 앱 비밀번호에서 16자리를 발급받아\n"
            "GitHub Secrets에 SMTP_PASS로 넣어주세요(띄어쓰기 없이).")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = formataddr((str(Header(MAIL_FROM_NAME, "utf-8")), MAIL_FROM))
    msg["To"] = ", ".join(MAIL_TO)
    if MAIL_CC:
        msg["Cc"] = ", ".join(MAIL_CC)
    msg.attach(MIMEText(text, "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))

    if SMTP_SECURITY == "ssl":
        server = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=60)
    else:
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60)
    try:
        server.ehlo()
        if SMTP_SECURITY == "starttls":
            server.starttls()
            server.ehlo()
        if SMTP_USER:
            server.login(SMTP_USER, SMTP_PASS)
        server.sendmail(MAIL_FROM, MAIL_TO + MAIL_CC, msg.as_string())
    finally:
        try:
            server.quit()
        except Exception:
            pass


# ──────────────────────────────────────────────────────────────
# 기간 정하기
# ──────────────────────────────────────────────────────────────
def resolve_period(mode: str, date_arg: str | None) -> tuple[bool, date, date]:
    """(주간인가, 시작일, 종료일).

    mode=auto 면 **한국 시각 기준 오늘이 월요일일 때만** 주간으로 간다.
    크론이 화~금 10시 / 월 13시 두 개라서, 요일만 보면 어느 쪽인지 갈린다.
    """
    kst_today = (datetime.utcnow() + timedelta(hours=9)).date()

    if mode == "auto":
        weekly = (kst_today.weekday() == 0)   # 0 = 월요일
    else:
        weekly = (mode == "weekly")

    if weekly:
        # --date 를 주면 '그 날이 속한 주'(월~일). 없으면 지난주.
        anchor = (datetime.strptime(date_arg, "%Y-%m-%d").date() if date_arg
                  else kst_today - timedelta(days=7))
        start = anchor - timedelta(days=anchor.weekday())
        return True, start, start + timedelta(days=6)

    day = (datetime.strptime(date_arg, "%Y-%m-%d").date() if date_arg
           else kst_today - timedelta(days=1))
    return False, day, day


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["auto", "daily", "weekly"], default="auto",
                    help="auto(기본)면 한국 기준 월요일에만 주간 리포트")
    ap.add_argument("--date", help="일별: 집계 날짜 / 주간: 그 날이 속한 주 (YYYY-MM-DD)")
    ap.add_argument("--dry-run", action="store_true", help="메일을 안 보내고 화면에만 출력")
    args = ap.parse_args()

    weekly, start, end = resolve_period(args.mode, args.date)
    n_days = (end - start).days + 1
    # 비교 기간: 주간은 그 전주, 일별은 전주 같은 요일
    p_start, p_end = start - timedelta(days=7), end - timedelta(days=7)

    print(f"{'주간' if weekly else '일별'} 리포트  ·  집계 {start} ~ {end} ({n_days}일)  ·  "
          f"비교 {p_start} ~ {p_end}", flush=True)
    print(f"받는 사람: {', '.join(MAIL_TO) or '(없음)'}  ·  "
          f"보내는 곳: {SMTP_HOST}:{SMTP_PORT} ({SMTP_SECURITY})", flush=True)

    client = get_client()
    since = p_start - timedelta(days=3)       # 비교 기간까지 덮고 경계 여유를 둔다
    ad_spend = load_table(client, "ad_spend_daily", since,
                          ["report_date", "channel", "source"])
    ga_daily = load_table(client, "ga_channel_daily", since,
                          ["report_date", "source_medium", "user_type"])
    master = load_table(client, "media_master")

    if master.empty:
        raise SystemExit(
            "media_master 표가 비어 있습니다.\n"
            "  대시보드 화면에는 매체 목록이 보이지만 그건 코드에 박아둔 기본값이고,\n"
            "  DB에는 아직 저장이 안 된 상태입니다. 이 스크립트는 DB만 읽습니다.\n"
            "\n"
            "  해결: 대시보드 → [채널 성과] 탭 → '⚙️ 매체 정의' 펼치기 → [저장] 클릭\n"
            "        한 번만 누르면 됩니다.")
    for c in ("sort_order", "utm_match", "spend_channel", "scope"):
        if c not in master.columns:
            master[c] = "" if c != "sort_order" else 100

    print(f"읽은 행: 광고비 {len(ad_spend):,} · GA {len(ga_daily):,} · 매체정의 {len(master):,}",
          flush=True)

    df = build_rows(ad_spend, ga_daily, master, start, end)
    if df.empty:
        print(f"[알림] {start}~{end} 에 집계할 데이터가 없습니다. 메일을 보내지 않습니다.\n"
              "       (대시보드에서 그 기간에 데이터가 있는지 확인해주세요)")
        return

    prev = totals(build_rows(ad_spend, ga_daily, master, p_start, p_end))
    series = daily_series(ad_spend, ga_daily, master, start, end) if weekly else None

    head, items, floor = build_actions(df, n_days)
    gap = spend_gap_note(ga_daily, ad_spend, start, end)
    html = render_html(weekly, start, end, df, prev, head, items, gap, floor, series)
    text = render_text(weekly, start, end, df, prev, head, items, series)

    t = totals(df)
    if weekly:
        subject = (f"[STCO 주간 성과] {start:%m/%d}~{end:%m/%d} · "
                   f"ROAS {t['roas']:,.0f}% · 매출 {t['rev']:,.0f}원")
    else:
        subject = (f"[STCO 일별 성과] {end:%m/%d} · "
                   f"ROAS {t['roas']:,.0f}% · 매출 {t['rev']:,.0f}원")

    if args.dry_run:
        print(subject)
        print()
        print(text)
        with open("daily_report_preview.html", "w", encoding="utf-8") as f:
            f.write(html)
        print("\n(HTML 미리보기를 daily_report_preview.html 로 저장했습니다)")
        return

    send_mail(subject, html, text)
    print(f"발송 완료: {subject} → {', '.join(MAIL_TO)}")


if __name__ == "__main__":
    main()
