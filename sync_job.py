"""메일 보내기 직전에 데이터를 최신으로 받아두는 스크립트 (GitHub Actions에서 돈다).

왜 필요한가
  대시보드(Streamlit)는 **누가 화면을 열어야** GA4·매체 API에서 어제 데이터를 받아온다.
  아침에 아무도 안 열면 DB에는 어제 데이터가 없고, 메일 스크립트(daily_report.py)는 DB만
  읽으니 '정액 계약(미리 깔아둔 일할 광고비)'만 있는 빈 메일이 나간다.
  (2026-10-01 메일: 브랜드검색·맨즈탭만 나오고 GFA·메타·크리테오·GA 매출이 전부 빠짐)

  그래서 메일 보내기 전에 이 스크립트가 대시보드와 **똑같은 함수**로 동기화를 한 번 돌린다.
  app.py를 그대로 불러다 쓰므로 동기화 규칙이 두 군데로 갈라질 일이 없다.

필요한 것
  .streamlit/secrets.toml — 워크플로가 GitHub Secret 'STREAMLIT_SECRETS' 내용으로 만들어준다.
  (Streamlit Cloud 앱 설정의 Secrets 칸 내용을 통째로 복사해 넣으면 된다)

한 매체가 실패해도 나머지는 계속 받고, 실패가 있어도 종료 코드는 0이다 — 동기화가
일부 안 돼도 메일은 나가야 하고, 무엇이 빠졌는지는 메일 쪽이 알려준다.
"""
import sys
import time
import traceback


def log(msg):
    print(msg, flush=True)


def main():
    t0 = time.time()
    try:
        import app
    except Exception:
        log("app.py를 불러오지 못했습니다 — 동기화를 건너뜁니다.")
        traceback.print_exc()
        return

    if app.get_supabase_client() is None:
        log("Supabase에 연결하지 못했습니다 (secrets.toml의 SUPABASE_URL / SUPABASE_KEY 확인). "
            "동기화를 건너뜁니다.")
        return

    def step(name, fn):
        s = time.time()
        try:
            out = fn()
            log(f"  ✅ {name} — {out}  ({time.time() - s:.0f}초)")
        except Exception as e:
            log(f"  ⚠️ {name} 실패 — {str(e)[:300]}")
        finally:
            try:
                app.st.cache_data.clear()     # 다음 단계가 방금 저장한 걸 읽게
            except Exception:
                pass

    utm = app.load_table("utm_channel_map")

    log("① GA4 채널 (구매·매출)")
    step("GA4 채널", lambda: app.sync_ga4_channel_daily(
        app.load_table("ga_channel_daily"), app.build_utm_channel_lookup(utm)))

    log("② GA4 소재")
    step("GA4 소재", lambda: app.sync_ga4_creative_daily(
        app.load_table("ga_creative_daily"),
        app.build_utm_channel_lookup(utm, include_extra=False)))

    log("③ 매체 광고비 (메타·구글·크리테오·네이버 등)")

    def _spend():
        n, saved, errors = app.sync_ad_spend(app.load_table("ad_spend_daily"), progress=log)
        for k, v in (errors or {}).items():
            log(f"     ⚠️ {k}: {str(v)[:200]}")
        return f"{n:,}행 " + ", ".join(f"{k} {v}" for k, v in (saved or {}).items())
    step("광고비", _spend)

    log("④ 매체 소재 실적")

    def _creative():
        n, saved, errors = app.sync_ad_creative(app.load_table("ad_creative_daily"), progress=log)
        for k, v in (errors or {}).items():
            log(f"     ⚠️ {k}: {str(v)[:200]}")
        return f"{n:,}행"
    step("소재 실적", _creative)

    log(f"동기화 끝 ({time.time() - t0:.0f}초)")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
    sys.exit(0)
