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

import base64
import hashlib
import io
import re
import urllib.parse
from datetime import date, datetime, timedelta

import numpy as np
import openpyxl
import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image

# 메인 타이틀 및 브라우저 탭 파비콘에 쓰는 막대그래프 아이콘(사용자가 준 PNG)을
# base64로 인라인 임베드 — 별도 이미지 파일 없이 app.py 하나로 배포되게 하기 위함.
# st.set_page_config()보다 먼저 정의해야 파비콘으로 바로 쓸 수 있다.
PAGE_TITLE_ICON_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAADAAAAAuCAIAAADhvA07AAAACXBIWXMAAA7DAAAOwwHHb6hkAAADu0lEQVRYR+2Yz47TMBDG4yTebrd/"
    "JC4sQkKIEwcQT8B78CwreAzeiTt3TtwQF1rBtk0T840nTRPHHm9IWO1ho1XqdeyZn78Ze9IqY0zykK70IcEQyyNQLCK5d8BdEksp5Z07"
    "slM5vvEv98hMTNO+NxzyRB4mLKajEGxV9vr57WuS5vpqPZvPQyuG0VzPLhZrDGAH7z7+gGhoh6TLMvX25SrP1OcPldY6y7K+cQ9QWZYw"
    "i9PAKGXC9nFaFMVBm0opsovFHEuTgjKzRD6ohnW73a5WqzTFcHfcGaiJlAVii+7o/oISUxnCUG9uIA9o0vks17mfKMWetgT7/X6xWHis"
    "JYmrEEetHkrpZGAAYoTu9CxJEGWCUonO0wud4g64vr+mV8gz/y5zbLVtn9oEyW2CRtiAkBINRAJcnwY9ge7OWBnISoT0IIWQT6RGq039"
    "LB45SxMIABjcx1yRk5qN2zvFjhy77Zb3GmUUUVghEsQKdL9XRKH7hSFvYYXoaCGJYkzRATED3eejFZqYR1SIk2jYAkePjioU2lx1/2gA"
    "14CUQzaDqDTwJHvqcJvOpFPbtTjyf1kh69eefvx3ateM1BnP+mGEQYXg7FhV5ohiLlm83e0WSy2NGPgsCITF43WiMJVo0BTb7XJFr0RT"
    "XR0gKtgKJdLGkdKnRAJBIKjVv2MICtz19TNqyDIOgT0DwSgSgup2lq1fvMYbhWCHuTHS+5IlTIw+chVieRhOnvw/aODRBbJBI5FkhRpu"
    "jteEe60DRMXdXoNe8iekwfI65xBMQ5jOSz7lM3KXXgdtEp/beFIUe8yQIzv06VkhXiju55d8OhTtHms++m2cVcrz5WEoRzPeE7JW9ti"
    "vFDVNZ+PjTLCv8XRO24+a9J85moly6ai1sd7YZ323D04FZDoaGA2f1CaBPJRTtnS0o9eJ5CmiFnGCK6pQOxa+9mSxqhcTVohkIYVweLc"
    "W7jJhiFh8B2sWBILnEtWjxK5uPKLPbe92+ytb7afCEkJmimN1eyhvD8fTX79d/NpucLID6Mun9VwzVTCK9I02Bu4pHe1qL+xnK1dd7TE"
    "lz/PnT9JXT7M/e3WB7/Y9xxj/u0ih+M3771ovai+9kHZ+sOJjGgcjLrmWNdUeHFzwgY7fNDabzeFwEJYBAPwytFwuLy8vMbfHQ78DnRXm"
    "bQ4UNGSjMAQORuF6jB4soygKeSU8ESheGspFx3GDIgNx3rTvzVrlieSyH86WUC4QP4oajdptuRjW9AMNszHpaGHbT+rnzsYegWJS/QWf"
    "cPB2ZPskBQAAAABJRU5ErkJggg=="
)


def _decode_page_icon():
    """base64 PNG를 PIL Image로 디코드해 st.set_page_config(page_icon=...)에 바로 넘긴다
    (emoji 문자열 대신 실제 이미지를 파비콘으로 쓰기 위함). 혹시 디코딩이 실패하면
    기존 이모지로 조용히 폴백해서 앱 실행 자체가 막히지 않게 한다."""
    try:
        return Image.open(io.BytesIO(base64.b64decode(PAGE_TITLE_ICON_B64)))
    except Exception:
        return "📊"


st.set_page_config(page_title="STCO 광고성과 대시보드", page_icon=_decode_page_icon(), layout="wide")

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

# 사이드바 '메뉴' 그룹 헤더 아이콘(사용자가 준 SVG, 초록 폴더+막대그래프)을 base64로 인라인 임베드.
# 별도 이미지 파일을 GitHub에 같이 올릴 필요 없이 app.py 하나로 배포되게 하기 위함.
NAV_GROUP_ICON_B64 = (
    "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgZmlsbD0ibm9uZSI+"
    "PGcgZmlsdGVyPSJ1cmwoI2EpIj48cGF0aCBmaWxsPSIjMUY5OTU1IiBkPSJNMi42NDYgNS43N0EyLjc3IDIuNzcgMCAwIDEgNS40"
    "MTUgM2g0LjJjLjUyNCAwIC45ODguMzM1IDEuMTU0LjgzLjExLjMzMi40Mi41NTUuNzY4LjU1NWg3LjcyNGMxLjAyIDAgMS44NDYu"
    "ODI2IDEuODQ2IDEuODQ2djExLjA3N0EzLjY5IDMuNjkgMCAwIDEgMTcuNDE1IDIxSDYuMzM4YTMuNjkgMy42OSAwIDAgMS0zLjY5"
    "Mi0zLjY5MnoiLz48L2c+PGcgZmlsdGVyPSJ1cmwoI2IpIj48cGF0aCBmaWxsPSIjMUY5OTU1IiBkPSJNNC4xNjUgOS45MDljLjE4"
    "LTIuMDMyIDEuOTgtMy42NzggNC4wMi0zLjY3OGgxMC43MjljMS45OTQgMCAzLjQ4OCAxLjU3NyAzLjM3MyAzLjU2M2wtLjQyMyA3"
    "LjM1NWMtLjEyIDIuMDc2LTEuOTQzIDMuNzkzLTQuMDI3IDMuNzkzSDYuODc3Yy0yLjA0IDAtMy41NDYtMS42NDctMy4zNjUtMy42"
    "Nzh6Ii8+PC9nPjxnIGZpbHRlcj0idXJsKCNjKSI+PHJlY3Qgd2lkdGg9IjExLjA3NyIgaGVpZ2h0PSI3LjM4NSIgZmlsbD0iI0Uy"
    "RTJFMiIgcng9IjEuODQ2IiB0cmFuc2Zvcm09Im1hdHJpeCgxIDAgLS4wODg1NCAuOTk2MDcgOC4xNDUgMTAuMzg1KSIvPjwvZz48"
    "ZyBmaWxsPSIjMUY5OTU1IiBmaWx0ZXI9InVybCgjZCkiPjxwYXRoIGQ9Ik0xNC45MDIgMTUuOTAyaC0xLjM1MWwuMzI3LTMuNjc4"
    "aDEuMzUxYzEuMTI3IDAgMS43Ni42OSAxLjY1NyAxLjgzOS0uMTAyIDEuMTQ4LS44NTcgMS44MzgtMS45ODQgMS44MzhtLS40MTgt"
    "Ljc0MmguNDU0Yy42MzIgMCAxLjAwNS0uMyAxLjA3Ni0xLjA5Ny4wNzEtLjc5OC0uMjQ4LTEuMTAzLS44OTEtMS4xMDNoLS40NDR6"
    "TTEwLjQ0IDE1LjkwMmgtLjkzOGwxLjU4Mi0zLjY3OGgxLjE1MmwuOTIzIDMuNjc4aC0uOTM0bC0uMTgtLjc4OEgxMC43NnptLjU4"
    "OC0xLjQ0OGguODY3bC0uMy0xLjMyMWgtLjAzMXoiLz48L2c+PGRlZnM+PGZpbHRlciBpZD0iYSIgd2lkdGg9IjIwLjMwOCIgaGVp"
    "Z2h0PSIxOS44NDYiIHg9IjIuMTg0IiB5PSIyLjUzOCIgY29sb3ItaW50ZXJwb2xhdGlvbi1maWx0ZXJzPSJzUkdCIiBmaWx0ZXJV"
    "bml0cz0idXNlclNwYWNlT25Vc2UiPjxmZUZsb29kIGZsb29kLW9wYWNpdHk9IjAiIHJlc3VsdD0iQmFja2dyb3VuZEltYWdlRml4"
    "Ii8+PGZlQ29sb3JNYXRyaXggaW49IlNvdXJjZUFscGhhIiByZXN1bHQ9ImhhcmRBbHBoYSIgdmFsdWVzPSIwIDAgMCAwIDAgMCAw"
    "IDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAxMjcgMCIvPjxmZU9mZnNldCBkeD0iLjQ2MiIgZHk9Ii40NjIiLz48ZmVHYXVzc2lhbkJs"
    "dXIgc3RkRGV2aWF0aW9uPSIuNDYyIi8+PGZlQ29tcG9zaXRlIGluMj0iaGFyZEFscGhhIiBvcGVyYXRvcj0ib3V0Ii8+PGZlQ29s"
    "b3JNYXRyaXggdmFsdWVzPSIwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwLjA4IDAiLz48ZmVCbGVuZCBpbjI9"
    "IkJhY2tncm91bmRJbWFnZUZpeCIgcmVzdWx0PSJlZmZlY3QxX2Ryb3BTaGFkb3dfNzYxXzQxNzgyIi8+PGZlQmxlbmQgaW49IlNv"
    "dXJjZUdyYXBoaWMiIGluMj0iZWZmZWN0MV9kcm9wU2hhZG93Xzc2MV80MTc4MiIgcmVzdWx0PSJzaGFwZSIvPjxmZUNvbG9yTWF0"
    "cml4IGluPSJTb3VyY2VBbHBoYSIgcmVzdWx0PSJoYXJkQWxwaGEiIHZhbHVlcz0iMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAw"
    "IDAgMCAwIDAgMTI3IDAiLz48ZmVPZmZzZXQgZHg9Ii40NjIiIGR5PSIuNDYyIi8+PGZlR2F1c3NpYW5CbHVyIHN0ZERldmlhdGlv"
    "bj0iLjIzMSIvPjxmZUNvbXBvc2l0ZSBpbjI9ImhhcmRBbHBoYSIgazI9Ii0xIiBrMz0iMSIgb3BlcmF0b3I9ImFyaXRobWV0aWMi"
    "Lz48ZmVDb2xvck1hdHJpeCB2YWx1ZXM9IjAgMCAwIDAgMSAwIDAgMCAwIDEgMCAwIDAgMCAxIDAgMCAwIDAuMiAwIi8+PGZlQmxl"
    "bmQgaW4yPSJzaGFwZSIgcmVzdWx0PSJlZmZlY3QyX2lubmVyU2hhZG93Xzc2MV80MTc4MiIvPjwvZmlsdGVyPjxmaWx0ZXIgaWQ9"
    "ImIiIHdpZHRoPSIyMi4wMjciIGhlaWdodD0iMTYuNTU3IiB4PSIuNzI4IiB5PSI1LjMwOCIgY29sb3ItaW50ZXJwb2xhdGlvbi1m"
    "aWx0ZXJzPSJzUkdCIiBmaWx0ZXJVbml0cz0idXNlclNwYWNlT25Vc2UiPjxmZUZsb29kIGZsb29kLW9wYWNpdHk9IjAiIHJlc3Vs"
    "dD0iQmFja2dyb3VuZEltYWdlRml4Ii8+PGZlQ29sb3JNYXRyaXggaW49IlNvdXJjZUFscGhhIiByZXN1bHQ9ImhhcmRBbHBoYSIg"
    "dmFsdWVzPSIwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAxMjcgMCIvPjxmZU9mZnNldCBkeD0iLTEuODQ2Ii8+"
    "PGZlR2F1c3NpYW5CbHVyIHN0ZERldmlhdGlvbj0iLjQ2MiIvPjxmZUNvbXBvc2l0ZSBpbjI9ImhhcmRBbHBoYSIgb3BlcmF0b3I9"
    "Im91dCIvPjxmZUNvbG9yTWF0cml4IHZhbHVlcz0iMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMC4yNSAwIi8+"
    "PGZlQmxlbmQgaW4yPSJCYWNrZ3JvdW5kSW1hZ2VGaXgiIHJlc3VsdD0iZWZmZWN0MV9kcm9wU2hhZG93Xzc2MV80MTc4MiIvPjxm"
    "ZUJsZW5kIGluPSJTb3VyY2VHcmFwaGljIiBpbjI9ImVmZmVjdDFfZHJvcFNoYWRvd183NjFfNDE3ODIiIHJlc3VsdD0ic2hhcGUi"
    "Lz48ZmVDb2xvck1hdHJpeCBpbj0iU291cmNlQWxwaGEiIHJlc3VsdD0iaGFyZEFscGhhIiB2YWx1ZXM9IjAgMCAwIDAgMCAwIDAg"
    "MCAwIDAgMCAwIDAgMCAwIDAgMCAwIDEyNyAwIi8+PGZlT2Zmc2V0IGR4PSIuNDYyIiBkeT0iLjQ2MiIvPjxmZUdhdXNzaWFuQmx1"
    "ciBzdGREZXZpYXRpb249Ii4yMzEiLz48ZmVDb21wb3NpdGUgaW4yPSJoYXJkQWxwaGEiIGsyPSItMSIgazM9IjEiIG9wZXJhdG9y"
    "PSJhcml0aG1ldGljIi8+PGZlQ29sb3JNYXRyaXggdmFsdWVzPSIwIDAgMCAwIDEgMCAwIDAgMCAxIDAgMCAwIDAgMSAwIDAgMCAw"
    "LjIgMCIvPjxmZUJsZW5kIGluMj0ic2hhcGUiIHJlc3VsdD0iZWZmZWN0Ml9pbm5lclNoYWRvd183NjFfNDE3ODIiLz48L2ZpbHRl"
    "cj48ZmlsdGVyIGlkPSJjIiB3aWR0aD0iMTUuMTExIiBoZWlnaHQ9IjExLjA0OCIgeD0iNS4wNjMiIHk9IjkiIGNvbG9yLWludGVy"
    "cG9sYXRpb24tZmlsdGVycz0ic1JHQiIgZmlsdGVyVW5pdHM9InVzZXJTcGFjZU9uVXNlIj48ZmVGbG9vZCBmbG9vZC1vcGFjaXR5"
    "PSIwIiByZXN1bHQ9IkJhY2tncm91bmRJbWFnZUZpeCIvPjxmZUNvbG9yTWF0cml4IGluPSJTb3VyY2VBbHBoYSIgcmVzdWx0PSJo"
    "YXJkQWxwaGEiIHZhbHVlcz0iMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMTI3IDAiLz48ZmVNb3JwaG9sb2d5"
    "IGluPSJTb3VyY2VBbHBoYSIgb3BlcmF0b3I9ImRpbGF0ZSIgcmFkaXVzPSIuOTIzIiByZXN1bHQ9ImVmZmVjdDFfZHJvcFNoYWRv"
    "d183NjFfNDE3ODIiLz48ZmVPZmZzZXQgZHg9Ii0uNzM4IiBkeT0iLjQ2MiIvPjxmZUdhdXNzaWFuQmx1ciBzdGREZXZpYXRpb249"
    "Ii40NjIiLz48ZmVDb21wb3NpdGUgaW4yPSJoYXJkQWxwaGEiIG9wZXJhdG9yPSJvdXQiLz48ZmVDb2xvck1hdHJpeCB2YWx1ZXM9"
    "IjAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAuMDQgMCIvPjxmZUJsZW5kIGluMj0iQmFja2dyb3VuZEltYWdl"
    "Rml4IiByZXN1bHQ9ImVmZmVjdDFfZHJvcFNoYWRvd183NjFfNDE3ODIiLz48ZmVCbGVuZCBpbj0iU291cmNlR3JhcGhpYyIgaW4y"
    "PSJlZmZlY3QxX2Ryb3BTaGFkb3dfNzYxXzQxNzgyIiByZXN1bHQ9InNoYXBlIi8+PGZlQ29sb3JNYXRyaXggaW49IlNvdXJjZUFs"
    "cGhhIiByZXN1bHQ9ImhhcmRBbHBoYSIgdmFsdWVzPSIwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAxMjcgMCIv"
    "PjxmZU9mZnNldCBkeD0iLjQ2MiIgZHk9Ii40NjIiLz48ZmVHYXVzc2lhbkJsdXIgc3RkRGV2aWF0aW9uPSIuMjMxIi8+PGZlQ29t"
    "cG9zaXRlIGluMj0iaGFyZEFscGhhIiBrMj0iLTEiIGszPSIxIiBvcGVyYXRvcj0iYXJpdGhtZXRpYyIvPjxmZUNvbG9yTWF0cml4"
    "IHZhbHVlcz0iMCAwIDAgMCAxIDAgMCAwIDAgMSAwIDAgMCAwIDEgMCAwIDAgMSAwIi8+PGZlQmxlbmQgaW4yPSJzaGFwZSIgcmVz"
    "dWx0PSJlZmZlY3QyX2lubmVyU2hhZG93Xzc2MV80MTc4MiIvPjwvZmlsdGVyPjxmaWx0ZXIgaWQ9ImQiIHdpZHRoPSI3LjM5NSIg"
    "aGVpZ2h0PSIzLjY3OCIgeD0iOS41MDIiIHk9IjEyLjIyNCIgY29sb3ItaW50ZXJwb2xhdGlvbi1maWx0ZXJzPSJzUkdCIiBmaWx0"
    "ZXJVbml0cz0idXNlclNwYWNlT25Vc2UiPjxmZUZsb29kIGZsb29kLW9wYWNpdHk9IjAiIHJlc3VsdD0iQmFja2dyb3VuZEltYWdl"
    "Rml4Ii8+PGZlQmxlbmQgaW49IlNvdXJjZUdyYXBoaWMiIGluMj0iQmFja2dyb3VuZEltYWdlRml4IiByZXN1bHQ9InNoYXBlIi8+"
    "PGZlQ29sb3JNYXRyaXggaW49IlNvdXJjZUFscGhhIiByZXN1bHQ9ImhhcmRBbHBoYSIgdmFsdWVzPSIwIDAgMCAwIDAgMCAwIDAg"
    "MCAwIDAgMCAwIDAgMCAwIDAgMCAxMjcgMCIvPjxmZU9mZnNldC8+PGZlR2F1c3NpYW5CbHVyIHN0ZERldmlhdGlvbj0iLjIzMSIv"
    "PjxmZUNvbXBvc2l0ZSBpbjI9ImhhcmRBbHBoYSIgazI9Ii0xIiBrMz0iMSIgb3BlcmF0b3I9ImFyaXRobWV0aWMiLz48ZmVDb2xv"
    "ck1hdHJpeCB2YWx1ZXM9IjAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAuMyAwIi8+PGZlQmxlbmQgaW4yPSJz"
    "aGFwZSIgcmVzdWx0PSJlZmZlY3QxX2lubmVyU2hhZG93Xzc2MV80MTc4MiIvPjwvZmlsdGVyPjwvZGVmcz48L3N2Zz4="
)


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
    # title 텍스트를 따로 지정하지 않은 차트는 title 객체에 font만 있고 text가 없어
    # 화면에 "undefined"라는 글자가 그대로 렌더링되는 경우가 있어, 빈 문자열로 명시해둔다.
    if fig.layout.title.text is None:
        fig.update_layout(title_text="")
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
            color: {THEME_COLORS["body"]} !important;
            font-weight: 600 !important;
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

        [data-testid="stPopover"] {{ width: fit-content !important; }}
        [data-testid="stPopover"] > div {{ width: fit-content !important; }}
        [data-testid="stPopover"] > div > button {{
            background-color: {THEME_COLORS["canvas"]} !important;
            color: {THEME_COLORS["foreground"]} !important;
            border: 1px solid {THEME_COLORS["border"]} !important;
            border-radius: 8px !important;
            font-weight: 500 !important;
            padding: 3px 8px !important;
            font-size: 12px !important;
            white-space: nowrap !important;
            min-width: 0 !important;
            width: fit-content !important;
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

        /* 사이드바 '메뉴' 내비게이션 전용 스타일 (네이버 광고관리자 스타일 참고):
           그룹 헤더는 아이콘+굵은 글씨+expander 기본 화살표(펼침/접힘)만 남기고 박스/배경 제거,
           페이지 항목은 버튼 박스가 아니라 '텍스트 링크'처럼 보이게(현재 페이지만 파란 굵은 글씨). */
        div.st-key-stco_nav div[data-testid="stExpander"] {{
            border: none !important;
            box-shadow: none !important;
            background: transparent !important;
        }}
        div.st-key-stco_nav div[data-testid="stExpander"] summary {{
            font-weight: 700 !important;
            font-size: 15px !important;
            color: {THEME_COLORS["foreground"]} !important;
            padding: 8px 4px !important;
        }}
        /* 그룹명 왼쪽에 커스텀 아이콘(SVG, 이모지 대신) 표시 */
        div.st-key-stco_nav div[data-testid="stExpander"] summary p::before {{
            content: "";
            display: inline-block;
            width: 18px;
            height: 18px;
            margin-right: 6px;
            vertical-align: -4px;
            background-image: url("data:image/svg+xml;base64,{NAV_GROUP_ICON_B64}");
            background-size: contain;
            background-repeat: no-repeat;
        }}
        div.st-key-stco_nav div[data-testid="stExpander"] summary:hover {{
            color: {THEME_COLORS["primary"]} !important;
        }}
        div.st-key-stco_nav div[data-testid="stExpanderDetails"] {{
            padding-left: 4px !important;
        }}
        div.st-key-stco_nav .stButton > button {{
            background: transparent !important;
            border: none !important;
            box-shadow: none !important;
            text-align: left !important;
            justify-content: flex-start !important;
            padding: 6px 14px !important;
            border-radius: 8px !important;
            font-weight: 400 !important;
            font-size: 14px !important;
            color: {THEME_COLORS["body"]} !important;
        }}
        div.st-key-stco_nav .stButton > button[kind="primary"] {{
            color: {THEME_COLORS["primary"]} !important;
            font-weight: 700 !important;
            background: transparent !important;
        }}
        div.st-key-stco_nav .stButton > button:hover {{
            background: {THEME_COLORS["surface"]} !important;
            color: {THEME_COLORS["primary"]} !important;
        }}
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
    "creative_performance": "creative_performance",
    "channel_audience_snapshot": "channel_audience_snapshot",
}

# 채널 요약 시트로 취급하지 않을 시트들
SHEET_SKIP_EXACT = {"매체통합", "GA-RAW", "RD_네이버"}
SHEET_SKIP_SUBSTR = ["_data", "_date", "확인용", "소재"]

# 채널믹스 목표 비중 (STCO 퍼포먼스마케팅 인수인계서 4절 기준: 발굴 메타50/구글20/네이버20, 회수 모비온10)
CHANNEL_MIX_TARGET = {
    "메타": 0.50,
    "구글": 0.20,
    "네이버": 0.20,
    "모비온": 0.10,
}

CHANNEL_GROUP_RULES = [
    ("메타", ["메타", "페이스북", "facebook", "meta", "인스타"]),
    ("구글", ["구글", "google"]),
    ("네이버", ["네이버", "naver", "gfa", "브검", "ssp"]),
    ("모비온", ["모비온", "mobon"]),
]


def map_channel_group(channel_name: str) -> str:
    """시트/채널명을 채널믹스 목표의 4개 그룹(메타/구글/네이버/모비온)으로 매핑.
    어디에도 안 걸리면 '기타'(크리테오·에디AI·카카오모먼트 등 목표 비중 배정이 없는 매체)."""
    name = str(channel_name).lower()
    for group, keywords in CHANNEL_GROUP_RULES:
        if any(kw.lower() in name for kw in keywords):
            return group
    return "기타"


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
    try:
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
    except Exception as e:
        # 테이블이 아직 Supabase에 없는 경우(예: creative_performance 신규 테이블 미생성) 등
        # API 에러가 나면 앱 전체가 죽지 않도록 빈 데이터로 취급하고 안내만 띄운다.
        st.sidebar.warning(f"'{TABLES.get(name, name)}' 테이블 조회 실패 — 해당 테이블이 Supabase에 없을 수 있습니다. ({e})")
        return pd.DataFrame()
    return pd.DataFrame(rows)


def delete_creative_performance_for_date(as_of_date_value):
    """오늘자 소재별 성과를 다시 저장하기 전에, 같은 날짜로 이미 저장돼있던 이전 스냅샷을
    통째로 지운다. upsert는 '새 업로드에 있는 행'만 갱신할 뿐 '새 업로드에서 빠진 행'은
    절대 지우지 않아서, 엑셀에서 시트를 숨기거나 소재를 뺐는데도 예전에 저장된 행이
    DB에 계속 남아 화면에 나오는 문제(예: 철 지난 숨김 시트 데이터)를 막기 위함이다."""
    client = get_supabase_client()
    if client is None:
        store = _local_store()
        df = store.get("creative_performance", pd.DataFrame())
        if not df.empty and "as_of_date" in df.columns:
            store["creative_performance"] = df[df["as_of_date"].astype(str) != str(as_of_date_value)]
        return
    try:
        client.table(TABLES["creative_performance"]).delete().eq(
            "as_of_date", str(as_of_date_value)
        ).execute()
    except Exception as e:
        st.sidebar.warning(f"기존 소재별 성과 스냅샷 삭제 실패(무시하고 계속 진행합니다): {e}")


def delete_channel_audience_for_date(as_of_date_value):
    """타겟팅별 성과용 채널×오디언스(신규/리타겟팅) 스냅샷도 creative_performance와 동일한 이유로
    저장 전에 같은 날짜의 이전 스냅샷을 지운다(upsert가 빠진 행을 알아서 지워주지 않아서)."""
    client = get_supabase_client()
    if client is None:
        store = _local_store()
        df = store.get("channel_audience_snapshot", pd.DataFrame())
        if not df.empty and "as_of_date" in df.columns:
            store["channel_audience_snapshot"] = df[df["as_of_date"].astype(str) != str(as_of_date_value)]
        return
    try:
        client.table(TABLES["channel_audience_snapshot"]).delete().eq(
            "as_of_date", str(as_of_date_value)
        ).execute()
    except Exception as e:
        st.sidebar.warning(f"기존 채널×오디언스 스냅샷 삭제 실패(무시하고 계속 진행합니다): {e}")


def save_table(name: str, df: pd.DataFrame, on_conflict: str, source_file: str):
    if df is None or df.empty:
        return 0
    df = df.copy()
    # 같은 업로드 안에 동일 키(예: creative_performance라면 같은 날짜+매체+소재명) 행이 중복되면
    # upsert 한 번의 요청 안에서 같은 행을 두 번 건드리게 되어 Postgres가 에러를 내므로, 저장 전에
    # 미리 정리해야 한다. 예전에는 단순 drop_duplicates(keep="last")로 나머지를 버렸는데, 같은
    # 소재가 여러 캠페인/그룹에 동시에 쓰이는 경우(예: '260630_슈즈'가 리타겟팅/프로스펙팅 캠페인에
    # 둘 다 있는 경우) 실제로는 두 행을 '합산'해야 할 실적인데 한쪽을 통째로 버려서 노출·광고비·매출이
    # 실제보다 훨씬 작게 저장되는 버그가 있었다. 그래서 중복 키가 있으면 숫자 컬럼은 합산하고,
    # 문자열/날짜 등 나머지 컬럼은 비어있지 않은 첫 값을 사용하도록 바꾼다.
    key_cols = [c.strip() for c in on_conflict.split(",")]
    if df.duplicated(subset=key_cols, keep=False).any():
        agg_map = {}
        for col in df.columns:
            if col in key_cols:
                continue
            if pd.api.types.is_numeric_dtype(df[col]):
                agg_map[col] = "sum"
            else:
                agg_map[col] = lambda s: next((v for v in s if pd.notna(v) and v != ""), None)
        df = df.groupby(key_cols, as_index=False).agg(agg_map)
    df["source_file"] = source_file
    df["uploaded_at"] = datetime.utcnow().isoformat()

    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].astype(str)
        elif df[col].dtype == object:
            df[col] = df[col].apply(lambda v: v.isoformat() if isinstance(v, (date, datetime)) else v)

    # NaN/inf는 표준 JSON에 없는 값이라 Supabase API 호출(JSON 직렬화) 자체가 통째로 실패한다
    # ("Out of range float values are not JSON compliant: nan"). 어떤 컬럼에서 오든(신규 image_url
    # 컬럼이 전부 비어 float64/NaN으로 잡히는 경우 등) 저장 직전에 전부 None으로 정리해 방지한다.
    # (float64 컬럼은 dtype을 object로 바꾸지 않으면 None이 다시 NaN으로 되돌아가므로 astype(object) 필수)
    df = df.replace([np.inf, -np.inf], np.nan)
    df = df.astype(object).where(pd.notnull(df), None)

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
    try:
        for i in range(0, len(records), 500):
            client.table(TABLES[name]).upsert(records[i : i + 500], on_conflict=on_conflict).execute()
    except Exception as e:
        # 예: creative_performance처럼 Supabase에 아직 테이블을 안 만든 경우 — 이 테이블만 건너뛰고
        # 나머지 저장(주간/월별/매체 등)은 정상 진행되도록 전체 저장 흐름을 막지 않는다.
        st.sidebar.error(f"'{TABLES.get(name, name)}' 저장 실패 — Supabase에 해당 테이블이 있는지 확인해주세요. ({e})")
        return 0
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


def match_col_pos(headers, include_all=None, include_any=None, exclude=None):
    """match_col과 같은 규칙이지만 컬럼 '이름' 대신 헤더 리스트 안에서의 위치(정수 인덱스)를 반환한다.
    소재 단위 raw 시트는 같은 이름의 컬럼(예: 자체 전환/매출/ROAS 다음에 GA 전환/매출/ROAS가
    똑같이 '전환'/'매출'/'ROAS'라는 이름으로 한 번 더 나오는 경우)이 있어서, 이름으로 data[name]을
    하면 중복 컬럼이 DataFrame으로 잡혀 계산이 깨진다. 위치 기반으로 첫 매칭 컬럼 하나만 정확히 집는다."""
    include_all, include_any, exclude = include_all or [], include_any or [], exclude or []
    for i, c in enumerate(headers):
        lc = clean_col(c).lower()
        if not lc:
            continue
        if any(ex in lc for ex in exclude):
            continue
        if include_all and not all(tok in lc for tok in include_all):
            continue
        if include_any and not any(tok in lc for tok in include_any):
            continue
        return i
    return None


def numcol_by_pos(block: pd.DataFrame, idx):
    """위치(정수 인덱스) 기반으로 숫자 컬럼 값을 꺼낸다 — 컬럼명이 중복돼도 안전하다."""
    if idx is None or idx >= block.shape[1]:
        return np.zeros(len(block))
    return pd.to_numeric(block.iloc[:, idx], errors="coerce").fillna(0).values


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


CREATIVE_NAME_HINTS = ["소재성과", "소재별성과"]  # 예: 페이스북_소재성과요약


def _infer_channel_from_sheet(sheet: str) -> str:
    """'(DA) 구글_실적최대화_data' → '구글', '페이스북_소재성과요약' → '페이스북' 처럼
    시트명 접두어에서 매체명만 뽑아낸다."""
    name = re.sub(r"^\([^)]*\)\s*", "", sheet)  # 앞의 "(DA) " 같은 괄호 접두어 제거
    name = name.split("_")[0].strip()
    return name or sheet


CREATIVE_SCAN_ROWS = 100  # 실제 파일에서 소재 상세표 헤더가 50~60번째 행 근처에 있는 경우가 있어 넉넉하게 잡음


def find_creative_sheets(xls: pd.ExcelFile):
    """소재 단위 데이터가 있는 시트를 찾는다.
    시트명 패턴(예: '_data' 접미사)에 의존하면 '(DA) GFA_data(자사몰)'처럼
    '_data'가 끝이 아니라 중간에 오는 등 실제 파일의 이름 규칙이 제각각이라 놓치기 쉽다.
    그래서 이름은 참고만 하고, 실제로 각 시트 안에 '광고소재' 컬럼이 있는지
    내용 기준으로 직접 확인한다 (숨겨진/그룹화된 행이 있어도 pandas는 다 읽으므로 문제 없음).
    """
    found = []
    for s in xls.sheet_names:
        if any(hint in s for hint in CREATIVE_NAME_HINTS):
            found.append(s)
            continue
        try:
            probe = pd.read_excel(xls, sheet_name=s, header=None, nrows=CREATIVE_SCAN_ROWS)
        except Exception:
            continue
        if find_header_row(probe, required=("광고소재",), scan=CREATIVE_SCAN_ROWS) is not None:
            found.append(s)
    return found


def _find_last_matching_row(raw: pd.DataFrame, required, scan=CREATIVE_SCAN_ROWS):
    """find_header_row와 달리 '가장 먼저' 매칭되는 행이 아니라 '가장 마지막으로' 매칭되는 행을 찾는다.
    실제 파일은 한 시트 안에 캠페인/그룹 집계표 → (일부 채널은) 소재 요약 집계표 → 진짜 소재별
    raw data 표 순서로 여러 표가 쌓여 있고, 우리가 원하는 건 항상 가장 마지막(가장 상세한) 표라서
    노출수·클릭수만 있는 앞쪽 집계표를 잘못 집지 않으려면 '마지막 매칭'을 써야 한다."""
    last = None
    for i in range(min(scan, len(raw))):
        row_text = " ".join(str(x) for x in raw.iloc[i].tolist())
        if all(tok in row_text for tok in required):
            last = i
    return last


# 소재별 성과 화면에는 '현재 실제로 운영 중인' 매체만 노출한다 (당근마켓/카카오모먼트/DV360/
# 구글에즈/MOBON/에디AI/네이버 SA·SSP 등은 현재 미운영이라 제외).
# GFA는 캠페인명에 붙는 '_PC'/'_MO' 접미사로 기기별(네이버 GFA PC / 네이버 GFA MO)로 쪼갠다.
CREATIVE_CHANNEL_WHITELIST_MAP = {
    "페이스북": "메타",
    "구글": "구글(P-MAX)",  # '(DA) 구글_실적최대화_data' 시트 = Performance Max
    "크리테오": "크리테오",
}


def _map_creative_channel(inferred_channel: str, campaign_series: pd.Series) -> pd.Series:
    """행별 매체탭 라벨을 계산. 현재 미운영 매체는 None을 반환해 화면/저장에서 제외되도록 한다."""
    idx = campaign_series.index
    if inferred_channel in CREATIVE_CHANNEL_WHITELIST_MAP:
        label = CREATIVE_CHANNEL_WHITELIST_MAP[inferred_channel]
        return pd.Series([label] * len(idx), index=idx, dtype=object)
    if inferred_channel == "GFA":
        camp = campaign_series.astype(str)
        is_pc = camp.str.contains("_PC", case=False, na=False)
        is_mo = camp.str.contains("_MO", case=False, na=False)
        result = pd.Series([None] * len(idx), index=idx, dtype=object)
        result[is_pc] = "네이버 GFA PC"
        result[~is_pc & is_mo] = "네이버 GFA MO"
        return result
    return pd.Series([None] * len(idx), index=idx, dtype=object)


# ──────────────────────────────────────────────────────────────
# 채널 그룹(타겟팅 오디언스) 단위 파싱 — 타겟팅별 성과용 (신규)
# 소재 상세표보다 앞에 나오는 '캠페인 분류 | 그룹 분류 | 노출수 ...' 집계표에서
# 그룹(광고그룹/타겟팅 오디언스) 단위로 신규/리타겟팅을 분류해 저장한다.
# ──────────────────────────────────────────────────────────────

# 사용자가 정리해준 '260804_광고 매체 그룹별 타겟팅 분류.xlsx' 기준 매핑.
# 키: (원본채널, 그룹명) 또는 [그룹 분류 컬럼이 없는 채널은] (원본채널, 캠페인분류명) → "신규"/"리타겟팅"
AUDIENCE_TYPE_MAP = {
    # 메타
    ("페이스북", "의류관심타겟"): "신규",
    ("페이스북", "패션관심타겟"): "신규",
    ("페이스북", "방문자180일"): "리타겟팅",
    # 네이버 GFA (자사몰 + 애드부스트)
    ("GFA", "패션관심타겟"): "신규",
    ("GFA", "쇼핑관심타겟"): "신규",
    ("GFA", "방문자180일"): "리타겟팅",
    ("GFA", "3050연령타겟"): "신규",
    ("GFA", "3050남성타겟"): "신규",
    # 크리테오 (그룹 분류 컬럼이 없어 캠페인 분류 자체가 그룹 역할)
    ("크리테오", "다이나믹_신규"): "신규",
    ("크리테오", "스태틱_신규"): "신규",
    ("크리테오", "다이나믹_리텐션"): "리타겟팅",
    ("크리테오", "스태틱_리텐션"): "리타겟팅",
    # 구글 실적최대화 (그룹 분류 없음 → 캠페인분류 전체가 신규)
    ("구글", "Pmax_온라인팀"): "신규",
}


def classify_audience_type(origin_channel: str, label) -> str:
    """(원본채널, 그룹명/캠페인분류명) → '신규' / '리타겟팅' / '미분류'.
    1) 사용자가 정리해준 매핑표를 먼저 조회
    2) 못 찾으면 이름에 흔히 쓰이는 키워드로 최대한 추정
    3) 그래도 애매하면 '미분류'로 남겨서 대시보드에서 바로 눈에 띄게 한다 (임의로 단정하지 않음).
    새 캠페인/그룹이 추가되면 위 AUDIENCE_TYPE_MAP에 항목을 추가해주면 된다."""
    name = str(label).strip()
    key = (origin_channel, name)
    if key in AUDIENCE_TYPE_MAP:
        return AUDIENCE_TYPE_MAP[key]
    if "재방문자제외" in name or "신규" in name or "관심타겟" in name:
        return "신규"
    if any(kw in name for kw in ["방문자", "리텐션", "리타겟"]):
        return "리타겟팅"
    return "미분류"


def parse_channel_group_sheet(xls: pd.ExcelFile, sheet: str, today: date):
    """소재 상세표보다 앞에 나오는 '캠페인 분류 | 그룹 분류 | 노출수 ...' 집계표를 파싱해서
    그룹(타겟팅 오디언스) 단위 성과 + 신규/리타겟팅 분류를 반환한다.
    그룹 분류 컬럼이 없는 채널(크리테오)은 캠페인 분류를 그룹처럼 취급한다."""
    raw = pd.read_excel(xls, sheet_name=sheet, header=None)
    hdr = find_header_row(raw, required=("노출수", "클릭수"), scan=30)
    if hdr is None:
        return None

    headers = [clean_col(h) for h in raw.iloc[hdr].tolist()]
    campaign_idx = match_col_pos(headers, include_any=["캠페인"])
    group_idx = match_col_pos(headers, include_any=["그룹"])
    if campaign_idx is None:
        return None
    label_idx = group_idx if group_idx is not None else campaign_idx

    # 이 집계표 바로 아래에는 컬럼 배치가 다른 '소재 상세표'(캠페인|그룹|광고소재|...)가 또 나오는데,
    # 헤더 텍스트가 똑같이 '노출수'+'클릭수'를 포함해서 그대로 이어 읽으면 소재명이 그룹명 자리에
    # 잘못 섞여 들어간다. 그래서 다음 헤더가 나오는 지점 '직전'까지만 이 표의 데이터로 본다.
    next_hdr = None
    for i in range(hdr + 1, len(raw)):
        row_text = " ".join(str(x) for x in raw.iloc[i].tolist())
        if "노출수" in row_text and "클릭수" in row_text:
            next_hdr = i
            break
    body = raw.iloc[hdr + 1 : next_hdr] if next_hdr is not None else raw.iloc[hdr + 1 :]
    if body.empty:
        return None

    label_series = body.iloc[:, label_idx].astype(str).str.strip()
    keep_mask = body.iloc[:, label_idx].notna() & (label_series != "") & (label_series.str.lower() != "nan")
    keep_mask &= ~label_series.str.contains("합계|TOTAL|총계", case=False, na=False)
    body = body[keep_mask]
    label_series = label_series[keep_mask]
    if body.empty:
        return None

    impr_idx = match_col_pos(headers, include_any=["노출"])
    clicks_idx = match_col_pos(headers, include_any=["클릭"])
    cost_ex_idx = match_col_pos(headers, include_all=["제외"], include_any=["광고비", "비용"], exclude=["포함"])
    cost_in_idx = match_col_pos(headers, include_all=["포함"], include_any=["광고비", "비용"], exclude=["제외"])
    if cost_ex_idx is None and cost_in_idx is None:
        cost_generic_idx = match_col_pos(headers, include_any=["광고비", "비용"])
        cost_ex_idx = cost_in_idx = cost_generic_idx
    signup_idx = match_col_pos(headers, include_any=["가입"])
    conv_idx = match_col_pos(headers, include_all=["전환"], exclude=["금액", "ga", "율"])
    rev_idx = match_col_pos(headers, include_any=["매출", "전환금액"], exclude=["ga", "객단가"])

    campaign_series = (
        body.iloc[:, campaign_idx].astype(str).str.strip() if campaign_idx is not None else label_series
    )

    origin_channel = _infer_channel_from_sheet(sheet)

    out = pd.DataFrame()
    out["channel"] = _map_creative_channel(origin_channel, campaign_series).values
    out["audience_type"] = [classify_audience_type(origin_channel, v) for v in label_series]
    out["impressions"] = numcol_by_pos(body, impr_idx)
    out["clicks"] = numcol_by_pos(body, clicks_idx)
    out["cost_excl_vat"] = numcol_by_pos(body, cost_ex_idx)
    out["cost_incl_vat"] = numcol_by_pos(body, cost_in_idx)
    out["signups"] = numcol_by_pos(body, signup_idx)
    out["conversions"] = numcol_by_pos(body, conv_idx)
    out["revenue"] = numcol_by_pos(body, rev_idx)
    out["as_of_date"] = today
    # 현재 미운영 매체(채널 매핑이 None인 행)는 제외
    out = out[out["channel"].notna()]
    if out.empty:
        return None
    out = out[
        (out["impressions"] > 0) | (out["clicks"] > 0) | (out["cost_incl_vat"] > 0)
        | (out["conversions"] > 0) | (out["revenue"] > 0)
    ]
    return out.reset_index(drop=True)


# 네이버 검색광고/쇼핑검색광고/브랜드검색광고 — 사용자 지시에 따라 캠페인/그룹 세부 분류 없이
# "채널 전체 = 리타겟팅"으로 우선 분류한다(신규 유입보다는 이미 브랜드/검색 의도가 있는 사용자가
# 대상이라는 판단). GFA/메타/크리테오처럼 그룹명 기준 신규·리타겟팅을 나누지 않고 시트 전체를
# 채널 합계 1행으로 요약한다.
#
# 데이터 소스는 "(SA)/(SSP)/(브검) 네이버_data"(원본 raw 표)가 아니라 "_data"가 없는 요약 시트의
# "■ 일일 데이터 → 당월 총합" 행을 쓴다. 두 시트가 같은 채널인데도 서로 다른 숫자를 담고 있는
# 경우가 있고(예: 브랜드검색광고는 _data 시트에 광고비가 0으로 비어있음), 사용자가 실제 주간
# 보고서에서 보는 숫자는 요약 시트의 "당월 총합"이라고 확인해줘서 그 쪽을 공식 소스로 쓴다.
NAVER_SEARCH_SHEET_CHANNEL_MAP = {
    "(SA)": "네이버 검색광고",
    "(SSP)": "네이버 쇼핑검색광고",
    "(브검)": "네이버 브랜드검색광고",
}


def _naver_search_channel_from_sheet(sheet: str):
    """요약 시트("_data"가 붙지 않은)만 대상으로 한다 — 원본 raw 시트는 제외."""
    if sheet.endswith("_data"):
        return None
    for prefix, label in NAVER_SEARCH_SHEET_CHANNEL_MAP.items():
        if sheet.startswith(prefix):
            return label
    return None


def parse_naver_search_sheet(xls: pd.ExcelFile, sheet: str, today: date):
    """(SA)/(SSP)/(브검) 네이버 요약 시트의 "■ 일일 데이터 → 당월 총합" 행 하나를
    채널 합계로 반환한다."""
    channel_label = _naver_search_channel_from_sheet(sheet)
    if channel_label is None:
        return None
    raw = pd.read_excel(xls, sheet_name=sheet, header=None)

    total_row = None
    for i in range(len(raw)):
        row_text = " ".join(str(x) for x in raw.iloc[i].tolist())
        if "당월" in row_text and "총합" in row_text:
            total_row = i
            break
    if total_row is None:
        return None

    hdr = None
    for i in range(total_row - 1, -1, -1):
        row_text = " ".join(str(x) for x in raw.iloc[i].tolist())
        if "노출수" in row_text and "클릭수" in row_text:
            hdr = i
            break
    if hdr is None:
        return None

    headers = [clean_col(h) for h in raw.iloc[hdr].tolist()]
    impr_idx = match_col_pos(headers, include_any=["노출"])
    clicks_idx = match_col_pos(headers, include_any=["클릭"])
    # 일부 시트는 "광고비(VAT제외)"처럼 '제외'가 명시돼 있고, 일부는 그냥 "광고비"(제외 의미)만
    # 있고 옆에 "광고비(VAT포함)"가 따로 있어 '제외' 키워드 없이도 구분해야 한다.
    cost_ex_idx = match_col_pos(headers, include_any=["광고비", "비용"], exclude=["포함"])
    cost_in_idx = match_col_pos(headers, include_all=["포함"], include_any=["광고비", "비용"])
    signup_idx = match_col_pos(headers, include_any=["가입"])
    conv_idx = match_col_pos(headers, include_all=["전환"], exclude=["금액", "ga", "율"])
    rev_idx = match_col_pos(headers, include_any=["매출", "전환금액"], exclude=["ga", "객단가"])

    row = raw.iloc[total_row]

    def _val(idx):
        if idx is None or idx >= len(row):
            return 0.0
        v = pd.to_numeric(row.iloc[idx], errors="coerce")
        return 0.0 if pd.isna(v) else float(v)

    out = pd.DataFrame([{
        "channel": channel_label,
        "audience_type": "리타겟팅",
        "impressions": _val(impr_idx),
        "clicks": _val(clicks_idx),
        "cost_excl_vat": _val(cost_ex_idx),
        "cost_incl_vat": _val(cost_in_idx),
        "signups": _val(signup_idx),
        "conversions": _val(conv_idx),
        "revenue": _val(rev_idx),
        "as_of_date": today,
    }])
    return out


def parse_creative_sheet(xls: pd.ExcelFile, sheet: str, today: date, hidden_rows: set = None):
    """소재 단위 시트 하나를 파싱. 시트마다 컬럼 구성이 조금씩 달라
    '소재명/광고소재/행 레이블' 컬럼과 노출/클릭/광고비/전환/매출 컬럼을 유연하게 매칭한다.
    hidden_rows: 엑셀에서 '행 숨기기'로 숨긴 행 번호(0-indexed) 집합. 해당 소재는 더 이상
    안 쓰는 것으로 보고 집계에서 제외한다(다시 보이게 하면 다음 업로드부터 자동으로 복귀)."""
    raw = pd.read_excel(xls, sheet_name=sheet, header=None)

    # 1순위: 노출수·클릭수·소재(광고소재/소재명 등)가 다 있는 '진짜 소재 상세표' 헤더 중 가장 마지막 것.
    # 컬럼명이 '광고소재'가 아니라 그냥 '소재'/'소재명'인 매체도 있어 '소재'로만 느슨하게 확인한다.
    hdr = _find_last_matching_row(raw, required=("노출수", "클릭수", "소재"))
    if hdr is None:
        # 2순위: '소재' 대신 '행 레이블' 같은 이름을 쓰는 피벗 요약 시트 (예: 페이스북_소재성과요약)
        hdr = find_header_row(raw, required=("노출수", "클릭수"), scan=CREATIVE_SCAN_ROWS)
    if hdr is None:
        return None

    headers = [clean_col(h) for h in raw.iloc[hdr].tolist()]
    creative_idx = None
    for cand in (["소재명"], ["광고소재"], ["행레이블"], ["소재"]):
        creative_idx = match_col_pos(headers, include_any=cand)
        if creative_idx is not None:
            break
    if creative_idx is None:
        return None

    body = raw.iloc[hdr + 1:]
    if body.empty:
        return None

    # 엑셀에서 '행 숨기기'로 숨겨둔 소재 행은 더 이상 안 쓰는 것으로 보고 제외
    if hidden_rows:
        body = body[~body.index.isin(hidden_rows)]
        if body.empty:
            return None

    # '총 합계'/TOTAL 행(캠페인·그룹 칸에 라벨이 오기도 함)과, 그 아래 붙는 '소재 이미지'
    # 미리보기 같은 성격이 다른 표는 소재명 칸 자체가 비어있거나(합계 행) 다른 컬럼 위치에
    # 있어서(이미지 표) 아래 notna/공백 필터에서 자연스럽게 걸러진다 — 별도 절단 없이 그대로 둔다.
    creative_series = body.iloc[:, creative_idx].astype(str).str.strip()
    keep_mask = body.iloc[:, creative_idx].notna() & (creative_series != "") & (creative_series.str.lower() != "nan")
    keep_mask &= ~creative_series.str.contains("합계|TOTAL|총계", case=False, na=False)
    body = body[keep_mask]
    creative_series = creative_series[keep_mask]
    if body.empty:
        return None

    impr_idx = match_col_pos(headers, include_any=["노출"])
    clicks_idx = match_col_pos(headers, include_any=["클릭"])
    cost_ex_idx = match_col_pos(headers, include_all=["제외"], include_any=["광고비", "비용"], exclude=["포함"])
    cost_in_idx = match_col_pos(headers, include_all=["포함"], include_any=["광고비", "비용"], exclude=["제외"])
    if cost_ex_idx is None and cost_in_idx is None:
        # VAT 구분 없이 '광고비' 한 컬럼만 있는 시트 대응 (포함/제외를 같은 값으로 취급)
        cost_generic_idx = match_col_pos(headers, include_any=["광고비", "비용"])
        cost_ex_idx = cost_in_idx = cost_generic_idx
    conv_idx = match_col_pos(headers, include_all=["전환"], exclude=["금액", "ga", "율"])
    rev_idx = match_col_pos(headers, include_any=["매출", "전환금액"], exclude=["ga", "객단가"])
    campaign_idx = match_col_pos(headers, include_any=["캠페인"])
    campaign_series = (
        body.iloc[:, campaign_idx] if campaign_idx is not None
        else pd.Series([""] * len(body), index=body.index)
    )

    out = pd.DataFrame()
    out["creative"] = creative_series.values
    out["channel"] = _map_creative_channel(_infer_channel_from_sheet(sheet), campaign_series).values
    out["impressions"] = numcol_by_pos(body, impr_idx)
    out["clicks"] = numcol_by_pos(body, clicks_idx)
    out["cost_excl_vat"] = numcol_by_pos(body, cost_ex_idx)
    out["cost_incl_vat"] = numcol_by_pos(body, cost_in_idx)
    out["conversions"] = numcol_by_pos(body, conv_idx)
    out["revenue"] = numcol_by_pos(body, rev_idx)
    out["as_of_date"] = today
    # 현재 미운영 매체(채널 매핑이 None인 행)는 소재별 성과에서 제외
    out = out[out["channel"].notna()]
    if out.empty:
        return None
    out = out[
        (out["impressions"] > 0) | (out["clicks"] > 0) | (out["cost_incl_vat"] > 0)
        | (out["conversions"] > 0) | (out["revenue"] > 0)
    ]
    return out.reset_index(drop=True)


# ──────────────────────────────────────────────────────────────
# 소재 이미지 추출/업로드 (신규)
# 각 채널 시트 끝에 실제로 embedded 되어 있는 '소재명 | 이미지' 표에서
# 이미지를 직접 추출해 Supabase Storage에 업로드하고, (원본채널, 소재명) → 공개 URL로 매핑한다.
# 소재 단위 성과 파싱(parse_creative_sheet)과는 별도 경로 — 실패해도 성과 데이터 저장에는 영향 없음.
# ──────────────────────────────────────────────────────────────
CREATIVE_IMAGE_BUCKET = "creative-images"

# 소재별 성과 화면의 매체탭(기기 분리 후) → 이미지가 들어있는 원본 채널명 역매핑
TAB_TO_ORIGIN_CHANNEL = {
    "메타": "페이스북",
    "구글(P-MAX)": "구글",
    "크리테오": "크리테오",
    "네이버 GFA PC": "GFA",
    "네이버 GFA MO": "GFA",
}

# Storage 저장 경로(폴더명)는 한글이 섞이면 일부 클라이언트/버전에서 URL 인코딩 문제로
# 업로드가 통째로 실패할 수 있어, 폴더명은 반드시 ASCII로 고정한다.
ORIGIN_CHANNEL_FOLDER = {
    "페이스북": "facebook",
    "구글": "google",
    "크리테오": "criteo",
    "GFA": "gfa",
}


def _creative_image_key(name) -> str:
    """소재명 표기 차이를 무시하고 매칭하기 위한 정규화 키.
    - 줄바꿈/공백 제거
    - 크리테오처럼 성과표의 소재명 끝에 '(크리에이티브ID)'가 붙지만 이미지 캡션표에는
      없는 경우가 있어, 끝에 붙은 '(숫자)'는 제거해서 매칭한다 ('(수정)'처럼 숫자가 아닌
      괄호 표기는 그대로 남겨 다른 소재와 구분되도록 유지)."""
    s = re.sub(r"\s+", "", str(name)).strip()
    s = re.sub(r"\(\d+\)$", "", s)
    return s


def extract_creative_images(file, sheets_and_channels) -> dict:
    """업로드된 엑셀 원본에서 시트별로 embedded 이미지를 찾아
    {(원본채널, 정규화된 소재명): (이미지bytes, 확장자)} 형태로 반환한다.
    시트 안에 '소재명(또는 광고소재/소재)'과 '이미지'가 같이 있는 헤더 행을 찾고,
    그 아래 각 행에 앵커된 이미지를 같은 행의 소재명 값과 매칭한다."""
    images = {}
    try:
        file.seek(0)
    except Exception:
        pass
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
    except Exception:
        return images
    finally:
        try:
            file.seek(0)
        except Exception:
            pass

    for sheet, origin_channel in sheets_and_channels:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        imgs = getattr(ws, "_images", [])
        if not imgs:
            continue
        try:
            raw = pd.read_excel(file, sheet_name=sheet, header=None)
        except Exception:
            continue
        finally:
            try:
                file.seek(0)
            except Exception:
                pass

        # '소재명 | 이미지' 헤더 행(진짜 성과 상세표와는 별개, 항상 시트 맨 끝 쪽에 있음) 탐색
        hdr = None
        for i in range(len(raw)):
            row_text = " ".join(str(x) for x in raw.iloc[i].tolist())
            if "이미지" in row_text and any(tok in row_text for tok in ("소재명", "광고소재", "소재")):
                hdr = i
        if hdr is None:
            continue

        headers = [clean_col(h) for h in raw.iloc[hdr].tolist()]
        name_idx = None
        for cand in (["소재명"], ["광고소재"], ["소재"]):
            name_idx = match_col_pos(headers, include_any=cand)
            if name_idx is not None:
                break
        if name_idx is None:
            continue

        for img in imgs:
            try:
                row_i = img.anchor._from.row  # 0-indexed, header=None으로 읽은 raw의 행 인덱스와 동일
                if row_i <= hdr or row_i >= len(raw):
                    continue
                name = str(raw.iloc[row_i, name_idx]).replace("\n", "").strip()
                if not name or name.lower() == "nan":
                    continue
                data = img._data()
                ext = (img.format or "png").lower()
                images[(origin_channel, _creative_image_key(name))] = (data, ext)
            except Exception:
                continue

    return images


def _safe_storage_name(name_key: str) -> str:
    """Supabase Storage의 객체 키 검증(isValidKey)은 영문/숫자/일부 기호만 허용하고
    한글 등 non-ASCII 문자는 URL 인코딩 여부와 무관하게 무조건 거부한다(서버가 디코딩한
    키 자체를 검사하기 때문에 percent-encoding으로는 우회되지 않음).
    그래서 한글은 아예 제거하고, 대신 원본 소재명의 해시를 붙여 서로 다른 소재끼리
    충돌하지 않게 한다(같은 소재는 항상 같은 해시 → upsert로 정상적으로 덮어써짐)."""
    ascii_part = re.sub(r"[^0-9A-Za-z_\-]", "", name_key)[:60]
    digest = hashlib.md5(name_key.encode("utf-8")).hexdigest()[:10]
    base = ascii_part or "creative"
    return f"{base}_{digest}"


def upload_creative_images(images: dict):
    """추출된 이미지를 Supabase Storage(버킷: creative-images)에 업로드하고
    ({(원본채널, 정규화된 소재명): 공개 URL} 딕셔너리, 에러 메시지 목록)을 반환한다.
    Supabase 미연결이거나 버킷이 없는 등 실패 시 조용히 건너뛴다(성과 저장 자체는 막지 않음)."""
    client = get_supabase_client()
    if client is None or not images:
        return {}, []
    urls = {}
    errors = []
    for (origin_channel, name_key), (data, ext) in images.items():
        folder = ORIGIN_CHANNEL_FOLDER.get(origin_channel, "etc")
        path = f"{folder}/{_safe_storage_name(name_key)}.{ext}"
        try:
            client.storage.from_(CREATIVE_IMAGE_BUCKET).upload(
                path, data, {"content-type": f"image/{ext}", "upsert": "true"}
            )
            urls[(origin_channel, name_key)] = client.storage.from_(CREATIVE_IMAGE_BUCKET).get_public_url(path)
        except Exception as e:
            if len(errors) < 3:  # 화면에 너무 길게 쌓이지 않도록 앞의 몇 개만 보관
                errors.append(f"{origin_channel}/{name_key}: {e}")
            continue
    return urls, errors


def attach_creative_images(creatives: pd.DataFrame, image_urls: dict) -> pd.DataFrame:
    """소재별 성과 데이터프레임(channel=탭 라벨, creative=소재명)에 image_url 컬럼을 붙인다."""
    if creatives is None or creatives.empty or not image_urls:
        return creatives
    creatives = creatives.copy()
    creatives["image_url"] = creatives.apply(
        lambda r: image_urls.get(
            (TAB_TO_ORIGIN_CHANNEL.get(r["channel"], r["channel"]), _creative_image_key(r["creative"]))
        ),
        axis=1,
    )
    return creatives


def get_hidden_rows(wb, sheet: str) -> set:
    """엑셀에서 마우스 우클릭 → '행 숨기기'로 숨긴 행의 번호(0-indexed, header=None으로 읽은
    raw DataFrame의 행 인덱스와 동일)를 반환한다. 소재 하나를 '이제 안 씀' 표시로 숨겨두면
    소재별 성과 집계에서 자동으로 제외하기 위한 용도 — 숨김을 풀면 다음 업로드부터 다시 잡힌다."""
    if wb is None or sheet not in wb.sheetnames:
        return set()
    ws = wb[sheet]
    hidden = set()
    for row_idx, dim in ws.row_dimensions.items():
        if dim.hidden:
            hidden.add(row_idx - 1)  # openpyxl 행 번호는 1-indexed
    return hidden


def get_visible_sheets(wb) -> set:
    """워크북 탭 자체가 숨겨진(우클릭 → 시트 숨기기) 시트 이름을 제외한, 눈에 보이는 시트 이름 집합.
    현재 안 쓰는 매체나 철 지난 시트를 통째로 탭 숨김 처리해두는 게 이 대시보드 운영자의
    실제 워크플로우라, '시트가 숨겨져 있으면 데이터도 안 가져온다'가 가장 직접적인 신호다.
    (wb가 없으면 None을 반환해 — 필터링 자체를 건너뛰고 기존처럼 전부 읽도록 — 안전하게 폴백한다.)"""
    if wb is None:
        return None
    return {name for name in wb.sheetnames if wb[name].sheet_state == "visible"}


def parse_workbook(file, today: date):
    xls = pd.ExcelFile(file)
    try:
        file.seek(0)
    except Exception:
        pass
    try:
        wb_for_hidden = openpyxl.load_workbook(file, data_only=True)
    except Exception:
        wb_for_hidden = None
    finally:
        try:
            file.seek(0)
        except Exception:
            pass
    visible_sheets = get_visible_sheets(wb_for_hidden)
    result = {
        "weekly": pd.DataFrame(),
        "monthly": pd.DataFrame(),
        "daily": pd.DataFrame(),
        "channel_snapshot": pd.DataFrame(),
        "channels": pd.DataFrame(),
        "ga": pd.DataFrame(),
        "creatives": pd.DataFrame(),
        "channel_audience": pd.DataFrame(),
        "channel_sheets_found": [],
        "channel_sheets_parsed": [],
        "creative_sheets_found": [],
    }
    if "매체통합" in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name="매체통합", header=None)
        bounds = find_sections(raw)
        result["monthly"] = parse_monthly(raw, bounds, today)
        result["weekly"] = parse_weekly(raw, bounds, today)
        result["daily"] = parse_daily(raw, bounds, today)
        result["channel_snapshot"] = parse_channel_snapshot(raw, bounds, result["monthly"])

    chan_frames = []
    channel_sheet_candidates = discover_channel_sheets(xls)
    if visible_sheets is not None:
        channel_sheet_candidates = [s for s in channel_sheet_candidates if s in visible_sheets]
    for s in channel_sheet_candidates:
        result["channel_sheets_found"].append(s)
        df = parse_channel_sheet(xls, s, today)
        if df is not None and len(df):
            chan_frames.append(df)
            result["channel_sheets_parsed"].append(s)
    if chan_frames:
        result["channels"] = pd.concat(chan_frames, ignore_index=True)

    creative_frames = []
    creative_sheets = find_creative_sheets(xls)
    if visible_sheets is not None:
        creative_sheets = [s for s in creative_sheets if s in visible_sheets]
    for s in creative_sheets:
        hidden_rows = get_hidden_rows(wb_for_hidden, s)
        df = parse_creative_sheet(xls, s, today, hidden_rows=hidden_rows)
        if df is not None and len(df):
            creative_frames.append(df)
    result["creatives"] = pd.concat(creative_frames, ignore_index=True) if creative_frames else pd.DataFrame()
    result["creative_sheets_found"] = creative_sheets

    # 타겟팅별 성과(신규 타겟팅 vs 리타겟팅)용 그룹(타겟팅 오디언스) 단위 데이터.
    # 소재 파싱과 같은 시트 목록을 쓰되, 소재 상세표가 아니라 그 앞의 그룹 집계표를 읽는다.
    audience_frames = []
    for s in creative_sheets:
        df = parse_channel_group_sheet(xls, s, today)
        if df is not None and len(df):
            audience_frames.append(df)
    # 네이버 검색광고/쇼핑검색광고/브랜드검색광고 — (SA)/(SSP)/(브검) 시트는 '광고소재' 컬럼이
    # 없어 creative_sheets 목록에 안 잡히므로 별도로 찾아서 리타겟팅 채널로 추가한다.
    naver_search_sheets = [s for s in xls.sheet_names if _naver_search_channel_from_sheet(s)]
    for s in naver_search_sheets:
        df = parse_naver_search_sheet(xls, s, today)
        if df is not None and len(df):
            audience_frames.append(df)
    result["channel_audience"] = pd.concat(audience_frames, ignore_index=True) if audience_frames else pd.DataFrame()

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
    "creative": "소재명",
    "creative_image": "소재",
    "image_url": "소재이미지URL",
    "impressions": "노출수",
    "clicks": "클릭수",
    "cost_excl_vat": "광고비(VAT제외)",
    "cost_incl_vat": "광고비(VAT+)",
    "signups": "회원가입",
    "signup_rate": "가입율(%)",
    "audience_type": "구분(신규/리타겟팅)",
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


DATE_PERIOD_OPTIONS = DATE_PRESETS + ["전체", "직접선택"]


def period_filter(min_d: date, max_d: date, key: str, default_preset: str = "이번달"):
    """날짜 프리셋 버튼 목록(누적 표의 preset_button_picker와 동일한 방식) + 직접선택 달력.
    버튼 아래에 실제로 적용된 날짜범위를 항상 캡션으로 보여준다. 반환값은 (start, end)."""
    preset = preset_button_picker(DATE_PERIOD_OPTIONS, key=f"{key}_dateperiod", default=default_preset)

    if preset == "전체":
        start, end = min_d, max_d
    elif preset == "직접선택":
        narrow_col, _spacer = st.columns([3, 9])
        with narrow_col:
            date_range = st.date_input(
                "기간 직접 선택", value=(min_d, max_d), min_value=min_d, max_value=max_d, key=f"{key}_manual",
            )
        if isinstance(date_range, tuple) and len(date_range) == 2:
            start, end = date_range
        else:
            start, end = min_d, max_d
    else:
        start, end = _preset_to_range(preset, min_d, max_d)

    st.caption(f"📆 {start:%Y-%m-%d} ~ {end:%Y-%m-%d}")
    return start, end


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
PCT2_COLS = {"ctr", "cvr", "bounce_rate", "ecommerce_cvr", "signup_rate"}
PCT0_COLS = {"roas", "ga_roas"}


KOR_WEEKDAY = ["월", "화", "수", "목", "금", "토", "일"]


def format_display(df: pd.DataFrame) -> pd.DataFrame:
    """화면/엑셀에 보여줄 때 쓰는 최종 포맷팅 (콤마, 소수점 자리수, 날짜 형식)."""
    df = df.copy()
    for c in df.columns:
        if c == "report_month":
            df[c] = pd.to_datetime(df[c]).dt.strftime("%Y-%m")
        elif c == "report_date":
            dt = pd.to_datetime(df[c])
            df[c] = dt.dt.strftime("%Y-%m-%d") + " (" + dt.dt.dayofweek.map(lambda i: KOR_WEEKDAY[int(i)]) + ")"
        elif c in ("week_start", "week_end", "as_of_month", "as_of_date"):
            df[c] = pd.to_datetime(df[c]).dt.strftime("%Y-%m-%d")
        elif c in MONEY_COLS:
            df[c] = pd.to_numeric(df[c], errors="coerce").map(lambda v: f"{v:,.0f}" if pd.notna(v) else "")
        elif c in PCT2_COLS:
            df[c] = pd.to_numeric(df[c], errors="coerce").map(lambda v: f"{v:.2f}%" if pd.notna(v) else "")
        elif c in PCT0_COLS:
            df[c] = pd.to_numeric(df[c], errors="coerce").map(lambda v: f"{v:.0f}%" if pd.notna(v) else "")
    return df


KOR_COLS_REV = {v: k for k, v in KOR_COLS.items()}


def render_html_table(table: pd.DataFrame, raw: pd.DataFrame = None):
    """pandas Styler(jinja2 의존) 없이 순수 HTML로 표를 그린다.
    ▲(상승)는 빨간색, ▼(하락)는 파란색 글씨로 표시하고, 인덱스는 표시하지 않는다.
    raw(원본 숫자 컬럼, TOTAL 행 제외 · table의 데이터 행과 같은 순서)를 넘기면 raw에 있는
    숫자 컬럼에 한해 헤더를 클릭해 오름차순/내림차순 정렬할 수 있다(TOTAL 행은 항상 고정)."""
    if table.empty:
        st.caption("데이터가 아직 없습니다.")
        return

    cols = list(table.columns)

    sortable_keys = {}
    if raw is not None:
        for c in cols:
            orig = KOR_COLS_REV.get(c)
            if orig and orig in raw.columns and pd.api.types.is_numeric_dtype(raw[orig]):
                sortable_keys[c] = orig

    if sortable_keys:
        thead = "".join(
            f'<th class="stco-sortable" data-key="{sortable_keys[c]}">{c} <span class="stco-sort-arrow">⇅</span></th>'
            if c in sortable_keys else f"<th>{c}</th>"
            for c in cols
        )
    else:
        thead = "".join(f"<th>{c}</th>" for c in cols)

    # TOTAL 행이 맨 위/맨 아래 중 어느 쪽에 있는지에 따라, 정렬 후에도 같은 위치에 고정한다.
    total_pin = "top" if (not table.empty and str(table.iloc[0][cols[0]]).strip() == "TOTAL") else "bottom"

    row_htmls = []
    raw_idx = 0  # table에서 TOTAL 행을 제외한 순번 == raw의 행 순번(raw가 있을 때)
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
            data_attr = ""
            if c in sortable_keys and not is_total and raw is not None and raw_idx < len(raw):
                sort_col = sortable_keys[c]
                raw_val = raw.iloc[raw_idx][sort_col]
                if pd.notna(raw_val):
                    data_attr = f' data-value="{float(raw_val)}"'
            cells.append(f'<td{colspan} style="{style}"{data_attr}>{text}</td>')
        row_class = ' class="stco-total-row"' if is_total else ""
        row_htmls.append(f"<tr{row_class}>{''.join(cells)}</tr>")
        if not is_total:
            raw_idx += 1

    table_style = f"""
    <style>
    .stco-table-wrap {{
        overflow-x:auto; border:1px solid {THEME_COLORS["border"]}; border-radius:10px; background:{THEME_COLORS["canvas"]};
        margin-bottom:18px;
    }}
    .stco-table {{
        width:100%; table-layout:auto; border-collapse:collapse; font-size:14px;
        font-family: {THEME_FONT_STACK};
    }}
    .stco-table th {{
        background:{THEME_COLORS["surface"]}; color:{THEME_COLORS["muted"]}; font-weight:600; padding:8px 12px;
        text-align:right; border-bottom:1px solid {THEME_COLORS["border"]}; white-space:nowrap;
    }}
    .stco-table th:first-child {{ text-align:left; border-top-left-radius:10px; width:1%; }}
    .stco-table th:last-child {{ border-top-right-radius:10px; }}
    .stco-table td {{
        padding:8px 12px; text-align:right; color:{THEME_COLORS["foreground"]};
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
    .stco-sortable {{ cursor:pointer; user-select:none; }}
    .stco-sortable:hover {{ color:{THEME_COLORS["foreground"]}; }}
    .stco-sort-arrow {{ font-size:11px; margin-left:2px; }}
    .stco-sortable.stco-sort-active {{ color:{THEME_COLORS["foreground"]}; }}
    .stco-sortable.stco-sort-active .stco-sort-arrow {{ font-weight:700; }}
    </style>
    """

    if not sortable_keys:
        # 정렬 기능이 필요 없는 표는 기존대로 st.markdown으로 그린다(고정 높이 없이 자연스럽게 흐름).
        html = f"""
        {table_style}
        <div class="stco-table-wrap">
        <table class="stco-table">
          <thead><tr>{thead}</tr></thead>
          <tbody>{''.join(row_htmls)}</tbody>
        </table>
        </div>
        """
        st.markdown(html, unsafe_allow_html=True)
        return

    # 정렬 기능이 있는 표는 st.markdown이 <script>를 실행하지 않기 때문에 iframe(components.html)으로 그린다.
    n_rows = len(table)
    iframe_height = min(max(90 + n_rows * 41, 150), 900)
    table_id = f"stco-table-{abs(hash(tuple(cols))) % 100000}"
    html = f"""
    <html><head>{table_style}</head>
    <body style="margin:0; font-family:{THEME_FONT_STACK};">
    <div class="stco-table-wrap">
    <table class="stco-table" id="{table_id}">
      <thead><tr>{thead}</tr></thead>
      <tbody>{''.join(row_htmls)}</tbody>
    </table>
    </div>
    <script>
    (function() {{
        var table = document.getElementById("{table_id}");
        var tbody = table.querySelector("tbody");
        var headers = table.querySelectorAll("th.stco-sortable");
        var pinTop = {str(total_pin == "top").lower()};
        var state = {{ key: null, dir: 1 }};

        headers.forEach(function(th) {{
            th.addEventListener("click", function() {{
                var key = th.getAttribute("data-key");
                var colIndex = Array.prototype.indexOf.call(th.parentNode.children, th);
                if (state.key === key) {{
                    state.dir = -state.dir;
                }} else {{
                    state.key = key;
                    state.dir = -1;  // 처음 클릭하면 내림차순(큰 값 먼저)
                }}
                headers.forEach(function(h) {{
                    h.classList.remove("stco-sort-active");
                    h.querySelector(".stco-sort-arrow").textContent = "⇅";
                }});
                th.classList.add("stco-sort-active");
                th.querySelector(".stco-sort-arrow").textContent = state.dir === 1 ? "▲" : "▼";

                var rows = Array.prototype.slice.call(tbody.querySelectorAll("tr"));
                var totalRows = rows.filter(function(r) {{ return r.classList.contains("stco-total-row"); }});
                var dataRows = rows.filter(function(r) {{ return !r.classList.contains("stco-total-row"); }});

                dataRows.sort(function(a, b) {{
                    var av = parseFloat(a.children[colIndex].getAttribute("data-value"));
                    var bv = parseFloat(b.children[colIndex].getAttribute("data-value"));
                    var aNum = isNaN(av) ? -Infinity : av;
                    var bNum = isNaN(bv) ? -Infinity : bv;
                    return (aNum - bNum) * state.dir;
                }});

                dataRows.forEach(function(r) {{ tbody.appendChild(r); }});
                if (pinTop) {{
                    for (var i = totalRows.length - 1; i >= 0; i--) {{
                        tbody.insertBefore(totalRows[i], tbody.firstChild);
                    }}
                }} else {{
                    totalRows.forEach(function(r) {{ tbody.appendChild(r); }});
                }}
            }});
        }});
    }})();
    </script>
    </body></html>
    """
    components.html(html, height=iframe_height, scrolling=True)


# 소재별 성과 표에서 헤더 클릭으로 정렬 가능한 컬럼: 화면표시(한글) 헤더 → 정렬 기준이 되는 원본 컬럼 키.
# st.markdown은 보안상 <script>를 실행하지 않아서(react-markdown이 그냥 무시함), 클릭 정렬처럼
# JS가 필요한 표는 반드시 st.components.v1.html(iframe)로 렌더링해야 한다.
CREATIVE_SORT_COLS = {
    "노출수": "impressions",
    "클릭수": "clicks",
    "CTR(%)": "ctr",
    "CPC": "cpc",
    "전환수": "conversions",
    "ROAS(%)": "roas",
}


def render_sortable_creative_table(show: pd.DataFrame, raw: pd.DataFrame):
    """소재별 성과 표 전용 렌더러. render_html_table과 모양은 같지만, 노출/클릭/CTR/CPC/전환/ROAS
    헤더를 클릭하면 오름차순↔내림차순으로 재정렬된다 (TOTAL 행은 항상 맨 위 고정).
    show: korify+format_display를 거친 화면표시용 표 (맨 위 TOTAL 행 포함, 문자열 포맷).
    raw : TOTAL 행을 제외한, show의 데이터 행과 같은 순서의 원본(raw, 숫자) 값 표
          (정렬 기준값을 문자열이 아니라 실제 숫자로 비교하기 위해 필요)."""
    if show.empty:
        st.caption("데이터가 아직 없습니다.")
        return

    cols = list(show.columns)
    thead_cells = []
    for c in cols:
        sort_key = CREATIVE_SORT_COLS.get(c)
        if sort_key:
            thead_cells.append(
                f'<th class="stco-sortable" data-key="{sort_key}">{c} <span class="stco-sort-arrow">⇅</span></th>'
            )
        else:
            thead_cells.append(f"<th>{c}</th>")
    thead = "".join(thead_cells)

    row_htmls = []
    raw_idx = 0  # show에서 TOTAL 행(맨 위 1개)을 제외한 순번 == raw의 행 순번
    for _, row in show.iterrows():
        first_text = str(row[cols[0]]).strip()
        is_total = first_text == "TOTAL"
        cells = []
        for c in cols:
            val = row[c]
            text = "" if pd.isna(val) else str(val)
            sort_key = CREATIVE_SORT_COLS.get(c)
            data_attr = ""
            if sort_key and not is_total and raw_idx < len(raw) and sort_key in raw.columns:
                raw_val = raw.iloc[raw_idx][sort_key]
                if pd.notna(raw_val):
                    data_attr = f' data-value="{float(raw_val)}"'
            cells.append(f"<td{data_attr}>{text}</td>")
        row_class = ' class="stco-total-row"' if is_total else ""
        row_htmls.append(f"<tr{row_class}>{''.join(cells)}</tr>")
        if not is_total:
            raw_idx += 1

    n_data_rows = max(len(show) - 1, 0)
    iframe_height = min(max(110 + n_data_rows * 156, 220), 900)

    html = f"""
    <style>
    body {{ margin:0; font-family:{THEME_FONT_STACK}; }}
    .stco-table-wrap {{
        overflow:auto; border:1px solid {THEME_COLORS["border"]}; border-radius:10px; background:{THEME_COLORS["canvas"]};
    }}
    .stco-table {{ width:auto; max-width:100%; border-collapse:collapse; font-size:14px; }}
    .stco-table th {{
        background:{THEME_COLORS["surface"]}; color:{THEME_COLORS["muted"]}; font-weight:600; padding:8px 14px;
        text-align:right; border-bottom:1px solid {THEME_COLORS["border"]}; white-space:nowrap;
        position:sticky; top:0; z-index:1;
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
    .stco-sortable {{ cursor:pointer; user-select:none; }}
    .stco-sortable:hover {{ color:{THEME_COLORS["foreground"]}; }}
    .stco-sort-arrow {{ font-size:11px; margin-left:2px; }}
    .stco-sortable.stco-sort-active {{ color:{THEME_COLORS["foreground"]}; }}
    .stco-sortable.stco-sort-active .stco-sort-arrow {{ font-weight:700; }}
    </style>
    <div class="stco-table-wrap">
    <table class="stco-table" id="stco-creative-table">
      <thead><tr>{thead}</tr></thead>
      <tbody>{''.join(row_htmls)}</tbody>
    </table>
    </div>
    <script>
    (function() {{
        var table = document.getElementById("stco-creative-table");
        var tbody = table.querySelector("tbody");
        var headers = table.querySelectorAll("th.stco-sortable");
        var state = {{ key: null, dir: 1 }};

        headers.forEach(function(th) {{
            th.addEventListener("click", function() {{
                var key = th.getAttribute("data-key");
                var colIndex = Array.prototype.indexOf.call(th.parentNode.children, th);
                if (state.key === key) {{
                    state.dir = -state.dir;
                }} else {{
                    state.key = key;
                    state.dir = -1;  // 처음 클릭하면 내림차순(큰 값 먼저)
                }}
                headers.forEach(function(h) {{
                    h.classList.remove("stco-sort-active");
                    h.querySelector(".stco-sort-arrow").textContent = "⇅";
                }});
                th.classList.add("stco-sort-active");
                th.querySelector(".stco-sort-arrow").textContent = state.dir === 1 ? "▲" : "▼";

                var rows = Array.prototype.slice.call(tbody.querySelectorAll("tr"));
                var totalRows = rows.filter(function(r) {{ return r.classList.contains("stco-total-row"); }});
                var dataRows = rows.filter(function(r) {{ return !r.classList.contains("stco-total-row"); }});

                dataRows.sort(function(a, b) {{
                    var av = parseFloat(a.children[colIndex].getAttribute("data-value"));
                    var bv = parseFloat(b.children[colIndex].getAttribute("data-value"));
                    var aNum = isNaN(av) ? -Infinity : av;
                    var bNum = isNaN(bv) ? -Infinity : bv;
                    return (aNum - bNum) * state.dir;
                }});

                totalRows.forEach(function(r) {{ tbody.appendChild(r); }});
                dataRows.forEach(function(r) {{ tbody.appendChild(r); }});
                for (var i = totalRows.length - 1; i >= 0; i--) {{
                    tbody.insertBefore(totalRows[i], tbody.firstChild);
                }}
            }});
        }});
    }})();
    </script>
    """
    components.html(html, height=iframe_height, scrolling=True)


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
        narrow_col, _spacer = st.columns([3, 9])
        with narrow_col:
            date_range = st.date_input("기간 직접 선택", value=(min_d, max_d), min_value=min_d, max_value=max_d, key=f"{key}_manual")
        start, end = date_range if isinstance(date_range, tuple) and len(date_range) == 2 else (min_d, max_d)
        view_all = d[(d[date_col].dt.date >= start) & (d[date_col].dt.date <= end)]
    else:  # "YYYY년"
        y = int(preset.replace("년", ""))
        view_all = d[d[date_col].dt.year == y]

    total = len(view_all)
    if need_pagination and total > PAGE_SIZE_OPTIONS[0]:
        narrow_ps_col, _ps_spacer = st.columns([2, 10])
        with narrow_ps_col:
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
            st.write(
                f"🎨 소재별 성과 인식: {len(result.get('creatives', []))}행 "
                f"({', '.join(result.get('creative_sheets_found', [])) or '해당 시트 없음'})"
            )
            audience_df = result.get("channel_audience", pd.DataFrame())
            unclassified_n = (
                int((audience_df["audience_type"] == "미분류").sum()) if not audience_df.empty else 0
            )
            st.write(f"🧭 타겟팅별 성과(신규/리타겟팅) 그룹 인식: {len(audience_df)}행")
            if unclassified_n:
                st.warning(
                    f"'미분류'로 남은 그룹 {unclassified_n}행이 있습니다 — 새 캠페인/그룹이 추가된 것일 수 있어요. "
                    "app.py의 AUDIENCE_TYPE_MAP에 분류를 추가해주세요."
                )
            missing = set(result["channel_sheets_found"]) - set(result["channel_sheets_parsed"])
            if missing:
                st.warning(f"인식 실패한 매체 시트: {', '.join(missing)}")
            status.update(label="분석 완료", state="complete")

        if st.sidebar.button("💾 전체 저장하기", type="primary"):
            creatives_df = result.get("creatives", pd.DataFrame())
            if not creatives_df.empty:
                with st.sidebar.status("🖼️ 소재 이미지 추출/업로드 중...", expanded=False):
                    sheets_and_channels = [
                        (s, _infer_channel_from_sheet(s)) for s in result.get("creative_sheets_found", [])
                    ]
                    images = extract_creative_images(file, sheets_and_channels)
                    image_urls, image_errors = upload_creative_images(images)
                    creatives_df = attach_creative_images(creatives_df, image_urls)
                    if images:
                        st.write(f"소재 이미지 {len(images)}개 인식, {len(image_urls)}개 업로드 성공")
                        if image_errors:
                            st.warning("업로드 실패 예시:\n" + "\n".join(image_errors))

            n1 = save_table("weekly_overview", result["weekly"], "week_start", file.name)
            n2 = save_table("monthly_overview", result["monthly"], "report_month", file.name)
            n3 = save_table("channel_monthly", result["channels"], "report_month,channel", file.name)
            n4 = save_table("channel_snapshot", result["channel_snapshot"], "as_of_month,channel", file.name)
            n5 = save_table("ga_source", result["ga"], "as_of_date,source_medium", file.name)
            n6 = save_table("daily_overview", result["daily"], "report_date", file.name)
            delete_creative_performance_for_date(today)
            n7 = save_table(
                "creative_performance", creatives_df,
                "as_of_date,channel,creative", file.name,
            )
            delete_channel_audience_for_date(today)
            n8 = save_table(
                "channel_audience_snapshot", result.get("channel_audience", pd.DataFrame()),
                "as_of_date,channel,audience_type", file.name,
            )
            st.cache_data.clear()
            st.sidebar.success(
                f"저장 완료! 주간 {n1} · 월별 {n2} · 일자별 {n6} · 매체(월) {n3} · 매체(당월) {n4} · "
                f"GA {n5}건 · 소재 {n7}건 · 퍼널(신규/리타겟) {n8}건"
            )
            st.rerun()

    st.sidebar.markdown("---")
    wk = load_table("weekly_overview")
    st.sidebar.metric("누적 주간 데이터", f"{len(wk):,} 주")
    if st.sidebar.button("🔄 새로고침 (캐시 비우기)"):
        st.cache_data.clear()
        st.rerun()


# ──────────────────────────────────────────────────────────────
# 채널믹스 목표 대비 (신규) — "매체별 성과" 페이지 안에서 호출
# ──────────────────────────────────────────────────────────────
def render_channel_mix(fc: pd.DataFrame):
    st.markdown("---")
    st.markdown("### 채널믹스 목표 대비 (발굴 메타50·구글20·네이버20 / 회수 모비온10)")
    st.caption(
        "목표 비중은 「STCO 퍼포먼스마케팅 인수인계서」 4절 기준. "
        "'기타'는 목표 비중 배정이 없는 매체(크리테오 등)입니다."
    )
    if fc is None or fc.empty:
        st.info("데이터가 아직 없습니다.")
        return

    grp = fc.copy()
    grp["채널그룹"] = grp["channel"].apply(map_channel_group)
    by_group = grp.groupby("채널그룹", as_index=False)["cost_incl_vat"].sum()
    total_cost = by_group["cost_incl_vat"].sum()
    if total_cost <= 0:
        st.info("광고비 데이터가 없습니다.")
        return

    all_groups = list(CHANNEL_MIX_TARGET.keys())
    if (by_group["채널그룹"] == "기타").any():
        all_groups = all_groups + ["기타"]
    by_group = (
        by_group.set_index("채널그룹").reindex(all_groups, fill_value=0).reset_index()
    )

    by_group["실제비중(%)"] = by_group["cost_incl_vat"] / total_cost * 100
    by_group["목표비중(%)"] = by_group["채널그룹"].map(lambda g: CHANNEL_MIX_TARGET.get(g, 0) * 100)
    by_group["차이(%p)"] = by_group["실제비중(%)"] - by_group["목표비중(%)"]

    chart_df = by_group.melt(
        id_vars="채널그룹", value_vars=["목표비중(%)", "실제비중(%)"], var_name="구분", value_name="비중(%)"
    )
    fig = px.bar(
        chart_df, x="채널그룹", y="비중(%)", color="구분", barmode="group", text_auto=".1f",
        title="채널그룹별 목표 vs 실제 예산 비중",
    )
    st.plotly_chart(theme_chart(fig), use_container_width=True)

    show = by_group.copy()
    show["광고비(원)"] = show["cost_incl_vat"].map(lambda v: f"{v:,.0f}")
    show["실제비중(%)"] = show["실제비중(%)"].map(lambda v: f"{v:.1f}%")
    show["목표비중(%)"] = show["목표비중(%)"].map(lambda v: f"{v:.1f}%")
    show["차이(%p)"] = show["차이(%p)"].map(lambda v: f"{v:+.1f}%p")
    render_html_table(show[["채널그룹", "광고비(원)", "목표비중(%)", "실제비중(%)", "차이(%p)"]])

    over = by_group[(by_group["채널그룹"] != "기타") & (by_group["차이(%p)"] > 5)]
    under = by_group[(by_group["채널그룹"] != "기타") & (by_group["차이(%p)"] < -5)]
    if len(over) or len(under):
        msgs = []
        for _, r in over.iterrows():
            msgs.append(f"- **{r['채널그룹']}** 목표 대비 +{r['차이(%p)']:.1f}%p 초과 집행 중 — 단계적 축소(10~15% 내) 검토")
        for _, r in under.iterrows():
            msgs.append(f"- **{r['채널그룹']}** 목표 대비 {r['차이(%p)']:.1f}%p 미달 — 단계적 증액 검토")
        st.markdown("\n".join(msgs))


# ──────────────────────────────────────────────────────────────
# 소재별 성과 (신규 페이지)
# ──────────────────────────────────────────────────────────────
# 현재 실제로 운영 중인 매체 탭만 고정 순서로 노출 (TOTAL이 맨 왼쪽)
CREATIVE_TABS = ["TOTAL", "네이버 GFA PC", "네이버 GFA MO", "메타", "구글(P-MAX)", "크리테오"]


def _render_creative_table(fc: pd.DataFrame):
    """선택된 탭(매체)의 필터링된 소재 데이터로 집계 테이블 + 판정 + 다운로드 버튼을 렌더링."""
    if fc.empty:
        st.info("선택한 기간/매체에 데이터가 없습니다.")
        return

    has_image = "image_url" in fc.columns

    # 주간 리포트가 '해당 월 1일~업로드일까지 누적' 형태라, 같은 달 안에 여러 번(매주) 업로드하면
    # 소재마다 as_of_date가 다른 여러 스냅샷이 쌓인다. 이걸 그대로 sum하면 매주 누적치를 또 더해서
    # 몇 배로 부풀려지므로, 기간 내 '가장 최근 스냅샷' 한 건만 남기고(이미 그 시점까지의 누적값이라
    # 합산이 아니라 대체) 집계한다.
    fc_latest = fc.sort_values("as_of_date").drop_duplicates(subset=["channel", "creative"], keep="last")

    agg_kwargs = dict(
        impressions=("impressions", "sum"), clicks=("clicks", "sum"),
        cost_excl_vat=("cost_excl_vat", "sum"), cost_incl_vat=("cost_incl_vat", "sum"),
        conversions=("conversions", "sum"), revenue=("revenue", "sum"),
    )
    if has_image:
        agg_kwargs["image_url"] = ("image_url", "first")
    agg = fc_latest.groupby(["channel", "creative"], as_index=False).agg(**agg_kwargs)
    agg = add_kpis(agg)

    MIN_SPEND = 50000  # 표본 기준: 광고비 5만원 미만은 판단 보류 (performance-marketing-analysis 스킬 기본값)
    total_cost, total_rev = agg["cost_incl_vat"].sum(), agg["revenue"].sum()
    account_avg_roas = (total_rev / total_cost * 100) if total_cost else 0

    def judge(row):
        if row["cost_incl_vat"] < MIN_SPEND:
            return "판단 보류(표본 부족)"
        if account_avg_roas <= 0:
            return "판단 보류"
        ratio = row["roas"] / account_avg_roas
        if ratio >= 1.2:
            return "우수"
        if ratio <= 0.7:
            return "부진"
        return "평균 수준"

    agg["판정"] = agg.apply(judge, axis=1)
    agg = agg.sort_values("cost_incl_vat", ascending=False)

    st.caption(f"계정 평균 ROAS(선택 기간): {account_avg_roas:,.0f}% · 광고비 {MIN_SPEND:,}원 미만은 표본 부족으로 판단 보류 처리")

    display_cols = ["channel", "creative"]
    if has_image:
        # render_html_table은 셀 값을 그대로 <td>에 넣으므로 <img> 태그 문자열이 실제 썸네일로 렌더링된다.
        agg["creative_image"] = agg["image_url"].map(
            lambda u: (
                f'<img src="{u}" style="height:140px;width:140px;border-radius:8px;object-fit:cover;">'
                if isinstance(u, str) and u else ""
            )
        )
        display_cols.append("creative_image")
    display_cols += ["impressions", "clicks", "ctr", "cpc", "cost_incl_vat",
                      "conversions", "cvr", "cpa", "revenue", "roas", "판정"]
    display_cols = [c for c in display_cols if c in agg.columns]

    show = format_display(agg[display_cols])

    # TOTAL 행: 선택된 탭/기간 전체 합계를 표 맨 위(헤더 바로 아래)에 음영 처리해서 보여준다.
    # render_html_table은 첫 컬럼이 정확히 "TOTAL"이면 자동으로 강조 스타일(음영/굵게)을 입힌다.
    total_base = pd.DataFrame([{
        "impressions": agg["impressions"].sum(), "clicks": agg["clicks"].sum(),
        "cost_excl_vat": agg["cost_excl_vat"].sum(), "cost_incl_vat": agg["cost_incl_vat"].sum(),
        "conversions": agg["conversions"].sum(), "revenue": agg["revenue"].sum(),
    }])
    total_kpis = add_kpis(total_base).iloc[0]
    total_row = {c: "" for c in display_cols}
    total_row["channel"] = "TOTAL"
    for c in ("impressions", "clicks", "ctr", "cpc", "cost_incl_vat", "conversions", "cvr", "cpa", "revenue", "roas"):
        if c in display_cols:
            total_row[c] = total_kpis[c]
    total_show = format_display(pd.DataFrame([total_row])[display_cols])
    show = pd.concat([total_show, show], ignore_index=True)

    render_sortable_creative_table(korify(show), agg)

    # 엑셀 다운로드에는 이미지 썸네일(HTML) 대신 원본 URL을 남긴다.
    dl_cols = [c for c in display_cols if c != "creative_image"]
    if has_image and "image_url" in agg.columns and "image_url" not in dl_cols:
        dl_cols.insert(dl_cols.index("creative") + 1, "image_url")
    st.download_button(
        "⬇️ 엑셀 다운로드 (소재별 성과)",
        data=to_excel_bytes(korify(format_display(agg[dl_cols]))),
        file_name="creative_performance.xlsx",
        key=f"dl_creative_{fc['channel'].iloc[0] if fc['channel'].nunique() == 1 else 'total'}",
    )


def render_creative_performance(creatives: pd.DataFrame):
    st.subheader("🎨 소재별 성과")
    st.caption("현재 운영 중인 매체(네이버 GFA PC/MO, 메타, 구글 P-MAX, 크리테오)의 소재 단위 데이터만 표시됩니다.")
    if creatives is None or creatives.empty:
        st.info(
            "소재별 데이터가 아직 없습니다. 엑셀에 '○○_소재성과요약' 시트 또는 "
            "'광고소재' 컬럼이 있는 '_data' 시트가 있는지 확인해주세요."
        )
        return

    creatives = creatives.copy()
    creatives["as_of_date"] = pd.to_datetime(creatives["as_of_date"]).dt.date
    min_d, max_d = creatives["as_of_date"].min(), creatives["as_of_date"].max()
    start, end = period_filter(min_d, max_d, key="creative")
    fc = creatives[(creatives["as_of_date"] >= start) & (creatives["as_of_date"] <= end)]
    # 예전 업로드분에 남아있을 수 있는 미운영 매체(당근마켓 등) 잔여 행은 TOTAL에서도 제외
    fc = fc[fc["channel"].isin(CREATIVE_TABS[1:])]

    tabs = st.tabs(CREATIVE_TABS)
    for tab_widget, tab_name in zip(tabs, CREATIVE_TABS):
        with tab_widget:
            tab_fc = fc if tab_name == "TOTAL" else fc[fc["channel"] == tab_name]
            _render_creative_table(tab_fc)


# ──────────────────────────────────────────────────────────────
# 사이드바 그룹 네비게이션 (신규 — st.tabs() 대체)
# ──────────────────────────────────────────────────────────────
NAV_GROUPS = {
    "성과 리포트": ["종합 대시보드", "매체별 성과", "타겟팅별 성과", "소재별 성과"],
    "GA 유입 리포트": ["GA 유입경로", "GA4 라이브 리포트", "채널별 유입 분석"],
    "운영 도구": ["UTM 빌더", "소재 로그", "예산 재배분", "마일스톤"],
    "가이드": ["가이드"],
}

# 아직 실제 데이터/로직이 없는 페이지들 — main()의 페이지 분기에서 이 목록에 있으면
# render_coming_soon()으로 "준비 중" 안내만 보여준다. 나중에 진짜 렌더 함수가 생기면
# main()에 elif 분기를 추가하고 여기서 이름을 지워주면 된다.
NAV_PAGES_COMING_SOON = {"마일스톤", "UTM 빌더", "소재 로그", "예산 재배분", "가이드", "채널별 유입 분석"}


def render_coming_soon(page_name: str):
    st.subheader(page_name)
    st.info(f"'{page_name}' 페이지는 아직 준비 중입니다. 데이터 연동이 되면 채워질 예정이에요.")


def render_nav() -> str:
    st.sidebar.markdown("---")
    default_page = NAV_GROUPS["성과 리포트"][0]
    if "nav_page" not in st.session_state:
        st.session_state["nav_page"] = default_page
    # key="stco_nav"로 감싸면 Streamlit이 이 블록에 "st-key-stco_nav" 클래스를 붙여주는데,
    # inject_theme()의 CSS가 그 클래스 안의 expander/button만 골라 '텍스트 링크' 스타일로 바꾼다
    # (사이드바 위쪽의 업로드/저장 버튼 등 다른 버튼들은 기존 스타일 그대로 유지됨).
    # container(key=...)는 비교적 최신 Streamlit에서만 지원하므로, 혹시 배포 환경 버전이 낮아
    # TypeError가 나더라도 내비게이션 자체는 동작하도록 안전하게 폴백한다(이 경우 스타일만 예전
    # 버튼 박스 모양으로 보임).
    try:
        nav_box = st.sidebar.container(key="stco_nav")
    except TypeError:
        nav_box = st.sidebar.container()
    with nav_box:
        for group, pages in NAV_GROUPS.items():
            # 아이콘은 텍스트가 아니라 CSS(summary p::before)로 붙인다 — NAV_GROUP_ICON_B64 참고.
            # 그룹이 여러 개가 되면 전부 펼쳐두면 사이드바가 너무 길어지니, 현재 선택된 페이지가
            # 속한 그룹만 자동으로 펼치고 나머지는 접어둔다.
            is_active_group = st.session_state["nav_page"] in pages
            with st.expander(group, expanded=is_active_group):
                for p in pages:
                    is_current = st.session_state["nav_page"] == p
                    if st.button(
                        p, key=f"nav_{p}", use_container_width=True,
                        type="primary" if is_current else "secondary",
                    ):
                        st.session_state["nav_page"] = p
                        st.rerun()
    return st.session_state["nav_page"]


# ──────────────────────────────────────────────────────────────
# 페이지별 렌더 함수 (예전 tab1~tab4의 내용을 그대로 옮김, 로직 변경 없음)
# ──────────────────────────────────────────────────────────────
def render_overview_page(weekly: pd.DataFrame, monthly: pd.DataFrame, daily: pd.DataFrame):
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
    else:
        st.info("주간 데이터가 아직 없습니다.")


def render_channel_page(channels: pd.DataFrame, snapshot: pd.DataFrame):
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

        by_channel_chart = by_channel.copy()
        by_channel_chart["roas_label"] = by_channel_chart["roas"].map(lambda v: f"{v:.0f}%")
        fig = px.bar(
            by_channel_chart, x="channel", y="roas", title="매체별 ROAS (%, 선택 기간 합산)", text="roas_label",
            labels={"channel": "매체", "roas": "ROAS(%)"},
        )
        st.plotly_chart(theme_chart(fig), use_container_width=True)

        # 광고비(VAT제외)는 CPA 계산에만 쓰고 화면/엑셀에는 광고비(VAT+)만 노출 — 컬럼이 많아
        # 표 폭이 들쭉날쭉해지는 것도 줄어든다. 컬럼 순서는 종합 대시보드/타겟팅별 성과와 동일하게
        # 노출수 → 클릭수 → CTR → CPC → 광고비 순으로 맞춘다.
        BC_COL_ORDER = ["channel", "impressions", "clicks", "ctr", "cpc", "cost_incl_vat",
                        "cpa", "conversions", "cvr", "revenue", "roas", "aov"]
        bc_cols = [c for c in BC_COL_ORDER if c in by_channel.columns]
        bc_cols += [c for c in by_channel.columns if c not in bc_cols and c != "cost_excl_vat"]
        bc_table = format_display(by_channel[bc_cols])
        # CPA는 광고비(VAT제외) 기준으로 계산되는데, bc_cols에서 그 컬럼을 표시에서 뺐다고
        # TOTAL 합산용 원본까지 잘라버리면 cost_excl_vat이 0 취급돼 CPA가 0으로 나온다.
        # 합산은 by_channel 전체 컬럼 기준으로 하고, 화면에 낼 항목만 bc_cols로 고른다.
        bc_total = build_total_row(by_channel, bc_cols, "channel", label_text="TOTAL")
        if bc_total:
            bc_table = pd.concat([bc_table, pd.DataFrame([bc_total])], ignore_index=True)
        render_html_table(korify(bc_table), raw=by_channel[bc_cols])
        st.download_button(
            "⬇️ 엑셀 다운로드 (매체별·월별)",
            data=to_excel_bytes(korify(format_display(by_channel[bc_cols]))), file_name="channel_performance.xlsx",
        )

        render_channel_mix(fc)  # ← 신규: 채널믹스 목표 대비
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


# "리타겟팅 매체(자사몰)"이 아니라 "리타겟팅 매체(네이버 스토어)"로 따로 빼는 채널.
# 네이버 쇼핑검색광고는 클릭 도착 페이지가 자사몰이 아니라 네이버 스토어(스마트스토어)라 구분한다.
TARGETING_STORE_CHANNELS = {"네이버 쇼핑검색광고"}

# 사용자가 정리해준 리포트 순서(비용순 정렬이 아니라 매체 관례상 고정 순서) — 목록에 없는
# 채널은 뒤에 광고비 내림차순으로 붙는다.
TARGETING_NEW_CHANNEL_ORDER = ["네이버 GFA PC", "네이버 GFA MO", "메타", "구글(P-MAX)", "네이버 맨즈탭"]
TARGETING_RETARGET_OWN_CHANNEL_ORDER = ["네이버 검색광고", "네이버 브랜드검색광고", "네이버 GFA PC", "네이버 GFA MO", "메타", "크리테오"]

TARGETING_CORE_COLS = ["channel", "impressions", "clicks", "ctr", "cpc", "cost_incl_vat",
                        "signups", "signup_rate", "conversions", "revenue", "roas"]


def _targeting_order_channels(df: pd.DataFrame, order_list: list) -> pd.DataFrame:
    df = df.copy()
    df["_ord"] = df["channel"].apply(lambda c: order_list.index(c) if c in order_list else len(order_list) + 1)
    return df.sort_values(["_ord", "cost_incl_vat"], ascending=[True, False]).drop(columns="_ord")


def _targeting_total_row(df: pd.DataFrame, label: str = "TOTAL") -> pd.DataFrame:
    row = {
        "channel": label,
        "impressions": df["impressions"].sum(), "clicks": df["clicks"].sum(),
        "cost_incl_vat": df["cost_incl_vat"].sum(), "signups": df["signups"].sum(),
        "conversions": df["conversions"].sum(), "revenue": df["revenue"].sum(),
    }
    row["ctr"] = (row["clicks"] / row["impressions"] * 100) if row["impressions"] else 0
    row["cpc"] = (row["cost_incl_vat"] / row["clicks"]) if row["clicks"] else 0
    row["signup_rate"] = (row["signups"] / row["clicks"] * 100) if row["clicks"] else 0
    row["roas"] = (row["revenue"] / row["cost_incl_vat"] * 100) if row["cost_incl_vat"] else 0
    return pd.DataFrame([row])[TARGETING_CORE_COLS]


def render_targeting_performance_page(audience: pd.DataFrame, creatives_fallback: pd.DataFrame = None):
    """타겟팅별 성과 — 신규 타겟팅 vs 리타겟팅을 서로 겹치지 않게(mutually exclusive) 나눠서 본다.
    ① TOTAL 비교(TOTAL = 신규 타겟팅 + 리타겟팅) → ② 매체별 성과(신규 타겟팅 매체(자사몰) /
    리타겟팅 매체(자사몰) / 리타겟팅 매체(네이버 스토어)) → ③ 매체별 ROAS 비교, 3단 구성.
    '발굴/회수/수확' 역할 구분이나 CAC 중심 판정은 쓰지 않는다 — 대행사 운영 + 매체 예산이
    이미 고정된 구조에는 그 프레임이 안 맞는다는 판단에 따른 것."""
    if audience.empty and (creatives_fallback is None or creatives_fallback.empty):
        st.info(
            "아직 데이터가 없습니다. 주간 리포트를 업로드하면 캠페인 그룹명을 기준으로 "
            "신규/리타겟팅이 자동 분류돼 채워집니다."
        )
        return

    audience = audience.copy()
    if not audience.empty:
        audience["as_of_date"] = pd.to_datetime(audience["as_of_date"]).dt.date
        min_d, max_d = audience["as_of_date"].min(), audience["as_of_date"].max()
    else:
        cf = creatives_fallback.copy()
        cf["as_of_date"] = pd.to_datetime(cf["as_of_date"]).dt.date
        min_d, max_d = cf["as_of_date"].min(), cf["as_of_date"].max()
    start, end = period_filter(min_d, max_d, key="targeting")
    fa = audience[(audience["as_of_date"] >= start) & (audience["as_of_date"] <= end)] if not audience.empty else audience
    # 소재별 성과와 같은 이유로, 같은 달 안에 여러 번 업로드해도 채널×오디언스별 '최신 스냅샷'만 사용
    # (누적 리포트를 그대로 더하면 몇 배로 부풀려짐).
    if not fa.empty:
        fa = fa.sort_values("as_of_date").drop_duplicates(subset=["channel", "audience_type"], keep="last")

    # 구글(P-MAX)은 리포트 원본에 그룹(오디언스)별 집계 자체가 비어있어(전부 0) 위 표에는 안 잡힌다.
    # 사용자가 확인해준 매핑상 PMax는 항상 전체 신규이므로, 이미 검증된 소재별 성과(creative_performance)
    # 합계를 '신규'로 채워 넣는다 (회원가입 수는 이 표에서 안 잡히는 값이라 0으로 남는다).
    if creatives_fallback is not None and not creatives_fallback.empty:
        already_has_google = (not fa.empty) and (fa["channel"] == "구글(P-MAX)").any()
        if not already_has_google:
            g = creatives_fallback.copy()
            g["as_of_date"] = pd.to_datetime(g["as_of_date"]).dt.date
            g = g[(g["channel"] == "구글(P-MAX)") & (g["as_of_date"] >= start) & (g["as_of_date"] <= end)]
            if not g.empty:
                g_latest = g.sort_values("as_of_date").drop_duplicates(subset=["channel", "creative"], keep="last")
                synth = pd.DataFrame([{
                    "channel": "구글(P-MAX)", "audience_type": "신규",
                    "impressions": g_latest["impressions"].sum(), "clicks": g_latest["clicks"].sum(),
                    "cost_excl_vat": g_latest["cost_excl_vat"].sum(), "cost_incl_vat": g_latest["cost_incl_vat"].sum(),
                    "signups": 0, "conversions": g_latest["conversions"].sum(), "revenue": g_latest["revenue"].sum(),
                }])
                fa = pd.concat([fa, synth], ignore_index=True) if not fa.empty else synth

    if fa.empty:
        st.info("선택한 기간에 데이터가 없습니다.")
        return

    unclassified = fa[fa["audience_type"] == "미분류"]
    if not unclassified.empty:
        st.warning(
            f"'미분류'로 남은 그룹이 {len(unclassified)}건 있습니다(신규 캠페인/그룹이 추가됐을 수 있어요). "
            "신규·리타겟팅 어느 쪽에도 안 잡히고 아래 표에 별도로만 표시됩니다."
        )

    agg = (
        fa.groupby(["channel", "audience_type"], as_index=False)
        .agg(impressions=("impressions", "sum"), clicks=("clicks", "sum"),
             cost_excl_vat=("cost_excl_vat", "sum"), cost_incl_vat=("cost_incl_vat", "sum"),
             signups=("signups", "sum"), conversions=("conversions", "sum"), revenue=("revenue", "sum"))
    )
    agg = add_kpis(agg)
    agg["signup_rate"] = np.where(agg["clicks"] > 0, agg["signups"] / agg["clicks"] * 100, 0)

    # 네이버 맨즈탭 — 아직 별도 리포트 연동 전이라 데이터가 없다. 항목은 항상 노출하되 0으로 표시
    # (나중에 전용 리포트를 업로드받으면 이 자리에 실제 값이 채워지게 될 예정).
    if "네이버 맨즈탭" not in agg["channel"].values:
        placeholder = pd.DataFrame([{
            "channel": "네이버 맨즈탭", "audience_type": "신규",
            "impressions": 0.0, "clicks": 0.0, "cost_excl_vat": 0.0, "cost_incl_vat": 0.0,
            "signups": 0.0, "conversions": 0.0, "revenue": 0.0,
            "ctr": 0.0, "cpc": 0.0, "cpa": 0.0, "cvr": 0.0, "roas": 0.0, "aov": 0.0, "signup_rate": 0.0,
        }])
        agg = pd.concat([agg, placeholder], ignore_index=True)

    # ── ① 타겟팅별 성과 TOTAL 비교 (TOTAL = 신규 타겟팅 + 리타겟팅) ──
    st.markdown("##### ① 타겟팅별 성과 TOTAL 비교")
    AUDIENCE_ORDER = ["신규", "리타겟팅", "미분류"]
    AUDIENCE_LABEL = {"신규": "신규 타겟팅", "리타겟팅": "리타겟팅", "미분류": "미분류"}
    audience_summary = (
        agg.groupby("audience_type", as_index=False)
        .agg(impressions=("impressions", "sum"), clicks=("clicks", "sum"),
             cost_excl_vat=("cost_excl_vat", "sum"), cost_incl_vat=("cost_incl_vat", "sum"),
             signups=("signups", "sum"), conversions=("conversions", "sum"), revenue=("revenue", "sum"))
    )
    audience_summary = add_kpis(audience_summary)
    audience_summary["signup_rate"] = np.where(
        audience_summary["clicks"] > 0, audience_summary["signups"] / audience_summary["clicks"] * 100, 0
    )
    audience_summary["_order"] = audience_summary["audience_type"].apply(
        lambda x: AUDIENCE_ORDER.index(x) if x in AUDIENCE_ORDER else 99
    )
    audience_summary = audience_summary.sort_values("_order")
    audience_summary = audience_summary.rename(columns={"audience_type": "channel"})
    audience_summary["channel"] = audience_summary["channel"].map(AUDIENCE_LABEL).fillna(audience_summary["channel"])
    show_top = format_display(audience_summary[TARGETING_CORE_COLS])
    total_top = format_display(_targeting_total_row(audience_summary, "TOTAL"))
    # ②의 신규/리타겟팅 매체 표와 같은 폭으로 보이도록 동일하게 2단 컬럼의 절반 폭에 맞춘다
    # (표 하나만 있다고 전체 페이지 폭으로 늘어나면 컬럼 간 간격이 과하게 벌어져 보인다).
    top_col, _spacer_col = st.columns(2)
    with top_col:
        render_html_table(korify(pd.concat([total_top, show_top], ignore_index=True)))
    st.caption("TOTAL = 신규 타겟팅 + 리타겟팅 (겹치지 않는 완전 분리 기준)")

    st.markdown("##### ② 타겟팅별 매체별 성과 비교")
    new_df = agg[agg["audience_type"].isin(["신규", "미분류"])]
    retarget_df = agg[agg["audience_type"] == "리타겟팅"]
    retarget_store_df = retarget_df[retarget_df["channel"].isin(TARGETING_STORE_CHANNELS)]
    retarget_own_df = retarget_df[~retarget_df["channel"].isin(TARGETING_STORE_CHANNELS)]

    col_new, col_re = st.columns(2)

    with col_new:
        st.markdown("###### 1) 신규 타겟팅 매체 (자사몰)")
        if new_df.empty:
            st.info("선택 기간에 신규 분류 데이터가 없습니다.")
        else:
            new_ordered = _targeting_order_channels(new_df, TARGETING_NEW_CHANNEL_ORDER)
            show_new = format_display(new_ordered[TARGETING_CORE_COLS])
            total_new = format_display(_targeting_total_row(new_ordered, "TOTAL"))
            render_html_table(korify(pd.concat([total_new, show_new], ignore_index=True)))
            st.download_button(
                "⬇️ 엑셀 다운로드 (신규 타겟팅 매체)",
                data=to_excel_bytes(korify(new_ordered[TARGETING_CORE_COLS])),
                file_name="targeting_new.xlsx",
            )

    with col_re:
        st.markdown("###### 2-1) 리타겟팅 매체 (자사몰)")
        if retarget_df.empty:
            st.info("선택 기간에 리타겟팅 분류 데이터가 없습니다.")
        else:
            own_ordered = _targeting_order_channels(retarget_own_df, TARGETING_RETARGET_OWN_CHANNEL_ORDER)
            show_own = format_display(own_ordered[TARGETING_CORE_COLS])
            # TOTAL은 자사몰만이 아니라 리타겟팅 전체(자사몰 + 네이버 스토어) 합계 — 사용자가
            # 확인해준 대로 ①의 TOTAL 비교 '리타겟팅' 행과 정확히 일치해야 한다.
            total_re = format_display(_targeting_total_row(retarget_df, "TOTAL"))
            render_html_table(korify(pd.concat([total_re, show_own], ignore_index=True)))

        st.markdown("###### 2-2) 리타겟팅 매체 (네이버 스토어)")
        if retarget_store_df.empty:
            st.info("아직 데이터가 없습니다.")
        else:
            store_ordered = retarget_store_df.sort_values("cost_incl_vat", ascending=False)
            render_html_table(korify(format_display(store_ordered[TARGETING_CORE_COLS])))

        if not retarget_df.empty:
            retarget_all = pd.concat([retarget_own_df, retarget_store_df], ignore_index=True)
            st.download_button(
                "⬇️ 엑셀 다운로드 (리타겟팅 매체, 자사몰+네이버 스토어)",
                data=to_excel_bytes(korify(retarget_all[["channel"] + TARGETING_CORE_COLS[1:]])),
                file_name="targeting_retarget.xlsx",
            )

    st.markdown("##### ③ 매체별 ROAS 성과 비교")
    chart_new, chart_re = st.columns(2)
    with chart_new:
        st.markdown(
            f"<div style='font-size:16px;font-weight:700;color:{THEME_COLORS['foreground']};margin-bottom:6px;'>신규 타겟팅 매체별 ROAS(%)</div>",
            unsafe_allow_html=True,
        )
        if not new_df.empty:
            new_chart_df = new_df.sort_values("roas", ascending=False)
            new_chart_df["roas_label"] = new_chart_df["roas"].map(lambda v: f"{v:.0f}%")
            fig_new = px.bar(
                new_chart_df, x="channel", y="roas", text="roas_label",
                labels={"channel": "매체", "roas": "ROAS(%)"},
            )
            st.plotly_chart(theme_chart(fig_new), use_container_width=True)
    with chart_re:
        st.markdown(
            f"<div style='font-size:16px;font-weight:700;color:{THEME_COLORS['foreground']};margin-bottom:6px;'>리타겟팅 매체별 ROAS(%) (자사몰 + 네이버 스토어)</div>",
            unsafe_allow_html=True,
        )
        if not retarget_df.empty:
            re_chart_df = retarget_df.sort_values("roas", ascending=False)
            re_chart_df["roas_label"] = re_chart_df["roas"].map(lambda v: f"{v:.0f}%")
            fig_re = px.bar(
                re_chart_df, x="channel", y="roas", text="roas_label",
                labels={"channel": "매체", "roas": "ROAS(%)"},
            )
            st.plotly_chart(theme_chart(fig_re), use_container_width=True)


def render_ga_page(ga: pd.DataFrame):
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


def render_ga4_page():
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
        scale = st.slider(
            "리포트 축소 비율", min_value=0.4, max_value=1.0, value=1.0, step=0.05, key="looker_scale",
        )
        # 폭을 줄이면 리포트 내부 표가 반응형으로 찌그러지며 컬럼이 잘리므로,
        # 원본 크기(1400x1000)로 그린 뒤 CSS로 통째로 축소(zoom-out)해서
        # 컬럼이 안 잘리고 더 많은 내용이 한 화면에 들어오게 한다.
        native_w, native_h = 1400, 1000
        disp_w, disp_h = int(native_w * scale), int(native_h * scale)
        st.markdown(
            f'<div style="width:{disp_w}px; height:{disp_h}px; overflow:hidden; margin:0 auto; border:1px solid #e5e8eb; border-radius:8px;">'
            f'<iframe src="{looker_embed_url}" width="{native_w}" height="{native_h}" '
            f'style="border:0; transform:scale({scale}); transform-origin:0 0;" allowfullscreen></iframe>'
            f'</div>',
            unsafe_allow_html=True,
        )


# ──────────────────────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────────────────────
def main():
    # st.title()은 raw HTML(이미지)을 못 받아서, 이모지(📊) 대신 커스텀 아이콘을 쓰려면
    # 직접 <h1>을 그려야 한다. inject_theme()의 전역 h1 스타일(글자크기/굵기/색)은
    # 태그 자체를 그대로 쓰기 때문에 동일하게 적용된다.
    st.markdown(
        f'<h1 style="display:flex;align-items:center;gap:10px;margin:0 0 0.5rem 0;">'
        f'<img src="data:image/png;base64,{PAGE_TITLE_ICON_B64}" style="height:34px;width:auto;" />'
        f'STCO 온라인팀 광고/마케팅 성과 대시보드</h1>',
        unsafe_allow_html=True,
    )
    render_upload_panel()

    weekly = load_table("weekly_overview")
    monthly = load_table("monthly_overview")
    daily = load_table("daily_overview")
    channels = load_table("channel_monthly")
    snapshot = load_table("channel_snapshot")
    ga = load_table("ga_source")
    creatives = load_table("creative_performance")
    audience = load_table("channel_audience_snapshot")

    if weekly.empty and monthly.empty:
        st.info("아직 저장된 데이터가 없습니다. 왼쪽 사이드바에서 주간 리포트 파일을 업로드하고 '전체 저장하기'를 눌러주세요.")
        return

    for df, col in [(weekly, "week_start"), (weekly, "week_end"), (monthly, "report_month"), (daily, "report_date")]:
        if not df.empty and col in df.columns:
            df[col] = pd.to_datetime(df[col]).dt.date

    page = render_nav()  # ← st.tabs() 대신 사이드바 그룹 네비게이션

    if page == "종합 대시보드":
        render_overview_page(weekly, monthly, daily)
    elif page == "매체별 성과":
        render_channel_page(channels, snapshot)
    elif page == "타겟팅별 성과":
        render_targeting_performance_page(audience, creatives_fallback=creatives)
    elif page == "소재별 성과":
        render_creative_performance(creatives)
    elif page == "GA 유입경로":
        render_ga_page(ga)
    elif page == "GA4 라이브 리포트":
        render_ga4_page()
    elif page in NAV_PAGES_COMING_SOON:
        render_coming_soon(page)


if __name__ == "__main__":
    main()
