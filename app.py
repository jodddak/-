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

# 그룹별 커스텀 아이콘(사용자가 준 SVG) — 성과 리포트/운영 도구/가이드 그룹 헤더 앞에 표시.
# GA 유입 리포트 그룹은 별도로 받은 아이콘이 없어서 기존 NAV_GROUP_ICON_B64를 그대로 쓴다.
NAV_ICON_PERFORMANCE_B64 = (
    "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgZmlsbD0ibm9u"
    "ZSI+PGcgZmlsdGVyPSJ1cmwoI2EpIj48cmVjdCB3aWR0aD0iNC43MzciIGhlaWdodD0iNC43MzciIHg9IjMiIHk9IjE2LjI2"
    "MyIgZmlsbD0idXJsKCNiKSIgcng9Ii45NDciLz48L2c+PGcgZmlsdGVyPSJ1cmwoI2MpIj48cGF0aCBmaWxsPSJ1cmwoI2Qp"
    "IiBkPSJNOS42MzIgMTQuMzY4YzAtLjUyMy40MjQtLjk0Ny45NDctLjk0N2gyLjg0MmMuNTIzIDAgLjk0Ny40MjQuOTQ3Ljk0"
    "N3Y1LjY4NWEuOTQ3Ljk0NyAwIDAgMS0uOTQ3Ljk0N0gxMC41OGEuOTQ3Ljk0NyAwIDAgMS0uOTQ3LS45NDd6Ii8+PC9nPjxn"
    "IGZpbHRlcj0idXJsKCNlKSI+PHBhdGggZmlsbD0idXJsKCNmKSIgZD0iTTE2LjI2MyAxMS41MjZjMC0uNTIzLjQyNC0uOTQ3"
    "Ljk0Ny0uOTQ3aDIuODQzYy41MjMgMCAuOTQ3LjQyNC45NDcuOTQ3djguNTI3YS45NDcuOTQ3IDAgMCAxLS45NDcuOTQ3SDE3"
    "LjIxYS45NDcuOTQ3IDAgMCAxLS45NDctLjk0N3oiLz48L2c+PGcgZmlsdGVyPSJ1cmwoI2cpIj48cGF0aCBmaWxsPSJ1cmwo"
    "I2gpIiBkPSJNMTguMTQ2IDcuNzUzQzE0Ljc2IDEyLjEzMyA3Ljc1MyAxMy4zMzQgNC4xNSAxMy40M2MtLjY0Ny4wMTctMS4x"
    "NS0uNTE3LTEuMTUtMS4xNjMgMC0uNjcxLjU0NS0xLjIxMSAxLjIxNC0xLjI2NiA1LjI1LS40MzUgMTAuMjM4LTMuNTM2IDEx"
    "LjU2OC01LjE1MWEuMDE0LjAxNCAwIDAgMCAwLS4wMThsLTEuMDY2LTEuMjhBLjk0Ny45NDcgMCAwIDEgMTUuNDQ0IDNoNC42"
    "MDljLjUyMyAwIC45NDcuNDI0Ljk0Ny45NDd2NC4zNDVjMCAuODQ0LTEuMDIgMS4yNjctMS42MTcuNjdsLTEuMjEtMS4yMTFh"
    "LjAyLjAyIDAgMCAwLS4wMjcuMDAyIi8+PC9nPjxkZWZzPjxsaW5lYXJHcmFkaWVudCBpZD0iYiIgeDE9IjUuMzY4IiB4Mj0i"
    "NS4zNjgiIHkxPSIxNi4yNjMiIHkyPSIyMSIgZ3JhZGllbnRVbml0cz0idXNlclNwYWNlT25Vc2UiPjxzdG9wIHN0b3AtY29s"
    "b3I9IiNGM0VFRUEiLz48c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiNFQ0U1REYiLz48L2xpbmVhckdyYWRpZW50Pjxs"
    "aW5lYXJHcmFkaWVudCBpZD0iZCIgeDE9IjEyIiB4Mj0iMTIiIHkxPSIxMy40MjEiIHkyPSIyMSIgZ3JhZGllbnRVbml0cz0i"
    "dXNlclNwYWNlT25Vc2UiPjxzdG9wIHN0b3AtY29sb3I9IiNGM0VFRUEiLz48c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9"
    "IiNFQ0U1REYiLz48L2xpbmVhckdyYWRpZW50PjxsaW5lYXJHcmFkaWVudCBpZD0iZiIgeDE9IjE4LjYzMiIgeDI9IjE4LjYz"
    "MiIgeTE9IjEwLjU3OSIgeTI9IjIxIiBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+PHN0b3Agc3RvcC1jb2xvcj0i"
    "I0YzRUVFQSIvPjxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iI0VDRTVERiIvPjwvbGluZWFyR3JhZGllbnQ+PGxpbmVh"
    "ckdyYWRpZW50IGlkPSJoIiB4MT0iMTIiIHgyPSIxMiIgeTE9IjEzLjQzMSIgeTI9IjMiIGdyYWRpZW50VW5pdHM9InVzZXJT"
    "cGFjZU9uVXNlIj48c3RvcCBzdG9wLWNvbG9yPSIjRjk4MDQ0Ii8+PHN0b3Agb2Zmc2V0PSIxIiBzdG9wLWNvbG9yPSIjRkVB"
    "OTdGIi8+PC9saW5lYXJHcmFkaWVudD48ZmlsdGVyIGlkPSJhIiB3aWR0aD0iNi42MzIiIGhlaWdodD0iNi42MzIiIHg9IjEu"
    "ODE2IiB5PSIxNS41NTMiIGNvbG9yLWludGVycG9sYXRpb24tZmlsdGVycz0ic1JHQiIgZmlsdGVyVW5pdHM9InVzZXJTcGFj"
    "ZU9uVXNlIj48ZmVGbG9vZCBmbG9vZC1vcGFjaXR5PSIwIiByZXN1bHQ9IkJhY2tncm91bmRJbWFnZUZpeCIvPjxmZUNvbG9y"
    "TWF0cml4IGluPSJTb3VyY2VBbHBoYSIgcmVzdWx0PSJoYXJkQWxwaGEiIHZhbHVlcz0iMCAwIDAgMCAwIDAgMCAwIDAgMCAw"
    "IDAgMCAwIDAgMCAwIDAgMTI3IDAiLz48ZmVPZmZzZXQgZHg9Ii0uMjM3IiBkeT0iLjIzNyIvPjxmZUdhdXNzaWFuQmx1ciBz"
    "dGREZXZpYXRpb249Ii40NzQiLz48ZmVDb21wb3NpdGUgaW4yPSJoYXJkQWxwaGEiIG9wZXJhdG9yPSJvdXQiLz48ZmVDb2xv"
    "ck1hdHJpeCB2YWx1ZXM9IjAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAuMDggMCIvPjxmZUJsZW5kIGlu"
    "Mj0iQmFja2dyb3VuZEltYWdlRml4IiByZXN1bHQ9ImVmZmVjdDFfZHJvcFNoYWRvd183NjFfNDE3ODMiLz48ZmVCbGVuZCBp"
    "bj0iU291cmNlR3JhcGhpYyIgaW4yPSJlZmZlY3QxX2Ryb3BTaGFkb3dfNzYxXzQxNzgzIiByZXN1bHQ9InNoYXBlIi8+PGZl"
    "Q29sb3JNYXRyaXggaW49IlNvdXJjZUFscGhhIiByZXN1bHQ9ImhhcmRBbHBoYSIgdmFsdWVzPSIwIDAgMCAwIDAgMCAwIDAg"
    "MCAwIDAgMCAwIDAgMCAwIDAgMCAxMjcgMCIvPjxmZU9mZnNldCBkeD0iLjc1OCIgZHk9Ii0uNDc0Ii8+PGZlR2F1c3NpYW5C"
    "bHVyIHN0ZERldmlhdGlvbj0iLjIzNyIvPjxmZUNvbXBvc2l0ZSBpbjI9ImhhcmRBbHBoYSIgazI9Ii0xIiBrMz0iMSIgb3Bl"
    "cmF0b3I9ImFyaXRobWV0aWMiLz48ZmVDb2xvck1hdHJpeCB2YWx1ZXM9IjAgMCAwIDAgMC44NTQ2NDIgMCAwIDAgMCAwLjg1"
    "NDY0MiAwIDAgMCAwIDAuODU0NjQyIDAgMCAwIDEgMCIvPjxmZUJsZW5kIGluMj0ic2hhcGUiIHJlc3VsdD0iZWZmZWN0Ml9p"
    "bm5lclNoYWRvd183NjFfNDE3ODMiLz48ZmVDb2xvck1hdHJpeCBpbj0iU291cmNlQWxwaGEiIHJlc3VsdD0iaGFyZEFscGhh"
    "IiB2YWx1ZXM9IjAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDEyNyAwIi8+PGZlT2Zmc2V0IGR4PSIuNDc0"
    "IiBkeT0iLjQ3NCIvPjxmZUdhdXNzaWFuQmx1ciBzdGREZXZpYXRpb249Ii4yMzciLz48ZmVDb21wb3NpdGUgaW4yPSJoYXJk"
    "QWxwaGEiIGsyPSItMSIgazM9IjEiIG9wZXJhdG9yPSJhcml0aG1ldGljIi8+PGZlQ29sb3JNYXRyaXggdmFsdWVzPSIwIDAg"
    "MCAwIDAuOTAyNjA3IDAgMCAwIDAgMC44NDIwMzEgMCAwIDAgMCAwLjgwMTY0NyAwIDAgMCAxIDAiLz48ZmVCbGVuZCBpbjI9"
    "ImVmZmVjdDJfaW5uZXJTaGFkb3dfNzYxXzQxNzgzIiByZXN1bHQ9ImVmZmVjdDNfaW5uZXJTaGFkb3dfNzYxXzQxNzgzIi8+"
    "PC9maWx0ZXI+PGZpbHRlciBpZD0iYyIgd2lkdGg9IjYuNjMyIiBoZWlnaHQ9IjkuNDc0IiB4PSI4LjQ0NyIgeT0iMTIuNzEi"
    "IGNvbG9yLWludGVycG9sYXRpb24tZmlsdGVycz0ic1JHQiIgZmlsdGVyVW5pdHM9InVzZXJTcGFjZU9uVXNlIj48ZmVGbG9v"
    "ZCBmbG9vZC1vcGFjaXR5PSIwIiByZXN1bHQ9IkJhY2tncm91bmRJbWFnZUZpeCIvPjxmZUNvbG9yTWF0cml4IGluPSJTb3Vy"
    "Y2VBbHBoYSIgcmVzdWx0PSJoYXJkQWxwaGEiIHZhbHVlcz0iMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAg"
    "MTI3IDAiLz48ZmVPZmZzZXQgZHg9Ii0uMjM3IiBkeT0iLjIzNyIvPjxmZUdhdXNzaWFuQmx1ciBzdGREZXZpYXRpb249Ii40"
    "NzQiLz48ZmVDb21wb3NpdGUgaW4yPSJoYXJkQWxwaGEiIG9wZXJhdG9yPSJvdXQiLz48ZmVDb2xvck1hdHJpeCB2YWx1ZXM9"
    "IjAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAuMDggMCIvPjxmZUJsZW5kIGluMj0iQmFja2dyb3VuZElt"
    "YWdlRml4IiByZXN1bHQ9ImVmZmVjdDFfZHJvcFNoYWRvd183NjFfNDE3ODMiLz48ZmVCbGVuZCBpbj0iU291cmNlR3JhcGhp"
    "YyIgaW4yPSJlZmZlY3QxX2Ryb3BTaGFkb3dfNzYxXzQxNzgzIiByZXN1bHQ9InNoYXBlIi8+PGZlQ29sb3JNYXRyaXggaW49"
    "IlNvdXJjZUFscGhhIiByZXN1bHQ9ImhhcmRBbHBoYSIgdmFsdWVzPSIwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAw"
    "IDAgMCAxMjcgMCIvPjxmZU9mZnNldCBkeD0iLjc1OCIgZHk9Ii0uNDc0Ii8+PGZlR2F1c3NpYW5CbHVyIHN0ZERldmlhdGlv"
    "bj0iLjIzNyIvPjxmZUNvbXBvc2l0ZSBpbjI9ImhhcmRBbHBoYSIgazI9Ii0xIiBrMz0iMSIgb3BlcmF0b3I9ImFyaXRobWV0"
    "aWMiLz48ZmVDb2xvck1hdHJpeCB2YWx1ZXM9IjAgMCAwIDAgMC44NTQ2NDIgMCAwIDAgMCAwLjg1NDY0MiAwIDAgMCAwIDAu"
    "ODU0NjQyIDAgMCAwIDEgMCIvPjxmZUJsZW5kIGluMj0ic2hhcGUiIHJlc3VsdD0iZWZmZWN0Ml9pbm5lclNoYWRvd183NjFf"
    "NDE3ODMiLz48ZmVDb2xvck1hdHJpeCBpbj0iU291cmNlQWxwaGEiIHJlc3VsdD0iaGFyZEFscGhhIiB2YWx1ZXM9IjAgMCAw"
    "IDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDEyNyAwIi8+PGZlT2Zmc2V0IGR4PSIuNDc0IiBkeT0iLjQ3NCIvPjxm"
    "ZUdhdXNzaWFuQmx1ciBzdGREZXZpYXRpb249Ii4yMzciLz48ZmVDb21wb3NpdGUgaW4yPSJoYXJkQWxwaGEiIGsyPSItMSIg"
    "azM9IjEiIG9wZXJhdG9yPSJhcml0aG1ldGljIi8+PGZlQ29sb3JNYXRyaXggdmFsdWVzPSIwIDAgMCAwIDAuOTAyNjA3IDAg"
    "MCAwIDAgMC44NDIwMzEgMCAwIDAgMCAwLjgwMTY0NyAwIDAgMCAxIDAiLz48ZmVCbGVuZCBpbjI9ImVmZmVjdDJfaW5uZXJT"
    "aGFkb3dfNzYxXzQxNzgzIiByZXN1bHQ9ImVmZmVjdDNfaW5uZXJTaGFkb3dfNzYxXzQxNzgzIi8+PC9maWx0ZXI+PGZpbHRl"
    "ciBpZD0iZSIgd2lkdGg9IjYuNjMyIiBoZWlnaHQ9IjEyLjMxNiIgeD0iMTUuMDc5IiB5PSI5Ljg2OCIgY29sb3ItaW50ZXJw"
    "b2xhdGlvbi1maWx0ZXJzPSJzUkdCIiBmaWx0ZXJVbml0cz0idXNlclNwYWNlT25Vc2UiPjxmZUZsb29kIGZsb29kLW9wYWNp"
    "dHk9IjAiIHJlc3VsdD0iQmFja2dyb3VuZEltYWdlRml4Ii8+PGZlQ29sb3JNYXRyaXggaW49IlNvdXJjZUFscGhhIiByZXN1"
    "bHQ9ImhhcmRBbHBoYSIgdmFsdWVzPSIwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAxMjcgMCIvPjxmZU9m"
    "ZnNldCBkeD0iLS4yMzciIGR5PSIuMjM3Ii8+PGZlR2F1c3NpYW5CbHVyIHN0ZERldmlhdGlvbj0iLjQ3NCIvPjxmZUNvbXBv"
    "c2l0ZSBpbjI9ImhhcmRBbHBoYSIgb3BlcmF0b3I9Im91dCIvPjxmZUNvbG9yTWF0cml4IHZhbHVlcz0iMCAwIDAgMCAwIDAg"
    "MCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMC4wOCAwIi8+PGZlQmxlbmQgaW4yPSJCYWNrZ3JvdW5kSW1hZ2VGaXgiIHJlc3Vs"
    "dD0iZWZmZWN0MV9kcm9wU2hhZG93Xzc2MV80MTc4MyIvPjxmZUJsZW5kIGluPSJTb3VyY2VHcmFwaGljIiBpbjI9ImVmZmVj"
    "dDFfZHJvcFNoYWRvd183NjFfNDE3ODMiIHJlc3VsdD0ic2hhcGUiLz48ZmVDb2xvck1hdHJpeCBpbj0iU291cmNlQWxwaGEi"
    "IHJlc3VsdD0iaGFyZEFscGhhIiB2YWx1ZXM9IjAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDEyNyAwIi8+"
    "PGZlT2Zmc2V0IGR4PSIuNzU4IiBkeT0iLS40NzQiLz48ZmVHYXVzc2lhbkJsdXIgc3RkRGV2aWF0aW9uPSIuMjM3Ii8+PGZl"
    "Q29tcG9zaXRlIGluMj0iaGFyZEFscGhhIiBrMj0iLTEiIGszPSIxIiBvcGVyYXRvcj0iYXJpdGhtZXRpYyIvPjxmZUNvbG9y"
    "TWF0cml4IHZhbHVlcz0iMCAwIDAgMCAwLjg1NDY0MiAwIDAgMCAwIDAuODU0NjQyIDAgMCAwIDAgMC44NTQ2NDIgMCAwIDAg"
    "MSAwIi8+PGZlQmxlbmQgaW4yPSJzaGFwZSIgcmVzdWx0PSJlZmZlY3QyX2lubmVyU2hhZG93Xzc2MV80MTc4MyIvPjxmZUNv"
    "bG9yTWF0cml4IGluPSJTb3VyY2VBbHBoYSIgcmVzdWx0PSJoYXJkQWxwaGEiIHZhbHVlcz0iMCAwIDAgMCAwIDAgMCAwIDAg"
    "MCAwIDAgMCAwIDAgMCAwIDAgMTI3IDAiLz48ZmVPZmZzZXQgZHg9Ii40NzQiIGR5PSIuNDc0Ii8+PGZlR2F1c3NpYW5CbHVy"
    "IHN0ZERldmlhdGlvbj0iLjIzNyIvPjxmZUNvbXBvc2l0ZSBpbjI9ImhhcmRBbHBoYSIgazI9Ii0xIiBrMz0iMSIgb3BlcmF0"
    "b3I9ImFyaXRobWV0aWMiLz48ZmVDb2xvck1hdHJpeCB2YWx1ZXM9IjAgMCAwIDAgMC45MDI2MDcgMCAwIDAgMCAwLjg0MjAz"
    "MSAwIDAgMCAwIDAuODAxNjQ3IDAgMCAwIDEgMCIvPjxmZUJsZW5kIGluMj0iZWZmZWN0Ml9pbm5lclNoYWRvd183NjFfNDE3"
    "ODMiIHJlc3VsdD0iZWZmZWN0M19pbm5lclNoYWRvd183NjFfNDE3ODMiLz48L2ZpbHRlcj48ZmlsdGVyIGlkPSJnIiB3aWR0"
    "aD0iMTkuODk1IiBoZWlnaHQ9IjEyLjMyNiIgeD0iMS44MTYiIHk9IjIuMjg5IiBjb2xvci1pbnRlcnBvbGF0aW9uLWZpbHRl"
    "cnM9InNSR0IiIGZpbHRlclVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+PGZlRmxvb2QgZmxvb2Qtb3BhY2l0eT0iMCIgcmVzdWx0"
    "PSJCYWNrZ3JvdW5kSW1hZ2VGaXgiLz48ZmVDb2xvck1hdHJpeCBpbj0iU291cmNlQWxwaGEiIHJlc3VsdD0iaGFyZEFscGhh"
    "IiB2YWx1ZXM9IjAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDEyNyAwIi8+PGZlT2Zmc2V0IGR4PSItLjIz"
    "NyIgZHk9Ii4yMzciLz48ZmVHYXVzc2lhbkJsdXIgc3RkRGV2aWF0aW9uPSIuNDc0Ii8+PGZlQ29tcG9zaXRlIGluMj0iaGFy"
    "ZEFscGhhIiBvcGVyYXRvcj0ib3V0Ii8+PGZlQ29sb3JNYXRyaXggdmFsdWVzPSIwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAw"
    "IDAgMCAwIDAgMCAwLjEyIDAiLz48ZmVCbGVuZCBpbjI9IkJhY2tncm91bmRJbWFnZUZpeCIgcmVzdWx0PSJlZmZlY3QxX2Ry"
    "b3BTaGFkb3dfNzYxXzQxNzgzIi8+PGZlQmxlbmQgaW49IlNvdXJjZUdyYXBoaWMiIGluMj0iZWZmZWN0MV9kcm9wU2hhZG93"
    "Xzc2MV80MTc4MyIgcmVzdWx0PSJzaGFwZSIvPjxmZUNvbG9yTWF0cml4IGluPSJTb3VyY2VBbHBoYSIgcmVzdWx0PSJoYXJk"
    "QWxwaGEiIHZhbHVlcz0iMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMTI3IDAiLz48ZmVPZmZzZXQgZHg9"
    "Ii0uNDc0IiBkeT0iLS40NzQiLz48ZmVHYXVzc2lhbkJsdXIgc3RkRGV2aWF0aW9uPSIuMjM3Ii8+PGZlQ29tcG9zaXRlIGlu"
    "Mj0iaGFyZEFscGhhIiBrMj0iLTEiIGszPSIxIiBvcGVyYXRvcj0iYXJpdGhtZXRpYyIvPjxmZUNvbG9yTWF0cml4IHZhbHVl"
    "cz0iMCAwIDAgMCAwLjkxMzUxOCAwIDAgMCAwIDAuNTEzMTA2IDAgMCAwIDAgMC4zMTc1NTYgMCAwIDAgMSAwIi8+PGZlQmxl"
    "bmQgaW4yPSJzaGFwZSIgcmVzdWx0PSJlZmZlY3QyX2lubmVyU2hhZG93Xzc2MV80MTc4MyIvPjxmZUNvbG9yTWF0cml4IGlu"
    "PSJTb3VyY2VBbHBoYSIgcmVzdWx0PSJoYXJkQWxwaGEiIHZhbHVlcz0iMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAg"
    "MCAwIDAgMTI3IDAiLz48ZmVPZmZzZXQgZHg9Ii40NzQiIGR5PSIuNDc0Ii8+PGZlR2F1c3NpYW5CbHVyIHN0ZERldmlhdGlv"
    "bj0iLjIzNyIvPjxmZUNvbXBvc2l0ZSBpbjI9ImhhcmRBbHBoYSIgazI9Ii0xIiBrMz0iMSIgb3BlcmF0b3I9ImFyaXRobWV0"
    "aWMiLz48ZmVDb2xvck1hdHJpeCB2YWx1ZXM9IjAgMCAwIDAgMC45OTc2NTIgMCAwIDAgMCAwLjU1MzYyNCAwIDAgMCAwIDAu"
    "NDY3NTc1IDAgMCAwIDEgMCIvPjxmZUJsZW5kIGluMj0iZWZmZWN0Ml9pbm5lclNoYWRvd183NjFfNDE3ODMiIHJlc3VsdD0i"
    "ZWZmZWN0M19pbm5lclNoYWRvd183NjFfNDE3ODMiLz48L2ZpbHRlcj48L2RlZnM+PC9zdmc+"
)

NAV_ICON_OPERATIONS_B64 = (
    "PHN2ZyB3aWR0aD0iMjQiIGhlaWdodD0iMjQiIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6"
    "Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPGcgZmlsdGVyPSJ1cmwoI2ZpbHRlcjBfZGlpXzc2MV80MTc4NykiPgo8cGF0aCBk"
    "PSJNMjAuNDI1NyA2Ljc1MDM2QzIwLjkwNjggNi4yNjkyNyAyMS43MTY1IDYuNDMyOTcgMjEuODE1OCA3LjEwNjA1QzIyLjA1"
    "MjcgOC43MTIzOSAyMS41NTM0IDEwLjQwNjIgMjAuMzE2OSAxMS42NDI2QzE4Ljg4NjEgMTMuMDczNCAxNi44NDI5IDEzLjUx"
    "NjcgMTUuMDMxNCAxMi45NzQ2QzE0LjY1NzYgMTIuODYyNyAxNC4yNDI4IDEyLjkzMzMgMTMuOTY2OSAxMy4yMDkyTDYuOTkx"
    "NTYgMjAuMTg0NUM2LjIzNjc0IDIwLjkzOTIgNS4wMTI5MyAyMC45MzkzIDQuMjU4MTUgMjAuMTg0NUMzLjUwMzM2IDE5LjQy"
    "OTcgMy41MDM0MSAxOC4yMDU5IDQuMjU4MTUgMTcuNDUxMUwxMS4yMzM1IDEwLjQ3NThDMTEuNTA5MyAxMC4xOTk5IDExLjU4"
    "IDkuNzg1MDUgMTEuNDY4MSA5LjQxMTI1QzEwLjkyNiA3LjU5OTc0IDExLjM2OTIgNS41NTY2IDEyLjgwMDEgNC4xMjU3M0Mx"
    "NC4wMzY1IDIuODg5MjkgMTUuNzMwMyAyLjM4OTk4IDE3LjMzNjYgMi42MjY4OUMxOC4wMDk3IDIuNzI2MTggMTguMTczNCAz"
    "LjUzNTg2IDE3LjY5MjMgNC4wMTY5NUwxNS44OTg5IDUuODEwMzVDMTUuNTA4NCA2LjIwMDg4IDE1LjUwODQgNi44MzQwNCAx"
    "NS44OTg5IDcuMjI0NTdMMTcuMjE4MSA4LjU0Mzc2QzE3LjYwODYgOC45MzQyOSAxOC4yNDE4IDguOTM0MjkgMTguNjMyMyA4"
    "LjU0Mzc2TDIwLjQyNTcgNi43NTAzNloiIGZpbGw9IiNDNkMzQkQiLz4KPC9nPgo8ZGVmcz4KPGZpbHRlciBpZD0iZmlsdGVy"
    "MF9kaWlfNzYxXzQxNzg3IiB4PSIxLjY5MjA4IiB5PSIxLjU2OTcyIiB3aWR0aD0iMjIuMTgwOSIgaGVpZ2h0PSIyMi4xODA5"
    "IiBmaWx0ZXJVbml0cz0idXNlclNwYWNlT25Vc2UiIGNvbG9yLWludGVycG9sYXRpb24tZmlsdGVycz0ic1JHQiI+CjxmZUZs"
    "b29kIGZsb29kLW9wYWNpdHk9IjAiIHJlc3VsdD0iQmFja2dyb3VuZEltYWdlRml4Ii8+CjxmZUNvbG9yTWF0cml4IGluPSJT"
    "b3VyY2VBbHBoYSIgdHlwZT0ibWF0cml4IiB2YWx1ZXM9IjAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDEy"
    "NyAwIiByZXN1bHQ9ImhhcmRBbHBoYSIvPgo8ZmVPZmZzZXQgZHk9IjEiLz4KPGZlR2F1c3NpYW5CbHVyIHN0ZERldmlhdGlv"
    "bj0iMSIvPgo8ZmVDb21wb3NpdGUgaW4yPSJoYXJkQWxwaGEiIG9wZXJhdG9yPSJvdXQiLz4KPGZlQ29sb3JNYXRyaXggdHlw"
    "ZT0ibWF0cml4IiB2YWx1ZXM9IjAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAuMTIgMCIvPgo8ZmVCbGVu"
    "ZCBtb2RlPSJub3JtYWwiIGluMj0iQmFja2dyb3VuZEltYWdlRml4IiByZXN1bHQ9ImVmZmVjdDFfZHJvcFNoYWRvd183NjFf"
    "NDE3ODciLz4KPGZlQmxlbmQgbW9kZT0ibm9ybWFsIiBpbj0iU291cmNlR3JhcGhpYyIgaW4yPSJlZmZlY3QxX2Ryb3BTaGFk"
    "b3dfNzYxXzQxNzg3IiByZXN1bHQ9InNoYXBlIi8+CjxmZUNvbG9yTWF0cml4IGluPSJTb3VyY2VBbHBoYSIgdHlwZT0ibWF0"
    "cml4IiB2YWx1ZXM9IjAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDEyNyAwIiByZXN1bHQ9ImhhcmRBbHBo"
    "YSIvPgo8ZmVPZmZzZXQgZHg9IjEiIGR5PSItMSIvPgo8ZmVHYXVzc2lhbkJsdXIgc3RkRGV2aWF0aW9uPSIwLjI1Ii8+Cjxm"
    "ZUNvbXBvc2l0ZSBpbjI9ImhhcmRBbHBoYSIgb3BlcmF0b3I9ImFyaXRobWV0aWMiIGsyPSItMSIgazM9IjEiLz4KPGZlQ29s"
    "b3JNYXRyaXggdHlwZT0ibWF0cml4IiB2YWx1ZXM9IjAgMCAwIDAgMC42NjI3NDUgMCAwIDAgMCAwLjYxOTYwOCAwIDAgMCAw"
    "IDAuNTgwMzkyIDAgMCAwIDEgMCIvPgo8ZmVCbGVuZCBtb2RlPSJub3JtYWwiIGluMj0ic2hhcGUiIHJlc3VsdD0iZWZmZWN0"
    "Ml9pbm5lclNoYWRvd183NjFfNDE3ODciLz4KPGZlQ29sb3JNYXRyaXggaW49IlNvdXJjZUFscGhhIiB0eXBlPSJtYXRyaXgi"
    "IHZhbHVlcz0iMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMTI3IDAiIHJlc3VsdD0iaGFyZEFscGhhIi8+"
    "CjxmZU9mZnNldCBkeD0iMSIvPgo8ZmVHYXVzc2lhbkJsdXIgc3RkRGV2aWF0aW9uPSIwLjEyNSIvPgo8ZmVDb21wb3NpdGUg"
    "aW4yPSJoYXJkQWxwaGEiIG9wZXJhdG9yPSJhcml0aG1ldGljIiBrMj0iLTEiIGszPSIxIi8+CjxmZUNvbG9yTWF0cml4IHR5"
    "cGU9Im1hdHJpeCIgdmFsdWVzPSIwIDAgMCAwIDAuNTUyOTQxIDAgMCAwIDAgMC41MzMzMzMgMCAwIDAgMCAwLjUwOTgwNCAw"
    "IDAgMCAxIDAiLz4KPGZlQmxlbmQgbW9kZT0ibm9ybWFsIiBpbjI9ImVmZmVjdDJfaW5uZXJTaGFkb3dfNzYxXzQxNzg3IiBy"
    "ZXN1bHQ9ImVmZmVjdDNfaW5uZXJTaGFkb3dfNzYxXzQxNzg3Ii8+CjwvZmlsdGVyPgo8L2RlZnM+Cjwvc3ZnPgo="
)

NAV_ICON_GUIDE_B64 = (
    "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgZmlsbD0ibm9u"
    "ZSI+PGcgY2xpcC1wYXRoPSJ1cmwoI2EpIj48ZyBmaWx0ZXI9InVybCgjYikiPjxwYXRoIGZpbGw9IiNGRERBOEEiIGQ9Ik02"
    "LjAzNCAzLjA1MmExLjEyMiAxLjEyMiAwIDAgMSAxLjkzMiAwbDMuMzgzIDUuNzkzYy40My43MzUtLjEwNyAxLjY1NS0uOTY3"
    "IDEuNjU1SDMuNjE4Yy0uODYgMC0xLjM5Ni0uOTItLjk2Ny0xLjY1NXoiLz48L2c+PGcgZmlsdGVyPSJ1cmwoI2MpIj48cGF0"
    "aCBmaWxsPSJ1cmwoI2QpIiBkPSJNMTMuNSAzLjQ0N2MwLS41MjMuNDI0LS45NDcuOTQ3LS45NDdoNi4xMDZjLjUyMyAwIC45"
    "NDcuNDI0Ljk0Ny45NDd2Ni4xMDZhLjk0Ny45NDcgMCAwIDEtLjk0Ny45NDdoLTYuMTA2YS45NDcuOTQ3IDAgMCAxLS45NDct"
    "Ljk0N3oiLz48L2c+PGcgZmlsdGVyPSJ1cmwoI2UpIj48Y2lyY2xlIGN4PSI3IiBjeT0iMTciIHI9IjQuNSIgZmlsbD0iIzlF"
    "RTNCMyIvPjwvZz48ZyBmaWx0ZXI9InVybCgjZikiPjxjaXJjbGUgY3g9IjE3LjUiIGN5PSIxNyIgcj0iNC41IiBmaWxsPSIj"
    "OUVFM0IzIi8+PC9nPjwvZz48ZGVmcz48ZmlsdGVyIGlkPSJiIiB3aWR0aD0iMTMiIGhlaWdodD0iMTIiIHg9Ii41IiB5PSIx"
    "LjUiIGNvbG9yLWludGVycG9sYXRpb24tZmlsdGVycz0ic1JHQiIgZmlsdGVyVW5pdHM9InVzZXJTcGFjZU9uVXNlIj48ZmVG"
    "bG9vZCBmbG9vZC1vcGFjaXR5PSIwIiByZXN1bHQ9IkJhY2tncm91bmRJbWFnZUZpeCIvPjxmZUNvbG9yTWF0cml4IGluPSJT"
    "b3VyY2VBbHBoYSIgcmVzdWx0PSJoYXJkQWxwaGEiIHZhbHVlcz0iMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAw"
    "IDAgMTI3IDAiLz48ZmVPZmZzZXQgZHk9IjEiLz48ZmVHYXVzc2lhbkJsdXIgc3RkRGV2aWF0aW9uPSIxIi8+PGZlQ29tcG9z"
    "aXRlIGluMj0iaGFyZEFscGhhIiBvcGVyYXRvcj0ib3V0Ii8+PGZlQ29sb3JNYXRyaXggdmFsdWVzPSIwIDAgMCAwIDAgMCAw"
    "IDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwLjEyIDAiLz48ZmVCbGVuZCBpbjI9IkJhY2tncm91bmRJbWFnZUZpeCIgcmVzdWx0"
    "PSJlZmZlY3QxX2Ryb3BTaGFkb3dfNzYxXzQxNzg2Ii8+PGZlQmxlbmQgaW49IlNvdXJjZUdyYXBoaWMiIGluMj0iZWZmZWN0"
    "MV9kcm9wU2hhZG93Xzc2MV80MTc4NiIgcmVzdWx0PSJzaGFwZSIvPjxmZUNvbG9yTWF0cml4IGluPSJTb3VyY2VBbHBoYSIg"
    "cmVzdWx0PSJoYXJkQWxwaGEiIHZhbHVlcz0iMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMTI3IDAiLz48"
    "ZmVPZmZzZXQgZHg9IjEiIGR5PSItMSIvPjxmZUdhdXNzaWFuQmx1ciBzdGREZXZpYXRpb249Ii41Ii8+PGZlQ29tcG9zaXRl"
    "IGluMj0iaGFyZEFscGhhIiBrMj0iLTEiIGszPSIxIiBvcGVyYXRvcj0iYXJpdGhtZXRpYyIvPjxmZUNvbG9yTWF0cml4IHZh"
    "bHVlcz0iMCAwIDAgMCAxIDAgMCAwIDAgMC44MDQzNjQgMCAwIDAgMCAwLjM1NzE5NyAwIDAgMCAxIDAiLz48ZmVCbGVuZCBp"
    "bjI9InNoYXBlIiByZXN1bHQ9ImVmZmVjdDJfaW5uZXJTaGFkb3dfNzYxXzQxNzg2Ii8+PGZlQ29sb3JNYXRyaXggaW49IlNv"
    "dXJjZUFscGhhIiByZXN1bHQ9ImhhcmRBbHBoYSIgdmFsdWVzPSIwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAg"
    "MCAxMjcgMCIvPjxmZU9mZnNldCBkeD0iMSIvPjxmZUdhdXNzaWFuQmx1ciBzdGREZXZpYXRpb249Ii40Ii8+PGZlQ29tcG9z"
    "aXRlIGluMj0iaGFyZEFscGhhIiBrMj0iLTEiIGszPSIxIiBvcGVyYXRvcj0iYXJpdGhtZXRpYyIvPjxmZUNvbG9yTWF0cml4"
    "IHZhbHVlcz0iMCAwIDAgMCAwLjg5OTUwOCAwIDAgMCAwIDAuNzMwODEyIDAgMCAwIDAgMC4zNDUyMjEgMCAwIDAgMSAwIi8+"
    "PGZlQmxlbmQgaW4yPSJlZmZlY3QyX2lubmVyU2hhZG93Xzc2MV80MTc4NiIgcmVzdWx0PSJlZmZlY3QzX2lubmVyU2hhZG93"
    "Xzc2MV80MTc4NiIvPjxmZUNvbG9yTWF0cml4IGluPSJTb3VyY2VBbHBoYSIgcmVzdWx0PSJoYXJkQWxwaGEiIHZhbHVlcz0i"
    "MCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMTI3IDAiLz48ZmVPZmZzZXQgZHg9Ii0xIiBkeT0iMSIvPjxm"
    "ZUdhdXNzaWFuQmx1ciBzdGREZXZpYXRpb249Ii40Ii8+PGZlQ29tcG9zaXRlIGluMj0iaGFyZEFscGhhIiBrMj0iLTEiIGsz"
    "PSIxIiBvcGVyYXRvcj0iYXJpdGhtZXRpYyIvPjxmZUNvbG9yTWF0cml4IHZhbHVlcz0iMCAwIDAgMCAxIDAgMCAwIDAgMC45"
    "MDYyMDYgMCAwIDAgMCAwLjY5MTgxOSAwIDAgMCAxIDAiLz48ZmVCbGVuZCBpbjI9ImVmZmVjdDNfaW5uZXJTaGFkb3dfNzYx"
    "XzQxNzg2IiByZXN1bHQ9ImVmZmVjdDRfaW5uZXJTaGFkb3dfNzYxXzQxNzg2Ii8+PC9maWx0ZXI+PGZpbHRlciBpZD0iYyIg"
    "d2lkdGg9IjEyIiBoZWlnaHQ9IjEyIiB4PSIxMS41IiB5PSIxLjUiIGNvbG9yLWludGVycG9sYXRpb24tZmlsdGVycz0ic1JH"
    "QiIgZmlsdGVyVW5pdHM9InVzZXJTcGFjZU9uVXNlIj48ZmVGbG9vZCBmbG9vZC1vcGFjaXR5PSIwIiByZXN1bHQ9IkJhY2tn"
    "cm91bmRJbWFnZUZpeCIvPjxmZUNvbG9yTWF0cml4IGluPSJTb3VyY2VBbHBoYSIgcmVzdWx0PSJoYXJkQWxwaGEiIHZhbHVl"
    "cz0iMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMTI3IDAiLz48ZmVPZmZzZXQgZHk9IjEiLz48ZmVHYXVz"
    "c2lhbkJsdXIgc3RkRGV2aWF0aW9uPSIxIi8+PGZlQ29tcG9zaXRlIGluMj0iaGFyZEFscGhhIiBvcGVyYXRvcj0ib3V0Ii8+"
    "PGZlQ29sb3JNYXRyaXggdmFsdWVzPSIwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwLjEyIDAiLz48ZmVC"
    "bGVuZCBpbjI9IkJhY2tncm91bmRJbWFnZUZpeCIgcmVzdWx0PSJlZmZlY3QxX2Ryb3BTaGFkb3dfNzYxXzQxNzg2Ii8+PGZl"
    "QmxlbmQgaW49IlNvdXJjZUdyYXBoaWMiIGluMj0iZWZmZWN0MV9kcm9wU2hhZG93Xzc2MV80MTc4NiIgcmVzdWx0PSJzaGFw"
    "ZSIvPjxmZUNvbG9yTWF0cml4IGluPSJTb3VyY2VBbHBoYSIgcmVzdWx0PSJoYXJkQWxwaGEiIHZhbHVlcz0iMCAwIDAgMCAw"
    "IDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMTI3IDAiLz48ZmVPZmZzZXQgZHg9IjEiIGR5PSItMSIvPjxmZUdhdXNzaWFu"
    "Qmx1ciBzdGREZXZpYXRpb249Ii41Ii8+PGZlQ29tcG9zaXRlIGluMj0iaGFyZEFscGhhIiBrMj0iLTEiIGszPSIxIiBvcGVy"
    "YXRvcj0iYXJpdGhtZXRpYyIvPjxmZUNvbG9yTWF0cml4IHZhbHVlcz0iMCAwIDAgMCAwLjg1NDY0MiAwIDAgMCAwIDAuODU0"
    "NjQyIDAgMCAwIDAgMC44NTQ2NDIgMCAwIDAgMSAwIi8+PGZlQmxlbmQgaW4yPSJzaGFwZSIgcmVzdWx0PSJlZmZlY3QyX2lu"
    "bmVyU2hhZG93Xzc2MV80MTc4NiIvPjxmZUNvbG9yTWF0cml4IGluPSJTb3VyY2VBbHBoYSIgcmVzdWx0PSJoYXJkQWxwaGEi"
    "IHZhbHVlcz0iMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMTI3IDAiLz48ZmVPZmZzZXQgZHg9IjEiLz48"
    "ZmVHYXVzc2lhbkJsdXIgc3RkRGV2aWF0aW9uPSIuNCIvPjxmZUNvbXBvc2l0ZSBpbjI9ImhhcmRBbHBoYSIgazI9Ii0xIiBr"
    "Mz0iMSIgb3BlcmF0b3I9ImFyaXRobWV0aWMiLz48ZmVDb2xvck1hdHJpeCB2YWx1ZXM9IjAgMCAwIDAgMC45MDI2MDcgMCAw"
    "IDAgMCAwLjg0MjAzMSAwIDAgMCAwIDAuODAxNjQ3IDAgMCAwIDEgMCIvPjxmZUJsZW5kIGluMj0iZWZmZWN0Ml9pbm5lclNo"
    "YWRvd183NjFfNDE3ODYiIHJlc3VsdD0iZWZmZWN0M19pbm5lclNoYWRvd183NjFfNDE3ODYiLz48L2ZpbHRlcj48ZmlsdGVy"
    "IGlkPSJlIiB3aWR0aD0iMTMiIGhlaWdodD0iMTMiIHg9Ii41IiB5PSIxMS41IiBjb2xvci1pbnRlcnBvbGF0aW9uLWZpbHRl"
    "cnM9InNSR0IiIGZpbHRlclVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+PGZlRmxvb2QgZmxvb2Qtb3BhY2l0eT0iMCIgcmVzdWx0"
    "PSJCYWNrZ3JvdW5kSW1hZ2VGaXgiLz48ZmVDb2xvck1hdHJpeCBpbj0iU291cmNlQWxwaGEiIHJlc3VsdD0iaGFyZEFscGhh"
    "IiB2YWx1ZXM9IjAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDEyNyAwIi8+PGZlT2Zmc2V0IGR5PSIxIi8+"
    "PGZlR2F1c3NpYW5CbHVyIHN0ZERldmlhdGlvbj0iMSIvPjxmZUNvbXBvc2l0ZSBpbjI9ImhhcmRBbHBoYSIgb3BlcmF0b3I9"
    "Im91dCIvPjxmZUNvbG9yTWF0cml4IHZhbHVlcz0iMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMC4xMiAw"
    "Ii8+PGZlQmxlbmQgaW4yPSJCYWNrZ3JvdW5kSW1hZ2VGaXgiIHJlc3VsdD0iZWZmZWN0MV9kcm9wU2hhZG93Xzc2MV80MTc4"
    "NiIvPjxmZUJsZW5kIGluPSJTb3VyY2VHcmFwaGljIiBpbjI9ImVmZmVjdDFfZHJvcFNoYWRvd183NjFfNDE3ODYiIHJlc3Vs"
    "dD0ic2hhcGUiLz48ZmVDb2xvck1hdHJpeCBpbj0iU291cmNlQWxwaGEiIHJlc3VsdD0iaGFyZEFscGhhIiB2YWx1ZXM9IjAg"
    "MCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDEyNyAwIi8+PGZlT2Zmc2V0IGR4PSIxIiBkeT0iLTEiLz48ZmVH"
    "YXVzc2lhbkJsdXIgc3RkRGV2aWF0aW9uPSIuNSIvPjxmZUNvbXBvc2l0ZSBpbjI9ImhhcmRBbHBoYSIgazI9Ii0xIiBrMz0i"
    "MSIgb3BlcmF0b3I9ImFyaXRobWV0aWMiLz48ZmVDb2xvck1hdHJpeCB2YWx1ZXM9IjAgMCAwIDAgMC41MjA3OTcgMCAwIDAg"
    "MCAwLjkzMTAyMSAwIDAgMCAwIDAuNjQ1NjQ4IDAgMCAwIDEgMCIvPjxmZUJsZW5kIGluMj0ic2hhcGUiIHJlc3VsdD0iZWZm"
    "ZWN0Ml9pbm5lclNoYWRvd183NjFfNDE3ODYiLz48ZmVDb2xvck1hdHJpeCBpbj0iU291cmNlQWxwaGEiIHJlc3VsdD0iaGFy"
    "ZEFscGhhIiB2YWx1ZXM9IjAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDEyNyAwIi8+PGZlT2Zmc2V0IGR4"
    "PSIxIi8+PGZlR2F1c3NpYW5CbHVyIHN0ZERldmlhdGlvbj0iLjQiLz48ZmVDb21wb3NpdGUgaW4yPSJoYXJkQWxwaGEiIGsy"
    "PSItMSIgazM9IjEiIG9wZXJhdG9yPSJhcml0aG1ldGljIi8+PGZlQ29sb3JNYXRyaXggdmFsdWVzPSIwIDAgMCAwIDAuMzI0"
    "ODc0IDAgMCAwIDAgMC43Nzc1MTMgMCAwIDAgMCAwLjQ2MjYzMyAwIDAgMCAxIDAiLz48ZmVCbGVuZCBpbjI9ImVmZmVjdDJf"
    "aW5uZXJTaGFkb3dfNzYxXzQxNzg2IiByZXN1bHQ9ImVmZmVjdDNfaW5uZXJTaGFkb3dfNzYxXzQxNzg2Ii8+PGZlQ29sb3JN"
    "YXRyaXggaW49IlNvdXJjZUFscGhhIiByZXN1bHQ9ImhhcmRBbHBoYSIgdmFsdWVzPSIwIDAgMCAwIDAgMCAwIDAgMCAwIDAg"
    "MCAwIDAgMCAwIDAgMCAxMjcgMCIvPjxmZU9mZnNldCBkeD0iLTEiIGR5PSIxIi8+PGZlR2F1c3NpYW5CbHVyIHN0ZERldmlh"
    "dGlvbj0iLjQiLz48ZmVDb21wb3NpdGUgaW4yPSJoYXJkQWxwaGEiIGsyPSItMSIgazM9IjEiIG9wZXJhdG9yPSJhcml0aG1l"
    "dGljIi8+PGZlQ29sb3JNYXRyaXggdmFsdWVzPSIwIDAgMCAwIDAuNjQ0MTU5IDAgMCAwIDAgMC45NTUzMDMgMCAwIDAgMCAw"
    "LjczODg1NSAwIDAgMCAxIDAiLz48ZmVCbGVuZCBpbjI9ImVmZmVjdDNfaW5uZXJTaGFkb3dfNzYxXzQxNzg2IiByZXN1bHQ9"
    "ImVmZmVjdDRfaW5uZXJTaGFkb3dfNzYxXzQxNzg2Ii8+PC9maWx0ZXI+PGZpbHRlciBpZD0iZiIgd2lkdGg9IjEzIiBoZWln"
    "aHQ9IjEzIiB4PSIxMSIgeT0iMTEuNSIgY29sb3ItaW50ZXJwb2xhdGlvbi1maWx0ZXJzPSJzUkdCIiBmaWx0ZXJVbml0cz0i"
    "dXNlclNwYWNlT25Vc2UiPjxmZUZsb29kIGZsb29kLW9wYWNpdHk9IjAiIHJlc3VsdD0iQmFja2dyb3VuZEltYWdlRml4Ii8+"
    "PGZlQ29sb3JNYXRyaXggaW49IlNvdXJjZUFscGhhIiByZXN1bHQ9ImhhcmRBbHBoYSIgdmFsdWVzPSIwIDAgMCAwIDAgMCAw"
    "IDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAxMjcgMCIvPjxmZU9mZnNldCBkeT0iMSIvPjxmZUdhdXNzaWFuQmx1ciBzdGREZXZp"
    "YXRpb249IjEiLz48ZmVDb21wb3NpdGUgaW4yPSJoYXJkQWxwaGEiIG9wZXJhdG9yPSJvdXQiLz48ZmVDb2xvck1hdHJpeCB2"
    "YWx1ZXM9IjAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAuMTIgMCIvPjxmZUJsZW5kIGluMj0iQmFja2dy"
    "b3VuZEltYWdlRml4IiByZXN1bHQ9ImVmZmVjdDFfZHJvcFNoYWRvd183NjFfNDE3ODYiLz48ZmVCbGVuZCBpbj0iU291cmNl"
    "R3JhcGhpYyIgaW4yPSJlZmZlY3QxX2Ryb3BTaGFkb3dfNzYxXzQxNzg2IiByZXN1bHQ9InNoYXBlIi8+PGZlQ29sb3JNYXRy"
    "aXggaW49IlNvdXJjZUFscGhhIiByZXN1bHQ9ImhhcmRBbHBoYSIgdmFsdWVzPSIwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAw"
    "IDAgMCAwIDAgMCAxMjcgMCIvPjxmZU9mZnNldCBkeD0iMSIgZHk9Ii0xIi8+PGZlR2F1c3NpYW5CbHVyIHN0ZERldmlhdGlv"
    "bj0iLjUiLz48ZmVDb21wb3NpdGUgaW4yPSJoYXJkQWxwaGEiIGsyPSItMSIgazM9IjEiIG9wZXJhdG9yPSJhcml0aG1ldGlj"
    "Ii8+PGZlQ29sb3JNYXRyaXggdmFsdWVzPSIwIDAgMCAwIDAuNTIwNzk3IDAgMCAwIDAgMC45MzEwMjEgMCAwIDAgMCAwLjY0"
    "NTY0OCAwIDAgMCAxIDAiLz48ZmVCbGVuZCBpbjI9InNoYXBlIiByZXN1bHQ9ImVmZmVjdDJfaW5uZXJTaGFkb3dfNzYxXzQx"
    "Nzg2Ii8+PGZlQ29sb3JNYXRyaXggaW49IlNvdXJjZUFscGhhIiByZXN1bHQ9ImhhcmRBbHBoYSIgdmFsdWVzPSIwIDAgMCAw"
    "IDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAxMjcgMCIvPjxmZU9mZnNldCBkeD0iMSIvPjxmZUdhdXNzaWFuQmx1ciBz"
    "dGREZXZpYXRpb249Ii40Ii8+PGZlQ29tcG9zaXRlIGluMj0iaGFyZEFscGhhIiBrMj0iLTEiIGszPSIxIiBvcGVyYXRvcj0i"
    "YXJpdGhtZXRpYyIvPjxmZUNvbG9yTWF0cml4IHZhbHVlcz0iMCAwIDAgMCAwLjMyNDg3NCAwIDAgMCAwIDAuNzc3NTEzIDAg"
    "MCAwIDAgMC40NjI2MzMgMCAwIDAgMSAwIi8+PGZlQmxlbmQgaW4yPSJlZmZlY3QyX2lubmVyU2hhZG93Xzc2MV80MTc4NiIg"
    "cmVzdWx0PSJlZmZlY3QzX2lubmVyU2hhZG93Xzc2MV80MTc4NiIvPjxmZUNvbG9yTWF0cml4IGluPSJTb3VyY2VBbHBoYSIg"
    "cmVzdWx0PSJoYXJkQWxwaGEiIHZhbHVlcz0iMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMTI3IDAiLz48"
    "ZmVPZmZzZXQgZHg9Ii0xIiBkeT0iMSIvPjxmZUdhdXNzaWFuQmx1ciBzdGREZXZpYXRpb249Ii40Ii8+PGZlQ29tcG9zaXRl"
    "IGluMj0iaGFyZEFscGhhIiBrMj0iLTEiIGszPSIxIiBvcGVyYXRvcj0iYXJpdGhtZXRpYyIvPjxmZUNvbG9yTWF0cml4IHZh"
    "bHVlcz0iMCAwIDAgMCAwLjY0NDE1OSAwIDAgMCAwIDAuOTU1MzAzIDAgMCAwIDAgMC43Mzg4NTUgMCAwIDAgMSAwIi8+PGZl"
    "QmxlbmQgaW4yPSJlZmZlY3QzX2lubmVyU2hhZG93Xzc2MV80MTc4NiIgcmVzdWx0PSJlZmZlY3Q0X2lubmVyU2hhZG93Xzc2"
    "MV80MTc4NiIvPjwvZmlsdGVyPjxsaW5lYXJHcmFkaWVudCBpZD0iZCIgeDE9IjE3LjUiIHgyPSIxNy41IiB5MT0iMi41IiB5"
    "Mj0iMTAuNSIgZ3JhZGllbnRVbml0cz0idXNlclNwYWNlT25Vc2UiPjxzdG9wIHN0b3AtY29sb3I9IiNGM0VFRUEiLz48c3Rv"
    "cCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiNFQ0U1REYiLz48L2xpbmVhckdyYWRpZW50PjxjbGlwUGF0aCBpZD0iYSI+PHBh"
    "dGggZmlsbD0iI2ZmZiIgZD0iTTAgMGgyNHYyNEgweiIvPjwvY2xpcFBhdGg+PC9kZWZzPjwvc3ZnPg=="
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


def kor_date_labels(dates, unit="day"):
    """date 값들을 한글 라벨(예: '8월 1일', '2026년 8월') 리스트로 변환.
    Plotly에 date 타입을 그대로 x축으로 주면 기본 영문(예: 'Aug 1')으로 표시되므로,
    차트를 그리기 전 이 라벨로 바꾸고 px.line/px.bar의 category_orders로 순서를 고정해서 쓴다."""
    if unit == "month":
        return [f"{d.year}년 {d.month}월" for d in dates]
    return [f"{d.month}월 {d.day}일" for d in dates]


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
        /* 그룹별 아이콘 지정 — render_nav()에서 그룹마다 container(key=f"navgrp_xxx")로 한 번 더
           감싸서 붙는 고유 클래스(st-key-navgrp_xxx)로 구분한다. :nth-of-type은 Streamlit이 위젯을
           각각 별도 element-container로 감싸는 구조라 전부 "1번째"로 잡혀서 못 쓴다.
           GA 유입 리포트는 별도 아이콘을 안 받아서 위 기본값(NAV_GROUP_ICON_B64) 그대로 유지. */
        div.st-key-navgrp_report div[data-testid="stExpander"] summary p::before {{
            background-image: url("data:image/svg+xml;base64,{NAV_ICON_PERFORMANCE_B64}");
        }}
        div.st-key-navgrp_ops div[data-testid="stExpander"] summary p::before {{
            background-image: url("data:image/svg+xml;base64,{NAV_ICON_OPERATIONS_B64}");
        }}
        div.st-key-navgrp_guide div[data-testid="stExpander"] summary p::before {{
            background-image: url("data:image/svg+xml;base64,{NAV_ICON_GUIDE_B64}");
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
    "inflow_revenue_daily": "inflow_revenue_daily",
    "ga_channel_inflow": "ga_channel_inflow",
    "agency_notes": "agency_notes",
    "channel_weekly": "channel_weekly",
    "utm_channel_map": "utm_channel_map",
    "channel_budget": "channel_budget",
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


def _parse_weekly_section(raw: pd.DataFrame, bounds, today: date) -> pd.DataFrame:
    """'■ 주간 데이터'(또는 '통합 주간별') 섹션 하나를 파싱하는 공용 코어.
    '매체통합' 시트의 통합 주간별 표뿐 아니라, 매체 개별 시트(네이버/GFA/메타/구글/크리테오 등)에
    똑같이 있는 '■ 주간 데이터' 표에도 그대로 쓴다 — 구조가 동일하기 때문."""
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
            # 주간 섹션은 같은 시트의 월간 섹션 첫 달과 같은 해에서 시작한다고 본다.
            # 컬럼 순서가 시트마다 달라서(예: '매체통합'은 날짜가 1번째 컬럼, 매체 개별
            # 시트는 0번째 컬럼인 경우 등) 위치를 하드코딩하지 않고 section_dataframe으로
            # 실제 날짜 컬럼을 다시 찾아서 연도를 뽑는다.
            year_val = None
            if "monthly" in bounds:
                m_data, m_date_idx = section_dataframe(raw, *bounds["monthly"], date_tokens=("월별", "기간"))
                if m_data is not None and not m_data.empty:
                    first_date = pd.to_datetime(m_data.iloc[0, m_date_idx], errors="coerce")
                    if pd.notna(first_date):
                        year_val = first_date.year
            year_state["year"] = year_val if year_val is not None else today.year
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
    out = out.sort_values("week_start").reset_index(drop=True)
    # parse_daily와 동일한 이유: 리포트 템플릿이 아직 안 지난 미래 주차까지 행(라벨)을 미리
    # 만들어두는 경우가 있어서, week_end<=today만으로는 다 안 걸러진다. 실제 값이 하나도 없는
    # (전부 0인) 말미 주차는 '아직 보고 안 된 주'로 보고 잘라낸다.
    metric_sum = out[["impressions", "clicks", "cost_excl_vat", "cost_incl_vat", "conversions", "revenue"]].sum(axis=1)
    nonzero_idx = metric_sum[metric_sum > 0].index
    if len(nonzero_idx):
        out = out.loc[: nonzero_idx.max()]
    return out.reset_index(drop=True)


def parse_weekly(raw: pd.DataFrame, bounds, today: date) -> pd.DataFrame:
    """'매체통합' 시트의 통합 주간별 표(전체 합산)를 파싱한다."""
    return _parse_weekly_section(raw, bounds, today)


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


# 대행사가 '매체통합' 시트 하단에 자유 텍스트로 남기는 운영 메모(특이사항/코멘트)를 찾을 때
# 상위 불릿(예: "*. 7월 신규 예산 반영 완료")으로 인정할 접두어.
AGENCY_NOTE_BULLET_PREFIXES = ("*", "-", "•", "※", "▶", "①", "②", "③", "④", "⑤")
# ">"로 시작하는 줄은 바로 위 상위 불릿에 딸린 하위/부연 설명으로 취급(들여쓰기해서 표시).
AGENCY_NOTE_SUB_PREFIX = ">"


def parse_agency_notes(raw: pd.DataFrame, today: date) -> pd.DataFrame:
    """'매체통합' 시트 하단에 대행사가 정리해주는 자유 텍스트 운영 코멘트/특이사항을 추출한다.
    상단 표(월간/주간/일별/매체별 현황)는 각 행이 숫자 위주 데이터라 첫 칸이 불릿 마커로
    시작하지 않는다는 점을 이용해, 각 행의 첫 번째 비어있지 않은 문자열 셀이 '*.', '-', '>' 같은
    불릿 마커로 시작하는 행만 순서대로 모아 하나의 텍스트로 합친다. '>'로 시작하는 줄은 바로 위
    상위 항목의 하위 설명으로 보고 들여써서 표시한다."""
    if raw is None or raw.empty:
        return pd.DataFrame()

    items = []  # (level, text) — level 0=상위 불릿, 1=하위(">")
    n_rows, n_cols = raw.shape
    for i in range(n_rows):
        first_str = None
        for j in range(n_cols):
            v = raw.iat[i, j]
            if isinstance(v, str) and v.strip():
                first_str = v.strip()
                break
        if not first_str or len(first_str) < 2:
            continue
        if first_str.startswith(AGENCY_NOTE_SUB_PREFIX):
            text = first_str.lstrip(AGENCY_NOTE_SUB_PREFIX + " ").strip()
            if text:
                items.append((1, text))
        elif first_str.startswith(AGENCY_NOTE_BULLET_PREFIXES):
            # "*." "* " 등 뒤에 붙는 마침표/공백까지 같이 제거 (예: "*. 쇼핑박스..." → "쇼핑박스...")
            text = first_str.lstrip("".join(AGENCY_NOTE_BULLET_PREFIXES) + ". ").strip()
            if text:
                items.append((0, text))

    if not items:
        return pd.DataFrame()

    note_text = "\n".join(f"    - {t}" if level == 1 else f"- {t}" for level, t in items)
    return pd.DataFrame([{"as_of_date": today, "note_text": note_text}])


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


def _drop_hidden_rows(raw: pd.DataFrame, hidden_rows: set | None) -> pd.DataFrame:
    """엑셀에서 행 자체를 숨겨둔(우클릭 → 행 숨기기) 행은 화면에 보이는 표와 무관한
    예전 버전/임시 데이터인 경우가 많다. 위치 기반(iloc) 파싱이 엉뚱한 행을 진짜 표로
    착각하지 않도록, 섹션을 찾기 전에 아예 원본에서 제거하고 행 번호를 다시 매긴다."""
    if not hidden_rows:
        return raw
    keep_idx = [i for i in range(len(raw)) if i not in hidden_rows]
    if len(keep_idx) == len(raw):
        return raw
    return raw.iloc[keep_idx].reset_index(drop=True)


def parse_channel_sheet(xls: pd.ExcelFile, sheet: str, today: date, hidden_rows: set | None = None):
    """매체 개별 시트(네이버/GFA/메타/구글/크리테오 등)에서 월간 데이터를 파싱한다.
    (monthly_df, weekly_df) 튜플을 반환한다 — weekly_df는 그 시트에 '■ 주간 데이터' 섹션이
    있을 때만 채워지고(대부분의 개별 매체 시트에 있음), 없으면 빈 DataFrame이다.
    hidden_rows: 엑셀에서 숨겨진 행 번호(raw의 0-indexed 위치) 집합 — 있으면 파싱 전에 미리
    제거한다. 시트 안에 예전 버전 표가 통째로 숨겨진 채 남아있어 잘못 잡히는 걸 막기 위함."""
    raw = pd.read_excel(xls, sheet_name=sheet, header=None)
    raw = _drop_hidden_rows(raw, hidden_rows)
    bounds = find_sections(raw)
    if "monthly" in bounds:
        data, date_idx = section_dataframe(raw, *bounds["monthly"], date_tokens=("기간", "월별"))
    else:
        # '■ 월간 데이터' 같은 섹션 제목이 없는 단순 시트는 기존 방식으로 폴백
        hdr = find_header_row(raw)
        if hdr is None:
            return None, pd.DataFrame()
        headers = raw.iloc[hdr].tolist()
        date_idx = next((i for i, h in enumerate(headers) if clean_col(h) in ("기간", "월별")), None)
        if date_idx is None:
            return None, pd.DataFrame()
        data = raw.iloc[hdr + 1 :].copy()
        data.columns = headers
        data = data[data.iloc[:, date_idx].notna()]

    weekly_out = pd.DataFrame()
    if "weekly" in bounds:
        weekly_out = _parse_weekly_section(raw, bounds, today)
        if not weekly_out.empty:
            weekly_out["channel"] = sheet

    if data is None or data.empty:
        return None, weekly_out
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
    return out.reset_index(drop=True), weekly_out


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


INFLOW_REVENUE_SHEET = "일별 GA,어드민 지표 비교"

# (엑셀 열 위치 0-based, 우리 컬럼명, 퍼센트 변환 배수) — 이름이 아니라 '위치'로 읽는 이유:
# 시트 헤더가 4행이고 그 위에 병합된 그룹 헤더(GA/자체 분석 툴/매체 기준/GA 기준)가 있어서
# "결제수·CVR·AOV·매출액·ROAS"가 두 번씩 반복되는데, pandas로 이름 기준 read_excel을 하면
# 뒤에 나온 컬럼이 앞 컬럼을 덮어써버려 데이터가 사라진다. 열 위치 고정이 훨씬 안전하다.
INFLOW_REVENUE_COL_MAP = [
    (0, "report_date", None),
    (1, "users", None),
    (2, "new_users", None),
    (3, "returning_users", None),
    (4, "signups", None),
    (5, "bounce_rate", 100),
    # 6번(인덱스6, 숫자 초 단위)은 실제로는 대부분 비어있고(372행 중 31행만 값 존재),
    # 7번(인덱스7, "3분 26초" 텍스트)이 거의 매일 채워져 있어서 이쪽을 파싱해서 쓴다 — 아래
    # avg_session_duration 처리 참고. 6번은 그래서 COL_MAP에서 건너뛴다.
    (8, "pageviews", None),
    (9, "admin_orders_total", None),
    (10, "admin_orders_real", None),
    (11, "admin_qty", None),
    (12, "admin_revenue", None),          # 회사 내부(어드민) 기준 매출
    (13, "price_rate", 100),
    (14, "admin_aov", None),
    (15, "admin_returns", None),
    (16, "return_rate", 100),
    # 17번(인덱스17)은 "매체 TOTAL" 아래 날짜 중복 컬럼이라 건너뜀
    (18, "impressions", None),
    (19, "clicks", None),
    (20, "ctr", 100),
    (21, "cpc", None),
    (22, "cost_incl_vat", None),
    (23, "conversions", None),            # 매체 리포트 기준 결제수
    (24, "cvr", 100),
    (25, "aov", None),
    (26, "revenue", None),                # 매체 리포트(보고서) 기준 매출
    (27, "roas", 100),
    (28, "ga_conversions", None),
    (29, "ga_cvr", 100),
    (30, "ga_aov", None),
    (31, "ga_revenue", None),             # GA 기준 매출
    (32, "ga_roas", 100),
    (33, "new_paying_customers", None),
    (34, "cac", None),
    # 35~39(인덱스)는 퍼널(1.유입~5.구매) 라벨 텍스트 칸이라 실데이터가 아니어서 제외
]


def parse_inflow_revenue_sheet(xls: pd.ExcelFile):
    """형이 별도로 정리해서 준 '일별 GA,어드민 지표 비교' 파일 전용 파서.
    5~11행에는 월별 요약행이 먼저 있고 그 아래(12행~)부터 실제 일별 데이터가 시작되는데,
    요약행은 날짜 칸이 '2026년 1월' 같은 문자열이라 datetime 파싱이 실패하는 것으로 자연스럽게
    걸러낸다. 아직 값이 없는 미래 날짜(연말까지 미리 만들어둔 빈 행)는 방문자수가 비어있는 것으로
    걸러낸다."""
    if INFLOW_REVENUE_SHEET not in xls.sheet_names:
        return pd.DataFrame()
    raw = pd.read_excel(xls, sheet_name=INFLOW_REVENUE_SHEET, header=None, skiprows=4)
    if raw.empty:
        return pd.DataFrame()

    out = pd.DataFrame(index=raw.index)
    for idx, name, scale in INFLOW_REVENUE_COL_MAP:
        if idx >= raw.shape[1]:
            out[name] = pd.NA
            continue
        col = raw.iloc[:, idx]
        if name == "report_date":
            out[name] = col
            continue
        col = pd.to_numeric(col, errors="coerce")
        if scale:
            col = col * scale
        out[name] = col

    # 평균 체류시간 — "3분 26초" 같은 텍스트(8번째 열, 인덱스7)를 초 단위 숫자로 변환.
    if raw.shape[1] > 7:
        dur_text = raw.iloc[:, 7].astype(str)
        dur_match = dur_text.str.extract(r"(?:(\d+)\s*분)?\s*(?:(\d+)\s*초)?")
        minutes = pd.to_numeric(dur_match[0], errors="coerce").fillna(0)
        seconds = pd.to_numeric(dur_match[1], errors="coerce").fillna(0)
        parsed = minutes * 60 + seconds
        parsed[~dur_text.str.contains("분|초", na=False)] = pd.NA
        out["avg_session_duration"] = parsed
    else:
        out["avg_session_duration"] = pd.NA

    out["report_date"] = pd.to_datetime(out["report_date"], errors="coerce")
    out = out.dropna(subset=["report_date", "users"]).copy()
    out["report_date"] = out["report_date"].dt.date
    return out.reset_index(drop=True)


# GA 내보내기 화면/버전마다 컬럼 순서·개수가 달라질 수 있어(뒤에 빈 컬럼이나 메모가 붙기도
# 하고, '거래'/'세션수' 같은 컬럼이 추가/누락되기도 함) 위치가 아니라 헤더 '이름'으로 찾는다.
# 후보를 여러 개 둔 건 GA4 내보내기 화면 언어/버전에 따라 라벨이 조금씩 달라질 수 있어서다.
GA_CHANNEL_INFLOW_HEADER_CANDIDATES = {
    "report_date": ["날짜"],
    "source_medium": ["세션 소스/매체", "세션 소스 / 매체", "소스/매체", "소스 / 매체"],
    "users": ["총 사용자", "사용자"],
    "new_users": ["새 사용자 수", "새 사용자"],
    "returning_users": ["재 사용자", "재방문 사용자", "재방문자"],
    "bounce_rate": ["이탈률"],
    "pageviews": ["조회수"],
    "avg_session_duration": ["평균 세션 시간", "평균 참여 시간", "평균 체류 시간"],
    "conversions": ["거래", "구매", "전환수"],
    "revenue": ["총수익", "수익", "매출"],
}
GA_CHANNEL_INFLOW_PCT_FIELDS = {"bounce_rate"}  # 0~1 비율로 들어오면 %로 쓰기 위해 100을 곱함


def _match_header_index(columns: list, candidates: list) -> int:
    """컬럼명 리스트에서 후보 이름과 '정확히' 일치하는 첫 컬럼의 위치를 찾는다. 정확히 일치하는
    것만 인정해서(부분일치 X) '조회수'와 '세션당 조회수'처럼 한쪽이 다른 쪽을 포함하는 헤더가
    섞여 있어도 엉뚱한 컬럼을 잡지 않는다."""
    cols_clean = [str(c).strip() for c in columns]
    for cand in candidates:
        for i, c in enumerate(cols_clean):
            if c == cand:
                return i
    return None


def parse_ga_channel_inflow_sheet(xls: pd.ExcelFile, channel_map: dict = None):
    """'GA 세션소스 매체&일자별 데이터' 같은 GA 원본 내보내기 파일 전용 파서.
    날짜×세션 소스/매체 단위로 하루 단위 데이터가 들어있다(하루에 소스/매체별로 한 행씩).
    '매체'(채널 그룹핑, 예: (SA) 네이버/구글(P-MAX) 등)는 UTM 소스/매체 원본에는 없는 값이라,
    별도로 학습해둔 utm_channel_map(build_utm_channel_lookup) 매핑을 channel_map으로 받아
    source_medium 문자열로 채운다(대소문자·공백 차이는 무시하고 매칭). 매핑에 없는 값은
    None으로 남겨 화면에서 '미매핑'으로 구분할 수 있게 한다."""
    if not xls.sheet_names:
        return pd.DataFrame()
    sheet = xls.sheet_names[0]
    raw = pd.read_excel(xls, sheet_name=sheet, header=0)
    if raw.empty:
        return pd.DataFrame()

    columns = list(raw.columns)
    idx_map = {
        name: _match_header_index(columns, candidates)
        for name, candidates in GA_CHANNEL_INFLOW_HEADER_CANDIDATES.items()
    }
    # 날짜/소스매체 컬럼을 못 찾으면 이 파일은 GA 채널 유입 파일이 아니라고 보고 빈 결과를 낸다
    # (같은 업로더에서 다른 파서들도 같이 시도되므로, 여기서 조용히 빠지는 게 맞다).
    if idx_map.get("report_date") is None or idx_map.get("source_medium") is None:
        return pd.DataFrame()

    out = pd.DataFrame(index=raw.index)
    for name, idx in idx_map.items():
        if idx is None:
            out[name] = pd.NA
            continue
        col = raw.iloc[:, idx]
        if name in ("report_date", "source_medium"):
            out[name] = col
            continue
        col = pd.to_numeric(col, errors="coerce")
        if name in GA_CHANNEL_INFLOW_PCT_FIELDS:
            col = col * 100
        out[name] = col

    # 날짜가 "2026. 1. 1." 같은 텍스트 포맷이라 고정 포맷으로 우선 파싱하고, 혹시 다른 형식이
    # 섞여 들어와도 죽지 않도록 실패한 행만 일반 파싱으로 한 번 더 시도한다.
    parsed = pd.to_datetime(out["report_date"], errors="coerce", format="%Y. %m. %d.")
    fallback = parsed.isna()
    if fallback.any():
        parsed.loc[fallback] = pd.to_datetime(out.loc[fallback, "report_date"], errors="coerce")
    out["report_date"] = parsed

    out = out.dropna(subset=["report_date", "source_medium"]).copy()
    out["report_date"] = out["report_date"].dt.date
    if channel_map:
        out["channel"] = out["source_medium"].map(
            lambda sm: channel_map.get(str(sm).strip().lower())
        )
    else:
        out["channel"] = None
    return out.reset_index(drop=True)


def parse_utm_channel_map(xls: pd.ExcelFile) -> pd.DataFrame:
    """'STCO_UTM 리스트' 같은 UTM 소스/매체 ↔ 보고서 매체명 매핑표 파일을 파싱한다.
    헤더 위치가 고정돼 있지 않을 수 있어(형이 셀을 밀거나 시트를 늘릴 수 있으니), '매체명'과
    '소스'/'매체'가 같이 있는 행을 헤더로 찾고, 그 아래 두 컬럼(매체명/소스·매체)이 둘 다
    채워진 행만 데이터로 읽는다. '비고' 컬럼이 있으면 참고용으로 같이 저장한다."""
    rows = []
    for sheet in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name=sheet, header=None)
        if raw.empty:
            continue
        header_row, chan_col, src_col, note_col = None, None, None, None
        for r in range(min(10, len(raw))):
            row_vals = raw.iloc[r].tolist()
            c_idx = next((i for i, v in enumerate(row_vals) if "매체명" in str(v)), None)
            s_idx = next((i for i, v in enumerate(row_vals) if "소스" in str(v) and "매체" in str(v)), None)
            if c_idx is not None and s_idx is not None:
                header_row, chan_col, src_col = r, c_idx, s_idx
                note_col = next((i for i, v in enumerate(row_vals) if "비고" in str(v)), None)
                break
        if header_row is None:
            continue
        for r in range(header_row + 1, len(raw)):
            channel = raw.iloc[r, chan_col] if chan_col < raw.shape[1] else None
            source_medium = raw.iloc[r, src_col] if src_col < raw.shape[1] else None
            if pd.isna(channel) or pd.isna(source_medium) or not str(channel).strip() or not str(source_medium).strip():
                continue
            note = raw.iloc[r, note_col] if note_col is not None and note_col < raw.shape[1] else None
            rows.append({
                "source_medium": str(source_medium).strip(),
                "channel": str(channel).strip(),
                "note": str(note).strip() if pd.notna(note) else None,
            })
    if not rows:
        return pd.DataFrame()
    out = pd.DataFrame(rows)
    # 같은 소스/매체가 여러 시트/행에 중복돼 있으면 마지막(가장 아래) 값을 채택한다.
    out = out.drop_duplicates(subset=["source_medium"], keep="last").reset_index(drop=True)
    return out


BUDGET_SCOPE_TITLES = {
    "자사몰": "자사몰 현황",
    # "외부몰": "외부몰 현황",  # 구조가 더 복잡해서(직입점/벤더 중첩) 다음 단계에서 추가 예정
}
# 실제로 시트에 등장하는 모든 섹션 제목(파싱 대상이 아닌 것도 포함) — 섹션 경계(어디까지가
# '자사몰 현황' 구간인지)를 정확히 자르기 위해 필요하다. 하나의 시트 안에 온라인사업팀/자사몰/
# 외부몰 현황이 위아래로 이어져 있을 수 있어서, 자사몰 구간이 다음 섹션(외부몰 현황) 내용까지
# 삼켜버리지 않도록 전체 목록으로 경계를 잡는다.
ALL_BUDGET_SECTION_TITLES = ["온라인사업팀 현황", "자사몰 현황", "외부몰 현황"]
BUDGET_YEAR_LABEL = "26년"
BUDGET_YEAR_NUM = 2000 + int(re.sub(r"\D", "", BUDGET_YEAR_LABEL))
BUDGET_SUBSECTION_DETAIL = "매체 세부내역"


def _find_budget_month_columns(raw: pd.DataFrame, start_row: int = 0, end_row: int | None = None):
    """[start_row, end_row) 범위 안에서 '1월'~'12월'이 순서대로 나열된 헤더 행을 찾아
    (헤더 행 번호, {월번호: 컬럼위치})를 반환한다. 10개월 이상 잡히는 첫 행을 헤더로 인정한다
    (연말 등 일부 컬럼이 빠져도 허용).

    섹션(온라인사업팀 현황/자사몰 현황/외부몰 현황 등)마다 표가 따로 그려져 있어서 헤더 행 위치와
    컬럼 배치가 섹션마다 다를 수 있다 — 그래서 시트 전체에서 한 번만 찾지 않고, 각 섹션 구간
    안에서 지역적으로(local) 찾는다."""
    end_row = len(raw) if end_row is None else min(end_row, len(raw))
    for r in range(max(start_row, 0), end_row):
        month_cols = {}
        for i, v in enumerate(raw.iloc[r].tolist()):
            m = re.match(r"^(\d{1,2})\s*월$", str(v).strip())
            if m and 1 <= int(m.group(1)) <= 12:
                month_cols[int(m.group(1))] = i
        if len(month_cols) >= 10:
            return r, month_cols
    return None, {}


def parse_channel_budget_sheet(xls: pd.ExcelFile) -> pd.DataFrame:
    """'◆26년 월별 예산 정리' 파일에서 '자사몰 현황' 섹션의 26년 매체별 월간 예산(광고선전비)을
    파싱한다. 형이 캡쳐해준 화면 구조를 보고 만든 버전 — 한 시트 안에 온라인사업팀 현황/자사몰
    현황/외부몰 현황이 위아래로 이어져 있고 섹션마다 표 레이아웃(헤더 행 위치, 라벨 컬럼 개수)이
    다를 수 있다고 보고, '1월'~'12월' 헤더를 시트 전체에서 한 번만 찾지 않고 섹션(자사몰 현황)
    구간 안에서 지역적으로 찾도록 만들었다. 매체 세부내역 각 행 + '자사몰' 전체 합계(광고선전비
    1월~12월) 행을 channel='TOTAL'로 같이 저장한다. 원본이 천원 단위(+VAT)라 1000을 곱해 원
    단위로 맞춘다(다른 테이블들과 단위를 통일하기 위함). 합성 데이터로는 검증했지만 실제 파일
    바이트 단위 검증은 아직이라, 형이 앱 업데이트 후 업로드해보고 여전히 안 잡히면 바로 알려주면
    좋겠다."""
    rows = []
    for sheet in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name=sheet, header=None)
        if raw.empty:
            continue

        # 시트 안에 등장하는 모든 섹션 제목의 위치를 먼저 전부 찾아둔다 (온라인사업팀 현황이
        # 자사몰 현황보다 앞에 있을 수 있고, 그 위쪽에 다른 요약 표가 더 있을 수도 있어서 —
        # 헤더/컬럼 위치를 시트 전체에서 한 번만 찾지 않고 섹션별로 다시 찾기 위한 기준점).
        title_positions = []  # [(row_idx, title), ...] 등장 순서대로
        for r in range(len(raw)):
            row_str = raw.iloc[r].astype(str)
            for t in ALL_BUDGET_SECTION_TITLES:
                if row_str.str.contains(t, na=False, regex=False).any():
                    title_positions.append((r, t))

        for scope, scope_title in BUDGET_SCOPE_TITLES.items():
            scope_starts = [r for r, t in title_positions if t == scope_title]
            if not scope_starts:
                continue
            start = scope_starts[0]
            # 다음 섹션 제목(어떤 제목이든)이 나오는 행 전까지가 이 scope 구간.
            later_titles = [r for r, t in title_positions if r > start]
            end = later_titles[0] if later_titles else len(raw)

            # 이 섹션 구간 안에서 지역적으로 '1월'~'12월' 헤더 행/컬럼 위치를 찾는다 — 섹션마다
            # 표 레이아웃(라벨 컬럼 개수 등)이 다를 수 있어 전역 헤더를 재사용하면 엉뚱한
            # 컬럼에서 값을 읽게 된다.
            header_row, month_cols = _find_budget_month_columns(raw, start_row=start, end_row=end)
            if header_row is None or not month_cols:
                continue
            label_col_end = min(month_cols.values())

            # 26년 블록: '26년' 텍스트가 나오는 행부터 '25년' 텍스트가 나오는 행 전까지.
            year_start_candidates = [
                r for r in range(start, end)
                if raw.iloc[r].astype(str).str.contains(BUDGET_YEAR_LABEL, na=False, regex=False).any()
            ]
            if not year_start_candidates:
                continue
            y_start = year_start_candidates[0]
            year_end_candidates = [
                r for r in range(y_start + 1, end)
                if raw.iloc[r].astype(str).str.contains("25년", na=False, regex=False).any()
            ]
            y_end = year_end_candidates[0] if year_end_candidates else end

            # scope 전체 합계: '광고선전비 (1월~12월)' 행 (매출/광고비 요약 블록에 있음, 매체
            # 세부내역보다 위쪽) → channel='TOTAL'로 저장.
            total_row_idx = [
                r for r in range(y_start, y_end)
                if raw.iloc[r].astype(str).str.contains(r"광고선전비\s*\(?\s*1\s*월\s*~\s*12\s*월", na=False, regex=True).any()
            ]
            if total_row_idx:
                r = total_row_idx[0]
                for m, c in month_cols.items():
                    v = pd.to_numeric(raw.iat[r, c], errors="coerce")
                    if pd.notna(v):
                        rows.append({"scope": scope, "channel": "TOTAL", "month": m, "budget_cost": float(v) * 1000})

            # 매체 세부내역: 라벨 행 다음부터 y_end까지, 각 행의 첫 문자열 셀(라벨 컬럼 범위 안)을
            # 매체명으로 쓴다. 숫자로 시작하는 값/빈 라벨/전부 빈 월별 값인 행은 건너뛴다.
            detail_start_candidates = [
                r for r in range(y_start, y_end)
                if raw.iloc[r].astype(str).str.contains(BUDGET_SUBSECTION_DETAIL, na=False, regex=False).any()
            ]
            if not detail_start_candidates:
                continue
            d_start = detail_start_candidates[0] + 1
            for r in range(d_start, y_end):
                label = None
                for c in range(label_col_end):
                    v = raw.iat[r, c]
                    if isinstance(v, str) and v.strip() and not re.match(r"^\d", v.strip()):
                        label = v.strip()
                if not label:
                    continue
                any_val = False
                for m, c in month_cols.items():
                    v = pd.to_numeric(raw.iat[r, c], errors="coerce")
                    if pd.notna(v):
                        any_val = True
                        rows.append({"scope": scope, "channel": label, "month": m, "budget_cost": float(v) * 1000})
                if not any_val:
                    continue

    if not rows:
        return pd.DataFrame()
    out = pd.DataFrame(rows)
    out["year"] = BUDGET_YEAR_NUM
    out = out.groupby(["scope", "channel", "year", "month"], as_index=False)["budget_cost"].sum()
    return out


def build_utm_channel_lookup(utm_map: pd.DataFrame) -> dict:
    """utm_channel_map 테이블(source_medium/channel)을 대소문자·공백 차이에 안전한
    lookup dict로 바꾼다 — GA 원본 값이 'Naver / cpc'처럼 대소문자가 섞여 올 수 있어서."""
    if utm_map is None or utm_map.empty:
        return {}
    return {
        str(sm).strip().lower(): ch
        for sm, ch in zip(utm_map["source_medium"], utm_map["channel"])
        if pd.notna(sm) and pd.notna(ch)
    }


def _ga_daily_agg(ga_channel_inflow: pd.DataFrame) -> pd.DataFrame:
    """ga_channel_inflow(일자×소스/매체, UTM 매핑 완료)를 날짜 단위로 합산해 ga_conversions/
    ga_revenue를 만든다. '일자별 누적' 표에 GA 컬럼을 붙일 때 쓴다. 우리가 실제로 운영 중인
    매체(UTM 매핑표에 있는 소스/매체)만 합산한다 — 자연유입/direct/referral/이메일 등은
    광고 성과가 아니라서 여기 섞이면 자체 ROAS랑 비교가 안 맞게 된다."""
    if ga_channel_inflow is None or ga_channel_inflow.empty:
        return pd.DataFrame(columns=["report_date", "ga_conversions", "ga_revenue"])
    g = ga_channel_inflow.copy()
    g = g[g["channel"].notna()]
    if g.empty:
        return pd.DataFrame(columns=["report_date", "ga_conversions", "ga_revenue"])
    g["report_date"] = pd.to_datetime(g["report_date"]).dt.date
    return g.groupby("report_date", as_index=False).agg(
        ga_conversions=("conversions", "sum"), ga_revenue=("revenue", "sum")
    )


def _ga_daily_agg_all(ga_channel_inflow: pd.DataFrame) -> pd.DataFrame:
    """_ga_daily_agg와 반대로, 매체 매핑 여부와 상관없이(자연유입/direct/referral/이메일 등
    포함) 날짜 단위로 전부 합산한다. '유입·매출 비교' 페이지의 GA-매출은 어드민(사이트 전체)
    매출과 비교하는 용도라, 광고 채널만 걸러내면 안 되고 GA가 보는 전체를 그대로 써야 한다."""
    if ga_channel_inflow is None or ga_channel_inflow.empty:
        return pd.DataFrame(columns=["report_date", "ga_conversions_all", "ga_revenue_all"])
    g = ga_channel_inflow.copy()
    g["report_date"] = pd.to_datetime(g["report_date"]).dt.date
    return g.groupby("report_date", as_index=False).agg(
        ga_conversions_all=("conversions", "sum"), ga_revenue_all=("revenue", "sum")
    )


def _ga_weekly_agg(ga_channel_inflow: pd.DataFrame, weekly: pd.DataFrame) -> pd.DataFrame:
    """'주간별 누적' 표의 각 주(week_start~week_end) 구간에 맞춰 ga_channel_inflow를
    합산한다. 리포트 원본 주간 섹션엔 GA 컬럼이 아예 없어서, 이 방식으로만 채울 수 있다.
    _ga_daily_agg와 마찬가지로 UTM 매핑된(=우리가 운영 중인) 매체만 합산한다."""
    if ga_channel_inflow is None or ga_channel_inflow.empty or weekly is None or weekly.empty:
        return pd.DataFrame(columns=["week_start", "ga_conversions", "ga_revenue"])
    g = ga_channel_inflow.copy()
    g = g[g["channel"].notna()]
    if g.empty:
        return pd.DataFrame(columns=["week_start", "ga_conversions", "ga_revenue"])
    g["report_date"] = pd.to_datetime(g["report_date"]).dt.date
    rows = []
    for _, w in weekly.iterrows():
        mask = (g["report_date"] >= w["week_start"]) & (g["report_date"] <= w["week_end"])
        rows.append({
            "week_start": w["week_start"],
            "ga_conversions": g.loc[mask, "conversions"].sum(),
            "ga_revenue": g.loc[mask, "revenue"].sum(),
        })
    return pd.DataFrame(rows)


def _ga_channel_agg(ga_channel_inflow: pd.DataFrame, start: date, end: date) -> pd.DataFrame:
    """선택된 기간 안에서, UTM 매핑이 된(channel이 채워진) 행만 매체 단위로 합산한다.
    '매체별 성과' 표에 (기간 선택에 맞춰 갱신되는) GA 컬럼을 붙일 때 쓴다."""
    if ga_channel_inflow is None or ga_channel_inflow.empty:
        return pd.DataFrame(columns=["channel", "ga_conversions", "ga_revenue"])
    g = ga_channel_inflow.copy()
    g["report_date"] = pd.to_datetime(g["report_date"]).dt.date
    g = g[(g["report_date"] >= start) & (g["report_date"] <= end) & g["channel"].notna()]
    if g.empty:
        return pd.DataFrame(columns=["channel", "ga_conversions", "ga_revenue"])
    return g.groupby("channel", as_index=False).agg(
        ga_conversions=("conversions", "sum"), ga_revenue=("revenue", "sum")
    )


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


def backfill_missing_creative_images(creatives: pd.DataFrame) -> pd.DataFrame:
    """이번 주 엑셀에 이미지가 안 박혀있거나 인식이 안 돼서 image_url이 비는 소재는,
    과거에 저장된 같은 매체+소재명의 최근 이미지 URL이 있으면 그대로 이어붙인다.
    Storage 경로가 소재명 해시로 고정돼 있어 파일 자체는 계속 남아있는데, 이번 주
    스냅샷 행에는 그 URL을 다시 채워주지 않으면 이미지가 빈 걸로 보이는 문제가 있었음
    (2026-08 발견). 소재명이 같으면 매체가 같은 것으로 보고 가장 최근 저장분의 URL을 쓴다."""
    if creatives is None or creatives.empty:
        return creatives
    creatives = creatives.copy()
    if "image_url" not in creatives.columns:
        creatives["image_url"] = None

    missing_mask = creatives["image_url"].isna() | (creatives["image_url"].astype(str) == "")
    if not missing_mask.any():
        return creatives

    prev = load_table("creative_performance")
    if prev.empty or "image_url" not in prev.columns:
        return creatives
    prev = prev[prev["image_url"].notna() & (prev["image_url"].astype(str) != "")]
    if prev.empty:
        return creatives
    if "as_of_date" in prev.columns:
        prev = prev.sort_values("as_of_date")
    lookup = prev.drop_duplicates(subset=["channel", "creative"], keep="last").set_index(
        ["channel", "creative"]
    )["image_url"]

    def _fill(row):
        cur = row["image_url"]
        if pd.notna(cur) and str(cur) != "":
            return cur
        return lookup.get((row["channel"], row["creative"]))

    creatives.loc[missing_mask, "image_url"] = creatives.loc[missing_mask].apply(_fill, axis=1)
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
        "channels_weekly": pd.DataFrame(),
        "ga": pd.DataFrame(),
        "creatives": pd.DataFrame(),
        "channel_audience": pd.DataFrame(),
        "channel_sheets_found": [],
        "channel_sheets_parsed": [],
        "creative_sheets_found": [],
        "agency_notes": pd.DataFrame(),
    }
    if "매체통합" in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name="매체통합", header=None)
        raw = _drop_hidden_rows(raw, get_hidden_rows(wb_for_hidden, "매체통합"))
        bounds = find_sections(raw)
        result["monthly"] = parse_monthly(raw, bounds, today)
        result["weekly"] = parse_weekly(raw, bounds, today)
        result["daily"] = parse_daily(raw, bounds, today)
        result["channel_snapshot"] = parse_channel_snapshot(raw, bounds, result["monthly"])
        result["agency_notes"] = parse_agency_notes(raw, today)

    chan_frames = []
    chan_weekly_frames = []
    channel_sheet_candidates = discover_channel_sheets(xls)
    if visible_sheets is not None:
        channel_sheet_candidates = [s for s in channel_sheet_candidates if s in visible_sheets]
    for s in channel_sheet_candidates:
        result["channel_sheets_found"].append(s)
        df, weekly_df = parse_channel_sheet(xls, s, today, hidden_rows=get_hidden_rows(wb_for_hidden, s))
        if df is not None and len(df):
            chan_frames.append(df)
            result["channel_sheets_parsed"].append(s)
        if weekly_df is not None and len(weekly_df):
            chan_weekly_frames.append(weekly_df)
    result["channels_weekly"] = (
        pd.concat(chan_weekly_frames, ignore_index=True) if chan_weekly_frames else pd.DataFrame()
    )
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
    # ↓ 유입·매출 비교(일별 GA·어드민 지표 비교) 전용 컬럼
    "returning_users": "재방문자",
    "pageviews": "페이지뷰",
    "admin_orders_total": "총 결제건수(어드민)",
    "admin_orders_real": "실 결제건수(어드민)",
    "admin_qty": "실 상품수량(어드민)",
    "admin_revenue": "어드민 매출(회사 내부 기준)",
    "price_rate": "판가율(%)",
    "admin_aov": "평균 구매금액(어드민)",
    "admin_returns": "반품건수",
    "return_rate": "반품율(%)",
    "ga_cvr": "GA-CVR(%)",
    "ga_aov": "GA-객단가",
    "new_paying_customers": "신규결제고객수",
    "cac": "CAC(고객획득비용)",
    "report_gap_pct": "보고서 매출 격차(%)",
    "ga_gap_pct": "GA 매출 격차(%)",
}


def korify(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns=KOR_COLS)


# 종합 대시보드(주간 추이)만 예전처럼 전체 프리셋을 쓰고, 나머지 페이지는 아래 축약판을 쓴다
# (형이 "오늘/어제/지난주/지난달/이번달/직접선택 6개만 남기고 줄여달라"고 요청 — 종합 대시보드는 제외).
DATE_PRESETS_FULL = [
    "오늘", "어제", "이번주", "지난주",
    "최근 7일(오늘 포함)", "최근 7일(오늘 제외)",
    "이번달", "지난달",
    "최근 30일(오늘 포함)", "최근 30일(오늘 제외)",
]
DATE_PRESETS_SHORT = ["오늘", "어제", "지난주", "지난달", "이번달"]


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


DATE_PERIOD_OPTIONS_FULL = DATE_PRESETS_FULL + ["전체", "직접선택"]
DATE_PERIOD_OPTIONS = DATE_PRESETS_SHORT + ["직접선택"]


def period_filter(min_d: date, max_d: date, key: str, default_preset: str = "이번달", options: list = None):
    """날짜 프리셋 버튼 목록(누적 표의 preset_button_picker와 동일한 방식) + 직접선택 달력.
    버튼 아래에 실제로 적용된 날짜범위를 항상 캡션으로 보여준다. 반환값은 (start, end).
    options를 안 주면 축약판(오늘/어제/지난주/지난달/이번달/직접선택) 6개만 쓴다."""
    preset = preset_button_picker(options or DATE_PERIOD_OPTIONS, key=f"{key}_dateperiod", default=default_preset)

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
    "returning_users", "pageviews", "admin_orders_total", "admin_orders_real", "admin_qty",
    "admin_revenue", "admin_aov", "admin_returns", "ga_aov", "new_paying_customers", "cac",
}
PCT2_COLS = {"ctr", "cvr", "bounce_rate", "ecommerce_cvr", "signup_rate", "price_rate", "return_rate", "ga_cvr"}
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


def render_html_table(table: pd.DataFrame, raw: pd.DataFrame = None, raw_label_map: dict = None):
    """pandas Styler(jinja2 의존) 없이 순수 HTML로 표를 그린다.
    ▲(상승)는 빨간색, ▼(하락)는 파란색 글씨로 표시하고, 인덱스는 표시하지 않는다.
    raw(원본 숫자 컬럼, TOTAL 행 제외 · table의 데이터 행과 같은 순서)를 넘기면 raw에 있는
    숫자 컬럼에 한해 헤더를 클릭해 오름차순/내림차순 정렬할 수 있다(TOTAL 행은 항상 고정).
    raw_label_map(화면표시 헤더 → 원본 컬럼명)을 따로 주면, 공용 KOR_COLS 대신 이 표 전용
    라벨(예: "매체-매출")로 정렬 매칭을 한다 — 페이지마다 다른 이름을 쓰고 싶을 때 사용."""
    if table.empty:
        st.caption("데이터가 아직 없습니다.")
        return

    cols = list(table.columns)

    sortable_keys = {}
    if raw is not None:
        for c in cols:
            orig = (raw_label_map or {}).get(c) or KOR_COLS_REV.get(c)
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
        # 주의: 문자열 내용이 4칸 이상 들여쓰기 되면 마크다운이 "코드 블록"으로 인식해 HTML을
        # 렌더링하지 않고 태그를 그대로 텍스트로 보여준다 — 그래서 내용은 항상 왼쪽에 맞춰 쓴다.
        html = (
            table_style
            + '<div class="stco-table-wrap">'
            + '<table class="stco-table">'
            + f"<thead><tr>{thead}</tr></thead>"
            + f"<tbody>{''.join(row_htmls)}</tbody>"
            + "</table>"
            + "</div>"
        )
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
            cw = result.get("channels_weekly", pd.DataFrame())
            st.write(f"📆 매체별 주간 데이터 인식: {len(cw)}행 (매체 {cw['channel'].nunique() if not cw.empty else 0}종)")
            st.write(f"🔎 GA 유입경로: {len(result['ga'])}건")
            st.write(
                f"🎨 소재별 성과 인식: {len(result.get('creatives', []))}행 "
                f"({', '.join(result.get('creative_sheets_found', [])) or '해당 시트 없음'})"
            )
            agency_notes_df = result.get("agency_notes", pd.DataFrame())
            st.write(
                f"📝 매체통합 시트 하단 운영 메모: "
                f"{'인식됨' if not agency_notes_df.empty else '인식 안 됨(없거나 형식 미확인)'}"
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
                    before_missing = int(
                        creatives_df["image_url"].isna().sum() if "image_url" in creatives_df.columns
                        else len(creatives_df)
                    )
                    creatives_df = backfill_missing_creative_images(creatives_df)
                    after_missing = int(creatives_df["image_url"].isna().sum())
                    backfilled_n = before_missing - after_missing
                    if backfilled_n > 0:
                        st.write(f"🔁 이번 파일엔 이미지가 없어 이전 저장분 이미지로 채운 소재: {backfilled_n}개")

            n1 = save_table("weekly_overview", result["weekly"], "week_start", file.name)
            n2 = save_table("monthly_overview", result["monthly"], "report_month", file.name)
            n3 = save_table("channel_monthly", result["channels"], "report_month,channel", file.name)
            n3w = save_table("channel_weekly", result.get("channels_weekly", pd.DataFrame()), "channel,week_start", file.name)
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
            n9 = save_table(
                "agency_notes", result.get("agency_notes", pd.DataFrame()),
                "as_of_date", file.name,
            )
            st.cache_data.clear()
            st.sidebar.success(
                f"저장 완료! 주간 {n1} · 월별 {n2} · 일자별 {n6} · 매체(월) {n3} · 매체(주간) {n3w} · 매체(당월) {n4} · "
                f"GA {n5}건 · 소재 {n7}건 · 퍼널(신규/리타겟) {n8}건 · 운영 메모 {n9}건"
            )
            st.rerun()

    st.sidebar.markdown("---")
    ga_file = st.sidebar.file_uploader(
        "② GA 유입 데이터 업로드 (유입·매출 비교 / GA 매체별 유입 경로 / UTM 매핑표 xlsx)",
        type=["xlsx", "xls"], key="ga_combined_uploader",
    )
    if ga_file is not None:
        with st.sidebar.status("파일 분석 중...", expanded=True) as status2:
            # 세 파일은 시트 구조가 서로 달라서, 하나의 업로더에서 세 파서를 다 시도해보고
            # 실제로 데이터가 나온 쪽만 채택한다(맞지 않는 파서는 항상 빈 결과를 내도록 이미
            # 각 파서 안에 시트명/필수컬럼 가드가 있어서 안전하게 겸용 가능).
            ga_xls = pd.ExcelFile(ga_file)
            inflow_df = parse_inflow_revenue_sheet(ga_xls)
            utm_map_df = parse_utm_channel_map(ga_xls)
            # 이번에 새로 올라온 매핑표 + 이미 저장돼있는 매핑표를 합쳐서 '매체' 채널 그룹핑에 쓴다
            # (새 파일에 없는 소스/매체는 기존에 저장된 매핑을 그대로 쓰기 위함).
            existing_utm_map = load_table("utm_channel_map")
            combined_utm_map = pd.concat([existing_utm_map, utm_map_df], ignore_index=True) if not utm_map_df.empty else existing_utm_map
            channel_lookup = build_utm_channel_lookup(combined_utm_map)
            ga_channel_df = parse_ga_channel_inflow_sheet(ga_xls, channel_map=channel_lookup)

            if inflow_df.empty and ga_channel_df.empty and utm_map_df.empty:
                st.warning(
                    "인식 가능한 데이터를 찾지 못했습니다. '일별 GA,어드민 지표 비교' 시트가 있는 "
                    "유입·매출 비교 파일이거나, 날짜/세션 소스/매체 컬럼이 있는 GA 매체별 유입 경로 "
                    "파일, 또는 '매체명'/'소스 / 매체' 컬럼이 있는 UTM 매핑표 파일인지 확인해주세요."
                )
            if not inflow_df.empty:
                st.write(
                    f"🚶 일별 유입·매출 데이터 인식: {len(inflow_df)}일 "
                    f"({inflow_df['report_date'].min()} ~ {inflow_df['report_date'].max()})"
                )
            if not utm_map_df.empty:
                st.write(f"🗺️ UTM 소스/매체 매핑 인식: {len(utm_map_df)}행")
            if not ga_channel_df.empty:
                mapped_n = ga_channel_df["channel"].notna().sum()
                st.write(
                    f"🔎 GA 매체별 유입 경로 인식: {len(ga_channel_df)}행 "
                    f"({ga_channel_df['report_date'].min()} ~ {ga_channel_df['report_date'].max()}, "
                    f"소스/매체 {ga_channel_df['source_medium'].nunique()}종, 매체 매핑 {mapped_n}행)"
                )
                if mapped_n == 0:
                    st.caption("※ '매체'(채널 그룹핑) 컬럼이 아직 비어있습니다 — UTM 매핑표를 같이/먼저 올리면 채워집니다.")
                elif mapped_n < len(ga_channel_df):
                    # source_medium에 문자열이 아닌 값(숫자로 잘못 인식된 셀 등)이 섞여 있으면
                    # sorted()/join()이 타입 비교 에러로 죽을 수 있어, 비교 전에 전부 문자열로 맞춘다.
                    unmapped = sorted(
                        ga_channel_df.loc[ga_channel_df["channel"].isna(), "source_medium"]
                        .dropna().astype(str).unique()
                    )
                    st.caption(f"※ 매핑표에 없는 소스/매체 {len(unmapped)}종은 미매핑 상태입니다: {', '.join(unmapped[:10])}{' 외' if len(unmapped) > 10 else ''}")
            status2.update(label="분석 완료", state="complete")

        if st.sidebar.button("💾 GA 유입 데이터 저장하기", type="primary", key="ga_combined_save_btn"):
            saved_msgs = []
            if not inflow_df.empty:
                n_inflow = save_table("inflow_revenue_daily", inflow_df, "report_date", ga_file.name)
                saved_msgs.append(f"유입·매출 비교 {n_inflow}일")
            if not utm_map_df.empty:
                n_utm = save_table("utm_channel_map", utm_map_df, "source_medium", ga_file.name)
                saved_msgs.append(f"UTM 매핑 {n_utm}행")
            if not ga_channel_df.empty:
                n_ga_channel = save_table(
                    "ga_channel_inflow", ga_channel_df, "report_date,source_medium", ga_file.name
                )
                saved_msgs.append(f"GA 매체별 유입 경로 {n_ga_channel}행")
            st.cache_data.clear()
            st.sidebar.success("저장 완료! " + (" · ".join(saved_msgs) if saved_msgs else "저장할 데이터가 없습니다."))
            st.rerun()

    st.sidebar.markdown("---")
    budget_file = st.sidebar.file_uploader(
        "③ 연간 예산 파일 업로드 (◆26년 월별 예산 정리 xlsx)",
        type=["xlsx", "xls"], key="budget_uploader",
    )
    if budget_file is not None:
        with st.sidebar.status("파일 분석 중...", expanded=True) as status3:
            budget_xls = pd.ExcelFile(budget_file)
            budget_df = parse_channel_budget_sheet(budget_xls)
            if budget_df.empty:
                st.warning(
                    "인식 가능한 예산 데이터를 찾지 못했습니다. '자사몰 현황' 섹션, '26년' 블록, "
                    "'매체 세부내역'이 있는 파일인지 확인해주세요. (1차 버전이라 원본 레이아웃과 "
                    "다르면 못 읽을 수 있습니다 — 이 경우 형에게 알려주세요.)"
                )
            else:
                total_n = budget_df[budget_df["channel"] == "TOTAL"]["budget_cost"].sum()
                st.write(
                    f"💰 자사몰 예산 인식: {len(budget_df)}행, "
                    f"매체 {budget_df.loc[budget_df['channel'] != 'TOTAL', 'channel'].nunique()}종, "
                    f"26년 TOTAL 연간 예산 {total_n:,.0f}원"
                )
            status3.update(label="분석 완료", state="complete")

        if st.sidebar.button("💾 예산 데이터 저장하기", type="primary", key="budget_save_btn"):
            if budget_df.empty:
                st.sidebar.warning("저장할 예산 데이터가 없습니다.")
            else:
                n_budget = save_table("channel_budget", budget_df, "scope,channel,year,month", budget_file.name)
                st.cache_data.clear()
                st.sidebar.success(f"저장 완료! 예산 {n_budget}행")
                st.rerun()

    st.sidebar.markdown("---")
    wk = load_table("weekly_overview")
    st.sidebar.metric("누적 주간 데이터", f"{len(wk):,} 주")
    if st.sidebar.button("🔄 새로고침 (캐시 비우기)"):
        st.cache_data.clear()
        st.rerun()


# ──────────────────────────────────────────────────────────────
# 예산 현황 (신규, 1차 버전) — '◆26년 월별 예산 정리' 파일 기반, 자사몰만
# ──────────────────────────────────────────────────────────────
def render_budget_page(monthly: pd.DataFrame, budget: pd.DataFrame):
    st.subheader("💰 예산 현황 (자사몰 · 1차 버전)")
    st.caption(
        "'◆26년 월별 예산 정리' 파일의 '자사몰 현황' 섹션(26년, 매체 세부내역) 기준입니다. "
        "형이 캡쳐해준 화면 구조를 보고 만든 1차 버전이라, 실제 파일로 아직 검증 전입니다 — "
        "숫자가 이상해 보이면 알려주세요. 매체별 예산의 매체명이 리포트 매체명과 아직 자동으로 "
        "안 맞춰져 있어서, 우선 자사몰 전체 목표 대비 실제 집행만 비교합니다."
    )
    if budget is None or budget.empty:
        st.info("아직 예산 데이터가 없습니다. 왼쪽 사이드바 '③ 연간 예산 파일 업로드'에서 파일을 올려주세요.")
        return

    b = budget[budget["scope"] == "자사몰"].copy()
    total_b = b[b["channel"] == "TOTAL"].sort_values("month")
    if total_b.empty:
        st.info("자사몰 TOTAL(광고선전비 1월~12월) 행을 못 찾았습니다. 파일 구조를 다시 확인해주세요.")
        return

    year = int(total_b["year"].iloc[0])
    actual_by_month = pd.Series(dtype=float)
    if monthly is not None and not monthly.empty:
        m = monthly.copy()
        m["report_month"] = pd.to_datetime(m["report_month"])
        m_year = m[m["report_month"].dt.year == year]
        if not m_year.empty:
            actual_by_month = m_year.groupby(m_year["report_month"].dt.month)["cost_incl_vat"].sum()

    rows = []
    for _, r in total_b.iterrows():
        mo = int(r["month"])
        budget_v = float(r["budget_cost"])
        actual_v = float(actual_by_month.get(mo, 0))
        rows.append({
            "월": f"{mo}월", "월번호": mo, "목표 예산": budget_v, "실제 집행": actual_v,
            "집행률": (actual_v / budget_v * 100) if budget_v else 0,
            "차이": actual_v - budget_v,
        })
    show = pd.DataFrame(rows).sort_values("월번호")

    st.markdown(f"##### {year}년 자사몰 월별 목표 예산 vs 실제 집행")
    fig = px.bar(
        show, x="월", y=["목표 예산", "실제 집행"], barmode="group",
        category_orders={"월": show["월"].tolist()}, labels={"value": "금액(원)", "variable": "구분"},
    )
    fig.update_yaxes(tickformat=",.0f")
    fig.for_each_trace(
        lambda t: t.update(hovertemplate=f"구분={t.name}<br>월=%{{x}}<br>금액=%{{y:,.0f}}원<extra></extra>")
    )
    st.plotly_chart(theme_chart(fig), use_container_width=True)

    show_disp = show[["월", "목표 예산", "실제 집행", "차이", "집행률"]].copy()
    for c in ["목표 예산", "실제 집행", "차이"]:
        show_disp[c] = show[c].map(lambda v: f"{v:,.0f}")
    show_disp["집행률"] = show["집행률"].map(lambda v: f"{v:.0f}%" if v else "-")
    render_html_table(show_disp)

    cur = show[show["실제 집행"] > 0].sort_values("월번호")
    if not cur.empty:
        last = cur.iloc[-1]
        if last["집행률"] >= 110:
            status = "초과 집행"
        elif last["집행률"] >= 80:
            status = "목표 근접"
        else:
            status = "저집행"
        st.markdown(
            f"**{last['월']} 기준 목표 예산 {last['목표 예산']:,.0f}원 대비 실제 {last['실제 집행']:,.0f}원 "
            f"집행({last['집행률']:.0f}%)로 {status}입니다.**"
        )

    detail = b[b["channel"] != "TOTAL"]
    if not detail.empty:
        st.markdown("---")
        st.markdown("##### 매체별 예산 (원본 참고용 — 리포트 매체명과 자동 매칭 전)")
        pivot = detail.pivot_table(index="channel", columns="month", values="budget_cost", aggfunc="sum")
        pivot = pivot.reindex(sorted(pivot.columns), axis=1)
        pivot.columns = [f"{c}월" for c in pivot.columns]
        pivot = pivot.reset_index().rename(columns={"channel": "매체명"})
        pivot_disp = pivot.copy()
        for c in pivot.columns[1:]:
            pivot_disp[c] = pivot[c].map(lambda v: f"{v:,.0f}" if pd.notna(v) else "-")
        render_html_table(pivot_disp)


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


def _render_creative_table(fc: pd.DataFrame, channel_name: str = None):
    """선택된 탭(매체)의 필터링된 소재 데이터로 집계 테이블 + 판정 + 다운로드 버튼을 렌더링.
    channel_name이 주어지면(=TOTAL 탭이 아니면) 표 위에 그 매체의 우수/부진 소재를 짚어주는
    코멘트를 한 줄 붙인다 — TOTAL 탭엔 매체가 섞여 있어 우수/부진 비교가 의미가 없어 생략한다."""
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

    if channel_name:
        judged = agg[agg["판정"] != "판단 보류(표본 부족)"]
        n_good = int((judged["판정"] == "우수").sum())
        n_mid = int((judged["판정"] == "평균 수준").sum())
        n_bad = int((judged["판정"] == "부진").sum())
        n_hold = len(agg) - len(judged)
        if len(judged) >= 2:
            best = judged.loc[judged["roas"].idxmax()]
            worst = judged.loc[judged["roas"].idxmin()]
            lines = [
                f"{channel_name} 소재 {len(judged)}개 중 우수 {n_good}개 · 평균 수준 {n_mid}개 · 부진 {n_bad}개입니다"
                + (f" (표본 부족 {n_hold}개는 판단 보류)." if n_hold else "."),
                f"<b>{best['creative']}</b>가 광고비 {best['cost_incl_vat']:,.0f}원으로 ROAS {best['roas']:,.0f}%를 기록해 가장 우수했고, "
                f"<b>{worst['creative']}</b>는 광고비 {worst['cost_incl_vat']:,.0f}원 대비 ROAS {worst['roas']:,.0f}%로 "
                f"계정 평균({account_avg_roas:,.0f}%) 대비 크게 낮아 가장 부진했습니다.",
            ]
            if n_bad:
                lines.append(_ops_next_action(f"부진 소재({n_bad}개)는 소재 교체 또는 예산 축소를 검토하는 것을 권장합니다."))
            st.markdown("<br>".join(lines), unsafe_allow_html=True)
        elif len(judged) == 1:
            st.markdown(
                f"{channel_name}은 판단 가능한(표본 충분) 소재가 <b>{judged.iloc[0]['creative']}</b> 1개뿐이라, "
                "우수/부진 비교는 소재가 더 쌓이면 확인하겠습니다.",
                unsafe_allow_html=True,
            )
        else:
            st.caption(f"{channel_name}은 아직 판단 가능한(표본 충분) 소재가 없습니다.")

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
            _render_creative_table(tab_fc, channel_name=None if tab_name == "TOTAL" else tab_name)


# ──────────────────────────────────────────────────────────────
# 사이드바 그룹 네비게이션 (신규 — st.tabs() 대체)
# ──────────────────────────────────────────────────────────────
NAV_GROUPS = {
    "성과 리포트": ["종합 대시보드", "매체별 성과", "타겟팅별 성과", "소재별 성과", "예산 현황"],
    "운영 코멘트": ["운영 코멘트"],
    "GA 유입 리포트": ["GA 매체별 유입 경로", "GA4 라이브 리포트", "유입·매출 비교"],
    "운영 도구": ["UTM 빌더", "소재 로그", "예산 재배분", "마일스톤"],
    "가이드": ["가이드"],
}

# 그룹별로 다른 CSS 클래스(st-key-navgrp_xxx)를 붙이기 위한 영문 키. GA 유입 리포트/운영 코멘트는
# 별도 아이콘을 안 받아서 키를 안 주고(None) 공통 기본 아이콘(NAV_GROUP_ICON_B64)을 그대로 쓴다.
NAV_GROUP_KEYS = {
    "성과 리포트": "report",
    "운영 코멘트": None,
    "GA 유입 리포트": None,
    "운영 도구": "ops",
    "가이드": "guide",
}

# 아직 실제 데이터/로직이 없는 페이지들 — main()의 페이지 분기에서 이 목록에 있으면
# render_coming_soon()으로 "준비 중" 안내만 보여준다. 나중에 진짜 렌더 함수가 생기면
# main()에 elif 분기를 추가하고 여기서 이름을 지워주면 된다.
NAV_PAGES_COMING_SOON = {"마일스톤", "UTM 빌더", "소재 로그", "예산 재배분", "가이드"}


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
            # 그룹별로 다른 아이콘을 쓰려면 CSS에서 그룹을 구분할 수 있어야 하는데, Streamlit은
            # 위젯마다 개별 element-container로 감싸버려서 :nth-of-type으로는 구분이 안 된다
            # (모든 expander가 "자기 부모 안에서는 1번째"라 전부 같은 규칙에 걸림). 그래서 그룹마다
            # container(key=...)로 한 번 더 감싸 고유 CSS 클래스(st-key-{key})를 붙여준다.
            group_key = NAV_GROUP_KEYS.get(group)
            try:
                group_box = st.container(key=f"navgrp_{group_key}") if group_key else st.container()
            except TypeError:
                group_box = st.container()
            with group_box, st.expander(group, expanded=is_active_group):
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
def render_overview_page(
    weekly: pd.DataFrame, monthly: pd.DataFrame, daily: pd.DataFrame,
    channels: pd.DataFrame = None, channels_weekly: pd.DataFrame = None,
    ga_channel_inflow: pd.DataFrame = None,
):
    weekly = _drop_trailing_zero_weeks(weekly)
    if weekly is None:
        weekly = pd.DataFrame()
    if not weekly.empty:
        st.subheader("🔎 기간 필터 (주간 기준)")
        min_d, max_d = weekly["week_start"].min(), weekly["week_end"].max()
        start, end = period_filter(min_d, max_d, key="weekly", options=DATE_PERIOD_OPTIONS_FULL)
        fw = weekly[(weekly["week_start"] >= start) & (weekly["week_start"] <= end)]
        fw = add_kpis(fw).sort_values("week_start")

        kpi_cards(fw)
        st.markdown("### 주간 추이")
        if fw.empty:
            st.caption("선택한 기간(주간 기준)에 해당하는 데이터가 없습니다.")
        else:
            week_label_order = kor_date_labels(fw["week_start"], "day")
            fw = fw.assign(주=week_label_order)
            c1, c2 = st.columns(2)
            with c1:
                chart_df = fw.rename(columns={"cost_incl_vat": "광고비(VAT포함)", "revenue": "매출"})
                fig = px.bar(
                    chart_df, x="주", y=["광고비(VAT포함)", "매출"], barmode="group",
                    title="주간 비용(VAT포함) vs 매출",
                    labels={"value": "금액(원)", "variable": "구분"},
                    category_orders={"주": week_label_order},
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
                    fw, x="주", y="roas", markers=True, title="주간 ROAS 추이 (%)",
                    labels={"roas": "ROAS(%)"},
                    category_orders={"주": week_label_order},
                )
                fig2.update_yaxes(tickformat=",.0f", ticksuffix="%")
                fig2.update_traces(hovertemplate="주=%{x}<br>ROAS=%{y:,.0f}%<extra></extra>")
                st.plotly_chart(theme_chart(fig2), use_container_width=True)

        # 플레이스홀더로 미리 만들어둔 빈 월(광고비 0원)까지 그리면 옛날~올해가 다 이어진 밋밋한
        # 0선이 껴서 정작 보고 싶은 최근 구간이 눌려 보인다 — 실제로 집행된(광고비>0) 달만 그린다.
        monthly_real = (
            monthly[monthly["cost_incl_vat"] > 0]
            if not monthly.empty and "cost_incl_vat" in monthly.columns
            else monthly
        )
        if not monthly_real.empty:
            st.markdown("---")
            st.markdown("### 월별 GA-ROAS vs 플랫폼 ROAS")
            fm_chart = add_kpis(monthly_real).sort_values("report_month").rename(
                columns={"roas": "플랫폼 ROAS", "ga_roas": "GA ROAS"}
            )
            month_label_order = kor_date_labels(fm_chart["report_month"], "month")
            fm_chart = fm_chart.assign(월=month_label_order)
            fig3 = px.line(
                fm_chart, x="월", y=["플랫폼 ROAS", "GA ROAS"], markers=True,
                labels={"value": "ROAS(%)", "variable": "기준"},
                title="플랫폼 리포팅 ROAS vs GA 기준 ROAS",
                category_orders={"월": month_label_order},
            )
            fig3.update_yaxes(tickformat=",.0f", ticksuffix="%")
            fig3.for_each_trace(
                lambda t: t.update(hovertemplate=f"기준={t.name}<br>월=%{{x}}<br>ROAS=%{{y:,.0f}}%<extra></extra>")
            )
            st.plotly_chart(theme_chart(fig3), use_container_width=True)
            st.caption("* 광고비가 집행된 달만 표시합니다 (미집행 빈 달 제외). GA-매출/GA-ROAS는 쇼핑검색 및 GFA 외부몰 데이터가 미집계될 수 있습니다 (원본 시트 주석 기준).")

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
        render_ops_comment_monthly(monthly, channels=channels)

        wk = weekly.copy()
        if not wk.empty:
            wk["week_no"] = wk["label"].astype(str).str.replace(r"\s*\(.*\)\s*$", "", regex=True).str.strip()
            wk["week_range"] = wk.apply(lambda r: f"{r['week_start']:%Y-%m-%d}~{r['week_end']:%Y-%m-%d}", axis=1)
            # 리포트 원본 주간 섹션엔 GA 컬럼이 없어서(월별 섹션에만 있음), GA 매체별 유입 경로
            # (UTM 매핑 완료분)를 주 단위로 직접 합산해 붙인다 — 업로드가 안 됐으면 조용히 생략.
            ga_wk = _ga_weekly_agg(ga_channel_inflow, wk)
            if not ga_wk.empty:
                wk = wk.merge(ga_wk, on="week_start", how="left")
                wk["ga_conversions"] = wk["ga_conversions"].fillna(0)
                wk["ga_revenue"] = wk["ga_revenue"].fillna(0)
                wk["ga_roas"] = np.where(wk["cost_incl_vat"] > 0, wk["ga_revenue"] / wk["cost_incl_vat"] * 100, 0)
        week_show_cols = ["week_range", "week_no", "impressions", "clicks", "ctr", "cpc",
                           "cost_excl_vat", "cost_incl_vat", "signups", "cpa", "conversions", "cvr", "revenue", "roas", "aov"]
        if not wk.empty and "ga_conversions" in wk.columns:
            week_show_cols += ["ga_conversions", "ga_revenue", "ga_roas"]
        week_numeric_cols = [c for c in week_show_cols if c not in ("week_no", "week_range")]
        render_cumulative_table(
            add_kpis(wk) if not wk.empty else wk,
            date_col="week_start", show_cols=week_show_cols, numeric_cols=week_numeric_cols,
            title="2) 주간별 누적", key="weekly_cum", mode="week",
        )
        if ga_channel_inflow is None or ga_channel_inflow.empty:
            st.caption("※ 'GA 매체별 유입 경로' 데이터를 업로드하면 GA-전환수/GA-매출/GA-ROAS 컬럼이 여기에도 채워집니다.")
        render_ops_comment_weekly(weekly, channels_weekly=channels_weekly)

        dy = daily.copy()
        if not dy.empty:
            ga_dy = _ga_daily_agg(ga_channel_inflow)
            if not ga_dy.empty:
                dy = dy.merge(ga_dy, on="report_date", how="left")
                dy["ga_conversions"] = dy["ga_conversions"].fillna(0)
                dy["ga_revenue"] = dy["ga_revenue"].fillna(0)
                dy["ga_roas"] = np.where(dy["cost_incl_vat"] > 0, dy["ga_revenue"] / dy["cost_incl_vat"] * 100, 0)
        day_show_cols = ["report_date", "impressions", "clicks", "ctr", "cpc", "cost_excl_vat", "cost_incl_vat",
                          "signups", "cpa", "conversions", "cvr", "revenue", "roas", "aov"]
        if not dy.empty and "ga_conversions" in dy.columns:
            day_show_cols += ["ga_conversions", "ga_revenue", "ga_roas"]
        day_numeric_cols = [c for c in day_show_cols if c != "report_date"]
        render_cumulative_table(
            add_kpis(dy) if not dy.empty else dy,
            date_col="report_date", show_cols=day_show_cols, numeric_cols=day_numeric_cols,
            title="3) 일자별 누적", key="daily_cum", mode="day",
        )
    else:
        st.info("주간 데이터가 아직 없습니다.")


def render_channel_page(channels: pd.DataFrame, snapshot: pd.DataFrame, ga_channel_inflow: pd.DataFrame = None):
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
        # 리포트 원본 매체별 GA 컬럼은 '당월 스냅샷' 기준이라 아래 별도 표로만 보여주고, 여기
        # 메인 표에는 UTM 매핑된 GA 매체별 유입 경로 데이터를 선택 기간(mstart~mend) 그대로
        # 합산해서 붙인다 — 기간을 바꾸면 이 GA 컬럼도 같이 갱신된다.
        ga_by_channel = _ga_channel_agg(ga_channel_inflow, mstart, mend)
        if not ga_by_channel.empty:
            by_channel = by_channel.merge(ga_by_channel, on="channel", how="left")
            by_channel["ga_conversions"] = by_channel["ga_conversions"].fillna(0)
            by_channel["ga_revenue"] = by_channel["ga_revenue"].fillna(0)
        by_channel = add_kpis(by_channel).sort_values("cost_incl_vat", ascending=False)
        if "ga_revenue" in by_channel.columns:
            by_channel["ga_roas"] = np.where(
                by_channel["cost_incl_vat"] > 0, by_channel["ga_revenue"] / by_channel["cost_incl_vat"] * 100, 0
            )

        by_channel_chart = by_channel.copy()
        by_channel_chart["roas_label"] = by_channel_chart["roas"].map(lambda v: f"{v:.0f}%")
        fig = px.bar(
            by_channel_chart, x="channel", y="roas", title="매체별 ROAS (%, 선택 기간 합산)", text="roas_label",
            labels={"channel": "매체", "roas": "ROAS(%)"},
        )
        fig.update_yaxes(tickformat=",.0f", ticksuffix="%")
        fig.update_traces(hovertemplate="매체=%{x}<br>ROAS=%{y:,.0f}%<extra></extra>")
        st.plotly_chart(theme_chart(fig), use_container_width=True)

        # 광고비(VAT제외)는 CPA 계산에만 쓰고 화면/엑셀에는 광고비(VAT+)만 노출 — 컬럼이 많아
        # 표 폭이 들쭉날쭉해지는 것도 줄어든다. 컬럼 순서는 종합 대시보드/타겟팅별 성과와 동일하게
        # 노출수 → 클릭수 → CTR → CPC → 광고비 순으로 맞춘다.
        BC_COL_ORDER = ["channel", "impressions", "clicks", "ctr", "cpc", "cost_incl_vat",
                        "cpa", "conversions", "cvr", "revenue", "roas", "aov",
                        "ga_conversions", "ga_revenue", "ga_roas"]
        bc_cols = [c for c in BC_COL_ORDER if c in by_channel.columns]
        bc_cols += [c for c in by_channel.columns if c not in bc_cols and c != "cost_excl_vat"]
        bc_table = format_display(by_channel[bc_cols])
        if "ga_revenue" not in by_channel.columns:
            st.caption("※ 'GA 매체별 유입 경로' 데이터를 업로드하면(UTM 매핑 필요) 매체별 GA-전환수/GA-매출/GA-ROAS도 여기 표시됩니다.")
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
        render_ops_comment_channel_narrative(
            by_channel, "roas",
            footnote="※ 선택 기간 합산·자체 ROAS 기준입니다. GA-ROAS 비교는 아래 '당월 매체별 GA 비교' 참고.",
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


def _ops_targeting_summary_comment(audience_summary: pd.DataFrame, total_row: pd.Series, start: date, end: date) -> str:
    """타겟팅별 성과 ① TOTAL 비교 표 아래에 쓰는 코멘트. 신규/리타겟팅 어느 쪽이 매출·ROAS를
    견인했는지 짚어준다. 표가 이미 선택된 기간(period_filter)을 기준으로 집계돼 있어, 코멘트도
    같은 기간을 그대로 따른다 — '지난주'로 보고 싶으면 위 기간 선택에서 '지난주'를 고르면 된다."""
    status = _ops_kpi_status(total_row["roas"])
    bits = [
        f"**{start:%Y-%m-%d}~{end:%Y-%m-%d} 기준 전체 ROAS {total_row['roas']:,.0f}%로 {status}**입니다 "
        f"(KPI {OPS_KPI_ROAS_LOW}~{OPS_KPI_ROAS_HIGH}%)."
    ]
    new_r = audience_summary[audience_summary["channel"] == "신규 타겟팅"]
    re_r = audience_summary[audience_summary["channel"] == "리타겟팅"]
    if not new_r.empty and not re_r.empty and total_row["revenue"]:
        n, r = new_r.iloc[0], re_r.iloc[0]
        if r["roas"] >= n["roas"]:
            lead_label, lead, other_label, other = "리타겟팅", r, "신규 타겟팅", n
        else:
            lead_label, lead, other_label, other = "신규 타겟팅", n, "리타겟팅", r
        rev_share = lead["revenue"] / total_row["revenue"] * 100
        ratio_txt = f"{(lead['roas'] / other['roas']):.1f}배 높습니다." if other["roas"] else "훨씬 높습니다."
        bits.append(
            f"{lead_label}이 매출의 {rev_share:,.0f}%를 차지하며 계정 성과를 견인했고, "
            f"ROAS도 {lead['roas']:,.0f}%로 {other_label}({other['roas']:,.0f}%) 대비 {ratio_txt}"
        )
    return " ".join(bits)


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
    total_row_raw = _targeting_total_row(audience_summary, "TOTAL").iloc[0]
    st.markdown(_ops_targeting_summary_comment(audience_summary, total_row_raw, start, end), unsafe_allow_html=True)

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
            fig_new.update_yaxes(tickformat=",.0f", ticksuffix="%")
            fig_new.update_traces(hovertemplate="매체=%{x}<br>ROAS=%{y:,.0f}%<extra></extra>")
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
            fig_re.update_yaxes(tickformat=",.0f", ticksuffix="%")
            fig_re.update_traces(hovertemplate="매체=%{x}<br>ROAS=%{y:,.0f}%<extra></extra>")
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


def render_inflow_revenue_page(df: pd.DataFrame, ga_channel_inflow: pd.DataFrame = None):
    """유입·매출 비교 — 형이 별도로 정리해서 준 '일별 GA·어드민 지표 비교' 파일 기반.
    채널별이 아니라 사이트 전체 일별 합산 기준. ① 방문자 추이(GA 총/신규) ② 매출 비교
    (어드민=회사 내부 기준 vs GA 기준, 어드민 대비 격차%)로 구성.
    '매체-매출'(광고 리포트 기준)은 이 비교에서 뺐다 — 어드민 매출은 사이트 전체 매출이고
    GA 매출도 자연유입 포함 전체인데, 매체-매출은 광고 채널 것만 걷은 거라 셋을 나란히 비교하면
    서로 다른 걸 비교하는 셈이 된다(형 확인). GA-매출은 원본 시트에 채워져 있으면 그 값을 쓰고,
    비어있으면(0/NaN) 'GA 매체별 유입 경로'(UTM 매핑 전 원본 전체 합산)로 채운다."""
    if df.empty:
        st.info(
            "아직 데이터가 없습니다. 왼쪽 사이드바 '② GA 유입 데이터 업로드'에서 "
            "'일별 GA,어드민 지표 비교' 시트가 있는 파일을 올려주세요."
        )
        return

    df = df.copy()
    df["report_date"] = pd.to_datetime(df["report_date"]).dt.date
    ga_all = _ga_daily_agg_all(ga_channel_inflow)
    if not ga_all.empty:
        # 원본 시트 값이 전부 0/NaN이면(int64) 실수 GA 매출을 대입할 때 타입 경고가 나므로,
        # 대입 전에 float으로 미리 맞춰둔다.
        df["ga_revenue"] = pd.to_numeric(df["ga_revenue"], errors="coerce").astype(float)
        df = df.merge(ga_all, on="report_date", how="left")
        needs_fill = df["ga_revenue"].isna() | (df["ga_revenue"] == 0)
        df.loc[needs_fill, "ga_revenue"] = df.loc[needs_fill, "ga_revenue_all"]
        df = df.drop(columns=["ga_revenue_all", "ga_conversions_all"], errors="ignore")
        # GA-ROAS(=GA매출/광고비)는 여기서 안 쓴다 — GA-매출은 자연유입 포함 사이트 전체 매출인데
        # 광고비는 매체(광고)에만 든 비용이라, 둘을 나누면 범위가 안 맞는 숫자가 나온다(형 확인).
    st.subheader("🔎 기간 필터")
    min_d, max_d = df["report_date"].min(), df["report_date"].max()
    start, end = period_filter(min_d, max_d, key="inflow", default_preset="이번달")
    fd = df[(df["report_date"] >= start) & (df["report_date"] <= end)].sort_values("report_date")

    if fd.empty:
        st.info("선택한 기간에 데이터가 없습니다.")
        return

    # ── ① 방문자 추이 (GA 기준) ──
    st.markdown("##### ① 방문자 추이 (GA 기준)")
    users_sum = fd["users"].sum()
    new_users_sum = fd["new_users"].sum()
    returning_users_sum = fd["returning_users"].sum()
    new_ratio = (new_users_sum / users_sum * 100) if users_sum else 0
    returning_ratio = (returning_users_sum / users_sum * 100) if users_sum else 0
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("총 방문자 합계", f"{users_sum:,.0f} 명")
    c2.metric("신규 방문자 합계", f"{new_users_sum:,.0f} 명")
    c3.metric("재방문자 합계", f"{returning_users_sum:,.0f} 명")
    c4.metric("신규 방문자 비중", f"{new_ratio:.1f} %")
    c5.metric("재방문자 비중", f"{returning_ratio:.1f} %")

    day_label_order = kor_date_labels(fd["report_date"], "day")
    fd = fd.assign(일자=day_label_order)
    visit_chart_df = fd.melt(
        id_vars=["일자"], value_vars=["users", "new_users"],
        var_name="구분", value_name="방문자수",
    )
    visit_chart_df["구분"] = visit_chart_df["구분"].map({"users": "총 방문자", "new_users": "신규 방문자"})
    fig_visit = px.line(
        visit_chart_df, x="일자", y="방문자수", color="구분", markers=True,
        category_orders={"일자": day_label_order},
    )
    fig_visit.update_yaxes(tickformat=",.0f")
    fig_visit.for_each_trace(
        lambda t: t.update(hovertemplate=f"구분={t.name}<br>일자=%{{x}}<br>방문자수=%{{y:,.0f}}명<extra></extra>")
    )
    st.plotly_chart(theme_chart(fig_visit), use_container_width=True)

    # ── ② 매출 비교 (어드민 vs GA) ──
    # '보고서(매체 리포트) 기준 매출'은 광고 채널에서 걷힌 매출만이라, 사이트 전체 매출인
    # 어드민·GA와 나란히 비교하면 서로 다른 범위를 비교하는 셈이라 이 표에서는 뺐다(형 확인).
    # 매체 리포트 기준 매출·ROAS는 '종합 대시보드'/'매체별 성과'에서 따로 확인할 수 있다.
    st.markdown("##### ② 매출 비교")
    st.caption("어드민 = 회사 내부(카페24 등 백엔드) 기준, 사이트 전체 매출 · GA = GA가 집계한 사이트 전체 매출(자연유입·direct 포함)")
    admin_sum = fd["admin_revenue"].sum()
    ga_sum = fd["ga_revenue"].sum()
    cost_sum = fd["cost_incl_vat"].sum()
    gap_ga = ((ga_sum - admin_sum) / admin_sum * 100) if admin_sum else 0

    r1, r3 = st.columns(2)
    r1.metric("어드민 매출(회사 내부 기준)", f"{admin_sum:,.0f} 원")
    r3.metric("GA 기준 매출", f"{ga_sum:,.0f} 원", f"{gap_ga:+.1f}% vs 어드민")
    st.caption("격차(%)는 어드민(회사 내부 기준) 매출 대비 초과/미달 비율입니다.")

    revenue_chart_df = fd.melt(
        id_vars=["일자"], value_vars=["admin_revenue", "ga_revenue"],
        var_name="구분", value_name="매출",
    )
    revenue_chart_df["구분"] = revenue_chart_df["구분"].map({
        "admin_revenue": "어드민(회사 내부)", "ga_revenue": "GA 기준",
    })
    fig_rev = px.line(
        revenue_chart_df, x="일자", y="매출", color="구분", markers=True,
        category_orders={"일자": day_label_order},
    )
    fig_rev.update_yaxes(tickformat=",.0f")
    fig_rev.for_each_trace(
        lambda t: t.update(hovertemplate=f"구분={t.name}<br>일자=%{{x}}<br>매출=%{{y:,.0f}}원<extra></extra>")
    )
    st.plotly_chart(theme_chart(fig_rev), use_container_width=True)

    # ── 일자별 상세 표 (정렬 가능) — 형이 정리한 엑셀 양식(사용자→신규방문자→재방문자→이탈률→
    # 페이지뷰→평균체류시간→어드민/매체/GA 매출→매체/GA-ROAS→어드민 대비 비교) 그대로 맞춘다.
    # "매출"/"ROAS"/"어드민 매출" 같은 이름은 다른 페이지에서도 공용으로 쓰는 라벨이라 KOR_COLS를
    # 바꾸면 거기까지 다 바뀌어버리므로, 이 표에서만 쓸 라벨을 따로 둔다.
    # GA-ROAS(=GA매출/광고비)는 이 표에서 뺐다 — GA-매출은 자연유입 포함 사이트 전체 매출인데
    # 광고비는 매체(광고)에만 쓴 비용이라 나누면 범위가 안 맞는 숫자가 나온다(형 확인). 매체 기준
    # ROAS는 종합 대시보드·매체별 성과에서 보는 걸로 하고, 여기는 어드민 vs GA 매출 비교만 한다.
    INFLOW_DETAIL_LABELS = {
        "report_date": "일자",
        "users": "총 방문자",
        "new_users": "신규방문자",
        "returning_users": "재방문자",
        "bounce_rate": "이탈률",
        "pageviews": "페이지뷰",
        "avg_session_duration": "평균 체류시간",
        "admin_revenue": "어드민-매출",
        "ga_revenue": "GA-매출",
        "ga_gap_pct": "어드민-GA 매출 비교",
    }
    INFLOW_DETAIL_LABELS_REV = {v: k for k, v in INFLOW_DETAIL_LABELS.items()}

    detail_cols = [
        "report_date", "users", "new_users", "returning_users", "bounce_rate", "pageviews",
        "avg_session_duration", "admin_revenue", "ga_revenue",
    ]
    detail = fd[detail_cols].copy()
    detail["ga_gap_pct"] = np.where(
        detail["admin_revenue"] > 0, (detail["ga_revenue"] - detail["admin_revenue"]) / detail["admin_revenue"] * 100, 0
    )
    final_cols = detail_cols + ["ga_gap_pct"]

    def _fmt_duration(v):
        # 초 단위 숫자를 "3분 26초" 형태로 되돌려 보여준다 (원본 엑셀 표기와 맞춤).
        if pd.isna(v):
            return "-"
        v = int(round(v))
        return f"{v // 60}분 {v % 60}초"

    show = format_display(detail[detail_cols])
    show["avg_session_duration"] = detail["avg_session_duration"].map(_fmt_duration)

    def _fmt_gap(v):
        # 매체 리포트/GA 데이터가 아직 안 올라온 최근 며칠은 NaN이라 "-"로 표시 (0%로 오해하지 않도록).
        if pd.isna(v):
            return "-"
        return f"{'▲' if v >= 0 else '▼'}{v:+.1f}%"

    show["ga_gap_pct"] = detail["ga_gap_pct"].map(_fmt_gap)

    returning_users_sum = fd["returning_users"].sum()  # (위 ①에서도 계산하지만 이 표 TOTAL 행에도 필요)
    total_row = {
        "report_date": "TOTAL",
        "users": f"{users_sum:,.0f}",
        "new_users": f"{new_users_sum:,.0f}",
        "returning_users": f"{returning_users_sum:,.0f}",
        "bounce_rate": f"{fd['bounce_rate'].mean():.2f}%" if fd["bounce_rate"].notna().any() else "-",
        "pageviews": f"{fd['pageviews'].sum():,.0f}",
        "avg_session_duration": _fmt_duration(fd["avg_session_duration"].mean()),
        "admin_revenue": f"{admin_sum:,.0f}",
        "ga_revenue": f"{ga_sum:,.0f}",
        "ga_gap_pct": f"{'▲' if gap_ga >= 0 else '▼'}{gap_ga:+.1f}%",
    }
    table = pd.concat([pd.DataFrame([total_row])[final_cols], show[final_cols]], ignore_index=True)
    table = table.rename(columns=INFLOW_DETAIL_LABELS)
    raw_numeric_cols = [
        "users", "new_users", "returning_users", "bounce_rate", "pageviews", "avg_session_duration",
        "admin_revenue", "ga_revenue",
    ]
    render_html_table(table, raw=detail[raw_numeric_cols], raw_label_map=INFLOW_DETAIL_LABELS_REV)

    # 이 페이지는 ROAS(성과) 판단용이 아니라 '어드민 장부 매출 vs GA 추적 매출'이 서로 맞게
    # 잡히고 있는지 보는 데이터 정합성 체크용이라, KPI 상태 표시 대신 격차 크기만 짚어준다.
    st.markdown("#### 💬 코멘트")
    comment_body = [
        f"총 방문자 {users_sum:,.0f}명(신규 비중 {new_ratio:.1f}%), 어드민 매출 대비 GA 매출 격차는 {_ops_fmt_pct(gap_ga)}입니다."
    ]
    st.markdown(" ".join(comment_body), unsafe_allow_html=True)
    if abs(gap_ga) >= 20:
        st.markdown(
            _ops_next_action("격차가 큰 편이니 GA 전자상거래 추적 설정, 어드민 매출 집계 범위(취소·환불 반영 여부)를 확인해보는 것을 권장합니다."),
            unsafe_allow_html=True,
        )

    dl_df = format_display(detail[detail_cols])
    dl_df["avg_session_duration"] = detail["avg_session_duration"].map(_fmt_duration)
    st.download_button(
        "⬇️ 엑셀 다운로드 (유입·매출 비교)",
        data=to_excel_bytes(dl_df.rename(columns=INFLOW_DETAIL_LABELS)),
        file_name="inflow_revenue_daily.xlsx",
    )


GA_CHANNEL_LABELS = {
    "source_medium": "세션 소스/매체",
    "channel": "매체",
    "users": "총 방문자",
    "new_users": "신규 방문자",
    "returning_users": "재 방문자",
    "bounce_rate": "이탈률",
    "pageviews": "조회수",
    "avg_session_duration": "평균 체류 시간",
    "conversions": "구매",
    "revenue": "매출",
}
GA_CHANNEL_LABELS_REV = {v: k for k, v in GA_CHANNEL_LABELS.items()}
GA_CHANNEL_DETAIL_COLS = [
    "source_medium", "channel", "users", "new_users", "returning_users",
    "bounce_rate", "pageviews", "avg_session_duration", "conversions", "revenue",
]
# ③ TOP 10 탭에서 쓰는 (컬럼명, 탭 라벨) 순서 — 형이 요청한 순서(총방문자→신규→재방문→구매→매출) 그대로.
GA_CHANNEL_TOP_METRICS = [
    ("users", "총 방문자"),
    ("new_users", "신규 방문자"),
    ("returning_users", "재방문자"),
    ("conversions", "구매"),
    ("revenue", "매출"),
]


def _fmt_duration_padded(v):
    # "00분 00초" 형태로 0을 채워서 보여준다 (형이 요청한 포맷).
    if pd.isna(v):
        return "-"
    v = int(round(v))
    return f"{v // 60:02d}분 {v % 60:02d}초"


def render_ga_channel_inflow_page(df: pd.DataFrame):
    """GA 매체별 유입 경로 — 세션 소스/매체 단위 일별 데이터를 기간 합산해서 보여준다.
    '매체'(채널 그룹핑, 예: 네이버 검색광고/구글(P-MAX) 등)는 형이 UTM 매핑을 완료해서
    다시 올리기 전까지는 비어있는 게 정상이다."""
    if df.empty:
        st.info(
            "아직 데이터가 없습니다. 왼쪽 사이드바 '② GA 유입 데이터 업로드'에서 "
            "파일을 올려주세요."
        )
        return

    df = df.copy()
    df["report_date"] = pd.to_datetime(df["report_date"]).dt.date
    st.subheader("🔎 기간 필터")
    min_d, max_d = df["report_date"].min(), df["report_date"].max()
    start, end = period_filter(min_d, max_d, key="ga_channel", default_preset="이번달")
    fd = df[(df["report_date"] >= start) & (df["report_date"] <= end)]

    if fd.empty:
        st.info("선택한 기간에 데이터가 없습니다.")
        return

    if fd["channel"].notna().any():
        st.caption("※ '매체' 컬럼은 UTM 매핑이 반영된 소스/매체만 채워져 있습니다.")
    else:
        st.caption("※ '매체'(채널 그룹핑) 매핑은 아직 반영 전이라 전부 비어있습니다 — 매핑 완료 파일을 다시 올리면 채워집니다.")

    # ── ① 방문자 추이 (GA 기준) — '유입·매출 비교' 페이지와 동일한 카드 구성 ──
    st.markdown("##### ① 방문자 추이 (GA 기준)")
    users_sum = fd["users"].sum()
    new_users_sum = fd["new_users"].sum()
    returning_users_sum = fd["returning_users"].sum()
    new_ratio = (new_users_sum / users_sum * 100) if users_sum else 0
    returning_ratio = (returning_users_sum / users_sum * 100) if users_sum else 0
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("총 방문자 합계", f"{users_sum:,.0f} 명")
    c2.metric("신규 방문자 합계", f"{new_users_sum:,.0f} 명")
    c3.metric("재방문자 합계", f"{returning_users_sum:,.0f} 명")
    c4.metric("신규 방문자 비중", f"{new_ratio:.1f} %")
    c5.metric("재방문자 비중", f"{returning_ratio:.1f} %")

    # 소스/매체별로 나뉜 데이터를 날짜 기준으로 다시 합쳐서(전체 소스/매체 합산) 일별 추이를 그린다.
    daily_trend = fd.groupby("report_date", as_index=False).agg(
        users=("users", "sum"), new_users=("new_users", "sum")
    ).sort_values("report_date")
    date_label_order = kor_date_labels(daily_trend["report_date"], "day")
    daily_trend["일자"] = date_label_order
    visit_chart_df = daily_trend.melt(
        id_vars=["일자"], value_vars=["users", "new_users"],
        var_name="구분", value_name="방문자수",
    )
    visit_chart_df["구분"] = visit_chart_df["구분"].map({"users": "총 방문자", "new_users": "신규 방문자"})
    fig_visit = px.line(
        visit_chart_df, x="일자", y="방문자수", color="구분", markers=True,
        category_orders={"일자": date_label_order},
    )
    fig_visit.update_yaxes(tickformat=",.0f")
    fig_visit.for_each_trace(
        lambda t: t.update(hovertemplate=f"구분={t.name}<br>일자=%{{x}}<br>방문자수=%{{y:,.0f}}명<extra></extra>")
    )
    st.plotly_chart(theme_chart(fig_visit), use_container_width=True)

    # ── ② 구매·매출 합계 ──
    st.markdown("##### ② 구매·매출 합계")
    conv_sum = fd["conversions"].sum()
    rev_sum = fd["revenue"].sum()
    r1, r2 = st.columns(2)
    r1.metric("구매 합계", f"{conv_sum:,.0f} 건")
    r2.metric("매출 합계", f"{rev_sum:,.0f} 원")

    # 선택 기간 안에서 소스/매체 단위로 합산(방문자·조회수·구매·매출은 합, 이탈률·체류시간은 평균).
    agg = (
        fd.groupby("source_medium", as_index=False)
        .agg(
            channel=("channel", "first"),
            users=("users", "sum"),
            new_users=("new_users", "sum"),
            returning_users=("returning_users", "sum"),
            bounce_rate=("bounce_rate", "mean"),
            pageviews=("pageviews", "sum"),
            avg_session_duration=("avg_session_duration", "mean"),
            conversions=("conversions", "sum"),
            revenue=("revenue", "sum"),
        )
        .sort_values("users", ascending=False)
        .reset_index(drop=True)
    )

    st.markdown("##### ③ 소스/매체별 TOP 10")
    metric_tabs = st.tabs([label for _, label in GA_CHANNEL_TOP_METRICS])
    for tab_widget, (metric_col, metric_label) in zip(metric_tabs, GA_CHANNEL_TOP_METRICS):
        with tab_widget:
            top10 = agg.sort_values(metric_col, ascending=False).head(10)
            fig_top = px.bar(
                top10, x="source_medium", y=metric_col,
                labels={"source_medium": "세션 소스/매체", metric_col: metric_label},
            )
            fig_top.update_yaxes(tickformat=",.0f")
            unit = "원" if metric_col == "revenue" else ("건" if metric_col == "conversions" else "명")
            fig_top.update_traces(
                hovertemplate=f"세션 소스/매체=%{{x}}<br>{metric_label}=%{{y:,.0f}}{unit}<extra></extra>"
            )
            st.plotly_chart(theme_chart(fig_top), use_container_width=True, key=f"ga_channel_top_{metric_col}")

    st.markdown("##### ④ 소스/매체별 상세 표")
    st.caption("TOTAL 행은 현재 페이지가 아니라 선택한 기간 전체(모든 소스/매체) 기준 합계입니다.")

    total = len(agg)
    narrow_ps_col, _ps_spacer = st.columns([2, 10])
    with narrow_ps_col:
        page_size = st.selectbox("페이지당 표시", PAGE_SIZE_OPTIONS, index=1, key="ga_channel_pagesize")
    total_pages = max(1, -(-total // page_size))
    page = render_pager(total_pages, key="ga_channel_pager") if total_pages > 1 else 1
    start_i, end_i = (page - 1) * page_size, page * page_size
    view = agg.iloc[start_i:end_i].copy()

    show = format_display(view[GA_CHANNEL_DETAIL_COLS])
    show["avg_session_duration"] = view["avg_session_duration"].map(_fmt_duration_padded)
    show["channel"] = view["channel"].fillna("미분류")

    total_row = {
        "source_medium": "TOTAL",
        "channel": "",
        "users": f"{users_sum:,.0f}",
        "new_users": f"{new_users_sum:,.0f}",
        "returning_users": f"{fd['returning_users'].sum():,.0f}",
        "bounce_rate": f"{fd['bounce_rate'].mean():.2f}%" if fd["bounce_rate"].notna().any() else "-",
        "pageviews": f"{fd['pageviews'].sum():,.0f}",
        "avg_session_duration": _fmt_duration_padded(fd["avg_session_duration"].mean()),
        "conversions": f"{conv_sum:,.0f}",
        "revenue": f"{rev_sum:,.0f}",
    }
    table = pd.concat([pd.DataFrame([total_row])[GA_CHANNEL_DETAIL_COLS], show[GA_CHANNEL_DETAIL_COLS]], ignore_index=True)
    table = table.rename(columns=GA_CHANNEL_LABELS)
    raw_numeric_cols = [
        "users", "new_users", "returning_users", "bounce_rate", "pageviews",
        "avg_session_duration", "conversions", "revenue",
    ]
    render_html_table(table, raw=view[raw_numeric_cols], raw_label_map=GA_CHANNEL_LABELS_REV)

    st.markdown("#### 💬 코멘트")
    comment_body = [
        f"선택 기간 총 방문자 {users_sum:,.0f}명(신규 비중 {new_ratio:.1f}%, 재방문 비중 {returning_ratio:.1f}%), "
        f"구매 {conv_sum:,.0f}건, 매출 {rev_sum:,.0f}원입니다."
    ]
    if not agg.empty:
        top1 = agg.iloc[0]
        comment_body.append(f"가장 유입이 많은 소스/매체는 {top1['source_medium']}({top1['users']:,.0f}명)입니다.")
    st.markdown(" ".join(comment_body))
    if not agg.empty:
        st.markdown(
            _ops_next_action(
                f"{top1['source_medium']} 유입이 다음 기간에도 유지되는지 확인하고, "
                "유입 대비 구매(전환)가 낮은 소스/매체가 있으면 랜딩페이지·타겟팅 점검을 권장합니다."
            ),
            unsafe_allow_html=True,
        )
    if agg["channel"].isna().all() if not agg.empty else False:
        st.caption("※ '매체'(채널 그룹핑) 매핑 전이라 소스/매체 단위로만 코멘트했습니다.")

    dl_df = format_display(agg[GA_CHANNEL_DETAIL_COLS])
    dl_df["avg_session_duration"] = agg["avg_session_duration"].map(_fmt_duration_padded)
    dl_df["channel"] = agg["channel"].fillna("미분류")
    st.download_button(
        "⬇️ 엑셀 다운로드 (GA 매체별 유입 경로)",
        data=to_excel_bytes(dl_df.rename(columns=GA_CHANNEL_LABELS)),
        file_name="ga_channel_inflow.xlsx",
    )


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
# 운영 코멘트 (신규) — 데이터를 바탕으로 자동 생성하는 실무 코멘트/제안
# performance-marketing-analysis 스킬(코멘트 작성 순서·소재 필터링·예산재배분 원칙)과
# 프로젝트 KPI 기준(판단 지표 GA-ROAS, 목표 200~300%)을 그대로 반영한 룰 기반 코멘트.
# ──────────────────────────────────────────────────────────────
OPS_KPI_ROAS_LOW = 200    # KPI 목표 하단(%) — GA 기준
OPS_KPI_ROAS_HIGH = 300   # KPI 목표 상단(%), 300%+가 목표
OPS_MIN_CHANNEL_SPEND = 100_000   # 매체 판단 보류 기준: 광고비 10만원 미만
OPS_CREATIVE_MIN_SPEND = 50_000   # 소재 판단 보류 기준(performance-marketing-analysis 스킬 기본값과 동일)


def _ops_kpi_status(roas: float) -> str:
    if roas >= OPS_KPI_ROAS_HIGH:
        return "목표 초과 달성"
    if roas >= OPS_KPI_ROAS_LOW:
        return "목표 구간 내"
    return "목표 미달"


def _ops_fmt_pp(v: float) -> str:
    # 표(render_html_table)와 같은 색 규칙: ▲(상승)는 빨강, ▼(하락)는 파랑.
    # HTML 문자열이라 이 값을 쓰는 st.markdown 호출엔 반드시 unsafe_allow_html=True를 줘야 한다.
    color = "#d93025" if v >= 0 else "#1a73e8"
    arrow = "▲" if v >= 0 else "▼"
    return f'<span style="color:{color}">{arrow}{abs(v):,.0f}%p</span>'


def _ops_fmt_pct(v: float) -> str:
    color = "#d93025" if v >= 0 else "#1a73e8"
    arrow = "▲" if v >= 0 else "▼"
    return f'<span style="color:{color}">{arrow}{abs(v):,.1f}%</span>'


def _ops_next_action(text: str) -> str:
    """'다음 액션' 라벨부터 그 뒤 제안 문장까지 전부 굵은 빨간 텍스트로 강조해서 눈에 띄게 한다.
    이 함수가 붙은 문장을 st.markdown()으로 렌더링할 땐 unsafe_allow_html=True가 필요하다."""
    return f'<b style="color:#d93025;">다음 액션: {text}</b>'


def _short_week_label(row) -> str:
    """weekly 행에서 'N월 N주차' 짧은 라벨을 뽑는다. 원본 엑셀 라벨에 '주차' 표기가 없으면
    week_start~week_end 날짜(MM/DD)로 대체한다."""
    m = re.search(r"(\d{1,2})월\s*(\d+)\s*주차", str(row.get("label", "")))
    if m:
        return f"{m.group(1)}월 {m.group(2)}주차"
    return f"{row['week_start']:%m/%d}~{row['week_end']:%m/%d}"


def _ops_driver_sentence(cur, prev, cur_label: str, prev_label: str, roas_col: str = "roas", roas_label: str = "ROAS") -> str:
    """단순히 'ROAS가 올랐다/내렸다'만 말하지 않고, 광고비·노출·클릭·매출이 어느 방향으로
    같이 움직였는지(볼륨 측) + CVR·객단가가 어떻게 바뀌어서 ROAS가 그렇게 움직였는지(효율 측)를
    묶어서 설명하는 문장을 만든다. 월별/주간별 코멘트에서 공용으로 쓴다.
    반환값에 <span>/<b> 태그가 섞여 있어 st.markdown()에 unsafe_allow_html=True가 필요하다."""

    def pct(a, b):
        return (a - b) / abs(b) * 100 if b else 0

    d_cost = cur.get("cost_incl_vat", 0) - prev.get("cost_incl_vat", 0)
    d_impr_pct = pct(cur.get("impressions", 0), prev.get("impressions", 0))
    d_clicks_pct = pct(cur.get("clicks", 0), prev.get("clicks", 0))
    d_rev_pct = pct(cur.get("revenue", 0), prev.get("revenue", 0))
    d_cvr_pct = pct(cur.get("cvr", 0), prev.get("cvr", 0))
    d_aov = cur.get("aov", 0) - prev.get("aov", 0)
    d_aov_pct = pct(cur.get("aov", 0), prev.get("aov", 0))
    d_roas_pp = cur.get(roas_col, 0) - prev.get(roas_col, 0)

    cost_word = "증가" if d_cost >= 0 else "감소"
    metrics = [("노출수", d_impr_pct), ("클릭수", d_clicks_pct), ("매출", d_rev_pct)]
    same_dir = [n for n, v in metrics if (v >= 0) == (d_cost >= 0)]
    diff_dir = [n for n, v in metrics if (v >= 0) != (d_cost >= 0)]

    if len(same_dir) == 3:
        vol_sentence = f"{cur_label}는 {prev_label} 대비 광고비가 {cost_word}하면서 노출수·클릭수·매출도 동반 {cost_word}했습니다."
    elif same_dir:
        vol_sentence = (
            f"{cur_label}는 {prev_label} 대비 광고비가 {cost_word}했고 {'·'.join(same_dir)}도 함께 {cost_word}했지만, "
            f"{'·'.join(diff_dir)}는 반대로 움직였습니다."
        )
    else:
        vol_sentence = f"{cur_label}는 {prev_label} 대비 광고비가 {cost_word}했지만, 노출수·클릭수·매출은 반대로 움직였습니다."

    eff_bits = []
    if abs(d_cvr_pct) >= 0.5:
        eff_bits.append(f"CVR이 {_ops_fmt_pct(d_cvr_pct)} {'상승' if d_cvr_pct >= 0 else '하락'}")
    if abs(d_aov_pct) >= 0.5:
        eff_bits.append(f"객단가가 {_ops_fmt_pct(d_aov_pct)}({d_aov:+,.0f}원) {'상승' if d_aov_pct >= 0 else '하락'}")

    roas_verb = "상승" if d_roas_pp >= 0 else "하락"
    if eff_bits:
        eff_sentence = "다만 " + ", ".join(eff_bits) + f"한 영향으로 {roas_label}는 {_ops_fmt_pp(d_roas_pp)} {roas_verb}했습니다."
    else:
        eff_sentence = f"{roas_label}는 {_ops_fmt_pp(d_roas_pp)} {roas_verb}했습니다."

    return vol_sentence + " " + eff_sentence


def _ops_top_mover_sentence(
    df: pd.DataFrame, key_col: str, period_col: str, cur_period, prev_period,
    cost_col: str = "cost_incl_vat", rev_col: str = "revenue", min_abs_cost: float = 0, min_abs_rev: float = 0,
) -> str:
    """전체(또는 그룹) 수치가 왜 그렇게 움직였는지, 어떤 매체/소재가 광고비·매출 변화를
    가장 크게 이끌었는지 짚어주는 문장 하나를 만든다. 매체별/소재별 코멘트에서 공용으로 쓴다.
    df엔 [key_col, period_col, cost_col, rev_col]가 있어야 한다. cur/prev 중 한쪽에만 있는
    key(이번 주 신규 집행/운영 중지된 매체 등)도 그 기간엔 0원으로 보고 비교 대상에 포함한다 —
    실제로 매체 운영을 중지/재개하는 것 자체가 흔한 원인이라 빠뜨리면 안 된다."""
    if df is None or df.empty:
        return ""
    cur_df = df[df[period_col] == cur_period].groupby(key_col)[[cost_col, rev_col]].sum()
    prev_df = df[df[period_col] == prev_period].groupby(key_col)[[cost_col, rev_col]].sum()
    all_keys = cur_df.index.union(prev_df.index)
    if len(all_keys) == 0:
        return ""
    cur_df = cur_df.reindex(all_keys, fill_value=0)
    prev_df = prev_df.reindex(all_keys, fill_value=0)
    d_cost = cur_df[cost_col] - prev_df[cost_col]
    d_rev = cur_df[rev_col] - prev_df[rev_col]

    bits = []
    cost_driver = None
    if len(d_cost) and d_cost.abs().max() > min_abs_cost:
        cost_driver = d_cost.abs().idxmax()
        v = d_cost[cost_driver]
        stop_note = "(운영 중지)" if cur_df.loc[cost_driver, cost_col] == 0 and v < 0 else (
            "(신규 집행)" if prev_df.loc[cost_driver, cost_col] == 0 and v > 0 else ""
        )
        bits.append(f"{cost_driver}에서 광고비가 {abs(v):,.0f}원 {'증가' if v >= 0 else '감소'}{stop_note}한 영향이 컸고")
    if len(d_rev) and d_rev.abs().max() > min_abs_rev:
        rev_driver = d_rev.abs().idxmax()
        v = d_rev[rev_driver]
        if rev_driver == cost_driver:
            bits.append(f"같은 곳에서 매출도 {abs(v):,.0f}원 {'증가' if v >= 0 else '감소'}하며 영향을 키웠습니다")
        else:
            connector = "대신 " if bits else ""
            bits.append(f"{connector}{rev_driver}에서 매출이 {abs(v):,.0f}원 {'증가' if v >= 0 else '감소'}한 영향이 컸습니다")
    if not bits:
        return ""
    if len(bits) == 1:
        return bits[0] + "."
    return bits[0] + ", " + bits[1] + "."


def _ops_channel_efficiency_driver_sentence(
    df: pd.DataFrame, key_col: str, period_col: str, cur_period, prev_period,
    cost_col: str = "cost_incl_vat", conv_col: str = "conversions", clicks_col: str = "clicks",
    rev_col: str = "revenue",
) -> str:
    """광고비·매출 증감(위 _ops_top_mover_sentence)만으로는 설명이 안 되는 ROAS 변화 —
    즉 광고비·매출이 둘 다 줄었는데도 ROAS는 오히려 오른 경우처럼, '효율' 쪽에서 원인을 찾아야
    할 때 쓴다. 두 기간 모두 광고비가 있었던 매체 중 ROAS(%)가 가장 많이 오른/내린 매체를 찾아
    CVR·객단가 변화로 설명한다."""
    if df is None or df.empty:
        return ""
    cur_g = df[df[period_col] == cur_period].groupby(key_col)[[cost_col, conv_col, clicks_col, rev_col]].sum()
    prev_g = df[df[period_col] == prev_period].groupby(key_col)[[cost_col, conv_col, clicks_col, rev_col]].sum()
    common = [k for k in cur_g.index.intersection(prev_g.index)
              if cur_g.loc[k, cost_col] > 0 and prev_g.loc[k, cost_col] > 0]
    if not common:
        return ""
    best_k, best_d_roas, best_stats = None, 0, None
    for k in common:
        cur_roas = cur_g.loc[k, rev_col] / cur_g.loc[k, cost_col] * 100
        prev_roas = prev_g.loc[k, rev_col] / prev_g.loc[k, cost_col] * 100
        d_roas = cur_roas - prev_roas
        if best_k is None or abs(d_roas) > abs(best_d_roas):
            cur_cvr = cur_g.loc[k, conv_col] / cur_g.loc[k, clicks_col] * 100 if cur_g.loc[k, clicks_col] else 0
            prev_cvr = prev_g.loc[k, conv_col] / prev_g.loc[k, clicks_col] * 100 if prev_g.loc[k, clicks_col] else 0
            cur_aov = cur_g.loc[k, rev_col] / cur_g.loc[k, conv_col] if cur_g.loc[k, conv_col] else 0
            prev_aov = prev_g.loc[k, rev_col] / prev_g.loc[k, conv_col] if prev_g.loc[k, conv_col] else 0
            best_k, best_d_roas = k, d_roas
            best_stats = {"d_cvr": cur_cvr - prev_cvr, "d_aov": cur_aov - prev_aov}
    if best_k is None or abs(best_d_roas) < 1:
        return ""
    bits = []
    if abs(best_stats["d_cvr"]) >= 0.01:
        bits.append(f"CVR이 {abs(best_stats['d_cvr']):.1f}%p {'상승' if best_stats['d_cvr'] >= 0 else '하락'}")
    if abs(best_stats["d_aov"]) >= 1:
        bits.append(f"객단가가 {abs(best_stats['d_aov']):,.0f}원 {'상승' if best_stats['d_aov'] >= 0 else '하락'}")
    eff = ", ".join(bits) if bits else "효율 지표가 개선"
    verb = "오른" if best_d_roas >= 0 else "내린"
    return f"이 중 {best_k}에서 {eff}하며 ROAS가 {abs(best_d_roas):.0f}%p {verb} 것이 전체 ROAS 변동의 주요 요인으로 보입니다."


WEEKDAY_KOR = ["월", "화", "수", "목", "금", "토", "일"]

_WEEKLY_ZERO_CHECK_COLS = ["impressions", "clicks", "cost_excl_vat", "cost_incl_vat", "conversions", "revenue"]


def _drop_trailing_zero_weeks(weekly: pd.DataFrame) -> pd.DataFrame:
    """parse_weekly()에서 이미 걸러내지만, Supabase에 예전 코드로 저장된 빈(전부 0) 미래 주
    placeholder 행이 그대로 남아있을 수도 있어서 화면에 쓸 때도 한 번 더 방어적으로 걸러준다.
    week_start 기준 정렬 후, 값이 하나도 없는 말미 행들을 제거한다."""
    if weekly is None or weekly.empty:
        return weekly
    w = weekly.sort_values("week_start").reset_index(drop=True)
    cols = [c for c in _WEEKLY_ZERO_CHECK_COLS if c in w.columns]
    if not cols:
        return w
    metric_sum = w[cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
    nonzero_idx = metric_sum[metric_sum > 0].index
    if len(nonzero_idx):
        w = w.loc[: nonzero_idx.max()]
    return w.reset_index(drop=True)


def _weekly_period_label(weekly: pd.DataFrame, label_prefix: str = "주간 코멘트 정리") -> str:
    """weekly(주간 통합 데이터)의 가장 최근 주를 '{prefix} - 26년 7월 5주차 : 07/27(월) ~08/02(일)'
    형식으로 만든다. 원본 엑셀의 'N월 N주차' 라벨 텍스트에서 'N주차' 숫자만 재사용하고,
    날짜 범위와 요일은 week_start/week_end로 직접 계산해 매주 자동으로 갱신되게 한다."""
    weekly = _drop_trailing_zero_weeks(weekly)
    if weekly is None or weekly.empty:
        return label_prefix
    w = weekly.sort_values("week_end").iloc[-1]
    week_start, week_end = w["week_start"], w["week_end"]
    if isinstance(week_start, str):
        week_start = pd.to_datetime(week_start).date()
    if isinstance(week_end, str):
        week_end = pd.to_datetime(week_end).date()
    year2 = f"{week_start.year % 100:02d}"
    month = week_start.month
    m = re.search(r"(\d+)\s*주차", str(w.get("label", "")))
    week_part = f"{m.group(1)}주차" if m else ""
    start_s = f"{week_start.month:02d}/{week_start.day:02d}({WEEKDAY_KOR[week_start.weekday()]})"
    end_s = f"{week_end.month:02d}/{week_end.day:02d}({WEEKDAY_KOR[week_end.weekday()]})"
    period = f"{year2}년 {month}월 {week_part} : {start_s} ~{end_s}".replace("  ", " ")
    return f"{label_prefix} - {period}"


def render_ops_comment_monthly(monthly: pd.DataFrame, heading: str = "#### 💬 월별 코멘트", channels: pd.DataFrame = None):
    """월별 총평 코멘트 — 종합 대시보드의 '1) 월별 누적' 표 아래와, 운영 코멘트 탭의 ①에서
    공용으로 쓴다. GA-ROAS를 KPI 판단 기준(200~300%)으로 쓰고, 자체 ROAS는 참고로 같이 보여준다.
    channels(매체별 월간 데이터)가 주어지면, '전체'만 보고 끝내지 않고 어느 매체가 광고비/매출
    변화를 가장 크게 이끌었는지 매체명까지 짚어주는 문장을 한 줄 더 붙인다."""
    st.markdown(heading)
    if monthly.empty:
        st.caption("월별 데이터가 아직 없습니다.")
        return
    m = monthly.copy()
    m["report_month"] = pd.to_datetime(m["report_month"]).dt.date
    m = add_kpis(m.sort_values("report_month")).reset_index(drop=True)
    cur = m.iloc[-1]
    prev = m.iloc[-2] if len(m) >= 2 else None
    cur_label = f"{cur['report_month'].year}년 {cur['report_month'].month}월"
    cur_roas, cur_garoas = cur.get("roas", 0), cur.get("ga_roas", 0)
    status = _ops_kpi_status(cur_garoas)

    st.markdown(f"**{cur_label} GA-ROAS {cur_garoas:,.0f}%로 {status}** (자체 ROAS {cur_roas:,.0f}%, KPI {OPS_KPI_ROAS_LOW}~{OPS_KPI_ROAS_HIGH}%)")

    body = []
    if cur_roas and cur_garoas and abs(cur_roas - cur_garoas) / max(cur_garoas, 1) > 0.3:
        body.append(
            f"자체 ROAS와 GA-ROAS 격차가 {_ops_fmt_pp(cur_roas - cur_garoas)}로 큰 편이라, "
            "두 지표 중 어느 쪽을 메인으로 볼지 다시 짚어볼 필요가 있습니다."
        )
    trend_down = None
    if prev is not None:
        prev_label = f"{prev['report_month'].year}년 {prev['report_month'].month}월"
        d_garoas = cur_garoas - prev.get("ga_roas", 0)
        trend_down = d_garoas < 0
        body.append(_ops_driver_sentence(cur, prev, cur_label, prev_label, roas_col="ga_roas", roas_label="GA-ROAS"))
        if channels is not None and not channels.empty and "report_month" in channels.columns:
            ch = channels.copy()
            ch["report_month"] = pd.to_datetime(ch["report_month"]).dt.date
            mover = _ops_top_mover_sentence(
                ch, key_col="channel", period_col="report_month",
                cur_period=cur["report_month"], prev_period=prev["report_month"],
            )
            if mover:
                body.append(mover)
            eff_mover = _ops_channel_efficiency_driver_sentence(
                ch, key_col="channel", period_col="report_month",
                cur_period=cur["report_month"], prev_period=prev["report_month"],
            )
            if eff_mover:
                body.append(eff_mover)
    if body:
        st.markdown(" ".join(body), unsafe_allow_html=True)

    if status == "목표 미달":
        if trend_down:
            body_next = (
                "목표 미달에 하락세까지 겹쳤습니다. 이번 달 남은 기간엔 ②매체별·③소재별 코멘트를 먼저 "
                "확인해 원인(매체/소재/타겟팅)을 좁히고, 다음 달 예산 조정 여부를 결정하는 것을 권장합니다."
            )
        else:
            body_next = (
                "목표 미달이지만 반등 조짐이 있어, 이번 달은 매체·소재 단위 원인 점검만 진행하고 "
                "다음 달 초까지 개선 여부를 지켜본 뒤 예산 조정을 검토하세요."
            )
    elif status == "목표 초과 달성":
        body_next = (
            "목표를 안정적으로 상회하고 있어, 다음 달엔 ②매체별 코멘트에서 목표 초과로 분류된 매체 위주로 "
            "10~20% 증액을 검토하는 것을 권장합니다."
        )
    else:
        body_next = "목표 구간 내에서 안정적으로 운영되고 있어, 다음 달에도 현재 운영 기조(매체 비중·소재 구성)를 유지하는 것을 권장합니다."
    st.markdown(_ops_next_action(body_next), unsafe_allow_html=True)


def render_ops_comment_weekly(
    weekly: pd.DataFrame, heading: str = "#### 💬 주간별 코멘트", channels_weekly: pd.DataFrame = None,
):
    """주간별 코멘트 — 종합 대시보드의 '2) 주간별 누적' 표 아래에 쓴다.
    주간 데이터엔 GA-ROAS가 없어(주간 리포트 시트 구조상 GA 비교 컬럼 없음) 자체 ROAS만 쓴다.
    channels_weekly(매체 개별 시트의 주간 데이터)가 주어지면, 어느 매체가 이번 주 광고비/매출
    변화를 가장 크게 이끌었는지 매체명까지 짚어주는 문장을 한 줄 더 붙인다."""
    st.markdown(heading)
    weekly = _drop_trailing_zero_weeks(weekly)
    if weekly is None or weekly.empty:
        st.caption("주간 데이터가 아직 없습니다.")
        return
    w = add_kpis(weekly.sort_values("week_start")).reset_index(drop=True)
    cur = w.iloc[-1]
    prev = w.iloc[-2] if len(w) >= 2 else None
    cur_label = f"{cur['week_start']:%Y-%m-%d}~{cur['week_end']:%Y-%m-%d}"
    cur_roas = cur.get("roas", 0)
    status = _ops_kpi_status(cur_roas)

    st.markdown(f"**{cur_label} 자체 ROAS {cur_roas:,.0f}%로 {status}** (KPI {OPS_KPI_ROAS_LOW}~{OPS_KPI_ROAS_HIGH}%, GA 비교는 월별 코멘트 참고)")

    trend_down = None
    if prev is not None:
        prev_label = f"{prev['week_start']:%Y-%m-%d}~{prev['week_end']:%Y-%m-%d}"
        d_roas = cur_roas - prev.get("roas", 0)
        trend_down = d_roas < 0
        st.markdown(
            _ops_driver_sentence(
                cur, prev, _short_week_label(cur), _short_week_label(prev), roas_col="roas", roas_label="ROAS"
            ),
            unsafe_allow_html=True,
        )
        if channels_weekly is not None and not channels_weekly.empty and "week_start" in channels_weekly.columns:
            cwk = channels_weekly.copy()
            cwk["week_start"] = pd.to_datetime(cwk["week_start"]).dt.date
            mover = _ops_top_mover_sentence(
                cwk, key_col="channel", period_col="week_start",
                cur_period=cur["week_start"], prev_period=prev["week_start"],
            )
            if mover:
                st.markdown(mover)
            eff_mover = _ops_channel_efficiency_driver_sentence(
                cwk, key_col="channel", period_col="week_start",
                cur_period=cur["week_start"], prev_period=prev["week_start"],
            )
            if eff_mover:
                st.markdown(eff_mover)
            cur_ch = sorted(cwk.loc[cwk["week_start"] == cur["week_start"], "channel"].unique())
            prev_ch = sorted(cwk.loc[cwk["week_start"] == prev["week_start"], "channel"].unique())
            st.caption(
                f"※ 매체별 주간 데이터 인식: 전체 {cwk['channel'].nunique()}개 매체 · "
                f"{_short_week_label(cur)} {len(cur_ch)}개({', '.join(cur_ch) or '없음'}) · "
                f"{_short_week_label(prev)} {len(prev_ch)}개({', '.join(prev_ch) or '없음'})"
            )
            # 매체마다 가장 최근에 잡힌 주가 언제인지(뒤처진 매체 탭이 있는지) 바로 보이게.
            latest_per_channel = (
                cwk.groupby("channel")["week_start"].max().sort_values(ascending=False)
            )
            latest_lines = "; ".join(f"{ch}: {wk}" for ch, wk in latest_per_channel.items())
            st.caption(f"※ 매체별 최신 주간 데이터 시점 — {latest_lines}")
        elif channels_weekly is not None and channels_weekly.empty:
            st.caption("※ 매체별 주간 데이터가 아직 없습니다 — 매체별 코멘트는 전체 합산 기준으로만 표시됩니다.")

    if trend_down is None:
        next_action = "아직 비교할 전주 데이터가 없어 다음 주 데이터가 쌓이면 추세를 판단할 수 있습니다."
    elif status == "목표 미달" and trend_down:
        next_action = "목표 미달+하락세가 겹쳤습니다. 이번 주 안에 소재/타겟팅을 먼저 점검하고, 다음 주에도 개선이 없으면 예산 축소를 검토하세요."
    elif status == "목표 미달":
        next_action = "목표는 미달이지만 반등 조짐이 있어, 다음 주까지는 현재 설정을 유지하며 지켜보는 것을 권장합니다."
    elif not trend_down and status != "목표 미달":
        next_action = "상승 추세라 다음 주도 현재 운영을 유지하며 지켜보는 것을 권장합니다."
    else:
        next_action = "하락 추세이니 다음 주엔 원인(소재 소진·시즌성 등)을 점검해보는 것을 권장합니다."
    st.markdown(_ops_next_action(next_action), unsafe_allow_html=True)


def _ops_channel_bucket_lines(df: pd.DataFrame, roas_col: str) -> list:
    """매체별 df('channel','cost_incl_vat', roas_col 컬럼 필요)를 목표초과/근접/미달로 나눠
    코멘트 문장 리스트를 만든다. 광고비가 OPS_MIN_CHANNEL_SPEND 미만인 매체는 판단 보류로 뺀다.
    그룹 전체의 min~max 범위 대신, 매체별 실제 수치를 각각 이름과 함께 보여준다."""
    if df.empty or roas_col not in df.columns:
        return []
    judged = df[df["cost_incl_vat"] >= OPS_MIN_CHANNEL_SPEND].copy()
    held = df[df["cost_incl_vat"] < OPS_MIN_CHANNEL_SPEND].copy()
    lines = []
    if not judged.empty:
        def _bucket(v):
            if v >= OPS_KPI_ROAS_HIGH:
                return "목표 초과"
            if v >= OPS_KPI_ROAS_LOW:
                return "목표 근접"
            return "목표 미달"

        judged["구분"] = judged[roas_col].map(_bucket)
        judged = judged.sort_values(roas_col, ascending=False)
        over = judged[judged["구분"] == "목표 초과"]
        near = judged[judged["구분"] == "목표 근접"]
        under = judged[judged["구분"] == "목표 미달"]

        def _names_with_roas(sub):
            return ", ".join(f"{r.channel}({getattr(r, roas_col):,.0f}%)" for r in sub.itertuples())

        if not over.empty:
            lines.append(
                f"**목표 초과 ({len(over)}개 매체)**: {_names_with_roas(over)} — 목표를 상회합니다. "
                + _ops_next_action("이번 주 중 10~20% 증액을 적용해보고, 다음 주 재측정 후 유지·추가 증액 여부를 결정하는 것을 권장합니다.")
            )
        if not near.empty:
            lines.append(
                f"**목표 근접 ({len(near)}개 매체)**: {_names_with_roas(near)} — 목표 구간 안쪽입니다. "
                + _ops_next_action("현재 예산·소재를 유지하며 다음 주까지 추세를 지켜보는 것을 권장합니다.")
            )
        if not under.empty:
            lines.append(
                f"**목표 미달 ({len(under)}개 매체)**: {_names_with_roas(under)} — 목표에 못 미칩니다. "
                + _ops_next_action(
                    "이번 주엔 소재 교체를 먼저 시도하고, 다음 주에도 개선이 없으면 총 예산의 10~15% "
                    "내에서 목표 초과 매체 쪽으로 단계적으로 이동하는 것을 권장합니다."
                )
            )
    if not held.empty:
        lines.append(f"판단 보류(광고비 {OPS_MIN_CHANNEL_SPEND:,}원 미만, 표본 부족): {', '.join(held['channel'])}")
    return lines


def render_ops_comment_channel_narrative(
    df: pd.DataFrame, roas_col: str, heading: str = "#### 💬 매체별 코멘트", footnote: str = "",
):
    """매체별 표 바로 아래 붙이는 짧은 서술형 코멘트(목표초과/근접/미달 분류 + 예산재배분 제안만).
    표는 이미 위에 있으므로 여기서는 표를 다시 그리지 않는다."""
    st.markdown(heading)
    lines = _ops_channel_bucket_lines(df, roas_col)
    if not lines:
        st.caption("판정 가능한 매체 데이터가 아직 없습니다.")
        return
    for line in lines:
        st.markdown(line, unsafe_allow_html=True)
    if footnote:
        st.caption(footnote)


def render_operation_comment_page(
    weekly: pd.DataFrame, monthly: pd.DataFrame, daily: pd.DataFrame,
    channels: pd.DataFrame, snapshot: pd.DataFrame, creatives: pd.DataFrame,
    audience: pd.DataFrame, inflow_revenue: pd.DataFrame, ga_channel_inflow: pd.DataFrame,
    agency_notes: pd.DataFrame = None,
):
    """⓪ 대행사 운영 메모 → ① 전체 총평 → ② 매체별 코멘트+예산재배분 제안 →
    ③ 소재별 제안 → ④ GA 유입 코멘트 순서로, 업로드된 데이터를 바탕으로 룰 기반 코멘트를
    자동 생성한다. 어디까지나 1차 초안이라 최종 운영 판단은 직접 검토가 필요하다는 점을 상단에 명시한다."""
    st.info(
        f"아래 코멘트는 업로드된 데이터를 기준으로 자동 생성한 1차 초안입니다. "
        f"판단 지표는 GA-ROAS, KPI 목표는 {OPS_KPI_ROAS_LOW}~{OPS_KPI_ROAS_HIGH}%입니다. "
        "최종 운영 판단은 직접 검토 후 반영해주세요."
    )

    # ── ⓪ 대행사 운영 메모 (매체통합 시트 하단 자유 텍스트) ──
    # 제목에 그 주의 연도/월/N주차/날짜범위(요일 포함)를 자동으로 붙인다 — 매주 업로드할 때마다
    # weekly(통합 주간별)의 가장 최근 주 기준으로 갱신된다.
    st.markdown(f"## {_weekly_period_label(weekly)}")
    if agency_notes is None or agency_notes.empty:
        st.caption("주간 리포트의 '매체통합' 시트 하단에서 인식된 운영 메모가 아직 없습니다.")
    else:
        latest_note = agency_notes.sort_values("as_of_date").iloc[-1]
        st.caption(f"기준일: {latest_note['as_of_date']} (주간 리포트 하단 코멘트를 그대로 정리한 것입니다)")
        st.markdown(latest_note["note_text"])
    st.caption("※ 실제 파일 구조 검증 전 휴리스틱 추출입니다 — 다음 업로드에서 누락/오인식이 있으면 알려주세요.")

    st.markdown("---")

    # ── ① 전체 총평 ──
    render_ops_comment_monthly(monthly, heading="## ① 전체 총평", channels=channels)

    st.markdown("---")

    # ── ② 매체별 코멘트 + 예산재배분 제안 ──
    st.markdown("## ② 매체별 코멘트 + 예산재배분 제안")
    if snapshot.empty:
        st.caption("매체별 당월 스냅샷 데이터가 아직 없습니다 (① 주간 리포트 업로드 필요).")
    else:
        snap = snapshot.copy()
        latest_month = snap["as_of_month"].max()
        snap = add_kpis(snap[snap["as_of_month"] == latest_month]).reset_index(drop=True)
        st.caption(f"기준월: {latest_month}")

        for line in _ops_channel_bucket_lines(snap, "ga_roas"):
            st.markdown(line, unsafe_allow_html=True)

        judged = snap[snap["cost_incl_vat"] >= OPS_MIN_CHANNEL_SPEND]
        gap_flag = judged[
            (judged["roas"] > 0) & (judged["ga_roas"] > 0)
            & ((judged["roas"] - judged["ga_roas"]).abs() / judged["ga_roas"] > 0.3)
        ]
        if not gap_flag.empty:
            st.caption(
                "⚠️ 자체 ROAS와 GA-ROAS 격차가 큰 매체: "
                + ", ".join(
                    f"{r.channel}(자체 {r.roas:,.0f}% vs GA {r.ga_roas:,.0f}%)"
                    for r in gap_flag.itertuples()
                )
            )

        show_cols = ["channel", "cost_incl_vat", "roas", "ga_roas"]
        show_cols = [c for c in show_cols if c in snap.columns]
        table = format_display(snap[show_cols].sort_values("cost_incl_vat", ascending=False))
        render_html_table(korify(table), raw=snap[show_cols])

    st.markdown("---")

    # ── ③ 소재별 제안 ──
    st.markdown("## ③ 소재별 제안")
    if creatives.empty:
        st.caption("소재별 성과 데이터가 아직 없습니다.")
    else:
        cr = creatives.copy()
        cr["as_of_date"] = pd.to_datetime(cr["as_of_date"]).dt.date
        cr_latest = cr.sort_values("as_of_date").drop_duplicates(subset=["channel", "creative"], keep="last")
        agg = cr_latest.groupby(["channel", "creative"], as_index=False).agg(
            impressions=("impressions", "sum"), clicks=("clicks", "sum"),
            cost_incl_vat=("cost_incl_vat", "sum"), conversions=("conversions", "sum"),
            revenue=("revenue", "sum"),
        )
        agg = add_kpis(agg)
        total_cost, total_rev = agg["cost_incl_vat"].sum(), agg["revenue"].sum()
        account_avg_roas = (total_rev / total_cost * 100) if total_cost else 0

        def _judge(row):
            if row["cost_incl_vat"] < OPS_CREATIVE_MIN_SPEND:
                return "판단 보류"
            if account_avg_roas <= 0:
                return "판단 보류"
            ratio = row["roas"] / account_avg_roas
            if ratio >= 1.2:
                return "우수"
            if ratio <= 0.7:
                return "부진"
            return "평균 수준"

        agg["판정"] = agg.apply(_judge, axis=1)
        good = agg[agg["판정"] == "우수"].sort_values("roas", ascending=False)
        bad = agg[agg["판정"] == "부진"].sort_values("roas")
        held_n = int((agg["판정"] == "판단 보류").sum())

        st.caption(f"계정 평균 ROAS: {account_avg_roas:,.0f}% · 광고비 {OPS_CREATIVE_MIN_SPEND:,}원 미만 {held_n}개 소재는 판단 보류")

        # 전체 소재 수치가 왜 그렇게 움직였는지, 어떤 소재가 광고비/매출 변화를 가장 크게
        # 이끌었는지 이전 스냅샷과 비교해 짚어준다(직전 저장 시점 대비).
        cr_dates = sorted(cr["as_of_date"].unique())
        if len(cr_dates) >= 2:
            cur_snap_date, prev_snap_date = cr_dates[-1], cr_dates[-2]
            cr_key = cr.copy()
            cr_key["key"] = cr_key["channel"].astype(str) + " " + cr_key["creative"].astype(str)
            creative_mover = _ops_top_mover_sentence(
                cr_key, key_col="key", period_col="as_of_date",
                cur_period=cur_snap_date, prev_period=prev_snap_date,
            )
            if creative_mover:
                st.markdown(f"직전 저장({prev_snap_date}) 대비 이번({cur_snap_date}) 소재 수치는 {creative_mover}")

        if not good.empty:
            top_good = ", ".join(f"{r.creative}({r.roas:,.0f}%)" for r in good.head(3).itertuples())
            st.markdown(
                f"**효율 우수 소재 ({len(good)}개)**: 계정 평균 대비 1.2배 이상 — 상위로는 {top_good} 등입니다. "
                + _ops_next_action("이번 주 중 예산을 증액하거나 동일 소재를 타 캠페인/타겟팅으로 확장 적용해보는 것을 권장합니다."),
                unsafe_allow_html=True,
            )
            st.dataframe(
                korify(format_display(good[["channel", "creative", "cost_incl_vat", "roas"]])),
                use_container_width=True, hide_index=True,
            )
        if not bad.empty:
            worst_bad = ", ".join(f"{r.creative}({r.roas:,.0f}%)" for r in bad.head(3).itertuples())
            st.markdown(
                f"**효율 부진 소재 ({len(bad)}개)**: 계정 평균 대비 0.7배 이하 — 하위로는 {worst_bad} 등입니다. "
                "다만 막 시작한 소재는 학습 기간(3~7일)을 감안해 성급히 끄지 않는 것을 권장합니다. "
                + _ops_next_action("이번 주엔 문구/이미지 교체를 먼저 시도하고, 다음 주까지 개선이 없으면 축소/OFF를 검토하세요."),
                unsafe_allow_html=True,
            )
            st.dataframe(
                korify(format_display(bad[["channel", "creative", "cost_incl_vat", "roas"]])),
                use_container_width=True, hide_index=True,
            )
        if good.empty and bad.empty:
            st.caption("우수/부진으로 구분될 만큼 표본이 확보된 소재가 아직 없습니다.")

    st.markdown("---")

    # ── ④ GA 유입 코멘트 ──
    st.markdown("## ④ GA 유입 코멘트")
    if inflow_revenue.empty:
        st.caption("유입·매출 비교 데이터가 아직 없습니다.")
    else:
        iv = inflow_revenue.copy()
        iv["report_date"] = pd.to_datetime(iv["report_date"]).dt.date
        iv["ym"] = iv["report_date"].map(lambda d: (d.year, d.month))
        monthly_iv = iv.groupby("ym", as_index=False).agg(users=("users", "sum"), new_users=("new_users", "sum"))
        monthly_iv = monthly_iv.sort_values("ym")
        if len(monthly_iv):
            cur_iv = monthly_iv.iloc[-1]
            cur_y, cur_mo = cur_iv["ym"]
            body = [f"{cur_y}년 {cur_mo}월 총 방문자 {cur_iv['users']:,.0f}명, 신규 방문자 {cur_iv['new_users']:,.0f}명입니다."]
            if len(monthly_iv) >= 2:
                prev_iv = monthly_iv.iloc[-2]
                d_users = cur_iv["users"] - prev_iv["users"]
                d_users_pct = (d_users / prev_iv["users"] * 100) if prev_iv["users"] else 0
                py, pmo = prev_iv["ym"]
                body.append(
                    f"전월({py}년 {pmo}월) 대비 방문자수는 {_ops_fmt_pct(d_users_pct)} "
                    f"({d_users:+,.0f}명) {'증가' if d_users >= 0 else '감소'}했습니다."
                )
            st.markdown(" ".join(body), unsafe_allow_html=True)

        if not ga_channel_inflow.empty:
            gc = ga_channel_inflow.copy()
            gc["report_date"] = pd.to_datetime(gc["report_date"]).dt.date
            max_d = gc["report_date"].max()
            min_d = max_d - timedelta(days=29)
            recent = gc[gc["report_date"] >= min_d]
            top_src = (
                recent.groupby("source_medium", as_index=False)["users"].sum()
                .sort_values("users", ascending=False).head(3)
            )
            if not top_src.empty:
                names = ", ".join(f"{r.source_medium}({r.users:,.0f}명)" for r in top_src.itertuples())
                st.markdown(f"최근 30일 기준 유입 상위 소스/매체는 {names} 순입니다.")
                st.markdown(
                    _ops_next_action(
                        f"상위 소스인 {top_src.iloc[0]['source_medium']} 유입을 다음 주에도 "
                        "유지되는지 확인하고, 유입은 있는데 구매로 안 이어지는 소스/매체가 있으면 랜딩페이지·"
                        "타겟팅을 점검하는 것을 권장합니다."
                    ),
                    unsafe_allow_html=True,
                )
            if recent["channel"].isna().all():
                st.caption(
                    "※ '매체'(채널 그룹핑) 매핑이 아직 안 되어 있어 소스/매체 단위로만 코멘트했습니다. "
                    "매핑 완료되면 채널 단위 코멘트로 보강할 수 있습니다."
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
    inflow_revenue = load_table("inflow_revenue_daily")
    ga_channel_inflow = load_table("ga_channel_inflow")
    agency_notes = load_table("agency_notes")
    channels_weekly = load_table("channel_weekly")
    budget = load_table("channel_budget")

    if weekly.empty and monthly.empty:
        st.info("아직 저장된 데이터가 없습니다. 왼쪽 사이드바에서 주간 리포트 파일을 업로드하고 '전체 저장하기'를 눌러주세요.")
        return

    for df, col in [(weekly, "week_start"), (weekly, "week_end"), (monthly, "report_month"), (daily, "report_date")]:
        if not df.empty and col in df.columns:
            df[col] = pd.to_datetime(df[col]).dt.date

    page = render_nav()  # ← st.tabs() 대신 사이드바 그룹 네비게이션

    if page == "종합 대시보드":
        render_overview_page(weekly, monthly, daily, channels, channels_weekly, ga_channel_inflow)
    elif page == "매체별 성과":
        render_channel_page(channels, snapshot, ga_channel_inflow)
    elif page == "타겟팅별 성과":
        render_targeting_performance_page(audience, creatives_fallback=creatives)
    elif page == "소재별 성과":
        render_creative_performance(creatives)
    elif page == "예산 현황":
        render_budget_page(monthly, budget)
    elif page == "운영 코멘트":
        render_operation_comment_page(
            weekly, monthly, daily, channels, snapshot, creatives, audience,
            inflow_revenue, ga_channel_inflow, agency_notes,
        )
    elif page == "GA 매체별 유입 경로":
        render_ga_channel_inflow_page(ga_channel_inflow)
    elif page == "GA4 라이브 리포트":
        render_ga4_page()
    elif page == "유입·매출 비교":
        render_inflow_revenue_page(inflow_revenue, ga_channel_inflow)
    elif page in NAV_PAGES_COMING_SOON:
        render_coming_soon(page)


if __name__ == "__main__":
    main()
