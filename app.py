"""법인카드 사용 적정성 점검 PoC — 1단계: 데이터 읽기·전처리·월별 조회."""
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from src.loader import load_all, monthly_summary

DEFAULT_FILE = Path(__file__).parent / "data" / "법인카드_가상데이터_202509-202608.xlsx"

DISPLAY_COLUMNS = [
    "승인번호", "승인일시", "요일", "휴일여부", "공휴일명", "구분", "카드번호", "직원명",
    "예산부서명", "예산부서", "G/L 계정", "G/L 계정명", "가맹점업종", "업종코드", "가맹점명",
    "가맹점주소", "전표금액", "취소구분", "상계상태", "상계승인번호", "전표 품목 텍스트",
    "결재문서제목", "국내해외코드명", "회계일자",
]

st.set_page_config(page_title="법인카드 점검 PoC", layout="wide")
st.title("법인카드 사용 적정성 점검 PoC")
st.caption("1단계: 데이터 읽기 · 전처리 · 월별 조회  |  모든 데이터는 교육용 가상 데이터입니다.")


@st.cache_data(show_spinner="엑셀을 읽는 중…")
def load(source) -> pd.DataFrame:
    return load_all(source)


uploaded = st.sidebar.file_uploader("SAP 추출 엑셀 업로드 (.xlsx)", type="xlsx")
if uploaded is not None:
    df = load(uploaded)
    st.sidebar.success(f"업로드 파일 사용: {uploaded.name}")
elif DEFAULT_FILE.exists():
    df = load(str(DEFAULT_FILE))
    st.sidebar.info(f"기본 파일 사용: data/{DEFAULT_FILE.name}")
else:
    st.warning("data/ 폴더에 엑셀이 없습니다. 왼쪽에서 파일을 업로드하세요.")
    st.stop()

# ── 전체 요약 ──────────────────────────────────────────
total = len(df)
n_cancel = int((df["취소구분"] == "취소").sum())
n_offset = int((~df["분석대상"]).sum())
target = df[df["분석대상"]]

c1, c2, c3, c4 = st.columns(4)
c1.metric("전체 원본 건수", f"{total:,}건")
c2.metric("취소건", f"{n_cancel:,}건")
c3.metric("상계 제외 (매입+취소)", f"{n_offset:,}건")
c4.metric("분석 대상", f"{len(target):,}건", f"{target['전표금액'].sum():,.0f}원", delta_color="off")

unmatched = df[df["상계상태"] == "취소미매칭"]
if not unmatched.empty:
    st.warning(f"짝이 되는 매입건을 찾지 못한 취소건 {len(unmatched)}건이 있습니다. 아래 목록에서 '취소미매칭'을 확인하세요.")

# ── 월별 요약 ──────────────────────────────────────────
st.subheader("월별 건수 · 금액 요약")
summary = monthly_summary(df)
left, right = st.columns([3, 2])
with left:
    st.dataframe(
        summary.style.format({c: "{:,}" for c in summary.columns if c != "승인월"}),
        hide_index=True, width="stretch",
    )
with right:
    fig = px.bar(summary, x="승인월", y="분석대상 금액", text="분석대상 건수",
                 labels={"분석대상 금액": "금액(원)", "승인월": ""})
    fig.update_traces(texttemplate="%{text}건", textposition="outside")
    fig.update_layout(height=330, margin=dict(t=20, b=10), yaxis_tickformat=",")
    st.plotly_chart(fig, width="stretch")

# ── 거래 목록 ──────────────────────────────────────────
st.subheader("거래 목록")
months = sorted(df["승인월"].unique())
f1, f2, f3 = st.columns([2, 1, 1])
month = f1.selectbox("승인월", ["전체"] + months, index=len(months))  # 기본: 최근월(점검대상월)
only_target = f2.checkbox("분석 대상만 (상계건 제외)", value=True)
only_holiday = f3.checkbox("휴일 사용만", value=False)

view = df if month == "전체" else df[df["승인월"] == month]
if only_target:
    view = view[view["분석대상"]]
if only_holiday:
    view = view[view["휴일여부"]]

st.write(f"**{month}** · {len(view):,}건 · 합계 {view['전표금액'].sum():,.0f}원")
st.dataframe(
    view[DISPLAY_COLUMNS].sort_values("승인일시"),
    hide_index=True, width="stretch", height=520,
    column_config={
        "승인일시": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm:ss"),
        "회계일자": st.column_config.DateColumn(format="YYYY-MM-DD"),
        "전표금액": st.column_config.NumberColumn(format="localized"),
        "승인번호": st.column_config.NumberColumn(format="plain"),
        "상계승인번호": st.column_config.NumberColumn(format="plain"),
    },
)

with st.expander("취소 상계 내역 보기"):
    pairs = df[df["상계상태"] != "정상"][
        ["승인번호", "승인일시", "카드번호", "가맹점명", "전표금액", "취소구분", "상계상태", "상계승인번호"]
    ].sort_values(["카드번호", "가맹점명", "승인일시"])
    st.dataframe(pairs, hide_index=True, width="stretch")
