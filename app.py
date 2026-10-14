"""법인카드 사용 적정성 점검 PoC — 1단계(데이터 정리) + 2단계(규칙 기반 선별) + 3단계(AI 코멘트)."""
import io
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from src.ai import DISCLAIMER, _sabun, annotate, claude_available, claude_comment, read_workplaces
from src.loader import load_all, monthly_summary
from src.rules import (STAGE2_RULES, Params, apply_rules, default_params, read_reference,
                       rule_names, select_for_month)

DEFAULT_FILE = Path(__file__).parent / "data" / "법인카드_가상데이터_202509-202608.xlsx"

DISPLAY_COLUMNS = [
    "승인번호", "승인일시", "요일", "휴일여부", "공휴일명", "구분", "카드번호", "직원명",
    "예산부서명", "예산부서", "G/L 계정", "G/L 계정명", "가맹점업종", "업종코드", "가맹점명",
    "가맹점주소", "전표금액", "취소구분", "상계상태", "상계승인번호", "전표 품목 텍스트",
    "결재문서제목", "국내해외코드명", "회계일자",
]

SELECT_COLUMNS = [
    "등급", "위험점수", "걸린 규칙", "AI 선별 사유", "AI 추가 확인사항", "승인번호", "승인일시", "직원명", "예산부서명", "구분",
    "가맹점명", "가맹점업종", "G/L 계정명", "전표금액", "전표 품목 텍스트", "결재문서제목", "회계일자",
]

AI_NOTICE = ("※ 선별 결과는 점검이 필요한 거래를 고르기 위한 참고자료이며, "
             "부적정 여부의 최종 판단은 감사부 담당자가 증빙 확인 후 내립니다.")

st.set_page_config(page_title="법인카드 점검 PoC", layout="wide")
st.title("법인카드 사용 적정성 점검 PoC")
st.caption("1단계 데이터 정리 · 2단계 규칙 기반 선별 · 3단계 AI 코멘트  |  모든 데이터는 교육용 가상 데이터입니다.")


def _src(source):
    return io.BytesIO(source) if isinstance(source, bytes) else source


@st.cache_data(show_spinner="엑셀을 읽는 중…")
def load(source):
    return load_all(_src(source)), read_reference(_src(source)), read_workplaces(_src(source))


uploaded = st.sidebar.file_uploader("SAP 추출 엑셀 업로드 (.xlsx)", type="xlsx")
if uploaded is not None:
    df, ref, workplaces = load(uploaded.getvalue())
    st.sidebar.success(f"업로드 파일 사용: {uploaded.name}")
elif DEFAULT_FILE.exists():
    df, ref, workplaces = load(str(DEFAULT_FILE))
    st.sidebar.info(f"기본 파일 사용: data/{DEFAULT_FILE.name}")
else:
    st.warning("data/ 폴더에 엑셀이 없습니다. 왼쪽에서 파일을 업로드하세요.")
    st.stop()

# ── 기준값 설정 (점검기준(가상) 시트 기본값, 화면에서 변경 가능) ───────────
base = default_params(ref)
names = rule_names(ref)
with st.sidebar.expander("점검 기준값 설정", expanded=False):
    night = st.slider("R02 정상 사용 시간대 (이 범위 밖이 심야)", 0, 24, (base.night_end, base.night_start),
                      help="기본: 06시 이상 23시 미만이 정상, 그 외(23시~06시)가 심야")
    split_gap = st.number_input("R04 분할결제 간격(분)", 1, 1440, base.split_gap_min)
    memo_len = st.number_input("R06 적요 최소 글자수", 1, 50, base.min_memo_len)
    settle = st.number_input("R15 정산 기한(일)", 1, 365, base.settle_days)
    st.markdown("**위험 가중치**")
    weights = dict(base.weights)
    for rid in STAGE2_RULES + ["R13"]:
        weights[rid] = st.number_input(f"{rid} {names.get(rid, '')}", 0, 10, int(base.weights.get(rid, 0)), key=f"w_{rid}")

params = Params(night_start=night[1], night_end=night[0], split_gap_min=int(split_gap),
                min_memo_len=int(memo_len), settle_days=int(settle), weights=weights)
scored = annotate(apply_rules(df, ref, params), workplaces, params.weights, names)
ALL_RULES = STAGE2_RULES[:-1] + ["R13", "R15"]
months = sorted(df["승인월"].unique())

tab_sel, tab_tx = st.tabs(["점검대상 선별", "거래 조회"])

# ══ 2단계: 점검대상 선별 ═════════════════════════════════════
with tab_sel:
    st.info(AI_NOTICE + "  \n※ " + DISCLAIMER + (" (현재: Claude API 키가 없어 규칙 기반 템플릿 문장으로 생성)"
                                                if not claude_available() else ""))
    sel_month = st.selectbox("점검대상월", months, index=len(months) - 1, key="sel_month")
    sel = select_for_month(scored, sel_month)
    n_month = int((scored["승인월"] == sel_month).sum())

    grade_cnt = sel["등급"].value_counts()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric(f"{sel_month} 분석대상", f"{n_month:,}건")
    m2.metric("선별 건수", f"{len(sel):,}건")
    m3.metric("등급 상 / 중 / 하",
              f"{grade_cnt.get('상', 0)} / {grade_cnt.get('중', 0)} / {grade_cnt.get('하', 0)}")
    m4.metric("선별 금액", f"{sel['전표금액'].sum():,.0f}원")
    st.caption("선별 대상 = 해당 월 승인 건 중 규칙에 1개 이상 걸린 건 + 해당 월에 회계처리된 정산지연(R15) 건. "
               "등급: R03 금지업종이면 무조건 '상', 그 외 점수 6↑ 상 · 3~5 중 · 1~2 하. "
               "R13(의미 불일치)은 AI 판단, 사전승인 문서번호가 적힌 휴일·심야 건은 AI가 저위험으로 조정")

    g1, g2 = st.columns([1, 3])
    grades = g1.multiselect("등급", ["상", "중", "하"], default=["상", "중", "하"])
    rule_filter = g2.multiselect("규칙", ALL_RULES, format_func=lambda r: f"{r} {names.get(r, '')}")
    view = sel[sel["등급"].isin(grades)]
    if rule_filter:
        view = view[view[rule_filter].any(axis=1)]

    st.caption("행을 클릭하면 아래에 건별 상세와 AI 코멘트가 나옵니다.")
    event = st.dataframe(
        view[SELECT_COLUMNS], hide_index=True, width="stretch", height=480,
        on_select="rerun", selection_mode="single-row", key="sel_table",
        column_config={
            "승인일시": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm"),
            "회계일자": st.column_config.DateColumn(format="YYYY-MM-DD"),
            "전표금액": st.column_config.NumberColumn(format="localized"),
            "승인번호": st.column_config.NumberColumn(format="plain"),
            "걸린 규칙": st.column_config.TextColumn(width="medium"),
            "AI 선별 사유": st.column_config.TextColumn(width="large"),
            "AI 추가 확인사항": st.column_config.TextColumn(width="large"),
        },
    )

    rows = event.selection.rows if event is not None else []
    if rows:
        r = view.iloc[rows[0]]
        st.subheader(f"건별 상세 · 승인번호 {r['승인번호']}")
        d1, d2 = st.columns([2, 3])
        with d1:
            detail = {k: r[k] for k in ["등급", "위험점수", "규칙점수", "AI 조정", "승인일시", "직원명", "예산부서명",
                                        "구분", "카드번호", "가맹점명", "가맹점업종", "가맹점주소", "G/L 계정명",
                                        "전표금액", "전표 품목 텍스트", "결재문서제목", "회계일자"]}
            st.dataframe(pd.Series(detail, name="값").astype(str), width="stretch")
        with d2:
            st.markdown(f"**걸린 규칙**  \n{r['걸린 규칙']}")
            st.markdown(f"**AI 선별 사유** ({r['AI 생성방식']})  \n{r['AI 선별 사유']}")
            st.markdown("**AI 추가 확인사항**")
            st.markdown("\n".join(f"- {x}" for x in r["AI 추가 확인사항"].split("; ") if x))
            if r["AI 조정 사유"]:
                st.success(f"AI 조정: {r['AI 조정 사유']} (규칙점수 {r['규칙점수']} → {r['위험점수']})")
            if r["R13"]:
                st.warning(f"R13 의미 불일치: {r['R13 사유']}")
            if claude_available():
                if st.button("Claude로 코멘트 다시 생성", key=f"claude_{r['승인번호']}"):
                    with st.spinner("Claude가 검토 중…"):
                        try:
                            st.session_state[f"cc_{r['승인번호']}"] = claude_comment(
                                r.to_dict(), workplaces.get(_sabun(r.get("참조"))))
                        except Exception as e:  # 네트워크·키 오류는 화면에 안내만
                            st.error(f"Claude 호출 실패: {e}")
                cc = st.session_state.get(f"cc_{r['승인번호']}")
                if cc:
                    st.markdown(f"**Claude 선별 사유**  \n{cc['선별사유']}")
                    st.markdown("\n".join(f"- {x}" for x in cc["추가확인사항"]))
                    if cc["의미불일치"]:
                        st.warning(f"Claude 의미 불일치 판단: {cc['의미불일치_사유']}")
                    if cc["저위험_조정권고"]:
                        st.success(f"Claude 저위험 조정 권고: {cc['조정_사유']}")
            else:
                st.caption("ANTHROPIC_API_KEY를 설정하면 이 건을 Claude에게 다시 검토받을 수 있습니다.")
            st.caption(DISCLAIMER)

    st.markdown("**규칙별 적중 건수** (해당 월 선별 기준)")
    hits = pd.DataFrame({
        "규칙ID": ALL_RULES,
        "점검항목": [names.get(r, "") for r in ALL_RULES],
        "가중치": [params.weights.get(r, 0) for r in ALL_RULES],
        "적중 건수": [int(sel[r].sum()) for r in ALL_RULES],
    })
    st.dataframe(hits, hide_index=True)

# ══ 1단계: 거래 조회 ═════════════════════════════════════════
with tab_tx:
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

    st.subheader("거래 목록")
    f1, f2, f3 = st.columns([2, 1, 1])
    month = f1.selectbox("승인월", ["전체"] + months, index=len(months), key="tx_month")  # 기본: 최근월
    only_target = f2.checkbox("분석 대상만 (상계건 제외)", value=True)
    only_holiday = f3.checkbox("휴일 사용만", value=False)

    tx = df if month == "전체" else df[df["승인월"] == month]
    if only_target:
        tx = tx[tx["분석대상"]]
    if only_holiday:
        tx = tx[tx["휴일여부"]]

    st.write(f"**{month}** · {len(tx):,}건 · 합계 {tx['전표금액'].sum():,.0f}원")
    st.dataframe(
        tx[DISPLAY_COLUMNS].sort_values("승인일시"),
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
