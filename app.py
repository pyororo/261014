"""법인카드 사용 적정성 점검 PoC — 1단계 데이터 정리 · 2단계 규칙 선별 · 3단계 AI 코멘트 · 4단계 기간 분석."""
import io
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from src.ai import (DISCLAIMER, _sabun, annotate, claude_available, claude_comment, claude_recommend,
                    read_workplaces)
from src.loader import load_all, monthly_summary, read_holidays
from src.patterns import TParams, aggregate, detect_all, monthly_by_user, prepare, recommend
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
st.caption("1단계 데이터 정리 · 2단계 규칙 기반 선별 · 3단계 AI 코멘트 · 4단계 기간 분석  |  모든 데이터는 교육용 가상 데이터입니다.")


def _src(source):
    return io.BytesIO(source) if isinstance(source, bytes) else source


@st.cache_data(show_spinner="엑셀을 읽는 중…")
def load(source):
    return (load_all(_src(source)), read_reference(_src(source)), read_workplaces(_src(source)),
            set(read_holidays(_src(source))["일자"]))


uploaded = st.sidebar.file_uploader("SAP 추출 엑셀 업로드 (.xlsx)", type="xlsx")
if uploaded is not None:
    df, ref, workplaces, holidays = load(uploaded.getvalue())
    st.sidebar.success(f"업로드 파일 사용: {uploaded.name}")
elif DEFAULT_FILE.exists():
    df, ref, workplaces, holidays = load(str(DEFAULT_FILE))
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

tab_sel, tab_trend, tab_tx = st.tabs(["점검대상 선별", "기간 분석", "거래 조회"])

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

# ══ 4단계: 기간 분석 ═════════════════════════════════════════
# 부서 8개 = 범주형 팔레트 8색, 부서 이름 순서로 고정 배정(필터로 색이 바뀌지 않게)
SERIES = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
DEPT_COLORS = {d: SERIES[i % len(SERIES)] for i, d in enumerate(sorted(df["예산부서명"].dropna().unique()))}
BASE_LAYOUT = dict(margin=dict(t=30, b=10, l=10, r=10), yaxis_tickformat=",", hovermode="x unified",
                   xaxis=dict(showgrid=False), yaxis=dict(gridcolor="rgba(128,128,128,0.2)"))


def _months_axis(fig):
    fig.update_xaxes(type="category", tickangle=0, labelalias={m: f"{m[2:4]}.{m[5:]}" for m in map(str, months)})
    return fig


with tab_trend:
    st.info("단월 점검에서는 건별로 한도 이내라 드러나지 않는 반복·급증·집중 패턴입니다.  \n※ " + DISCLAIMER
            + (" (현재: 규칙 기반 템플릿 문장)" if not claude_available() else ""))
    tp_base = TParams()
    with st.sidebar.expander("기간 패턴 기준값 설정", expanded=False):
        tparams = TParams(
            repeat_months=st.number_input("T-R1 기간(개월)", 1, 12, tp_base.repeat_months),
            repeat_count=st.number_input("T-R1 반복 건수 기준", 2, 50, tp_base.repeat_count),
            surge_ratio=st.number_input("T-R2 급증 배수", 1.0, 10.0, tp_base.surge_ratio, step=0.5),
            month_end_share=st.slider("T-R3 월말 5영업일 비중", 0.1, 0.9, tp_base.month_end_share, 0.05),
            dept_ratio=st.number_input("T-R3 부서 월평균 배수", 1.0, 5.0, tp_base.dept_ratio, step=0.1),
            fixed_months=st.number_input("T-R4 연속 개월 수", 2, 12, tp_base.fixed_months),
        )

    d = prepare(df)
    c1, c2 = st.columns([1, 2])
    end_month = pd.Period(c1.selectbox("기준월", months, index=len(months) - 1, key="trend_end"), "M")
    span = c2.radio("조회 기간", [3, 6, 12], index=2, horizontal=True, format_func=lambda n: f"최근 {n}개월")
    period = [m for m in pd.period_range(end=end_month, periods=span, freq="M") if m in set(d["월"])]
    dp = d[d["월"].isin(period)]

    patterns = detect_all(d, end_month, holidays, tparams)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric(f"조회 기간 ({period[0].strftime('%y.%m')}~{period[-1].strftime('%y.%m')})", f"{len(period)}개월")
    k2.metric("사용액", f"{dp['전표금액'].sum():,.0f}원")
    k3.metric("건수", f"{len(dp):,}건")
    k4.metric("감지된 기간 패턴", f"{len(patterns)}건")

    # ── 추이 ──
    monthly = dp.groupby("월").agg(금액=("전표금액", "sum"), 건수=("전표금액", "size")).reset_index()
    monthly["월"] = monthly["월"].astype(str)
    monthly["3개월 이동평균"] = monthly["금액"].rolling(3, min_periods=1).mean()
    left, right = st.columns(2)
    with left:
        st.markdown("**월별 사용액** (막대) · 3개월 이동평균 (선)")
        fig = go.Figure()
        fig.add_bar(x=monthly["월"], y=monthly["금액"], name="월 사용액", marker_color=SERIES[0],
                    customdata=monthly["건수"], hovertemplate="%{y:,.0f}원 · %{customdata}건<extra></extra>")
        fig.add_scatter(x=monthly["월"], y=monthly["3개월 이동평균"], name="3개월 이동평균", mode="lines",
                        line=dict(color="#6b6b6b", width=2), hovertemplate="이동평균 %{y:,.0f}원<extra></extra>")
        fig.update_layout(height=320, legend=dict(orientation="h", y=1.12), **BASE_LAYOUT)
        st.plotly_chart(_months_axis(fig), width="stretch")
    with right:
        st.markdown("**부서별 월 사용액**")
        dept_m = dp.groupby(["월", "예산부서명"])["전표금액"].sum().reset_index()
        dept_m["월"] = dept_m["월"].astype(str)
        fig = px.line(dept_m, x="월", y="전표금액", color="예산부서명", markers=True,
                      color_discrete_map=DEPT_COLORS, labels={"전표금액": "", "월": "", "예산부서명": "부서"})
        fig.update_traces(line=dict(width=2), marker=dict(size=8), hovertemplate="%{y:,.0f}원")
        fig.update_layout(height=320, legend=dict(orientation="h", y=-0.2, title=None), **BASE_LAYOUT)
        st.plotly_chart(_months_axis(fig), width="stretch")

    # ── 패턴 목록 ──
    st.subheader(f"기간 패턴 (T-R1~T-R4, 기준월 {end_month})")
    if patterns.empty:
        st.write("기준을 넘는 패턴이 없습니다.")
    else:
        st.dataframe(
            patterns.drop(columns=["_가중"]), hide_index=True, width="stretch",
            column_config={c: st.column_config.TextColumn(width="large")
                           for c in ["지표", "AI 해석", "AI 추가 확인사항", "관련 승인번호"]},
        )
        for _, pt in patterns.iterrows():
            with st.expander(f"{pt['No']}. {pt['패턴']} · {pt['대상']}"):
                st.markdown(f"**AI 해석**  \n{pt['AI 해석']}")
                st.markdown("**AI 추가 확인사항**  \n" + "\n".join(f"- {x}" for x in pt["AI 추가 확인사항"].split(" / ")))
                ids = [int(i) for i in pt["관련 승인번호"].split(", ")]
                user = pt["대상"].split(" (")[0]
                if pt["패턴"].startswith("T-R2"):
                    s_ = monthly_by_user(d).loc[user].reindex(pd.period_range(end=end_month, periods=12, freq="M"), fill_value=0)
                    recent = s_.index >= end_month - (tparams.surge_recent - 1)
                    fig = go.Figure(go.Bar(x=s_.index.astype(str), y=s_.values,
                                           marker_color=[SERIES[1] if r else "#9aa0a6" for r in recent],
                                           hovertemplate="%{y:,.0f}원<extra></extra>"))
                    fig.update_layout(height=260, title=f"{user} 월별 사용액 (주황 = 최근 {tparams.surge_recent}개월)",
                                      **BASE_LAYOUT)
                    st.plotly_chart(_months_axis(fig), width="stretch")
                else:
                    x = d[d["승인번호"].isin(ids)].sort_values("승인일시")
                    if pt["패턴"].startswith(("T-R1", "T-R4")):
                        x = d[(d["사용자"] == user) & (d["가맹점명"] == pt["관련가맹점"])]
                        cnt = x.groupby("월").agg(건수=("전표금액", "size"), 금액=("전표금액", "sum"))
                        cnt = cnt.reindex(pd.period_range(end=end_month, periods=12, freq="M"), fill_value=0)
                        fig = go.Figure(go.Bar(x=cnt.index.astype(str), y=cnt["건수"], marker_color=SERIES[0],
                                               customdata=cnt["금액"],
                                               hovertemplate="%{y}건 · %{customdata:,.0f}원<extra></extra>"))
                        fig.update_layout(height=260, title=f"{user} · {pt['관련가맹점']} 월별 이용 건수", **BASE_LAYOUT)
                        st.plotly_chart(_months_axis(fig), width="stretch")
                    st.dataframe(x[["승인번호", "승인일시", "요일", "사용자", "가맹점명", "G/L 계정명", "전표금액",
                                    "전표 품목 텍스트"]].sort_values("승인일시"),
                                 hide_index=True, width="stretch",
                                 column_config={"승인번호": st.column_config.NumberColumn(format="plain"),
                                                "전표금액": st.column_config.NumberColumn(format="localized"),
                                                "승인일시": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm")})

    # ── 공통 집계 ──
    st.subheader("공통 집계")
    dims = {"직원(사용자)": "사용자", "부서": "예산부서명", "계정": "G/L 계정명", "업종": "가맹점업종"}
    dim_label = st.radio("집계 기준", list(dims), horizontal=True, key="agg_dim")
    agg = aggregate(d, dims[dim_label], period)
    money = [c for c in agg.columns if "금액" in c or "평균" in c or "합계" in c]
    st.dataframe(agg, hide_index=True, width="stretch",
                 column_config={**{c: st.column_config.NumberColumn(format="localized") for c in money},
                                "전월 대비": st.column_config.NumberColumn(format="percent"),
                                "전년동월 대비": st.column_config.NumberColumn(format="percent")})
    st.caption("전년동월 대비는 13개월 이상 데이터가 쌓이면 계산됩니다(현재 12개월).")

    # ── AI 추가 점검항목 추천 ──
    st.subheader("AI 추천 — 기존 점검기준 외 추가 검토 항목")
    recs = recommend(d, scored, patterns)
    st.dataframe(recs, hide_index=True, width="stretch",
                 column_config={"근거 및 기대효과": st.column_config.TextColumn(width="large")})
    if claude_available():
        if st.button("Claude에게 추가 점검항목 추천 받기", key="claude_rec"):
            summary_text = ("[월별 집계]\n" + monthly.to_string(index=False)
                            + "\n\n[부서별 집계]\n" + aggregate(d, "예산부서명", period).to_string(index=False)
                            + "\n\n[감지된 기간 패턴]\n"
                            + (patterns[["패턴", "대상", "지표"]].to_string(index=False) if not patterns.empty else "없음")
                            + f"\n\n[{end_month} 선별 결과]\n"
                            + select_for_month(scored, str(end_month))["걸린 규칙"].value_counts().head(15).to_string())
            with st.spinner("Claude가 집계표를 검토 중…"):
                try:
                    st.session_state["claude_recs"] = claude_recommend(summary_text)
                except Exception as e:
                    st.error(f"Claude 호출 실패: {e}")
        if st.session_state.get("claude_recs"):
            st.dataframe(pd.DataFrame(st.session_state["claude_recs"]), hide_index=True, width="stretch")
    else:
        st.caption("규칙 기반 추천입니다. ANTHROPIC_API_KEY를 설정하면 Claude에게 집계표를 주고 추천을 받을 수 있습니다.")

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
