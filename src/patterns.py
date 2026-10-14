"""4단계: 기간(시계열) 패턴 T-R1~T-R4 + 공통 집계 + AI 추가 점검항목 추천 (CLAUDE.md 2-3, 2-5)."""
from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


@dataclass
class TParams:
    repeat_months: int = 3       # T-R1 기간(개월)
    repeat_count: int = 8        # T-R1 건수 기준
    surge_ratio: float = 2.0     # T-R2 배수
    surge_recent: int = 3        # T-R2 최근 기간(개월)
    surge_base: int = 6          # T-R2 비교 기간(개월)
    month_end_days: int = 5      # T-R3 월말 영업일 수
    month_end_share: float = 0.5 # T-R3 월말 비중
    dept_ratio: float = 1.3      # T-R3 부서 월평균 배수
    fixed_months: int = 3        # T-R4 연속 개월 수


def prepare(df: pd.DataFrame) -> pd.DataFrame:
    """분석대상만 남기고 기간 분석용 컬럼(사용자, 월)을 붙인다. 공용카드는 '(부서공용) 부서명'."""
    d = df[df["분석대상"]].copy()
    d["사용자"] = d["직원명"].fillna("(부서공용) " + d["예산부서명"])
    d["월"] = d["승인일자"].dt.to_period("M")
    return d


def _ids(x: pd.DataFrame) -> str:
    return ", ".join(str(i) for i in x.sort_values("승인일시")["승인번호"])


def _won(x) -> str:
    return f"{int(round(x)):,}원"


# ── 패턴 규칙 ──────────────────────────────────────────────


def t_r1_repeat(d: pd.DataFrame, end: pd.Period, p: TParams) -> pd.DataFrame:
    """동일 사용자·가맹점, 최근 N개월 건수 ≥ 기준."""
    start = end - (p.repeat_months - 1)
    w = d[(d["월"] >= start) & (d["월"] <= end)]
    rows = []
    for (user, mer), x in w.groupby(["사용자", "가맹점명"]):
        if len(x) < p.repeat_count:
            continue
        wd = x["요일"].value_counts()
        evening = (x["승인일시"].dt.hour >= 17).sum()
        concentrated = wd.iloc[0] >= len(x) * 0.7
        memo = x["전표 품목 텍스트"].fillna("").value_counts()
        rows.append({
            "패턴": "T-R1 동일가맹점 반복", "대상": f"{user} ({x['예산부서명'].iloc[0]})", "관련가맹점": mer,
            "기간": f"{start}~{end}",
            "지표": f"{p.repeat_months}개월 {len(x)}회, 합계 {_won(x['전표금액'].sum())}, {wd.index[0]}요일 {wd.iloc[0]}회",
            "기준": f"{p.repeat_months}개월 {p.repeat_count}회 이상",
            "AI 해석": (f"{p.repeat_months}개월 동안 같은 가맹점을 {len(x)}회 이용"
                      + (f", 그중 {wd.iloc[0]}회가 {wd.index[0]}요일" if concentrated else "")
                      + (f"이고 {evening}회가 저녁 시간대" if evening >= len(x) * 0.7 else "")
                      + f". 주된 적요는 '{memo.index[0]}'"
                      + ("로 상대방이 특정되지 않음" if "ㅇㅇㅇ" not in memo.index[0] else "")
                      + ". 건당 금액은 한도 이내라 단월 점검에서는 드러나지 않는 패턴."),
            "AI 추가 확인사항": "접대 상대방 명단 제출 요청 / 가맹점과 사용자의 특수관계(가족 운영 등) 여부 / 자택 인근 여부",
            "관련 승인번호": _ids(x), "_가중": concentrated,
        })
    return pd.DataFrame(rows)


def monthly_by_user(d: pd.DataFrame) -> pd.DataFrame:
    return d.groupby(["사용자", "월"])["전표금액"].sum().unstack(fill_value=0)


def t_r2_surge(d: pd.DataFrame, end: pd.Period, p: TParams) -> pd.DataFrame:
    """최근 N개월 월평균 ÷ 직전 M개월 월평균 > 배수."""
    ms = monthly_by_user(d)
    recent = [end - i for i in range(p.surge_recent)]
    base = [end - i for i in range(p.surge_recent, p.surge_recent + p.surge_base)]
    if any(m not in ms.columns for m in recent + base):
        return pd.DataFrame()
    r, b = ms[recent].mean(axis=1), ms[base].mean(axis=1)
    rows = []
    for user in ms.index[(b > 0) & (r / b.where(b > 0) > p.surge_ratio)]:
        x = d[(d["사용자"] == user) & d["월"].isin(recent)]
        acc = x.groupby("G/L 계정명")["전표금액"].sum().sort_values(ascending=False)
        top_mer = x["가맹점명"].value_counts()
        rows.append({
            "패턴": "T-R2 사용액 급증", "대상": f"{user} ({x['예산부서명'].iloc[0]})", "관련가맹점": "-",
            "기간": f"직전 {p.surge_base}개월 vs {recent[-1]}~{end}",
            "지표": f"월평균 {_won(b[user])} → {_won(r[user])} ({r[user] / b[user]:.1f}배)",
            "기준": f"{p.surge_ratio:g}배 초과",
            "AI 해석": (f"최근 {p.surge_recent}개월 사용액이 직전 평균의 {r[user] / b[user]:.1f}배. "
                      f"증가분의 {acc.iloc[0] / acc.sum():.0%}가 {acc.index[0]}이고, 가맹점 {len(top_mer)}곳에 분산. "
                      "건별로는 한도 이내라 개별 규칙에 거의 걸리지 않음. 담당 업무·거래처 변동이 없다면 이례적."),
            "AI 추가 확인사항": "담당 채널·영업목표 변경 여부 / 증가분의 상대방 분포(특정 거래처 집중 여부) / 부서장 사전 인지 여부",
            "관련 승인번호": _ids(x), "_가중": False,
        })
    return pd.DataFrame(rows)


def _last_business_days(month: pd.Period, holidays: set, n: int) -> list:
    days = pd.bdate_range(month.start_time, month.end_time)
    return [x for x in days if x.normalize() not in holidays][-n:]


def t_r3_month_end(d: pd.DataFrame, holidays: set, p: TParams) -> pd.DataFrame:
    """분기말 월, 부서별 월말 N영업일 비중 > 기준 AND 월 사용액 ≥ 부서 월평균 × 배수."""
    dm = d.groupby(["예산부서명", "월"])["전표금액"].sum().unstack(fill_value=0)
    avg = dm.mean(axis=1)
    rows = []
    for m in [c for c in dm.columns if c.month in (3, 6, 9, 12)]:
        first = _last_business_days(m, holidays, p.month_end_days)[0]
        for dept in dm.index:
            x = d[(d["예산부서명"] == dept) & (d["월"] == m)]
            total = x["전표금액"].sum()
            tail = x[x["승인일자"] >= first]
            share = tail["전표금액"].sum() / total if total else 0
            if share > p.month_end_share and total >= avg[dept] * p.dept_ratio:
                acc = tail.groupby("G/L 계정명")["전표금액"].sum().sort_values(ascending=False)
                gift = tail[tail["가맹점업종"] == "상품권"]
                rows.append({
                    "패턴": "T-R3 연말·분기말 집중", "대상": dept,
                    "관련가맹점": ", ".join(tail["가맹점명"].value_counts().index[:3]), "기간": str(m),
                    "지표": (f"월말 {p.month_end_days}영업일 {_won(tail['전표금액'].sum())} / 월 {_won(total)} "
                           f"({share:.0%}), 부서 월평균의 {total / avg[dept]:.1f}배"),
                    "기준": f"분기말 월, 월말 비중 {p.month_end_share:.0%} 초과 & 부서 월평균 {p.dept_ratio:g}배 이상",
                    "AI 해석": (f"{m.month}월 마지막 {p.month_end_days}영업일에 사용액의 {share:.0%}가 몰림(주로 {acc.index[0]})"
                              + (f". 상품권 {_won(gift['전표금액'].sum())}이 '{gift['전표 품목 텍스트'].iloc[0]}'(으)로 처리됨"
                                 if len(gift) else "")
                              + ". 예산 소진 목적의 선구매" + (" 또는 상품권 현금화" if len(gift) else "") + " 가능성."),
                    "AI 추가 확인사항": "구매 물품 실물·재고 확인 / 상품권 수령자·사용처 / 해당 부서 연간 예산 잔액 추이",
                    "관련 승인번호": _ids(tail), "_가중": False,
                })
    return pd.DataFrame(rows)


def t_r4_fixed(d: pd.DataFrame, p: TParams) -> pd.DataFrame:
    """동일 사용자·가맹점·금액이 N개월 이상 연속 매월 발생."""
    rows = []
    for (user, mer, amt), x in d.groupby(["사용자", "가맹점명", "전표금액"]):
        months = sorted(x["월"].unique())
        if len(months) < p.fixed_months:
            continue
        best, run, run_start, best_range = 1, 1, months[0], (months[0], months[0])
        for a, b in zip(months, months[1:]):
            run, run_start = (run + 1, run_start) if b == a + 1 else (1, b)
            if run > best:
                best, best_range = run, (run_start, b)
        if best < p.fixed_months:
            continue
        last_day = (x["승인일자"] == x["승인일자"] + pd.offsets.MonthEnd(0)).sum()
        memo = x["전표 품목 텍스트"].fillna("").value_counts().index[0]
        rows.append({
            "패턴": "T-R4 정액 반복결제", "대상": f"{user} ({x['예산부서명'].iloc[0]})", "관련가맹점": mer,
            "기간": f"{best_range[0]}~{best_range[1]}",
            "지표": f"{best}개월 연속 매월 {_won(amt)} (총 {len(x)}회)", "기준": f"{p.fixed_months}개월 연속",
            "AI 해석": (f"{'매월 말일 ' if last_day >= len(x) * 0.5 else '매월 '}같은 가맹점에 같은 금액 결제, 적요는 '{memo}'. "
                      "실제 식사 시점과 결제 시점이 달라 사용 내역 확인이 어렵고 장부거래 가능성."),
            "AI 추가 확인사항": "가맹점 선결제 잔액·장부 사용내역 / 실제 식사 일자·참석자 / 월말 예산 맞추기 여부",
            "관련 승인번호": _ids(x), "_가중": False,
        })
    return pd.DataFrame(rows)


def detect_all(d: pd.DataFrame, end: pd.Period, holidays: set, p: TParams) -> pd.DataFrame:
    """T-R3은 분기말 월 기준이라 end까지의 전체 기간을 보고, 나머지는 end 기준 최근 기간을 본다."""
    upto = d[d["월"] <= end]
    parts = [t_r1_repeat(upto, end, p), t_r2_surge(upto, end, p), t_r3_month_end(upto, holidays, p), t_r4_fixed(upto, p)]
    out = pd.concat([x for x in parts if not x.empty], ignore_index=True)
    if out.empty:
        return out
    out.insert(0, "No", range(1, len(out) + 1))
    return out


# ── 공통 집계 ──────────────────────────────────────────────


def aggregate(d: pd.DataFrame, dim: str, months: list) -> pd.DataFrame:
    """차원별 월합계·건수·건당평균·3개월 이동평균·전월/전년동월 대비 증감률 (마지막 달 기준)."""
    g = d[d["월"].isin(months)].groupby([dim, "월"])["전표금액"].agg(["sum", "count"]).reset_index()
    amt = g.pivot(index=dim, columns="월", values="sum").reindex(columns=months, fill_value=0).fillna(0)
    cnt = g.pivot(index=dim, columns="월", values="count").reindex(columns=months, fill_value=0).fillna(0)
    last = months[-1]
    out = pd.DataFrame({
        f"{last} 금액": amt[last],
        f"{last} 건수": cnt[last].astype(int),
        "건당평균": (amt[last] / cnt[last].where(cnt[last] > 0)).round(0),
        "3개월 이동평균": amt[months[-3:]].mean(axis=1).round(0),
        "전월 대비": (amt[last] / amt[months[-2]].where(amt[months[-2]] > 0) - 1) if len(months) > 1 else None,
        f"기간 합계({len(months)}개월)": amt.sum(axis=1),
    })
    prev_year = last - 12
    out["전년동월 대비"] = float("nan") if prev_year not in set(d["월"]) else (
        amt[last] / d[d["월"] == prev_year].groupby(dim)["전표금액"].sum().reindex(out.index) - 1)
    return out.sort_values(f"기간 합계({len(months)}개월)", ascending=False).reset_index()


# ── AI 추가 점검항목 추천 (규칙 기반 대체) ─────────────────────


def recommend(d: pd.DataFrame, scored: pd.DataFrame, patterns: pd.DataFrame) -> pd.DataFrame:
    """월별 집계 숫자를 보고 기존 규칙에 없는 확인 항목을 제안한다 (키가 없을 때의 템플릿 버전)."""
    recs = []
    adj = scored[scored.get("AI 조정", 0) < 0] if "AI 조정" in scored else scored.iloc[0:0]
    if len(adj):
        recs.append(("주말 행사 '사전승인 결재번호' 실존 대조",
                     f"휴일·심야 사용 {len(adj)}건이 적요의 사전승인 번호로 소명 가능 처리됐지만, 번호 실존 여부는 "
                     "데이터만으로 확인 불가. 그룹웨어 결재 목록과 번호 대조를 정기 항목으로 추가 권장."))
    if not patterns.empty and (patterns["패턴"].str.startswith("T-R1")).any():
        recs.append(("반복 가맹점의 특수관계 확인",
                     "한 직원이 한 가맹점을 장기간 반복 이용하는 패턴이 있음. 가맹점 사업자와 임직원의 관계"
                     "(임직원 가족 사업장 목록 대조) 확인을 추가 권장."))
    taxi = d[(d["가맹점업종"] == "택시") & (d["승인일시"].dt.hour >= 22)]
    if len(taxi):
        recs.append(("22시 이후 택시 ↔ 근태(야근) 기록 대조",
                     f"22시 이후 택시 {len(taxi)}건은 현재 정상으로 보고 있으나, 야근 기록 없는 날의 심야 택시는 "
                     "사적사용 가능성. D-HR 연계 시 자동 대조 가능."))
    if not patterns.empty and (patterns["패턴"].str.startswith("T-R3")).any():
        recs.append(("부서 예산 집행률과 연말 집중 사용 연계",
                     "T-R3는 사용 비중만 봄. 분기말 직전 예산 잔액이 많은 부서의 분기말 사용을 우선 점검하면 효율적."))
    memo = d["전표 품목 텍스트"].fillna("").str.split().str.len()
    short = (memo <= 2).groupby(d["사용자"]).mean().sort_values(ascending=False)
    recs.append(("'사무용품·식대' 등 범용 적요 비율 모니터링",
                 f"적요가 두 단어 이하인 거래가 전체의 {(memo <= 2).mean():.0%}. 직원별 비율 상위는 "
                 + ", ".join(f"{u}({v:.0%})" for u, v in short.head(3).items())
                 + ". 증빙 관리가 느슨한 사용자를 조기 파악 가능."))
    return pd.DataFrame(recs, columns=["추천 점검항목", "근거 및 기대효과"]).assign(No=lambda x: range(1, len(x) + 1))[
        ["No", "추천 점검항목", "근거 및 기대효과"]]
