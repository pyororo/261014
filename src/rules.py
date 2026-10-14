"""2단계: 규칙 기반 선별 (CLAUDE.md 2-2 중 R01~R07·R09·R15) + 위험점수·등급 (2-4)."""
from __future__ import annotations

import re
from dataclasses import dataclass, field

import pandas as pd

RULE_SHEET = "점검기준(가상)"
ACCOUNT_SHEET = "계정코드"
INDUSTRY_SHEET = "업종코드"

# 2단계에서 구현하는 규칙
STAGE2_RULES = ["R01", "R02", "R03", "R04", "R05", "R06", "R07", "R09", "R15"]

# 접대비 적요의 상대방 표기: ㅇㅇㅇ, 직함, n인/n명
TITLES = ["대표", "사장", "회장", "부회장", "전무", "상무", "이사", "본부장", "지점장", "실장",
          "부장", "차장", "과장", "팀장", "대리", "원장", "국장", "센터장", "단장", "고객"]
COUNTERPART_RE = re.compile(r"ㅇㅇㅇ|\d+\s*[인명]|" + "|".join(TITLES))


@dataclass
class Params:
    """화면에서 바꿀 수 있는 기준값. 기본값은 점검기준(가상) 시트."""
    night_start: int = 23          # R02 심야 시작 시각(이상)
    night_end: int = 6             # R02 심야 종료 시각(미만)
    split_gap_min: int = 60        # R04 분할결제 간격(분)
    min_memo_len: int = 5          # R06 적요 최소 글자수
    settle_days: int = 60          # R15 정산 기한(일)
    weights: dict = field(default_factory=dict)


@dataclass
class RefTables:
    rules: pd.DataFrame            # 규칙ID, 점검항목, 위험가중치 …
    accounts: pd.DataFrame         # G/L 계정, 허용업종(set), 건당한도
    industries: pd.DataFrame       # 업종코드, 가맹점업종, 제한구분


def read_reference(source) -> RefTables:
    rules = pd.read_excel(source, sheet_name=RULE_SHEET, dtype=str)
    rules["위험가중치"] = pd.to_numeric(rules["위험가중치"], errors="coerce").fillna(0).astype(int)

    acc = pd.read_excel(source, sheet_name=ACCOUNT_SHEET, dtype=str)
    acc = acc.rename(columns={"건당 한도(원)": "건당한도", "허용 가맹점업종": "허용업종"})
    acc["건당한도"] = pd.to_numeric(acc["건당한도"])
    acc["허용업종"] = acc["허용업종"].fillna("").map(
        lambda s: {x.strip() for x in s.split(",") if x.strip()}
    )

    ind = pd.read_excel(source, sheet_name=INDUSTRY_SHEET, dtype=str)
    return RefTables(rules=rules, accounts=acc, industries=ind)


def default_params(ref: RefTables) -> Params:
    return Params(weights=dict(zip(ref.rules["규칙ID"], ref.rules["위험가중치"])))


def rule_names(ref: RefTables) -> dict:
    return dict(zip(ref.rules["규칙ID"], ref.rules["점검항목"]))


# ── 개별 규칙: 분석대상 df를 받아 같은 인덱스의 bool Series를 돌려준다 ──────────


def r01_holiday(df, p):
    return df["휴일여부"]


def r02_night(df, p):
    h = df["승인일시"].dt.hour
    return (h >= p.night_start) | (h < p.night_end)


def r03_restricted(df, p):
    return df["제한구분"].isin(["금지", "사전승인"])


def r04_split(df, p):
    """동일 카드+가맹점+승인일, 연속 간격 ≤ N분인 묶음이 2건 이상이고 합계 > 건당한도."""
    hit = pd.Series(False, index=df.index)
    gap = pd.Timedelta(minutes=p.split_gap_min)
    for _, g in df.groupby(["카드번호", "가맹점명", "승인일자"]):
        if len(g) < 2:
            continue
        g = g.sort_values("승인일시")
        cluster = (g["승인일시"].diff() > gap).cumsum()
        for _, c in g.groupby(cluster):
            limit = c["건당한도"].iloc[0]
            if len(c) >= 2 and pd.notna(limit) and c["전표금액"].sum() > limit:
                hit[c.index] = True
    return hit


def r05_over_limit(df, p):
    return df["건당한도"].notna() & (df["전표금액"] > df["건당한도"])


def r06_poor_memo(df, p):
    memo = df["전표 품목 텍스트"].fillna("").str.strip()
    too_short = memo.str.len() < p.min_memo_len
    is_entertain = df["G/L 계정명"].fillna("").str.contains("접대비")
    no_counterpart = ~memo.str.contains(COUNTERPART_RE)
    return too_short | (is_entertain & no_counterpart)


def r07_mismatch(df, p):
    def bad(row):
        allowed = row["허용업종"]
        return isinstance(allowed, set) and row["가맹점업종"] not in allowed
    return df.apply(bad, axis=1)


def r09_duplicate(df, p):
    # 취소로 상계된 건은 이미 분석대상에서 빠져 있으므로 남은 건끼리만 비교
    keys = ["카드번호", "가맹점명", "전표금액", "승인일자"]
    return df.duplicated(keys, keep=False)


def r15_late(df, p):
    return (df["회계일자"] - df["승인일자"]).dt.days > p.settle_days


RULE_FUNCS = {
    "R01": r01_holiday, "R02": r02_night, "R03": r03_restricted, "R04": r04_split,
    "R05": r05_over_limit, "R06": r06_poor_memo, "R07": r07_mismatch,
    "R09": r09_duplicate, "R15": r15_late,
}


def apply_rules(df: pd.DataFrame, ref: RefTables, p: Params) -> pd.DataFrame:
    """분석대상 거래에 규칙을 적용해 규칙별 플래그, 걸린 규칙, 위험점수, 등급을 붙인다."""
    base = df[df["분석대상"]].copy()
    base = base.merge(ref.accounts[["G/L 계정", "허용업종", "건당한도"]], on="G/L 계정", how="left")
    base = base.merge(ref.industries[["업종코드", "제한구분"]], on="업종코드", how="left")
    base.index = df.index[df["분석대상"]]

    names = rule_names(ref)
    for rid in STAGE2_RULES:
        base[rid] = RULE_FUNCS[rid](base, p).astype(bool)

    flags = base[STAGE2_RULES]
    base["걸린 규칙"] = flags.apply(
        lambda r: " / ".join(f"{rid} {names.get(rid, '')}" for rid in STAGE2_RULES if r[rid]), axis=1
    )
    base["걸린 규칙수"] = flags.sum(axis=1)
    base["위험점수"] = sum(flags[rid] * p.weights.get(rid, 0) for rid in STAGE2_RULES).astype(int)
    base["금지업종"] = base["R03"] & (base["제한구분"] == "금지")
    base["등급"] = base.apply(_grade, axis=1)
    return base


def _grade(row) -> str:
    if row["금지업종"]:
        return "상"
    s = row["위험점수"]
    if s >= 6:
        return "상"
    if s >= 3:
        return "중"
    if s >= 1:
        return "하"
    return ""


def select_for_month(scored: pd.DataFrame, month: str) -> pd.DataFrame:
    """점검대상월 선별 목록: 그 달 승인 건 + 그 달에 회계처리된 정산지연(R15) 건."""
    acct_month = scored["회계일자"].dt.strftime("%Y-%m")
    in_month = (scored["승인월"] == month) | (scored["R15"] & (acct_month == month))
    out = scored[in_month & (scored["위험점수"] > 0)].copy()
    out["_등급순"] = out["등급"].map({"상": 0, "중": 1, "하": 2})
    out = out.sort_values(["_등급순", "위험점수", "승인일시"], ascending=[True, False, True])
    return out.drop(columns="_등급순")
