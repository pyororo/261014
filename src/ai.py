"""3단계: AI 선별 사유 · 추가 확인사항 · R13 의미 불일치 · 오탐(FP) 저위험 조정 (CLAUDE.md 2-5).

Claude API 키가 있으면 건별 코멘트를 Claude가 생성하고, 없으면 규칙 기반 템플릿 문장으로 대체한다.
R13 판단과 FP 조정은 전체 거래에 적용해야 하므로 기본은 규칙 기반 판단을 쓰고,
건별 상세 화면에서 Claude에게 다시 판단을 받아볼 수 있다.
"""
from __future__ import annotations

import os
import re

import pandas as pd

from src.rules import _grade

MODEL = "claude-opus-5-5"
EMPLOYEE_SHEET = "직원마스터"

# 결재문서제목에 특정 비용 항목이 적혀 있으면 그 항목에 맞는 업종이어야 함
TITLE_KEYWORDS = {
    "하이패스": {"주유소", "주차장"},
    "통행료": {"주유소", "주차장"},
    "주유": {"주유소"},
    "주차": {"주차장"},
    "택시": {"택시"},
    "항공": {"항공사"},
    "숙박": {"숙박"},
    "도서": {"서적"},
}
TRAVEL_WORDS = re.compile(r"출장|이동|택시|귀가|방문|외근|교통")
PREAPPROVAL_RE = re.compile(r"사전승인|GW-\d{6}-\d+")
SOFT_RULES = {"R01", "R02"}  # 사전승인 문서가 있으면 소명 가능한 규칙

DISCLAIMER = "AI 결과는 최종 판단이 아닌 참고자료입니다."


# ── 기준 데이터 ─────────────────────────────────────────────


def read_workplaces(source) -> dict:
    """사번 → 근무지 ('서울 중구' 형식)."""
    emp = pd.read_excel(source, sheet_name=EMPLOYEE_SHEET, dtype=str)
    return dict(zip(emp["사번"], emp["근무지"]))


def _sabun(ref_text) -> str | None:
    # 참조 컬럼 '20024/이누리' 에서 사번 추출
    if isinstance(ref_text, str) and "/" in ref_text:
        return ref_text.split("/")[0].strip()
    return None


def _area(addr) -> str:
    # '서울 중구 예시로 32' → '서울 중구'
    parts = str(addr).split()
    return " ".join(parts[:2]) if len(parts) >= 2 else str(addr)


# ── R13 의미 불일치 (규칙 기반 판단) ───────────────────────────


def r13_rule_check(row, workplaces: dict) -> str:
    """앞뒤가 안 맞으면 이유 문장을, 맞으면 빈 문자열을 돌려준다."""
    title = str(row.get("결재문서제목") or "")
    memo = str(row.get("전표 품목 텍스트") or "")
    industry = row.get("가맹점업종")

    for kw, ok in TITLE_KEYWORDS.items():
        if kw in title and industry not in ok:
            return f"결재문서제목('{title}')과 실제 결제 업종({industry})·적요('{memo}')가 서로 다름"

    if "여비교통비" in str(row.get("G/L 계정명")):
        workplace = workplaces.get(_sabun(row.get("참조")))
        same_area = workplace is not None and _area(row.get("가맹점주소")) == workplace
        if "회식" in memo or (same_area and industry == "일반음식점" and not TRAVEL_WORDS.search(memo)):
            where = f"근무지({workplace}) 안" if same_area else "출장 표시 없이"
            return f"{where}의 식사('{memo}')를 여비교통비로 처리 — 복리후생비 등으로 계정 오분류 의심"
    return ""


# ── 문장 생성 (템플릿) ─────────────────────────────────────────


def _won(x) -> str:
    return f"{int(x):,}원"


def reason_phrases(row) -> list[str]:
    p = []
    if row.get("R01"):
        day = row["공휴일명"] if isinstance(row.get("공휴일명"), str) else f"{row['요일']}요일"
        p.append(f"휴일({day}) 사용")
    if row.get("R02"):
        p.append(f"심야 {row['승인일시']:%H:%M} 승인")
    if row.get("R03"):
        p.append(f"{row['가맹점업종']}은(는) {row['제한구분']} 업종")
    if row.get("R04"):
        p.append(f"같은 날 같은 가맹점에서 나눠 결제해 합계가 계정 한도({_won(row['건당한도'])})를 넘음")
    if row.get("R05"):
        p.append(f"건당 한도 {_won(row['건당한도'])} 초과")
    if row.get("R06"):
        memo = str(row.get("전표 품목 텍스트") or "").strip()
        if not memo:
            p.append("적요 공란")
        elif len(memo) < 5:
            p.append(f"적요 부실('{memo}')")
        else:
            p.append(f"접대 상대방 미기재('{memo}')")
    if row.get("R07"):
        p.append(f"{row['G/L 계정명']} 계정에 허용되지 않은 업종({row['가맹점업종']})")
    if row.get("R09"):
        p.append("같은 날 같은 금액이 취소 없이 2번 승인")
    if row.get("R15"):
        days = (row["회계일자"] - row["승인일자"]).days
        p.append(f"승인 후 {days}일 지나 정산")
    if row.get("R13"):
        p.append(row["R13 사유"])
    return p


CHECK_ITEMS = {
    "R01": "휴일 사용 사유와 사전승인 문서번호 확인",
    "R02": "심야 사용 사유와 실제 사용자 확인",
    "R04": "나눠 결제한 사유, 실제 이용 인원과 영수증 확인",
    "R05": "한도 초과 사유와 사전 품의 여부 확인",
    "R07": "실제 사용 목적에 맞는 계정으로 재분류 검토",
    "R09": "카드사 매출전표로 이중 결제 여부 확인, 취소 누락이면 환수",
    "R15": "정산 지연 사유와 증빙 보관 여부 확인",
    "R13": "결재문서·영수증으로 실제 사용 내역 대조",
}


def check_items(row) -> list[str]:
    items = []
    for rid in ["R03", "R01", "R02", "R04", "R05", "R06", "R07", "R09", "R13", "R15"]:
        if not row.get(rid):
            continue
        if rid == "R03":
            items.append("금지업종 사용 경위 소명 및 환수 검토" if row["제한구분"] == "금지"
                         else "사전 품의(결재) 문서 확인")
        elif rid == "R06":
            items.append("접대 상대방(소속·직함·인원)과 목적 확인" if "접대비" in str(row["G/L 계정명"])
                         else "적요 보완 요청(사용 목적·참석자)")
        else:
            items.append(CHECK_ITEMS[rid])
    return list(dict.fromkeys(items))


# ── 전체 거래에 3단계 적용 ──────────────────────────────────────


def annotate(scored: pd.DataFrame, workplaces: dict, weights: dict, names: dict) -> pd.DataFrame:
    """R13·FP 조정을 반영해 위험점수·등급을 다시 계산하고, 사유·추가확인 문장을 붙인다."""
    df = scored.copy()
    df["R13 사유"] = df.apply(lambda r: r13_rule_check(r, workplaces), axis=1)
    df["R13"] = df["R13 사유"] != ""
    df["규칙점수"] = df["위험점수"] + df["R13"] * weights.get("R13", 0)

    # 오탐 조정: 걸린 규칙이 휴일·심야뿐이고 적요에 사전승인 문서번호가 있으면 규칙당 1점 감점(최소 1점)
    memo = df["전표 품목 텍스트"].fillna("")
    hit_cols = [c for c in ["R01", "R02", "R03", "R04", "R05", "R06", "R07", "R09", "R13", "R15"] if c in df]
    only_soft = df[hit_cols].sum(axis=1) == df[list(SOFT_RULES)].sum(axis=1)
    fp = memo.str.contains(PREAPPROVAL_RE) & only_soft & (df["규칙점수"] > 0)
    df["AI 조정"] = 0
    df.loc[fp, "AI 조정"] = -df.loc[fp, list(SOFT_RULES)].sum(axis=1)
    df["위험점수"] = (df["규칙점수"] + df["AI 조정"]).clip(lower=df["규칙점수"].clip(upper=1))
    df["AI 조정 사유"] = ""
    df.loc[fp, "AI 조정 사유"] = "적요에 사전승인 문서번호가 기재되어 소명 가능 → 저위험으로 조정"

    order = ["R01", "R02", "R03", "R04", "R05", "R06", "R07", "R09", "R13", "R15"]
    df["걸린 규칙"] = df.apply(
        lambda r: " / ".join(f"{rid} {names.get(rid, '')}" for rid in order if r.get(rid)), axis=1)
    df["등급"] = df.apply(_grade, axis=1)

    df["AI 선별 사유"] = df.apply(_reason_sentence, axis=1)
    df["AI 추가 확인사항"] = df.apply(lambda r: "; ".join(check_items(r)), axis=1)
    df["AI 생성방식"] = "규칙 기반 템플릿"
    return df


def _reason_sentence(row) -> str:
    phrases = reason_phrases(row)
    if not phrases:
        return ""
    who = row["직원명"] if isinstance(row.get("직원명"), str) else "사용자 미기재(공용카드)"
    s = f"{who}의 {row['가맹점명']}({row['가맹점업종']}) {_won(row['전표금액'])} 결제: " + ", ".join(phrases) + "."
    if row.get("AI 조정 사유"):
        s += f" 다만 {row['AI 조정 사유']}."
    return s


# ── Claude API (선택) ─────────────────────────────────────────


def claude_available() -> bool:
    if not os.environ.get("ANTHROPIC_API_KEY"):
        return False
    try:
        import anthropic  # noqa: F401
    except ImportError:
        return False
    return True


GUIDELINE = """[법인카드 사용지침 (가상 — 교육용)]
1. 법인카드는 업무 목적에 한하여 사용하며, 사적 사용은 금지한다.
2. 다음 업종은 사용을 금지한다: 유흥주점, 단란주점, 노래방, 안마·마사지, 상품권, 귀금속.
3. 골프장 사용은 사전에 품의(결재)를 받은 경우에 한한다.
4. 휴일·심야(23시~06시) 사용 시 전표 적요에 사유와 사전승인 문서번호를 기재한다.
5. 접대비는 상대방(소속·직함·인원)과 목적을 적요에 기재하며, 건당 50만원을 초과할 수 없다.
6. 한도 회피를 위한 분할결제를 금지한다.
7. 계정과목은 실제 사용 목적에 맞게 처리한다(예: 근무지 내 부서 식사는 여비교통비가 아님).
8. 사용분은 승인일로부터 60일 이내에 정산한다.
9. 공용카드 사용 시 실제 사용자를 전표에 기재한다.
10. 휴가 기간 중에는 법인카드를 사용하지 않는다."""

SYSTEM = (
    "당신은 보험회사 감사부의 법인카드 점검을 돕는 보조자입니다. 규칙 엔진이 고른 거래 1건과 사용지침을 보고 "
    "감사 담당자가 읽을 짧은 한국어 문장을 씁니다. 선별 사유는 걸린 규칙과 거래 정보를 엮어 한두 문장으로, "
    "추가 확인사항은 담당자가 증빙으로 확인할 항목을 2~4개로 씁니다. 결재문서제목·적요·계정명·업종·주소가 서로 "
    "맞지 않으면 의미 불일치로 표시하고, 적요에 사전승인 문서번호가 있어 소명 가능한 건이면 저위험 조정을 권고합니다. "
    "당신의 판단은 최종 판단이 아닌 참고자료입니다.\n\n" + GUIDELINE
)

TX_FIELDS = ["승인번호", "승인일시", "요일", "공휴일명", "직원명", "예산부서명", "구분", "가맹점명", "가맹점업종",
             "가맹점주소", "G/L 계정명", "전표금액", "전표 품목 텍스트", "결재문서제목", "회계일자", "걸린 규칙", "위험점수"]


def claude_comment(row, workplace: str | None = None) -> dict:
    """선별 건 1건에 대해 Claude가 사유·추가확인사항·R13·FP 판단을 생성한다."""
    import anthropic
    from pydantic import BaseModel

    class Comment(BaseModel):
        선별사유: str
        추가확인사항: list[str]
        의미불일치: bool
        의미불일치_사유: str
        저위험_조정권고: bool
        조정_사유: str

    tx = "\n".join(f"- {k}: {row.get(k)}" for k in TX_FIELDS)
    if workplace:
        tx += f"\n- 사용자 근무지: {workplace}"

    client = anthropic.Anthropic()
    response = client.beta.messages.parse(
        model=MODEL,
        max_tokens=4000,
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
        output_config={"effort": "low"},
        system=SYSTEM,
        messages=[{"role": "user", "content": f"다음 거래를 검토해 주세요.\n{tx}"}],
        output_format=Comment,
    )
    if response.stop_reason == "refusal" or response.parsed_output is None:
        raise RuntimeError("Claude가 응답을 생성하지 못했습니다.")
    return response.parsed_output.model_dump()
