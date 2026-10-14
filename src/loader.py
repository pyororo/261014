"""1단계: 엑셀 읽기 · 전처리 (CLAUDE.md 2-1 전처리 규칙 1~4)."""
from __future__ import annotations

import pandas as pd

TX_SHEET = "전체"
HOLIDAY_SHEET = "공휴일"

# 숫자처럼 보여도 앞자리 0·자릿수를 지키기 위해 텍스트로 읽는 컬럼
TEXT_COLUMNS = ["G/L 계정", "승인시간", "예산부서", "업종코드", "가맹점사업자번호"]

WEEKDAY_KO = ["월", "화", "수", "목", "금", "토", "일"]


def read_transactions(source) -> pd.DataFrame:
    """'전체' 시트를 읽는다. source는 파일 경로 또는 업로드된 파일 객체."""
    df = pd.read_excel(source, sheet_name=TX_SHEET, dtype={c: str for c in TEXT_COLUMNS})
    df["승인시간"] = df["승인시간"].str.strip().str.zfill(6)
    return df


def read_holidays(source) -> pd.DataFrame:
    hol = pd.read_excel(source, sheet_name=HOLIDAY_SHEET)
    hol["일자"] = pd.to_datetime(hol["일자"]).dt.normalize()
    return hol


def add_datetime_and_holiday(df: pd.DataFrame, holidays: pd.DataFrame) -> pd.DataFrame:
    """승인일+승인시간 → 승인일시, 주말 또는 공휴일 시트에 있는 날 → 휴일."""
    df = df.copy()
    df["승인일자"] = pd.to_datetime(df["승인일"].astype(str), format="%Y%m%d")
    df["승인일시"] = pd.to_datetime(
        df["승인일"].astype(str) + df["승인시간"], format="%Y%m%d%H%M%S"
    )
    df["승인월"] = df["승인일자"].dt.strftime("%Y-%m")
    df["요일"] = df["승인일자"].dt.weekday.map(lambda i: WEEKDAY_KO[i])

    holiday_names = holidays.set_index("일자")["명칭"]
    df["공휴일명"] = df["승인일자"].map(holiday_names)
    is_weekend = df["승인일자"].dt.weekday >= 5
    df["휴일여부"] = is_weekend | df["공휴일명"].notna()
    df["회계일자"] = pd.to_datetime(df["회계일자"])
    return df


def offset_cancellations(df: pd.DataFrame) -> pd.DataFrame:
    """취소건(음수)을 같은 카드·가맹점·절대금액의 매입건과 짝지어 표시한다.

    짝은 아직 짝이 없는 매입 중 승인일시가 가장 가까운 건으로 고른다.
    (추출 데이터의 취소 시각은 분 단위로 잘려 매입보다 몇 초 앞설 수 있어 선후는 따지지 않음)
    결과 컬럼:
      상계상태: 정상 / 취소상계(매입) / 취소상계(취소) / 취소미매칭
      상계승인번호: 짝이 된 상대 거래의 승인번호
      분석대상: 상계된 매입·취소건은 False (분석에서 제외)
    """
    df = df.sort_values("승인일시", kind="stable").copy()
    df["상계상태"] = "정상"
    df["상계승인번호"] = pd.NA

    cancels = df[df["취소구분"] == "취소"]
    purchases = df[df["취소구분"] != "취소"]
    used: set = set()

    for idx, c in cancels.iterrows():
        cand = purchases[
            (purchases["카드번호"] == c["카드번호"])
            & (purchases["가맹점명"] == c["가맹점명"])
            & (purchases["전표금액"] == -c["전표금액"])
            & (~purchases.index.isin(used))
        ]
        if cand.empty:
            df.at[idx, "상계상태"] = "취소미매칭"
            continue
        p_idx = (cand["승인일시"] - c["승인일시"]).abs().idxmin()
        used.add(p_idx)
        df.at[idx, "상계상태"] = "취소상계(취소)"
        df.at[p_idx, "상계상태"] = "취소상계(매입)"
        df.at[idx, "상계승인번호"] = df.at[p_idx, "승인번호"]
        df.at[p_idx, "상계승인번호"] = c["승인번호"]

    df["상계승인번호"] = df["상계승인번호"].astype("Int64")
    # 짝이 없는 취소건도 매입이 아니므로 분석 대상에서는 뺀다(화면에서 따로 확인)
    df["분석대상"] = df["상계상태"] == "정상"
    return df.sort_index()


def load_all(source) -> pd.DataFrame:
    """엑셀 한 개를 읽어 1단계 전처리를 모두 적용한 거래 테이블을 돌려준다."""
    tx = read_transactions(source)
    hol = read_holidays(source)
    tx = add_datetime_and_holiday(tx, hol)
    return offset_cancellations(tx)


def monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
    """월별 원본 건수·상계 제외 건수·분석대상 건수·금액."""
    g = df.groupby("승인월")
    out = pd.DataFrame(
        {
            "원본 건수": g.size(),
            "취소건수": g.apply(lambda x: (x["취소구분"] == "취소").sum()),
            "상계 제외 건수": g.apply(lambda x: (~x["분석대상"]).sum()),
            "분석대상 건수": g["분석대상"].sum(),
            "분석대상 금액": g.apply(lambda x: x.loc[x["분석대상"], "전표금액"].sum()),
            "휴일 사용 건수": g.apply(lambda x: (x["분석대상"] & x["휴일여부"]).sum()),
        }
    ).astype("int64")
    out.index.name = "승인월"
    return out.reset_index()
