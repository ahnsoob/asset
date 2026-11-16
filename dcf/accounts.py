"""DART 재무제표 계정 → DCF 표준 항목 매핑.

매칭 순서: (1) XBRL account_id 정확히 일치  →  (2) 정규화한 계정명(account_nm) 정규식 fullmatch.
회사마다 계정명이 조금씩 다르므로 필요하면 이 파일의 패턴을 추가하면 된다.
"""
from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass(frozen=True)
class AccountSpec:
    key: str
    label: str
    statements: tuple[str, ...]      # sj_div: BS, IS, CIS, CF
    ids: tuple[str, ...] = ()
    names: tuple[str, ...] = ()      # 정규화된 계정명에 대한 정규식 (fullmatch)
    sum_all: bool = False            # True: 매칭되는 모든 행 합산 (예: 감가상각비 + 무형자산상각비)
    absolute: bool = False           # True: 부호 무시 (현금흐름표의 유출 항목 등)
    exclude: tuple[str, ...] = ()    # 이름 매칭 시 제외할 정규식 (search)


IS = ("IS", "CIS")
BS = ("BS",)
CF = ("CF",)

SPECS: tuple[AccountSpec, ...] = (
    # ---------------------------------------------------------- 손익계산서
    AccountSpec("revenue", "매출액", IS,
                ids=("ifrs-full_Revenue", "ifrs_Revenue"),
                names=(r"매출액", r"수익\(매출액\)", r"매출", r"영업수익", r"수익")),
    AccountSpec("operating_income", "영업이익", IS,
                ids=("dart_OperatingIncomeLoss",),
                names=(r"영업이익(\(손실\))?", r"영업손익", r"영업손실(\(이익\))?")),
    AccountSpec("pretax_income", "법인세차감전순이익", IS,
                ids=("ifrs-full_ProfitLossBeforeTax", "ifrs_ProfitLossBeforeTax"),
                names=(r"법인세(비용)?차감전(계속영업)?(순)?(이익|손익)(\(손실\))?",
                       r"법인세비용차감전순손익")),
    AccountSpec("income_tax", "법인세비용", IS,
                ids=("ifrs-full_IncomeTaxExpenseContinuingOperations", "ifrs_IncomeTaxExpenseContinuingOperations"),
                names=(r"법인세비용(\(수익\))?", r"법인세(비용)?(\(이익\))?")),
    AccountSpec("net_income", "당기순이익", IS,
                ids=("ifrs-full_ProfitLoss", "ifrs_ProfitLoss"),
                names=(r"당기순이익(\(손실\))?", r"당기순손익", r"당기순손실(\(이익\))?")),

    # ---------------------------------------------------------- 현금흐름표
    AccountSpec("depreciation_amortization", "감가상각비+상각비", CF,
                ids=("ifrs-full_AdjustmentsForDepreciationExpense",
                     "ifrs-full_AdjustmentsForAmortisationExpense",
                     "ifrs-full_AdjustmentsForDepreciationAndAmortisationExpense",
                     "dart_AdjustmentsForDepreciationExpense",
                     "dart_AdjustmentsForAmortisationExpense"),
                names=(r".*감가상각비.*", r".*상각비"),
                exclude=(r"환입", r"손상", r"처분", r"대손"),
                sum_all=True, absolute=True),
    AccountSpec("capex_ppe", "유형자산의 취득", CF,
                ids=("ifrs-full_PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities",),
                names=(r"유형자산의?(취득|증가)",),
                absolute=True),
    AccountSpec("capex_intangible", "무형자산의 취득", CF,
                ids=("ifrs-full_PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities",),
                names=(r"무형자산의?(취득|증가)",),
                absolute=True),
    AccountSpec("interest_paid", "이자의 지급", CF,
                ids=("ifrs-full_InterestPaidClassifiedAsOperatingActivities",
                     "ifrs-full_InterestPaidClassifiedAsFinancingActivities"),
                names=(r"이자(의)?지급(액)?", r"이자지급"),
                absolute=True),

    # ---------------------------------------------------------- 재무상태표
    AccountSpec("current_assets", "유동자산", BS,
                ids=("ifrs-full_CurrentAssets",), names=(r"유동자산(합계)?",)),
    AccountSpec("current_liabilities", "유동부채", BS,
                ids=("ifrs-full_CurrentLiabilities",), names=(r"유동부채(합계)?",)),
    AccountSpec("cash", "현금및현금성자산", BS,
                ids=("ifrs-full_CashAndCashEquivalents",), names=(r"현금및현금(성)?자산",)),
    AccountSpec("short_term_investments", "단기금융자산", BS,
                ids=("dart_ShortTermDepositsNotClassifiedAsCashEquivalents",
                     "ifrs-full_ShorttermDepositsNotClassifiedAsCashEquivalents"),
                names=(r"단기금융상품", r"단기투자자산", r"단기.*금융자산", r"단기매매증권"),
                sum_all=True),
    AccountSpec("short_term_borrowings", "단기차입금", BS,
                ids=("ifrs-full_ShorttermBorrowings",), names=(r"단기차입금",)),
    AccountSpec("current_portion_debt", "유동성장기부채/사채", BS,
                ids=("ifrs-full_CurrentPortionOfLongtermBorrowings", "dart_CurrentPortionOfBonds"),
                names=(r"유동성장기(차입금|부채)", r"유동성사채"),
                sum_all=True),
    AccountSpec("bonds", "사채", BS,
                ids=("dart_BondsIssued", "ifrs-full_BondsIssued"), names=(r"사채",)),
    AccountSpec("long_term_borrowings", "장기차입금", BS,
                ids=("ifrs-full_LongtermBorrowings", "dart_LongTermBorrowingsGross"),
                names=(r"장기차입금",)),
    AccountSpec("lease_liabilities", "리스부채", BS,
                ids=("ifrs-full_CurrentLeaseLiabilities", "ifrs-full_NoncurrentLeaseLiabilities",
                     "ifrs-full_LeaseLiabilities"),
                names=(r"(유동|비유동)?리스부채",),
                sum_all=True),
    AccountSpec("total_equity", "자본총계", BS,
                ids=("ifrs-full_Equity",), names=(r"자본총계",)),
    AccountSpec("owners_equity", "지배기업소유주지분", BS,
                ids=("ifrs-full_EquityAttributableToOwnersOfParent",),
                names=(r"지배기업(의)?(소유주|소유지분).*", r"지배주주지분")),
    AccountSpec("nci", "비지배지분", BS,
                ids=("ifrs-full_NoncontrollingInterests",), names=(r"비지배지분",)),
)

SPEC_BY_KEY = {s.key: s for s in SPECS}

_NUMBERING = re.compile(r"^\s*(?:[IVXⅠ-Ⅻ]+|\d+|\(\d+\)|[가-하])\s*[.)]\s*")


def normalize_name(name: str) -> str:
    """'Ⅰ. 유동자산' → '유동자산', '영업이익 (손실)' → '영업이익(손실)'."""
    name = _NUMBERING.sub("", name or "")
    return re.sub(r"\s+", "", name)
