"""사업보고서 원문(document.xml)의 XBRL 태그에서 주석 수치를 찾는다.

fnlttSinglAcntAll 은 본문 재무제표만 주므로, 현금흐름표에 감가상각비가 따로 표시되지 않는 회사는
D&A 를 알 수 없다. 원문 XML 의 표 셀에는 아래처럼 XBRL 태그가 붙어 있어 주석까지 읽을 수 있다.

  <TE ACODE="ifrs-full_DepreciationAndAmortisationExpense"
      ACONTEXT="CFY2025dFY_ifrs-full_ConsolidatedAndSeparateFinancialStatementsAxis_ifrs-full_ConsolidatedMember_...">
      359,725,574,807</TE>

ACONTEXT 의 앞부분은 기간(CFY=당기, PFY=전기, d=기간/e=시점), 그 뒤는 축(연결/별도, 자산 종류 등)이다.
축이 연결/별도 구분뿐인(= 합계) 셀만 쓴다.
"""
from __future__ import annotations

import html
import re
from collections import Counter
from dataclasses import dataclass

_CELL = re.compile(r'<TE\b[^>]*?\bACODE="([^"]+)"[^>]*?\bACONTEXT="([^"]+)"[^>]*>(.*?)</TE>', re.S)
_TAG = re.compile(r"<[^>]+>")

_CS_AXIS = "ifrs-full_ConsolidatedAndSeparateFinancialStatementsAxis"
_MEMBER = {"CFS": "ifrs-full_ConsolidatedMember", "OFS": "ifrs-full_SeparateMember"}
# 표 서식용 축 (값의 의미를 바꾸지 않음)
_NEUTRAL_AXES = (
    "ifrs-full_CarryingAmountAccumulatedDepreciationAmortisationAndImpairmentAndGrossCarryingAmountAxis"
    "_dart_ReportedAmountMember",
)

# (방법 이름, 설명, 합산할 XBRL 코드 묶음 — 앞의 묶음부터 시도)
DA_METHODS: tuple[tuple[str, str, tuple[tuple[str, ...], ...]], ...] = (
    ("expense_by_nature", "비용의 성격별 분류 — 감가상각비 및 상각비 (손익계산서와 같은 범위)", (
        ("ifrs-full_DepreciationAndAmortisationExpense",),
        ("ifrs-full_DepreciationExpense", "ifrs-full_AmortisationExpense"),
    )),
    ("cash_flow_adjustment", "현금흐름표 조정내역 — 감가상각비·무형자산상각비 가산 (중단영업 포함 가능)", (
        ("ifrs-full_AdjustmentsForDepreciationAndAmortisationExpense",),
        ("ifrs-full_AdjustmentsForDepreciationExpense", "ifrs-full_AdjustmentsForAmortisationExpense"),
        ("ifrs-full_AdjustmentsForDepreciationExpense",),
    )),
    ("asset_rollforward", "유형·무형자산 변동표 — 감가상각비 + 상각비", (
        ("ifrs-full_DepreciationPropertyPlantAndEquipment", "ifrs-full_AmortisationIntangibleAssetsOtherThanGoodwill"),
        ("ifrs-full_DepreciationPropertyPlantAndEquipment",),
    )),
)


@dataclass
class NoteValue:
    value: float          # 원 단위, 양수
    method: str           # DA_METHODS 의 방법 이름
    description: str
    codes: tuple[str, ...]
    period: str           # "CFY" | "PFY"
    scale: float          # 원문 표시 단위 → 원 배율 (1, 1e3, 1e6)

    @property
    def source(self) -> str:
        unit = {1: "원", 1e3: "천원", 1e6: "백만원"}.get(self.scale, f"×{self.scale:g}")
        period = "당기" if self.period == "CFY" else "전기"
        return f"[주석] {self.description} ({' + '.join(self.codes)}, {period}, 원문 단위 {unit})"


def _amount(raw: str) -> float | None:
    s = html.unescape(_TAG.sub("", raw)).strip().replace(",", "").replace(" ", "")
    if s in ("", "-", "−"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    if s.startswith(("-", "−")):
        neg, s = True, s[1:]
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def parse_facts(xml: str) -> list[tuple[str, str, float]]:
    """원문 XML → [(XBRL 코드, context, 값)] (원문 표시 단위 그대로)"""
    out = []
    for code, ctx, raw in _CELL.findall(xml):
        v = _amount(raw)
        if v is not None:
            out.append((code, ctx, v))
    return out


def _plain_context(ctx: str, period: str, year: int, fs_div: str, kind: str = "d") -> bool:
    """축이 연결/별도 구분(과 서식용 축)뿐인 '합계' context 인지."""
    base = f"{period}{year}{kind}FY"
    if not ctx.startswith(base):
        return False
    rest = ctx[len(base):]
    for ax in _NEUTRAL_AXES:
        rest = rest.replace("_" + ax, "")
    if not rest:
        return True                               # 축 없음 (연결/별도 구분 없는 문서)
    return rest == f"_{_CS_AXIS}_{_MEMBER[fs_div]}"


def _first_values(facts, period: str, year: int, fs_div: str) -> dict[str, float]:
    vals: dict[str, float] = {}
    for code, ctx, v in facts:
        if code not in vals and _plain_context(ctx, period, year, fs_div):
            vals[code] = v
    return vals


def detect_scale(facts, api_rows: list[dict], year: int, fs_div: str) -> float:
    """원문 표시 단위를 API(원 단위) 값과 대조해 추정. 일치하는 계정이 없으면 1(원)."""
    votes: Counter = Counter()
    for kind in ("d", "e"):
        doc = {}
        for code, ctx, v in facts:
            if code not in doc and _plain_context(ctx, "CFY", year, fs_div, kind):
                doc[code] = v
        for r in api_rows:
            code = r.get("account_id")
            if code not in doc or not doc[code]:
                continue
            try:
                api = float(str(r.get("thstrm_amount", "")).replace(",", ""))
            except ValueError:
                continue
            if not api:
                continue
            ratio = abs(api / doc[code])
            for s in (1.0, 1e3, 1e6):
                if abs(ratio / s - 1) < 0.01:
                    votes[s] += 1
    return votes.most_common(1)[0][0] if votes else 1.0


def find_da(xml: str, year: int, fs_div: str = "CFS", period: str = "CFY",
            api_rows: list[dict] | None = None, revenue: float | None = None) -> NoteValue | None:
    """감가상각비+무형자산상각비 (사용권자산 상각 제외 — CAPEX 에 사용권자산 취득이 없으므로).

    XBRL 태그로 먼저 찾고, 태그가 없는 옛 원문이면 '비용의 성격별 분류' 표 텍스트를 읽는다.
    """
    facts = parse_facts(xml)
    vals = _first_values(facts, period, year, fs_div) if facts else {}
    if not vals:
        return find_da_text(xml, fs_div, period, revenue)
    scale = detect_scale(facts, api_rows or [], year if period == "CFY" else year + 1, fs_div)
    for method, desc, groups in DA_METHODS:
        for codes in groups:
            if all(c in vals for c in codes):
                total = sum(abs(vals[c]) for c in codes) * scale
                if total > 0 and _plausible(total, revenue):
                    return NoteValue(total, method, desc, codes, period, scale)
    return find_da_text(xml, fs_div, period, revenue)


# ---------------------------------------------------------------- 태그 없는 옛 보고서용 (표 텍스트)
_TITLE = re.compile(r"<TITLE\b[^>]*>(.*?)</TITLE>", re.S)
_TABLE = re.compile(r"<TABLE\b.*?</TABLE>", re.S)
_ROW = re.compile(r"<TR\b.*?</TR>", re.S)
_CELLS = re.compile(r"<T[DEHU]\b[^>]*>(.*?)</T[DEHU]>", re.S)
_UNIT = re.compile(r"단위\s*[:：]\s*(백만원|천원|원)")
_UNITS = {"원": 1.0, "천원": 1e3, "백만원": 1e6}
_DA_COMBINED = re.compile(r"(총)?(유형자산|유·무형자산|유무형자산)?감가상각비(및|와|,|ㆍ|·)?(무형자산)?(상각비)?(합계)?"
                          r"|감가상각비및(무형자산)?상각비|유무형자산상각비|감가상각및상각비")
_DEP_ONLY = re.compile(r"(유형자산)?감가상각비")
_AMORT_ONLY = re.compile(r"무형자산(의)?상각비")


def _text(fragment: str) -> str:
    return re.sub(r"\s+", " ", html.unescape(_TAG.sub(" ", fragment))).strip()


_TITLE_TAG = re.compile(r"<TITLE\b([^>]*)>(.*?)</TITLE>", re.S)
_SECTION_CODE = {"CFS": "D-0-3-3-0", "OFS": "D-0-3-5-0"}   # 사업보고서 목차 코드: 연결/별도 재무제표 주석


def _notes_section(xml: str, fs_div: str) -> str:
    """'3. 연결재무제표 주석'(CFS) 또는 '5. 재무제표 주석'(OFS) 구간.

    목차 제목에는 AASSOCNOTE 코드가 붙어 있다 (주석 안의 '1. 회사의 개요' 같은 하위 제목에는 없음).
    """
    titles = [(m.start(), m.group(1), _text(m.group(2))) for m in _TITLE_TAG.finditer(xml)]
    start = None
    for i, (pos, attrs, t) in enumerate(titles):
        t0 = t.replace(" ", "")
        if _SECTION_CODE[fs_div] in attrs or (
                "AASSOCNOTE" in attrs and t0.endswith("재무제표주석") and ("연결" in t0) == (fs_div == "CFS")):
            start = i
            break
    if start is None:
        return ""
    pos = titles[start][0]
    end = next((p for p, attrs, _ in titles[start + 1:] if "AASSOCNOTE" in attrs), len(xml))
    return xml[pos:end]


def find_da_text(xml: str, fs_div: str = "CFS", period: str = "CFY",
                 revenue: float | None = None) -> NoteValue | None:
    """XBRL 태그가 없는 원문에서 '비용의 성격별 분류' 주석 표를 읽어 감가상각비+상각비를 찾는다."""
    sec = _notes_section(xml, fs_div)
    if not sec:
        return None
    want = "당기" if period == "CFY" else "전기"
    for occ in re.finditer(r"성격별", sec):
        window = sec[occ.start(): occ.start() + 30000]
        for tm in _TABLE.finditer(window):
            # 표 바로 앞 문구 ('당기', '(단위 : 천원)' 은 1행짜리 작은 표에 들어 있기도 함)
            context = _text(window[max(0, tm.start() - 6000): tm.start()])[-300:]
            rows = [[_text(c) for c in _CELLS.findall(r)] for r in _ROW.findall(tm.group(0))]
            if not rows:
                continue
            flat = " ".join(" ".join(r) for r in rows)
            units = _UNIT.findall(context + " " + flat)
            scale = _UNITS[units[-1]] if units else None

            # 기간: 표 머리글에 당기/전기 열이 함께 있으면 열 위치로, 아니면 표 앞 문구로
            col = None
            for r in rows[:3]:
                labels = [c.replace(" ", "") for c in r]
                if any(l.startswith("당기") for l in labels) and any(l.startswith("전기") for l in labels):
                    period_cols = [l[:2] for l in labels if l.startswith(("당기", "전기"))]
                    col = period_cols.index(want)
                    break
            if col is None:
                marks = re.findall(r"(?<![가-힣])(당기|전기)(?![가-힣])", context)
                if not marks or marks[-1] != want:
                    continue

            def value(r):
                nums = [v for v in (_amount(c) for c in r[1:]) if v is not None]
                idx = col or 0
                return abs(nums[idx]) if len(nums) > idx else None

            combined = dep = amort = None
            for r in rows:
                if not r:
                    continue
                label = re.sub(r"\s+", "", r[0])
                if "사용권" in label or "손상" in label or "환입" in label:
                    continue
                if combined is None and ("상각비" in label and ("및" in label or "무형" in label and "감가" in label)) \
                        and _DA_COMBINED.fullmatch(label):
                    combined = value(r)
                elif dep is None and _DEP_ONLY.fullmatch(label):
                    dep = value(r)
                elif amort is None and _AMORT_ONLY.fullmatch(label):
                    amort = value(r)
            raw = combined if combined else ((dep or 0) + (amort or 0) if dep else None)
            if not raw:
                continue
            scale = scale or _guess_scale(raw, revenue)
            if scale is None or not _plausible(raw * scale, revenue):
                continue
            codes = ("표: 감가상각비 및 무형자산상각비",) if combined else \
                (("표: 감가상각비 + 무형자산상각비",) if amort else ("표: 감가상각비",))
            return NoteValue(raw * scale, "expense_by_nature_text",
                             "비용의 성격별 분류 표 (XBRL 태그 없는 원문)", codes, period, scale)
    return None


def _plausible(value: float, revenue: float | None) -> bool:
    return revenue is None or revenue <= 0 or 0.0005 < value / revenue < 0.6


def _guess_scale(raw: float, revenue: float | None) -> float | None:
    if not revenue:
        return None
    ok = [s for s in (1.0, 1e3, 1e6) if 0.005 < raw * s / revenue < 0.4]
    return ok[0] if len(ok) == 1 else None
