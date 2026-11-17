"""DART 응답 → 연도별 표준 재무데이터."""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from .accounts import SPEC_BY_KEY, SPECS, AccountSpec, normalize_name
from .dart import DartClient, REPRT_CODES
from .issues import DataIssue
from .notes import find_da


def parse_amount(value) -> float | None:
    """'1,234' / '-1,234' / '(1,234)' / '' / '-' → float | None"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(",", "")
    if s in ("", "-", "N/A"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


@dataclass
class YearFinancials:
    year: int
    values: dict[str, float | None] = field(default_factory=dict)
    sources: dict[str, str] = field(default_factory=dict)   # 항목 → 사용한 계정 (감사 추적용)
    filled: dict[str, str] = field(default_factory=dict)    # 본문 외 공시(주석 등)에서 채운 항목 → 방법

    def get(self, key: str, default: float | None = None) -> float | None:
        v = self.values.get(key)
        return default if v is None else v

    def g0(self, key: str) -> float:
        return self.get(key, 0.0)

    # ---------------------------------------------------------- 파생 항목
    @property
    def revenue(self) -> float | None:
        return self.get("revenue")

    @property
    def ebit(self) -> float | None:
        return self.get("operating_income")

    @property
    def da(self) -> float | None:
        return self.get("depreciation_amortization")

    @property
    def ebitda(self) -> float | None:
        if self.ebit is None or self.da is None:
            return None
        return self.ebit + self.da

    @property
    def capex(self) -> float | None:
        a, b = self.get("capex_ppe"), self.get("capex_intangible")
        if a is None and b is None:
            return None
        return (a or 0.0) + (b or 0.0)

    @property
    def borrowings(self) -> float:
        """이자부 차입금 (리스부채 제외)"""
        return sum(self.g0(k) for k in
                   ("short_term_borrowings", "current_portion_debt", "bonds", "long_term_borrowings"))

    @property
    def total_debt(self) -> float:
        return self.borrowings + self.g0("lease_liabilities")

    @property
    def cash_and_investments(self) -> float:
        return self.g0("cash") + self.g0("short_term_investments")

    @property
    def nwc(self) -> float | None:
        """영업 순운전자본 = (유동자산 − 현금성/단기금융자산) − (유동부채 − 단기 이자부부채)"""
        ca, cl = self.get("current_assets"), self.get("current_liabilities")
        if ca is None or cl is None:
            return None
        op_ca = ca - self.cash_and_investments
        op_cl = cl - self.g0("short_term_borrowings") - self.g0("current_portion_debt")
        # 유동리스부채는 리스부채 합계에 포함되어 있어 따로 알 수 없으므로 무시 (보수적 근사)
        return op_ca - op_cl

    @property
    def tax_rate(self) -> float | None:
        pt, tax = self.get("pretax_income"), self.get("income_tax")
        if pt is None or tax is None or pt <= 0:
            return None
        return tax / pt


def _match_rows(rows: list[dict], spec: AccountSpec) -> list[dict]:
    rows = [r for r in rows if r.get("sj_div") in spec.statements]

    def excluded(r):
        n = normalize_name(r.get("account_nm", ""))
        return any(re.search(p, n) for p in spec.exclude)

    if spec.sum_all:
        hits, seen = [], set()
        for r in rows:
            n = normalize_name(r.get("account_nm", ""))
            by_id = r.get("account_id") in spec.ids
            by_name = any(re.fullmatch(p, n) for p in spec.names) and not excluded(r)
            if (by_id or by_name) and id(r) not in seen:
                seen.add(id(r))
                hits.append(r)
        return hits

    for acc_id in spec.ids:
        for r in rows:
            if r.get("account_id") == acc_id:
                return [r]
    for pattern in spec.names:
        for r in rows:
            if re.fullmatch(pattern, normalize_name(r.get("account_nm", ""))) and not excluded(r):
                return [r]
    return []


def extract_year(rows: list[dict], year: int, column: str = "thstrm_amount") -> YearFinancials:
    fin = YearFinancials(year=year)
    # IS 와 CIS 가 모두 있으면 IS 우선 (CIS 는 포괄손익 항목까지 섞여 있음)
    has_is = any(r.get("sj_div") == "IS" for r in rows)
    if has_is:
        rows = [r for r in rows if r.get("sj_div") != "CIS"]
    for spec in SPECS:
        hits = _match_rows(rows, spec)
        amounts = [(r, parse_amount(r.get(column))) for r in hits]
        amounts = [(r, a) for r, a in amounts if a is not None]
        if not amounts:
            fin.values[spec.key] = None
            continue
        total = sum(abs(a) if spec.absolute else a for _, a in amounts)
        fin.values[spec.key] = total
        fin.sources[spec.key] = " + ".join(
            f"[{r.get('sj_div')}] {r.get('account_nm')} ({r.get('account_id')})" for r, _ in amounts)
    return fin


def shares_outstanding(rows: list[dict], share_class: str = "total") -> float | None:
    """stockTotqySttus 응답 → 유통주식수 (발행주식총수 − 자기주식). share_class: total | common"""
    def row_for(label):
        for r in rows:
            if normalize_name(r.get("se", "")) == label:
                return r
        return None

    def outstanding(r):
        if r is None:
            return None
        issued = parse_amount(r.get("istc_totqy"))
        treasury = parse_amount(r.get("tesstk_co")) or 0.0
        if issued is None:
            return parse_amount(r.get("distb_stock_co"))
        return issued - treasury

    if share_class == "common":
        return outstanding(row_for("보통주"))
    total = outstanding(row_for("합계"))
    if total is None:
        parts = [outstanding(row_for(x)) for x in ("보통주", "우선주")]
        parts = [p for p in parts if p is not None]
        total = sum(parts) if parts else None
    return total


@dataclass
class CompanyData:
    corp_code: str
    corp_name: str
    stock_code: str
    fs_div: str
    years: list[YearFinancials]          # 과거 → 최근
    shares: float | None
    warnings: list[str] = field(default_factory=list)
    issues: list[DataIssue] = field(default_factory=list)

    @property
    def latest(self) -> YearFinancials:
        return self.years[-1]


# 기준연도에 계정이 없으면 0 으로 계산되는 항목 (차입금이 없는 회사 등)
ZERO_OK = ("short_term_investments", "short_term_borrowings", "current_portion_debt", "bonds",
           "long_term_borrowings", "lease_liabilities", "nci")


def fill_da_from_notes(client: DartClient, corp_code: str, years: list[YearFinancials],
                       rows_by_year: dict[int, list[dict]], fs_div: str, notes: list[str]) -> list[DataIssue]:
    """현금흐름표에 감가상각비가 없는 연도를 사업보고서 원문 주석에서 채운다.

    1순위: 해당 연도 사업보고서의 당기 값 (그 해 손익계산서와 같은 범위)
    2순위: 다음 연도 사업보고서의 전기 값 (재작성됐을 수 있음)
    API 가 가리키는 보고서가 일부만 고친 정정 공시라 주석이 없으면 같은 연도의 다른 사업보고서도 본다.
    """
    issues: list[DataIssue] = []
    docs: dict[str, str | None] = {}
    listed: dict[int, list[str]] = {}

    def fetch_doc(rcept: str, year: int):
        if rcept not in docs:
            try:
                docs[rcept] = client.document_xml(rcept)
            except Exception as e:  # noqa: BLE001  (주석 보완은 선택 기능 — 실패해도 계속)
                notes.append(f"{year}년 사업보고서 원문({rcept})을 받지 못함: {e}")
                docs[rcept] = None
        return docs[rcept]

    def candidates(report_year: int):
        rows = rows_by_year.get(report_year) or []
        first = next((r.get("rcept_no") for r in rows if r.get("rcept_no")), None)
        if first:
            yield first, rows
        if report_year not in listed:
            try:
                listed[report_year] = client.annual_reports(corp_code, report_year)
            except Exception:  # noqa: BLE001
                listed[report_year] = []
        for rc in listed[report_year]:
            if rc != first:
                yield rc, rows

    for yf in years:
        if yf.values.get("depreciation_amortization") is not None:
            continue
        found, used = None, None
        for report_year, period in ((yf.year, "CFY"), (yf.year + 1, "PFY")):
            for rcept, rows in candidates(report_year):
                xml = fetch_doc(rcept, report_year)
                if xml:
                    found = find_da(xml, yf.year, fs_div, period, api_rows=rows, revenue=yf.revenue)
                if found:
                    used = rcept
                    break
            if found:
                break
        if not found:
            continue
        yf.values["depreciation_amortization"] = found.value
        yf.sources["depreciation_amortization"] = f"{found.source} — 사업보고서 접수번호 {used}"
        yf.filled["depreciation_amortization"] = found.method
        extra = ""
        if found.method == "cash_flow_adjustment":
            extra = " 중단영업이 있으면 손익계산서 범위와 다를 수 있습니다."
        elif found.period == "PFY":
            extra = f" {yf.year}년 보고서에서 못 찾아 {yf.year + 1}년 보고서의 전기 값 사용 (재작성됐을 수 있음)."
        issues.append(DataIssue(
            "filled", f"감가상각비 ({yf.year})", f"{found.value / 1e8:,.0f}억원",
            f"현금흐름표 본문에 없음 → {found.source}.{extra}",
            [f"hist:depreciation_amortization:{yf.year}"]))
    return issues


def load_company(client: DartClient, query: str, base_year: int, n_years: int = 5,
                 fs_div: str = "auto", share_class: str = "total", use_notes: bool = True) -> CompanyData:
    corp = client.find_corp(query)
    code = corp["corp_code"]
    reprt = REPRT_CODES["annual"]
    notes: list[str] = []
    issues: list[DataIssue] = []

    def fetch(year, div):
        return client.financial_statements(code, year, reprt, div)

    chosen = fs_div
    if fs_div == "auto":
        chosen = "CFS" if fetch(base_year, "CFS") else "OFS"
        if chosen == "OFS":
            notes.append("연결재무제표가 없어 별도재무제표(OFS)를 사용합니다.")

    years: list[YearFinancials] = []
    rows_by_year: dict[int, list[dict]] = {}
    for y in range(base_year - n_years + 1, base_year + 1):
        rows = fetch(y, chosen)
        if not rows:
            notes.append(f"{y}년 사업보고서 재무제표가 없어 제외합니다.")
            continue
        rows_by_year[y] = rows
        years.append(extract_year(rows, y))
    if not years:
        raise ValueError(f"{corp['corp_name']}: {base_year}년 기준 재무제표를 찾을 수 없습니다.")

    if use_notes and any(y.values.get("depreciation_amortization") is None for y in years):
        issues += fill_da_from_notes(client, code, years, rows_by_year, chosen, notes)

    shares = shares_outstanding(client.share_status(code, base_year, reprt), share_class)
    if shares is None:
        notes.append("주식총수 현황에서 유통주식수를 찾지 못했습니다. --shares 로 직접 지정하세요.")

    for yf in years:
        missing = [k for k in ("revenue", "operating_income", "current_assets", "current_liabilities")
                   if yf.values.get(k) is None]
        if missing:
            notes.append(f"{yf.year}년 필수 항목 누락: {', '.join(missing)}")

    last = years[-1]
    zero = [k for k in ZERO_OK if last.values.get(k) is None]
    if zero:
        issues.append(DataIssue(
            "zero", ", ".join(SPEC_BY_KEY[k].label for k in zero), "0",
            f"{last.year}년 재무상태표에 해당 계정이 없어 0 으로 계산 (실제로 없는 경우가 대부분)",
            [f"hist:{k}:{last.year}" for k in zero]))
    return CompanyData(corp_code=code, corp_name=corp["corp_name"], stock_code=corp["stock_code"],
                       fs_div=chosen, years=years, shares=shares, warnings=notes, issues=issues)
