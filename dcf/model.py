"""FCFF 기반 DCF 밸류에이션 엔진.

FCFF = EBIT × (1 − t) + D&A − CAPEX − ΔNWC
EV   = Σ PV(FCFF) + PV(TV)
주주가치 = EV + 현금·단기금융자산 − 차입금 − 리스부채 − 비지배지분
"""
from __future__ import annotations

from dataclasses import dataclass, field
from statistics import mean

from .extract import CompanyData
from .issues import DataIssue, sort_issues

KOREA_TOP_TAX_RATE = 0.264   # 법인세 최고세율 24% + 지방소득세 10%


def _clip(x, lo, hi):
    return max(lo, min(hi, x))


def _per_year(value, n: int) -> list[float] | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return [float(value)] * n
    vals = [float(v) for v in value]
    if len(vals) < n:                 # 짧으면 마지막 값으로 채움
        vals += [vals[-1]] * (n - len(vals))
    return vals[:n]


@dataclass
class Assumptions:
    """None 인 항목은 공시 데이터(과거 실적)에서 자동 산출한다. 숫자 또는 연도별 리스트로 지정 가능."""
    projection_years: int = 5
    revenue_growth: float | list[float] | None = None
    ebit_margin: float | list[float] | None = None
    da_pct_revenue: float | list[float] | None = None
    capex_pct_revenue: float | list[float] | None = None
    nwc_pct_revenue: float | list[float] | None = None
    tax_rate: float | None = None
    terminal_growth: float = 0.02
    exit_ev_ebitda: float | None = None      # 지정 시 영구성장 대신 EV/EBITDA 배수법으로 TV 계산

    # WACC
    risk_free_rate: float = 0.030            # 국고채 10년
    equity_risk_premium: float = 0.060
    beta: float = 1.0
    size_premium: float = 0.0
    pre_tax_cost_of_debt: float | None = None
    target_debt_weight: float | None = None  # D/(D+E). None 이면 시가(주가 입력 시) 또는 장부가 기준
    wacc_override: float | None = None

    share_price: float | None = None         # 현재 주가 (시가총액 가중치 & 상승여력 계산용)
    shares_outstanding: float | None = None  # None 이면 공시(주식총수 현황)에서
    mid_year: bool = True
    history_window: int = 3                  # 비율 평균을 낼 최근 연도 수
    # 사용자가 직접 지정한 필드 이름. None = 알 수 없음 (시장 데이터 기본값 경고를 생략)
    provided: frozenset[str] | None = None

    @classmethod
    def from_dict(cls, d: dict) -> "Assumptions":
        known = {k: v for k, v in d.items() if k in cls.__dataclass_fields__ and k != "provided"}
        return cls(**known, provided=frozenset(k for k, v in known.items() if v is not None))


@dataclass
class ResolvedAssumptions:
    years: list[int]
    revenue_growth: list[float]
    ebit_margin: list[float]
    da_pct: list[float]
    capex_pct: list[float]
    nwc_pct: list[float]
    tax_rate: float
    terminal_growth: float
    exit_ev_ebitda: float | None
    risk_free_rate: float
    equity_risk_premium: float
    beta: float
    size_premium: float
    cost_of_equity: float
    pre_tax_cost_of_debt: float
    debt_weight: float
    wacc: float
    wacc_override: float | None
    shares: float
    mid_year: bool
    history_window: int = 3
    notes: list[str] = field(default_factory=list)
    # 항목 → 산출 근거. "history": 과거 실적에서 그대로 산출(Excel 에서 수식으로 연결),
    # "market"/"book": 자본구조 가중치 기준. 없으면 직접 입력·기본값·제한(clip) 적용값.
    derived: dict[str, str] = field(default_factory=dict)
    issues: list[DataIssue] = field(default_factory=list)


@dataclass
class DCFResult:
    company: CompanyData
    a: ResolvedAssumptions
    base_revenue: float
    base_nwc: float
    rows: list[dict]                 # 연도별 추정치
    terminal_fcff: float
    terminal_value: float
    pv_terminal_value: float
    pv_fcff_sum: float
    enterprise_value: float
    cash_and_investments: float
    total_debt: float
    nci: float
    equity_value: float
    value_per_share: float
    share_price: float | None
    sensitivity: dict = field(default_factory=dict)

    @property
    def issues(self) -> list[DataIssue]:
        """데이터 점검 항목 (가정 → 기본값 → 조정 → 보완 → 0 처리 순)"""
        return sort_issues(self.company.issues + self.a.issues)

    @property
    def upside(self) -> float | None:
        if not self.share_price:
            return None
        return self.value_per_share / self.share_price - 1


# ---------------------------------------------------------------- 가정 산출
def resolve_assumptions(data: CompanyData, a: Assumptions) -> ResolvedAssumptions:
    n = a.projection_years
    notes: list[str] = []
    issues: list[DataIssue] = []

    def issue(kind, item, value, reason, *targets):
        issues.append(DataIssue(kind, item, value, reason, list(targets)))
    hist = [y for y in data.years if y.revenue]
    if not hist:
        raise ValueError("매출액이 있는 연도가 없습니다. (금융업 등은 FCFF DCF 에 적합하지 않을 수 있음)")
    recent = hist[-a.history_window:]
    last = data.latest
    derived: dict[str, str] = {}
    # 모든 연도에 매출이 있어야 Excel 의 '최근 N년 평균' 열과 같은 연도를 쓰게 된다
    linkable = len(hist) == len(data.years)

    def avg_ratio(fn, label, default, lo, hi, key):
        pairs = [(y.year, fn(y)) for y in recent]
        vals = [v for _, v in pairs if v is not None]
        if not vals:
            notes.append(f"{label}: 공시 데이터 부족 → 기본값 {default:.1%}")
            issue("assumed", label, f"{default:.1%} (전 기간)",
                  f"최근 {len(recent)}년 공시에서 산출 불가 → 기본값", f"driver:{key}")
            return default
        m = mean(vals)
        if linkable and lo <= m <= hi and len(vals) == len(recent):
            derived[key] = "history"
        if len(vals) < len(recent):
            miss = ", ".join(str(yr) for yr, v in pairs if v is None)
            issue("adjusted", label, f"{_clip(m, lo, hi):.1%}",
                  f"{miss}년 데이터 없음 → 나머지 {len(vals)}개 연도 평균", f"driver:{key}")
        if not lo <= m <= hi:
            issue("adjusted", label, f"{_clip(m, lo, hi):.1%}",
                  f"평균 {m:.1%} 가 허용 범위({lo:.0%}~{hi:.0%}) 밖 → 제한", f"driver:{key}")
        return _clip(m, lo, hi)

    # 매출 성장률: 과거 CAGR 에서 영구성장률로 선형 수렴
    growth = _per_year(a.revenue_growth, n)
    if growth is None:
        first, lastr = hist[0], hist[-1]
        span = lastr.year - first.year
        if span > 0 and first.revenue > 0 and lastr.revenue > 0:
            cagr = (lastr.revenue / first.revenue) ** (1 / span) - 1
        else:
            cagr = a.terminal_growth
            notes.append("매출 CAGR 계산 불가 → 영구성장률 사용")
            issue("assumed", "매출 성장률", f"{cagr:.1%} (영구성장률)",
                  "과거 매출이 1개 연도뿐이거나 음수라 CAGR 계산 불가", "driver:revenue_growth")
        g0 = _clip(cagr, -0.10, 0.30)
        if g0 != cagr:
            notes.append(f"과거 매출 CAGR {cagr:.1%} → {g0:.1%} 로 제한")
            issue("adjusted", "매출 성장률 (1년차)", f"{g0:.1%}",
                  f"과거 매출 CAGR {cagr:.1%} 가 −10%~30% 범위 밖 → 제한", "driver:revenue_growth")
        elif span > 0 and linkable:
            derived["revenue_growth"] = "history"
        gT = a.terminal_growth
        growth = [g0 + (gT - g0) * (i / (n - 1) if n > 1 else 0) for i in range(n)]

    margin = _per_year(a.ebit_margin, n) or [avg_ratio(
        lambda y: y.ebit / y.revenue if y.ebit is not None else None,
        "영업이익률", 0.10, -0.5, 0.6, "ebit_margin")] * n
    da = _per_year(a.da_pct_revenue, n)
    capex = _per_year(a.capex_pct_revenue, n) or [avg_ratio(
        lambda y: y.capex / y.revenue if y.capex is not None else None,
        "CAPEX/매출", 0.05, 0.0, 0.6, "capex_pct")] * n
    if da is None:
        da_vals = [y.da / y.revenue for y in recent if y.da is not None]
        if da_vals:
            da = [_clip(mean(da_vals), 0.0, 0.6)] * n
            if linkable and len(da_vals) == len(recent) and 0.0 <= mean(da_vals) <= 0.6:
                derived["da_pct"] = "history"
            if len(da_vals) < len(recent):
                miss = ", ".join(str(y.year) for y in recent if y.da is None)
                issue("adjusted", "D&A / 매출", f"{da[0]:.1%}",
                      f"{miss}년 감가상각비 없음 → 나머지 {len(da_vals)}개 연도 평균", "driver:da_pct")
        else:
            da = list(capex)
            notes.append("감가상각비를 찾지 못해 D&A = CAPEX 로 가정")
            issue("assumed", "D&A / 매출", f"= CAPEX/매출 {capex[0]:.1%}" + (" (연도별 CAPEX 와 동일)" if len(set(capex)) > 1 else ""),
                  "현금흐름표·주석 어디서도 감가상각비를 찾지 못함 → D&A = CAPEX (순투자 0) 가정. "
                  "--da 로 직접 지정 권장", "driver:da_pct")
    nwc = _per_year(a.nwc_pct_revenue, n) or [avg_ratio(
        lambda y: y.nwc / y.revenue if y.nwc is not None else None,
        "순운전자본/매출", 0.10, -0.5, 1.0, "nwc_pct")] * n

    if a.tax_rate is not None:
        tax = a.tax_rate
    else:
        rates = [y.tax_rate for y in recent if y.tax_rate is not None]
        if rates and 0.05 <= mean(rates) <= 0.40:
            tax = mean(rates)
            if linkable and len(rates) == len(recent):
                derived["tax_rate"] = "history"
        else:
            tax = KOREA_TOP_TAX_RATE
            notes.append(f"유효세율 산출 불가/이상치 → 법정 최고세율 {tax:.1%} 사용")
            why = (f"최근 평균 유효세율 {mean(rates):.1%} 가 5%~40% 밖" if rates
                   else "법인세비용 또는 세전이익(양수)을 찾지 못함")
            issue("assumed", "법인세율", f"{tax:.1%} (법정 최고세율)", why + " → 법정세율", "assump:tax_rate")

    # 주식수
    shares = a.shares_outstanding or data.shares
    if not shares:
        raise ValueError("유통주식수를 알 수 없습니다. --shares 로 지정하세요.")

    # WACC
    ke = a.risk_free_rate + a.beta * a.equity_risk_premium + a.size_premium
    if a.pre_tax_cost_of_debt is not None:
        kd = a.pre_tax_cost_of_debt
    else:
        kd = None
        if last.get("interest_paid") and len(data.years) >= 2:
            prev = data.years[-2]
            avg_debt = (last.borrowings + prev.borrowings) / 2
            if avg_debt > 0:
                kd = _clip(last.get("interest_paid") / avg_debt, a.risk_free_rate, a.risk_free_rate + 0.08)
                derived["pre_tax_cost_of_debt"] = "history"   # Excel: MAX(rf, MIN(rf+8%, 이자/평균차입금))
        if kd is None:
            kd = a.risk_free_rate + 0.015
            notes.append(f"이자비용/차입금으로 타인자본비용 산출 불가 → 무위험이자율+1.5% = {kd:.2%}")
            issue("assumed", "세전 타인자본비용 (Kd)", f"{kd:.2%} (Rf + 1.5%)",
                  "이자의 지급 또는 평균 차입금이 없어 산출 불가", "assump:kd")

    debt = last.total_debt
    if a.target_debt_weight is not None:
        wd = a.target_debt_weight
    else:
        if a.share_price:
            equity_mv = a.share_price * shares
            basis = "시가총액"
            derived["debt_weight"] = "market"
        else:
            equity_mv = last.get("owners_equity") or last.get("total_equity") or 0.0
            basis = "장부가 자본"
            derived["debt_weight"] = "book"
            if a.wacc_override is None:
                issue("default", "부채 비중 D/(D+E)",
                      f"{(debt / (debt + equity_mv) if debt + equity_mv > 0 else 0):.1%} (장부가 기준)",
                      "주가(--price)·목표 자본구조(--debt-weight) 미입력 → 장부가 자본으로 가중", "assump:wd")
        wd = debt / (debt + equity_mv) if (debt + equity_mv) > 0 else 0.0
        notes.append(f"자본구조 가중치: {basis} 기준 D/(D+E) = {wd:.1%}")
    if a.provided is not None and a.wacc_override is None:
        market = [("risk_free_rate", "무위험이자율 (Rf)", "rf", f"{a.risk_free_rate:.2%}"),
                  ("equity_risk_premium", "시장위험프리미엄 (ERP)", "erp", f"{a.equity_risk_premium:.2%}"),
                  ("beta", "베타 (β)", "beta", f"{a.beta:.2f}")]
        for fld, label, name, shown in market:
            if fld not in a.provided:
                issue("default", label, shown, "공시에 없는 시장 데이터 — 입력하지 않아 기본값 사용",
                      f"assump:{name}")
    wacc = a.wacc_override if a.wacc_override is not None else (1 - wd) * ke + wd * kd * (1 - tax)
    if wacc <= a.terminal_growth and a.exit_ev_ebitda is None:
        raise ValueError(f"WACC({wacc:.2%}) 가 영구성장률({a.terminal_growth:.2%}) 이하입니다.")

    base_year = data.latest.year
    return ResolvedAssumptions(
        years=[base_year + i + 1 for i in range(n)],
        revenue_growth=growth, ebit_margin=margin, da_pct=da, capex_pct=capex, nwc_pct=nwc,
        tax_rate=tax, terminal_growth=a.terminal_growth, exit_ev_ebitda=a.exit_ev_ebitda,
        risk_free_rate=a.risk_free_rate, equity_risk_premium=a.equity_risk_premium, beta=a.beta,
        size_premium=a.size_premium, cost_of_equity=ke, pre_tax_cost_of_debt=kd, debt_weight=wd,
        wacc=wacc, wacc_override=a.wacc_override,
        shares=shares, mid_year=a.mid_year, history_window=a.history_window, notes=notes,
        derived=derived, issues=issues,
    )


# ---------------------------------------------------------------- 밸류에이션
def _value(data: CompanyData, r: ResolvedAssumptions, wacc: float, g: float) -> dict:
    last = data.latest
    rev_prev = last.revenue
    nwc_prev = r.nwc_pct[0] * rev_prev        # 기준연도 NWC 도 같은 비율로 정규화 (1차년도 왜곡 방지)
    base_nwc = nwc_prev
    rows, pv_sum = [], 0.0
    for i, year in enumerate(r.years):
        t = i + 1
        rev = rev_prev * (1 + r.revenue_growth[i])
        ebit = rev * r.ebit_margin[i]
        nopat = ebit * (1 - r.tax_rate)
        da = rev * r.da_pct[i]
        capex = rev * r.capex_pct[i]
        nwc = rev * r.nwc_pct[i]
        d_nwc = nwc - nwc_prev
        fcff = nopat + da - capex - d_nwc
        period = t - 0.5 if r.mid_year else t
        df = (1 + wacc) ** -period
        pv = fcff * df
        pv_sum += pv
        rows.append(dict(year=year, revenue=rev, growth=r.revenue_growth[i], ebit=ebit,
                         ebit_margin=r.ebit_margin[i], taxes=ebit * r.tax_rate, nopat=nopat,
                         da=da, ebitda=ebit + da, capex=capex, nwc=nwc, delta_nwc=d_nwc,
                         fcff=fcff, period=period, discount_factor=df, pv_fcff=pv))
        rev_prev, nwc_prev = rev, nwc

    n = len(r.years)
    final = rows[-1]
    if r.exit_ev_ebitda is not None:
        tv = final["ebitda"] * r.exit_ev_ebitda
        tv_fcff = None
        tv_period = n                                # 배수법: N년 말 시점 가치
    else:
        rev_t = final["revenue"] * (1 + g)
        ebit_t = rev_t * r.ebit_margin[-1]
        tv_fcff = (ebit_t * (1 - r.tax_rate) + rev_t * r.da_pct[-1] - rev_t * r.capex_pct[-1]
                   - (rev_t * r.nwc_pct[-1] - final["nwc"]))
        tv = tv_fcff / (wacc - g)
        tv_period = n - 0.5 if r.mid_year else n     # Gordon 식은 첫 현금흐름 한 기간 전 시점의 가치
    pv_tv = tv * (1 + wacc) ** -tv_period

    ev = pv_sum + pv_tv
    cash = last.cash_and_investments
    debt = last.total_debt
    nci = last.g0("nci")
    equity = ev + cash - debt - nci
    return dict(rows=rows, base_nwc=base_nwc, tv_fcff=tv_fcff, tv=tv, pv_tv=pv_tv, pv_sum=pv_sum,
                ev=ev, cash=cash, debt=debt, nci=nci, equity=equity, per_share=equity / r.shares)


def run_dcf(data: CompanyData, a: Assumptions | None = None,
            wacc_steps=(-0.01, -0.005, 0.0, 0.005, 0.01),
            g_steps=(-0.01, -0.005, 0.0, 0.005, 0.01)) -> DCFResult:
    a = a or Assumptions()
    r = resolve_assumptions(data, a)
    v = _value(data, r, r.wacc, r.terminal_growth)

    waccs = [r.wacc + s for s in wacc_steps]
    gs = [r.terminal_growth + s for s in g_steps]
    table = [[(_value(data, r, w, g)["per_share"] if w > g or r.exit_ev_ebitda else None)
              for g in gs] for w in waccs]

    return DCFResult(
        company=data, a=r, base_revenue=data.latest.revenue, base_nwc=v["base_nwc"], rows=v["rows"],
        terminal_fcff=v["tv_fcff"], terminal_value=v["tv"], pv_terminal_value=v["pv_tv"],
        pv_fcff_sum=v["pv_sum"], enterprise_value=v["ev"], cash_and_investments=v["cash"],
        total_debt=v["debt"], nci=v["nci"], equity_value=v["equity"], value_per_share=v["per_share"],
        share_price=a.share_price,
        sensitivity={"wacc": waccs, "growth": gs, "per_share": table},
    )
