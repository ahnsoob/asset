"""콘솔 출력용 리포트."""
from __future__ import annotations

import unicodedata

from .model import DCFResult

UNIT = 1e8  # 억원


def _ljust(s: str, width: int) -> str:
    """한글(전각) 폭을 2 로 계산해 왼쪽 정렬."""
    w = sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in s)
    return s + " " * max(0, width - w)


def _f(x, unit=UNIT):
    return "—" if x is None else f"{x / unit:,.0f}"


def _p(x):
    return "—" if x is None else f"{x:.1%}"


def format_report(res: DCFResult) -> str:
    d, a = res.company, res.a
    out = []
    w = out.append
    w(f"■ {d.corp_name} ({d.stock_code})  FCFF DCF  |  기준 {d.latest.year}년 "
      f"{'연결' if d.fs_div == 'CFS' else '별도'}  |  단위: 억원")
    w("")
    w("[과거 실적 — 공시 추출]")
    w(f"{'':14}" + "".join(f"{y.year:>10}" for y in d.years))
    for label, fn in [
        ("매출액", lambda y: _f(y.revenue)), ("영업이익", lambda y: _f(y.ebit)),
        ("영업이익률", lambda y: _p(y.ebit / y.revenue) if y.ebit is not None and y.revenue else "—"),
        ("감가상각비", lambda y: _f(y.da)), ("CAPEX", lambda y: _f(y.capex)),
        ("순운전자본", lambda y: _f(y.nwc)), ("총차입금", lambda y: _f(y.total_debt)),
        ("현금·단기금융", lambda y: _f(y.cash_and_investments)), ("유효세율", lambda y: _p(y.tax_rate)),
    ]:
        w(_ljust(label, 14) + "".join(f"{fn(y):>10}" for y in d.years))
    w("")
    w("[추정]")
    w(f"{'':14}" + "".join(f"{r['year']:>10}" for r in res.rows))
    for label, key, pct in [("매출 성장률", "growth", True), ("매출액", "revenue", False),
                            ("영업이익률", "ebit_margin", True), ("영업이익", "ebit", False),
                            ("NOPAT", "nopat", False), ("(+) D&A", "da", False),
                            ("(−) CAPEX", "capex", False), ("(−) ΔNWC", "delta_nwc", False),
                            ("FCFF", "fcff", False), ("PV(FCFF)", "pv_fcff", False)]:
        w(_ljust(label, 14) + "".join(f"{(_p(r[key]) if pct else _f(r[key])):>10}" for r in res.rows))
    w("")
    w("[할인율]")
    w(f"  Ke = {a.risk_free_rate:.2%} + {a.beta:.2f} × {a.equity_risk_premium:.2%}"
      + (f" + {a.size_premium:.2%}" if a.size_premium else "") + f" = {a.cost_of_equity:.2%}")
    w(f"  Kd(세전) {a.pre_tax_cost_of_debt:.2%}, 세율 {a.tax_rate:.1%}, D/(D+E) {a.debt_weight:.1%}"
      f"  →  WACC {a.wacc:.2%}")
    tv_method = (f"Exit EV/EBITDA {a.exit_ev_ebitda:.1f}x" if a.exit_ev_ebitda
                 else f"영구성장률 {a.terminal_growth:.1%}")
    w(f"  잔존가치: {tv_method}, 기중할인 {'적용' if a.mid_year else '미적용'}")
    w("")
    w("[밸류에이션]")
    w(f"  FCFF 현재가치 합계      {_f(res.pv_fcff_sum):>12}")
    w(f"  잔존가치 현재가치       {_f(res.pv_terminal_value):>12}  (EV 의 {res.pv_terminal_value / res.enterprise_value:.0%})")
    w(f"  기업가치 (EV)           {_f(res.enterprise_value):>12}")
    w(f"  (+) 현금·단기금융자산   {_f(res.cash_and_investments):>12}")
    w(f"  (−) 차입금·리스부채     {_f(res.total_debt):>12}")
    w(f"  (−) 비지배지분          {_f(res.nci):>12}")
    w(f"  주주가치                {_f(res.equity_value):>12}")
    w(f"  유통주식수              {a.shares:>12,.0f} 주")
    w(f"  ▶ 주당 내재가치         {res.value_per_share:>12,.0f} 원")
    if res.share_price:
        w(f"    현재 주가             {res.share_price:>12,.0f} 원  (상승여력 {res.upside:+.1%})")
    w("")
    s = res.sensitivity
    w("[민감도: 주당가치(원), 행=WACC 열=g]")
    w(f"{'':>8}" + "".join(f"{g:>10.1%}" for g in s["growth"]))
    for wv, row in zip(s["wacc"], s["per_share"]):
        w(f"{wv:>8.1%}" + "".join(f"{('—' if v is None else f'{v:,.0f}'):>10}" for v in row))
    notes = a.notes + d.warnings
    if notes:
        w("")
        w("[참고]")
        for n in notes:
            w(f"  - {n}")
    return "\n".join(out)
