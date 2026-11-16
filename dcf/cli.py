"""명령행 인터페이스.

예)
  export DART_API_KEY=xxxxxxxx
  python -m dcf 삼성전자 --year 2024 --beta 1.1 --price 55000 --excel samsung_dcf.xlsx
  python -m dcf 샘플전자 --year 2024 --offline --cache-dir examples/sample_cache
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import date

from .dart import DartClient, DartError
from .excel import export_excel
from .extract import load_company
from .model import Assumptions, run_dcf
from .report import format_report


def _pct(s: str) -> float:
    """'3.5' 또는 '3.5%' → 0.035, '0.035' → 0.035"""
    s = s.strip().rstrip("%")
    v = float(s)
    return v / 100 if abs(v) >= 1 else v


def _pct_list(s: str):
    parts = [p for p in s.split(",") if p.strip()]
    vals = [_pct(p) for p in parts]
    return vals[0] if len(vals) == 1 else vals


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="python -m dcf", description="DART 공시 기반 DCF 밸류에이션")
    p.add_argument("company", help="회사명, 종목코드(6자리) 또는 DART 고유번호(8자리)")
    p.add_argument("--year", type=int, default=date.today().year - 1, help="기준 사업연도 (기본: 작년)")
    p.add_argument("--history", type=int, default=5, help="불러올 과거 연도 수 (기본 5)")
    p.add_argument("--fs", choices=["auto", "CFS", "OFS"], default="auto", help="연결/별도 (기본 auto)")
    p.add_argument("--share-class", choices=["total", "common"], default="total",
                   help="주당가치 분모: total=보통주+우선주, common=보통주만")

    g = p.add_argument_group("가정 (비우면 공시 데이터로 자동 산출)")
    g.add_argument("--config", help="가정 JSON 파일 (Assumptions 필드명). 명령행 인자가 우선")
    g.add_argument("--years", type=int, help="추정 기간 (기본 5)")
    g.add_argument("--growth", type=_pct_list, help="매출 성장률. 예: 8 또는 10,8,6,4,3")
    g.add_argument("--margin", type=_pct_list, help="영업이익률")
    g.add_argument("--capex", type=_pct_list, help="CAPEX/매출")
    g.add_argument("--da", type=_pct_list, help="감가상각비/매출")
    g.add_argument("--nwc", type=_pct_list, help="순운전자본/매출")
    g.add_argument("--tax", type=_pct, help="법인세율")
    g.add_argument("--tg", type=_pct, help="영구성장률 (기본 2%%)")
    g.add_argument("--exit-multiple", type=float, help="Exit EV/EBITDA 배수 (지정 시 영구성장법 대신 사용)")
    g.add_argument("--rf", type=_pct, help="무위험이자율 (기본 3.0%%)")
    g.add_argument("--erp", type=_pct, help="시장위험프리미엄 (기본 6.0%%)")
    g.add_argument("--beta", type=float, help="베타 (기본 1.0)")
    g.add_argument("--size-premium", type=_pct, help="규모 프리미엄")
    g.add_argument("--kd", type=_pct, help="세전 타인자본비용")
    g.add_argument("--debt-weight", type=_pct, help="목표 D/(D+E)")
    g.add_argument("--wacc", type=_pct, help="WACC 직접 지정")
    g.add_argument("--price", type=float, help="현재 주가 (원)")
    g.add_argument("--shares", type=float, help="유통주식수 직접 지정")
    g.add_argument("--no-mid-year", action="store_true", help="기말 할인 사용")

    o = p.add_argument_group("데이터/출력")
    o.add_argument("--api-key", help="DART API 키 (기본: 환경변수 DART_API_KEY)")
    o.add_argument("--cache-dir", default=".dart_cache")
    o.add_argument("--offline", action="store_true", help="캐시만 사용")
    o.add_argument("--no-notes", action="store_true",
                   help="감가상각비가 본문에 없을 때 사업보고서 원문 주석을 조회하지 않음")
    o.add_argument("--excel", help="Excel 파일로 저장 (수식 포함)")
    o.add_argument("--json", help="결과를 JSON 으로 저장")
    return p


CLI_TO_FIELD = {
    "years": "projection_years", "growth": "revenue_growth", "margin": "ebit_margin",
    "capex": "capex_pct_revenue", "da": "da_pct_revenue", "nwc": "nwc_pct_revenue", "tax": "tax_rate",
    "tg": "terminal_growth", "exit_multiple": "exit_ev_ebitda", "rf": "risk_free_rate",
    "erp": "equity_risk_premium", "beta": "beta", "size_premium": "size_premium",
    "kd": "pre_tax_cost_of_debt", "debt_weight": "target_debt_weight", "wacc": "wacc_override",
    "price": "share_price", "shares": "shares_outstanding",
}


def assumptions_from_args(args) -> Assumptions:
    cfg = {}
    if args.config:
        with open(args.config, encoding="utf-8") as f:
            cfg = json.load(f)
    for cli, fld in CLI_TO_FIELD.items():
        v = getattr(args, cli)
        if v is not None:
            cfg[fld] = v
    if args.no_mid_year:
        cfg["mid_year"] = False
    return Assumptions.from_dict(cfg)


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    client = DartClient(api_key=args.api_key, cache_dir=args.cache_dir, offline=args.offline)
    try:
        data = load_company(client, args.company, args.year, args.history, args.fs, args.share_class,
                            use_notes=not args.no_notes)
        res = run_dcf(data, assumptions_from_args(args))
    except (DartError, ValueError) as e:
        print(f"오류: {e}", file=sys.stderr)
        return 1

    print(format_report(res))
    if args.excel:
        export_excel(res, args.excel)
        print(f"\nExcel 저장: {args.excel}")
    if args.json:
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(result_to_dict(res), f, ensure_ascii=False, indent=2)
        print(f"JSON 저장: {args.json}")
    return 0


def result_to_dict(res) -> dict:
    a = res.a
    return {
        "company": {"name": res.company.corp_name, "stock_code": res.company.stock_code,
                    "corp_code": res.company.corp_code, "base_year": res.company.latest.year,
                    "fs_div": res.company.fs_div},
        "historical": [{"year": y.year, **y.values} for y in res.company.years],
        "data_issues": [i.to_dict() for i in res.issues],
        "assumptions": {k: getattr(a, k) for k in a.__dataclass_fields__ if k != "issues"},
        "projection": res.rows,
        "valuation": {
            "pv_fcff_sum": res.pv_fcff_sum, "terminal_value": res.terminal_value,
            "pv_terminal_value": res.pv_terminal_value, "enterprise_value": res.enterprise_value,
            "cash_and_investments": res.cash_and_investments, "total_debt": res.total_debt,
            "nci": res.nci, "equity_value": res.equity_value, "value_per_share": res.value_per_share,
            "share_price": res.share_price, "upside": res.upside,
        },
        "sensitivity": res.sensitivity,
    }


if __name__ == "__main__":
    sys.exit(main())
