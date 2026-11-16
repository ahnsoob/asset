import json
from pathlib import Path

import pytest

from dcf.cli import main
from dcf.dart import DartClient, DartError
from dcf.excel import export_excel
from dcf.extract import extract_year, load_company, parse_amount, shares_outstanding
from dcf.model import Assumptions, run_dcf

ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "examples" / "sample_cache"
B = 1e8


@pytest.fixture(scope="module")
def sample():
    client = DartClient(cache_dir=SAMPLE, offline=True)
    return load_company(client, "샘플전자", 2024)


def row(sj, name, amount, acc_id="-표준계정코드 미사용-"):
    return {"sj_div": sj, "account_nm": name, "account_id": acc_id, "thstrm_amount": amount}


# ------------------------------------------------------------------ parsing
@pytest.mark.parametrize("raw,expected", [
    ("1,234", 1234.0), ("-1,234", -1234.0), ("(1,234)", -1234.0),
    ("", None), ("-", None), (None, None), (5, 5.0),
])
def test_parse_amount(raw, expected):
    assert parse_amount(raw) == expected


def test_extract_sample(sample):
    y = sample.latest
    assert y.revenue == 11200 * B
    assert y.ebit == 1060 * B
    assert y.da == (470 + 50) * B                 # 대손상각비 제외
    assert y.capex == (680 + 65) * B              # 음수 유출액 → 절대값
    assert y.get("lease_liabilities") == 190 * B  # 유동 + 비유동
    assert y.total_debt == (560 + 270 + 1000 + 400 + 190) * B
    assert y.nwc == ((5750 - 1250 - 400) - (3400 - 560 - 270)) * B
    assert sample.shares == 53_000_000
    assert [yf.year for yf in sample.years] == [2020, 2021, 2022, 2023, 2024]


def test_name_fallback_and_numbering():
    rows = [
        row("BS", "Ⅰ. 유동자산", "1,000"),
        row("BS", "Ⅰ. 유동부채", "600"),
        row("BS", "현금및현금성자산", "100"),
        row("IS", "매출액", "5,000"),
        row("IS", "영업이익", "500"),
        row("IS", "계속영업이익", "999"),            # 영업이익으로 잡히면 안 됨
        row("CF", "유형자산의취득", "(300)"),
    ]
    y = extract_year(rows, 2024)
    assert y.get("current_assets") == 1000
    assert y.revenue == 5000
    assert y.ebit == 500
    assert y.capex == 300


def test_is_preferred_over_cis():
    rows = [row("CIS", "당기순이익", "1", "ifrs-full_ProfitLoss"),
            row("IS", "당기순이익", "2", "ifrs-full_ProfitLoss")]
    assert extract_year(rows, 2024).get("net_income") == 2


def test_shares_outstanding():
    rows = [{"se": "보통주", "istc_totqy": "1,000", "tesstk_co": "100"},
            {"se": "우선주", "istc_totqy": "200", "tesstk_co": "-"}]
    assert shares_outstanding(rows, "common") == 900
    assert shares_outstanding(rows, "total") == 1100      # 합계 행 없으면 합산


# ------------------------------------------------------------------ model
def test_dcf_matches_closed_form(sample):
    """성장·마진 상수, 기말 할인일 때 손으로 계산한 값과 일치하는지."""
    g, m, t, w, tg = 0.05, 0.10, 0.25, 0.09, 0.02
    a = Assumptions(projection_years=3, revenue_growth=g, ebit_margin=m, tax_rate=t,
                    da_pct_revenue=0.04, capex_pct_revenue=0.05, nwc_pct_revenue=0.10,
                    terminal_growth=tg, wacc_override=w, mid_year=False)
    res = run_dcf(sample, a)
    rev0 = sample.latest.revenue
    revs = [rev0 * (1 + g) ** k for k in range(4)]
    fcff = [revs[k] * m * (1 - t) + revs[k] * 0.04 - revs[k] * 0.05 - (revs[k] - revs[k - 1]) * 0.10
            for k in (1, 2, 3)]
    pv = sum(f / (1 + w) ** (k + 1) for k, f in enumerate(fcff))
    rev4 = revs[3] * (1 + tg)
    tv_fcff = rev4 * m * (1 - t) + rev4 * 0.04 - rev4 * 0.05 - (rev4 - revs[3]) * 0.10
    pv_tv = tv_fcff / (w - tg) / (1 + w) ** 3
    assert res.enterprise_value == pytest.approx(pv + pv_tv)
    y = sample.latest
    expected_equity = pv + pv_tv + y.cash_and_investments - y.total_debt - y.g0("nci")
    assert res.equity_value == pytest.approx(expected_equity)
    assert res.value_per_share == pytest.approx(expected_equity / sample.shares)


def test_auto_assumptions_from_history(sample):
    res = run_dcf(sample, Assumptions())
    a = res.a
    assert a.revenue_growth[0] == pytest.approx((11200 / 8000) ** 0.25 - 1)
    assert a.revenue_growth[-1] == pytest.approx(0.02)       # 영구성장률로 수렴
    assert a.ebit_margin[0] == pytest.approx((870 / 9700 + 950 / 10300 + 1060 / 11200) / 3)
    assert 0.2 < a.tax_rate < 0.3
    assert res.value_per_share > 0
    mid = len(res.sensitivity["wacc"]) // 2
    assert res.sensitivity["per_share"][mid][mid] == pytest.approx(res.value_per_share)


def test_exit_multiple(sample):
    res = run_dcf(sample, Assumptions(exit_ev_ebitda=8.0))
    assert res.terminal_value == pytest.approx(res.rows[-1]["ebitda"] * 8.0)


def test_wacc_below_growth_raises(sample):
    with pytest.raises(ValueError):
        run_dcf(sample, Assumptions(wacc_override=0.015, terminal_growth=0.02))


# ------------------------------------------------------------------ client
class FakeResp:
    def __init__(self, payload):
        self.payload = payload

    def raise_for_status(self):
        pass

    def json(self):
        return self.payload


class FakeSession:
    def __init__(self, responder):
        self.responder, self.calls = responder, []

    def get(self, url, params=None, timeout=None):
        self.calls.append((url, dict(params)))
        return FakeResp(self.responder(url, params))


def test_client_cache_and_ofs_fallback(tmp_path):
    (tmp_path / "corpCode.json").write_text(json.dumps(
        [{"corp_code": "00000001", "corp_name": "별도만회사", "stock_code": "000001", "modify_date": ""}]))

    def responder(url, params):
        if "stockTotqySttus" in url:
            return {"status": "000", "list": [{"se": "합계", "istc_totqy": "10", "tesstk_co": "0"}]}
        if "list.json" in url:                                   # 주석 조회용 공시 목록
            return {"status": "013", "message": "조회된 데이타가 없습니다."}
        if params["fs_div"] == "CFS":
            return {"status": "013", "message": "조회된 데이타가 없습니다."}
        return {"status": "000", "list": [row("IS", "매출액", "100"), row("IS", "영업이익", "10"),
                                          row("BS", "유동자산", "50"), row("BS", "유동부채", "20")]}

    session = FakeSession(responder)
    client = DartClient(api_key="dummy", cache_dir=tmp_path, session=session)
    data = load_company(client, "000001", 2024, n_years=2)
    assert data.fs_div == "OFS"
    assert len(data.years) == 2 and data.shares == 10
    n_calls = len(session.calls)
    assert all("crtfc_key" in p for _, p in session.calls)
    assert not any("crtfc_key" in f.name for f in tmp_path.iterdir())

    load_company(client, "000001", 2024, n_years=2)          # 두 번째는 전부 캐시
    assert len(session.calls) == n_calls


def test_client_error_status(tmp_path):
    session = FakeSession(lambda u, p: {"status": "010", "message": "등록되지 않은 키입니다."})
    client = DartClient(api_key="bad", cache_dir=tmp_path, session=session)
    with pytest.raises(DartError, match="010"):
        client.financial_statements("00000001", 2024)


def test_offline_cache_miss(tmp_path):
    with pytest.raises(DartError, match="오프라인"):
        DartClient(cache_dir=tmp_path, offline=True).financial_statements("00000001", 2024)


# ------------------------------------------------------------------ outputs
def test_excel_export(sample, tmp_path):
    import openpyxl
    res = run_dcf(sample, Assumptions(share_price=30000))
    path = tmp_path / "out.xlsx"
    export_excel(res, str(path), recalc=False)
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames == ["Check", "Summary", "Historical", "Assumptions", "DCF", "Sensitivity", "Sources"]
    assert wb["DCF"]["B38"].value == "=B36*1000000/B37"
    assert "wacc" in wb.defined_names


def test_cli_offline(tmp_path, capsys):
    out = tmp_path / "r.json"
    code = main(["샘플전자", "--year", "2024", "--offline", "--cache-dir", str(SAMPLE),
                 "--growth", "10,8,6,4,3", "--beta", "1.2", "--json", str(out)])
    assert code == 0
    assert "주당 내재가치" in capsys.readouterr().out
    j = json.loads(out.read_text(encoding="utf-8"))
    assert j["assumptions"]["revenue_growth"] == pytest.approx([0.10, 0.08, 0.06, 0.04, 0.03])
    assert j["assumptions"]["beta"] == 1.2
