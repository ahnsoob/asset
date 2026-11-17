"""주석(원문) 감가상각비 추출 + 데이터 점검 항목 테스트."""
import io
import json
import zipfile

import openpyxl
import pytest

from dcf.cli import main
from dcf.dart import DartClient, _decode
from dcf.excel import export_excel
from dcf.extract import load_company
from dcf.model import Assumptions, run_dcf
from dcf.notes import find_da, find_da_text
from dcf.report import format_report

CTX = "{p}{y}dFY_ifrs-full_ConsolidatedAndSeparateFinancialStatementsAxis_ifrs-full_{m}Member"
RPT = "_ifrs-full_CarryingAmountAccumulatedDepreciationAmortisationAndImpairmentAndGrossCarryingAmountAxis_dart_ReportedAmountMember"


def te(code, ctx, value):
    return f'<TE ACODE="{code}" ACONTEXT="{ctx}" ADECIMAL="0">{value}</TE>'


def xbrl_doc(unit=1):
    """천원 단위(unit=1000) 로 표시된 원문 흉내: 본문 매출 + 주석 D&A (연결/별도, 당기/전기, 자산별 축)"""
    c = lambda p, y, m="Consolidated": CTX.format(p=p, y=y, m=m)  # noqa: E731
    return "<DOCUMENT>" + "".join([
        te("ifrs-full_Revenue", c("CFY", 2024), f"{1000_000 // unit:,}"),
        # 자산 종류별 축이 붙은 셀은 무시돼야 함
        te("ifrs-full_DepreciationPropertyPlantAndEquipment",
           c("CFY", 2024) + "_ifrs-full_ClassesOfPropertyPlantAndEquipmentAxis_ifrs-full_BuildingsMember", "(999)"),
        te("ifrs-full_DepreciationPropertyPlantAndEquipment", c("CFY", 2024), f"({70_000 // unit:,})"),
        te("ifrs-full_AmortisationIntangibleAssetsOtherThanGoodwill", c("CFY", 2024), f"({5_000 // unit:,})"),
        te("ifrs-full_AdjustmentsForDepreciationExpense", c("CFY", 2024) + RPT, f"{90_000 // unit:,}"),
        te("ifrs-full_AdjustmentsForAmortisationExpense", c("CFY", 2024) + RPT, f"{10_000 // unit:,}"),
        te("ifrs-full_DepreciationAndAmortisationExpense", c("CFY", 2024) + RPT, f"{80_000 // unit:,}"),
        te("ifrs-full_DepreciationAndAmortisationExpense", c("PFY", 2023) + RPT, f"{60_000 // unit:,}"),
        te("ifrs-full_DepreciationAndAmortisationExpense", c("CFY", 2024, "Separate") + RPT, f"{50_000 // unit:,}"),
    ]) + "</DOCUMENT>"


API_ROWS = [{"account_id": "ifrs-full_Revenue", "thstrm_amount": "1000000", "rcept_no": "20250301000001"}]


def test_xbrl_priority_period_and_member():
    x = xbrl_doc()
    v = find_da(x, 2024, "CFS", "CFY", api_rows=API_ROWS)
    assert v.value == 80_000 and v.method == "expense_by_nature"      # 성격별 분류가 1순위
    assert find_da(x, 2023, "CFS", "PFY", api_rows=API_ROWS).value == 60_000
    assert find_da(x, 2024, "OFS", "CFY", api_rows=API_ROWS).value == 50_000


def test_xbrl_fallback_methods():
    x = xbrl_doc().replace("ifrs-full_DepreciationAndAmortisationExpense", "x_removed")
    v = find_da(x, 2024, "CFS", "CFY", api_rows=API_ROWS)
    assert v.value == 100_000 and v.method == "cash_flow_adjustment"
    x = x.replace("ifrs-full_AdjustmentsFor", "x_removed")
    v = find_da(x, 2024, "CFS", "CFY", api_rows=API_ROWS)
    assert v.value == 75_000 and v.method == "asset_rollforward"     # 자산별 축(999) 제외


def test_xbrl_scale_detected_from_api():
    v = find_da(xbrl_doc(unit=1000), 2024, "CFS", "CFY", api_rows=API_ROWS)
    assert v.value == 80_000 and v.scale == 1000


TEXT_DOC = """<DOCUMENT>
<TITLE ATOC="Y" AASSOCNOTE="D-0-3-3-0">3. 연결재무제표 주석</TITLE>
<TITLE ATOC="Y">1. 회사의 개요 (연결)</TITLE><P>...</P>
<TITLE ATOC="Y">25. 비용의 성격별 분류 (연결)</TITLE>
<P>당기와 전기 중 발생한 비용의 성격별 분류에 대한 정보는 다음과 같습니다.</P>
<TABLE><TR><TD>당기</TD></TR></TABLE>
<TABLE><TR><TD>(단위 : 천원)</TD></TR></TABLE>
<TABLE><TR><TH> </TH><TH>공시금액</TH></TR>
<TR><TD>인건비</TD><TD>693,650</TD></TR>
<TR><TD>감가상각비 및 무형자산상각비</TD><TD>47,750</TD></TR>
<TR><TD>사용권자산상각비</TD><TD>1,111</TD></TR></TABLE>
<TABLE><TR><TD>전기</TD></TR></TABLE>
<TABLE><TR><TD>(단위 : 천원)</TD></TR></TABLE>
<TABLE><TR><TD>감가상각비 및 무형자산상각비</TD><TD>29,680</TD></TR></TABLE>
<TITLE ATOC="Y" AASSOCNOTE="D-0-3-4-0">4. 재무제표</TITLE>
<TITLE ATOC="Y" AASSOCNOTE="D-0-3-5-0">5. 재무제표 주석</TITLE>
<P>비용의 성격별 분류</P><TABLE><TR><TH>구분</TH><TH>당기</TH><TH>전기</TH></TR>
<TR><TD>감가상각비</TD><TD>30,000</TD><TD>20,000</TD></TR>
<TR><TD>무형자산상각비</TD><TD>1,000</TD><TD>900</TD></TR></TABLE>
<P>(단위 : 백만원)</P>
<TITLE ATOC="Y" AASSOCNOTE="D-0-3-6-0">6. 배당에 관한 사항</TITLE>
</DOCUMENT>"""


def test_text_fallback_separate_tables_per_period():
    v = find_da(TEXT_DOC, 2023, "CFS", "CFY", revenue=400_000_000)
    assert v.value == 47_750 * 1000 and v.method == "expense_by_nature_text"   # 사용권자산 제외
    assert find_da(TEXT_DOC, 2023, "CFS", "CFY", revenue=400e12) is None        # 매출 대비 비현실적이면 버림
    assert find_da_text(TEXT_DOC, "CFS", "PFY").value == 29_680 * 1000


def test_text_fallback_period_columns_and_unit_guess():
    # 별도 주석: 당기/전기가 열로 나뉜 표, 단위 문구가 표 뒤에 있어 매출로 단위를 추정
    v = find_da_text(TEXT_DOC, "OFS", "CFY", revenue=400_000e6)
    assert v.value == 31_000e6
    assert find_da_text(TEXT_DOC, "OFS", "PFY", revenue=400_000e6).value == 20_900e6
    assert find_da_text(TEXT_DOC, "OFS", "CFY", revenue=None) is None   # 단위를 알 수 없으면 포기


def test_decode_cp949():
    raw = '<?xml version="1.0" encoding="utf-8"?><A>감가상각비</A>'.encode("cp949")
    assert "감가상각비" in _decode(raw)


# ------------------------------------------------------------------ load_company 통합
class Resp:
    def __init__(self, payload):
        self.payload = payload
        self.content = payload if isinstance(payload, bytes) else json.dumps(payload).encode()

    def raise_for_status(self):
        pass

    def json(self):
        if isinstance(self.payload, bytes):
            raise ValueError
        return self.payload


class Session:
    def __init__(self, with_doc=True):
        self.with_doc, self.calls = with_doc, []

    def get(self, url, params=None, timeout=None):
        self.calls.append(url)
        if "stockTotqySttus" in url:
            return Resp({"status": "000", "list": [{"se": "합계", "istc_totqy": "1000", "tesstk_co": "0"}]})
        if "list.json" in url:
            return Resp({"status": "013", "list": []})
        if "document.xml" in url:
            if not self.with_doc:
                return Resp({"status": "014", "message": "파일이 존재하지 않습니다."})
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w") as z:
                z.writestr("doc.xml", xbrl_doc())
            return Resp(buf.getvalue())
        y = int(params["bsns_year"])
        r = lambda sj, n, a, i="-": {"sj_div": sj, "account_nm": n, "account_id": i,  # noqa: E731
                                     "thstrm_amount": a, "rcept_no": f"{y + 1}0301000001"}
        rows = [r("IS", "매출액", "1000000", "ifrs-full_Revenue"), r("IS", "영업이익", "200000"),
                r("IS", "법인세비용차감전순이익", "200000"), r("IS", "법인세비용", "44000"),
                r("BS", "유동자산", "500000"), r("BS", "유동부채", "300000"), r("BS", "현금및현금성자산", "100000"),
                r("BS", "자본총계", "900000"), r("CF", "유형자산의취득", "(150000)")]
        return Resp({"status": "000", "list": rows})


def test_load_company_fills_da_from_notes(tmp_path):
    (tmp_path / "corpCode.json").write_text(json.dumps(
        [{"corp_code": "00000002", "corp_name": "주석회사", "stock_code": "000002", "modify_date": ""}]))
    client = DartClient(api_key="dummy", cache_dir=tmp_path, session=Session())
    data = load_company(client, "000002", 2024, n_years=1, fs_div="CFS")
    y = data.latest
    assert y.da == 80_000 and y.filled["depreciation_amortization"] == "expense_by_nature"
    assert "주석" in y.sources["depreciation_amortization"]
    kinds = [(i.kind, i.item) for i in data.issues]
    assert ("filled", "감가상각비 (2024)") in kinds
    assert list((tmp_path).glob("document__*.xml.gz"))          # 원문은 압축 캐시

    res = run_dcf(data, Assumptions.from_dict({"beta": 1.1}))
    assert not any(i.item == "D&A / 매출" and i.kind == "assumed" for i in res.issues)
    # 시장 데이터: beta 는 입력, Rf·ERP 는 기본값으로 표시
    defaults = {i.item for i in res.issues if i.kind == "default"}
    assert "무위험이자율 (Rf)" in defaults and "베타 (β)" not in defaults


def test_issues_when_notes_unavailable(tmp_path):
    (tmp_path / "corpCode.json").write_text(json.dumps(
        [{"corp_code": "00000002", "corp_name": "주석회사", "stock_code": "000002", "modify_date": ""}]))
    client = DartClient(api_key="dummy", cache_dir=tmp_path, session=Session(with_doc=False))
    data = load_company(client, "000002", 2024, n_years=1, fs_div="CFS")
    assert data.latest.da is None
    res = run_dcf(data, Assumptions.from_dict({}))
    kinds = [i.kind for i in res.issues]
    assert kinds == sorted(kinds, key=["assumed", "default", "adjusted", "filled", "zero"].index)  # 중요도 순
    da = next(i for i in res.issues if i.item == "D&A / 매출")
    assert da.kind == "assumed" and "driver:da_pct" in da.targets
    assert any(i.kind == "assumed" and "Kd" in i.item for i in res.issues)      # 이자지급 없음

    text = format_report(res)
    assert text.index("[데이터 점검]") < text.index("[과거 실적")              # 맨 위에 표시
    assert "⚠ [가정] D&A / 매출" in text

    path = tmp_path / "o.xlsx"
    export_excel(res, str(path), recalc=False)
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames[0] == "Check" and wb.active.title == "Check"
    assert "D&A / 매출" in [c.value for c in wb["Check"]["B"]]
    assert wb["DCF"]["C7"].fill.fgColor.rgb.endswith("F8CBAD")                 # 가정 셀 강조
    assert "D&A" in wb["DCF"]["C7"].comment.text
    assert "Check" in wb["Summary"]["A3"].value


def test_cli_json_has_issues(tmp_path):
    out = tmp_path / "r.json"
    from pathlib import Path
    sample = Path(__file__).resolve().parents[1] / "examples" / "sample_cache"
    assert main(["샘플전자", "--year", "2024", "--offline", "--cache-dir", str(sample),
                 "--json", str(out)]) == 0
    j = json.loads(out.read_text(encoding="utf-8"))
    assert {"kind", "item", "value", "reason"} <= set(j["data_issues"][0])


@pytest.mark.parametrize("flag", [["--no-notes"]])
def test_cli_no_notes_flag(flag):
    from dcf.cli import build_parser
    assert build_parser().parse_args(["x", *flag]).no_notes is True
