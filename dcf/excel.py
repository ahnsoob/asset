"""DCF 계산 과정을 수식이 살아있는 Excel 모델로 출력.

계산 흐름 (모두 셀 수식으로 연결):
  Historical (DART 공시 원본 → 파생 지표 → 최근 N년 평균/CAGR)
      → Assumptions (WACC, 세율, 순차입금 …)
      → DCF (매출 → EBIT → NOPAT → FCFF → 현재가치 → EV → 주당가치)
      → Sensitivity (WACC × g, 수식)
      → Summary

- 파란 글씨: 직접 입력값 (바꾸면 전체가 재계산됨)
- 초록 글씨: 다른 시트를 참조하는 연결 수식
- 검정 글씨: 계산 수식 / 공시 원본 값
- 금액 단위: 백만원 (주당가치·주가는 원)
"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as C
from openpyxl.workbook.defined_name import DefinedName

from .accounts import SPECS
from .model import DCFResult

UNIT = 1_000_000  # 백만원
NUM = '#,##0;[Red]-#,##0'
PCT = '0.0%;[Red]-0.0%'
INPUT = Font(color="0000CC")
LINK = Font(color="007A33")
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14)
NOTE = Font(italic=True, color="666666")
HEAD_FILL = PatternFill("solid", fgColor="DDE5F0")
SECTION_FILL = PatternFill("solid", fgColor="EEF2F7")
KEY_FILL = PatternFill("solid", fgColor="FFF2CC")
RAW_FILL = PatternFill("solid", fgColor="F7F7F7")
THIN = Border(top=Side(style="thin"))
DOUBLE = Border(top=Side(style="thin"), bottom=Side(style="double"))


def _name(wb, name, sheet, cell):
    col = "".join(c for c in cell if c.isalpha())
    row = "".join(c for c in cell if c.isdigit())
    wb.defined_names[name] = DefinedName(name, attr_text=f"'{sheet}'!${col}${row}")


def _m(x):
    return None if x is None else x / UNIT


def _section(ws, row, text, width):
    ws.cell(row, 1, text).font = BOLD
    for col in range(1, width + 1):
        ws.cell(row, col).fill = SECTION_FILL


def export_excel(res: DCFResult, path: str, recalc: bool = True) -> None:
    wb = Workbook()
    hist = _build_historical(wb, res)
    _build_assumptions(wb, res, hist)
    dcf = _build_dcf(wb, res, hist)
    _build_sensitivity(wb, res, dcf)
    _build_sources(wb, res)
    _build_summary(wb, res, dcf)
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True    # Excel 에서 열 때 전체 재계산
    wb.save(path)
    if recalc:
        recalculate(path)


def recalculate(path: str) -> bool:
    """LibreOffice 가 있으면 수식 결과값을 파일에 채워 넣는다.

    openpyxl 은 수식만 저장하므로, 미리보기(모바일·드라이브 등)에서도 숫자가 보이게 하려는 것.
    LibreOffice 가 없거나 실패하면 원본(수식만 있는 파일)을 그대로 둔다.
    """
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False
    src = Path(path).resolve()
    with tempfile.TemporaryDirectory() as tmp:
        try:
            subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, str(src)],
                           check=True, capture_output=True, timeout=180,
                           env={**os.environ, "HOME": tmp})
        except (subprocess.SubprocessError, OSError):
            return False
        out = Path(tmp) / src.name
        if not out.exists():
            return False
        shutil.copyfile(out, src)
    return True


# ====================================================================== Historical
def _build_historical(wb, res: DCFResult) -> dict:
    d, a = res.company, res.a
    ws = wb.active
    ws.title = "Historical"
    years = d.years
    H = len(years)
    first, lastc, avgc = 2, 1 + H, 2 + H
    L, AV = C(lastc), C(avgc)
    win = min(a.history_window, H)
    win_from = C(lastc - win + 1)

    ws["A1"] = f"{d.corp_name} 과거 실적 — DART 공시 추출 ({'연결' if d.fs_div == 'CFS' else '별도'})"
    ws["A1"].font = TITLE
    ws["A2"] = "단위: 백만원 | 회색 영역 = 공시 원본 값, 아래 파생 지표는 수식"
    ws["A2"].font = NOTE
    ws.cell(3, 1, "항목")
    for j, y in enumerate(years):
        ws.cell(3, first + j, y.year)
    ws.cell(3, avgc, f"최근 {win}년 평균")
    ws.cell(3, avgc + 1, "계산식")
    for col in range(1, avgc + 2):
        ws.cell(3, col).font, ws.cell(3, col).fill = BOLD, HEAD_FILL

    rows: dict[str, int] = {}
    r = 4
    _section(ws, r, "① 공시 원본 (fnlttSinglAcntAll)", avgc + 1)
    for spec in SPECS:
        r += 1
        rows[spec.key] = r
        ws.cell(r, 1, spec.label)
        for j, y in enumerate(years):
            c = ws.cell(r, first + j, _m(y.values.get(spec.key)))
            c.number_format, c.fill = NUM, RAW_FILL
        ws.cell(r, avgc + 1, f"{'/'.join(spec.statements)} 계정 매칭 (Sources 시트 참조)").font = NOTE

    r += 2
    _section(ws, r, "② 파생 지표 (수식)", avgc + 1)

    def R(key, x):
        return f"{x}{rows[key]}"

    derived_specs = [
        # key, label, fmt, formula(x, p, j), avg?, 설명
        ("ebitda", "EBITDA", NUM,
         lambda x, p, j: f'=IF(OR({R("operating_income", x)}="",{R("depreciation_amortization", x)}=""),"",'
                         f'{R("operating_income", x)}+{R("depreciation_amortization", x)})', False,
         "영업이익 + 감가상각비"),
        ("capex", "CAPEX 합계", NUM,
         lambda x, p, j: f'=IF(AND({R("capex_ppe", x)}="",{R("capex_intangible", x)}=""),"",'
                         f'N({R("capex_ppe", x)})+N({R("capex_intangible", x)}))', False,
         "유형자산 취득 + 무형자산 취득"),
        ("cash_inv", "현금 및 단기금융자산", NUM,
         lambda x, p, j: f'=N({R("cash", x)})+N({R("short_term_investments", x)})', False,
         "현금및현금성자산 + 단기금융자산"),
        ("borrowings", "차입금 (리스 제외)", NUM,
         lambda x, p, j: f'=N({R("short_term_borrowings", x)})+N({R("current_portion_debt", x)})'
                         f'+N({R("bonds", x)})+N({R("long_term_borrowings", x)})', False,
         "단기차입금 + 유동성장기부채 + 사채 + 장기차입금"),
        ("total_debt", "총차입금 (리스 포함)", NUM,
         lambda x, p, j: f'={x}{{borrowings}}+N({R("lease_liabilities", x)})', False,
         "차입금 + 리스부채"),
        ("nwc", "영업 순운전자본", NUM,
         lambda x, p, j: f'=IF(OR({R("current_assets", x)}="",{R("current_liabilities", x)}=""),"",'
                         f'({R("current_assets", x)}-{x}{{cash_inv}})-({R("current_liabilities", x)}'
                         f'-N({R("short_term_borrowings", x)})-N({R("current_portion_debt", x)})))', False,
         "(유동자산 − 현금·단기금융) − (유동부채 − 단기차입금 − 유동성장기부채)"),
        (None, None, None, None, None, None),
        ("growth", "매출 성장률", PCT,
         lambda x, p, j: "" if j == 0 else
         f'=IF(OR({R("revenue", p)}="",{R("revenue", x)}=""),"",{R("revenue", x)}/{R("revenue", p)}-1)', False,
         "당기 매출 / 전기 매출 − 1"),
        ("ebit_margin", "영업이익률", PCT,
         lambda x, p, j: f'=IF(OR({R("operating_income", x)}="",{R("revenue", x)}=""),"",'
                         f'{R("operating_income", x)}/{R("revenue", x)})', True, "영업이익 / 매출"),
        ("da_pct", "D&A / 매출", PCT,
         lambda x, p, j: f'=IF(OR({R("depreciation_amortization", x)}="",{R("revenue", x)}=""),"",'
                         f'{R("depreciation_amortization", x)}/{R("revenue", x)})', True, "감가상각비 / 매출"),
        ("capex_pct", "CAPEX / 매출", PCT,
         lambda x, p, j: f'=IF(OR({x}{{capex}}="",{R("revenue", x)}=""),"",{x}{{capex}}/{R("revenue", x)})',
         True, "CAPEX / 매출"),
        ("nwc_pct", "순운전자본 / 매출", PCT,
         lambda x, p, j: f'=IF(OR({x}{{nwc}}="",{R("revenue", x)}=""),"",{x}{{nwc}}/{R("revenue", x)})',
         True, "순운전자본 / 매출"),
        ("tax_rate", "유효세율", PCT,
         lambda x, p, j: f'=IF(OR({R("pretax_income", x)}="",{R("income_tax", x)}=""),"",'
                         f'IF({R("pretax_income", x)}>0,{R("income_tax", x)}/{R("pretax_income", x)},""))',
         True, "법인세비용 / 법인세차감전이익"),
        ("kd", "이자지급 / 평균차입금", PCT,
         lambda x, p, j: "" if j == 0 else
         f'=IF(OR({R("interest_paid", x)}="",({x}{{borrowings}}+{p}{{borrowings}})=0),"",'
         f'{R("interest_paid", x)}/(({x}{{borrowings}}+{p}{{borrowings}})/2))', False,
         "이자의 지급 / 기초·기말 평균 차입금"),
    ]
    for key, label, fmt, fn, has_avg, desc in derived_specs:
        r += 1
        if key is None:
            continue
        rows[key] = r
        ws.cell(r, 1, label)
        for j in range(H):
            x, p = C(first + j), C(first + j - 1)
            f = fn(x, p, j)
            ws.cell(r, first + j, f or None).number_format = fmt
        if has_avg:
            c = ws.cell(r, avgc, f'=IFERROR(AVERAGE({win_from}{r}:{L}{r}),"")')
            c.number_format, c.font, c.fill = fmt, BOLD, KEY_FILL
        ws.cell(r, avgc + 1, desc).font = NOTE

    # {borrowings} 같은 행 참조 자리표시자를 실제 행 번호로 치환
    for row in ws.iter_rows(min_row=rows["ebitda"], max_row=r, min_col=first, max_col=lastc):
        for c in row:
            if isinstance(c.value, str) and "{" in c.value:
                v = c.value
                for k, rr in rows.items():
                    v = v.replace("{" + k + "}", str(rr))
                c.value = v

    r += 1
    rows["cagr"] = r
    ws.cell(r, 1, "매출 CAGR")
    span_years = f"({L}3-{C(first)}3)"
    c = ws.cell(r, avgc, f'=IFERROR(({L}{rows["revenue"]}/{C(first)}{rows["revenue"]})^(1/{span_years})-1,"")')
    c.number_format, c.font, c.fill = PCT, BOLD, KEY_FILL
    ws.cell(r, avgc + 1, f"({years[-1].year} 매출 / {years[0].year} 매출)^(1/연수) − 1").font = NOTE

    ws.column_dimensions["A"].width = 26
    for col in range(first, avgc + 1):
        ws.column_dimensions[C(col)].width = 13
    ws.column_dimensions[C(avgc + 1)].width = 60
    ws.freeze_panes = "B4"
    return {"sheet": ws.title, "rows": rows, "last": L, "avg": AV}


# ====================================================================== Assumptions
def _build_assumptions(wb, res: DCFResult, hist: dict):
    d, a = res.company, res.a
    ws = wb.create_sheet("Assumptions")
    hs, hr, HL, HA = f"'{hist['sheet']}'", hist["rows"], hist["last"], hist["avg"]
    der = a.derived

    ws["A1"] = f"{d.corp_name} ({d.stock_code}) — 할인율 및 가치 조정 가정"
    ws["A1"].font = TITLE
    ws["A2"] = "파란 글씨 = 직접 입력, 초록 글씨 = Historical 시트 연결 수식, 검정 = 계산"
    ws["A2"].font = NOTE

    owners_row = hr["owners_equity"] if d.latest.get("owners_equity") else hr["total_equity"]
    kd = (f"=MAX(rf,MIN(rf+0.08,{hs}!{HL}{hr['kd']}))", "link",
          "Historical: 이자지급/평균차입금 (Rf ~ Rf+8% 로 제한)") \
        if der.get("pre_tax_cost_of_debt") == "history" else (a.pre_tax_cost_of_debt, "input", "Rf + 1.5% 또는 직접 입력")
    tax = (f"={hs}!{HA}{hr['tax_rate']}", "link", "Historical: 유효세율 최근 평균") \
        if der.get("tax_rate") == "history" else (a.tax_rate, "input", "직접 입력 / 법정세율")
    if der.get("debt_weight") == "market":
        wd = ("=debt/(debt+price*shares/1000000)", "calc", "시가총액 기준 D/(D+E)")
    elif der.get("debt_weight") == "book":
        wd = (f"=debt/(debt+{hs}!{HL}{owners_row})", "link", "장부가 자본 기준 D/(D+E)")
    else:
        wd = (a.debt_weight, "input", "목표 자본구조")
    wacc = (a.wacc_override, "input", "직접 입력 (override)") if a.wacc_override is not None else \
        ("=(1-wd)*ke+wd*kd*(1-tax_rate)", "calc", "E/(D+E)×Ke + D/(D+E)×Kd×(1−t)")

    sections = [
        ("① 자기자본비용 (CAPM)", [
            ("rf", "무위험이자율 (Rf)", a.risk_free_rate, "input", PCT, "국고채 10년 수익률"),
            ("erp", "시장위험프리미엄 (ERP)", a.equity_risk_premium, "input", PCT, ""),
            ("beta", "베타 (β)", a.beta, "input", "0.00", "회귀 베타 또는 동종업계 평균"),
            ("size_premium", "규모 프리미엄", a.size_premium, "input", PCT, ""),
            ("ke", "자기자본비용 (Ke)", "=rf+beta*erp+size_premium", "calc", PCT, "Rf + β × ERP + 규모 프리미엄"),
        ]),
        ("② 타인자본비용 및 WACC", [
            ("kd", "세전 타인자본비용 (Kd)", kd[0], kd[1], PCT, kd[2]),
            ("tax_rate", "법인세율 (t)", tax[0], tax[1], PCT, tax[2]),
            ("kd_after", "세후 타인자본비용", "=kd*(1-tax_rate)", "calc", PCT, "Kd × (1 − t)"),
            ("wd", "부채 비중 D/(D+E)", wd[0], wd[1], PCT, wd[2]),
            ("wacc", "WACC", wacc[0], wacc[1], PCT, wacc[2]),
        ]),
        ("③ 잔존가치", [
            ("tg", "영구성장률 (g)", a.terminal_growth, "input", PCT, ""),
            ("exit_multiple", "Exit EV/EBITDA (0 = 영구성장법)", a.exit_ev_ebitda or 0, "input", '0.0"x"', ""),
            ("mid_year", "기중 할인 (1 = 예, 0 = 아니오)", 1 if a.mid_year else 0, "input", "0", ""),
        ]),
        ("④ 기업가치 → 주주가치 조정 (기준연도 말)", [
            ("cash", "현금 및 단기금융자산", f"={hs}!{HL}{hr['cash_inv']}", "link", NUM, "Historical 연결"),
            ("debt", "차입금 + 사채 + 리스부채", f"={hs}!{HL}{hr['total_debt']}", "link", NUM, "Historical 연결"),
            ("nci", "비지배지분", f"=N({hs}!{HL}{hr['nci']})", "link", NUM, "Historical 연결"),
            ("shares", "유통주식수 (주)", a.shares, "input", NUM, "주식총수 현황: 발행주식 − 자기주식 (공시)"),
            ("price", "현재 주가 (원)", res.share_price or 0, "input", NUM, "0 이면 상승여력 미표시"),
        ]),
    ]
    r = 3
    for title, items in sections:
        r += 1
        _section(ws, r, title, 3)
        for name, label, val, kind, fmt, note in items:
            r += 1
            ws.cell(r, 1, label)
            c = ws.cell(r, 2, val)
            c.number_format = fmt
            c.font = {"input": INPUT, "link": LINK}.get(kind, Font())
            if name == "wacc":
                c.font, c.fill = BOLD, KEY_FILL
                ws.cell(r, 1).font = BOLD
            ws.cell(r, 3, note).font = NOTE
            _name(wb, name, ws.title, f"B{r}")
        r += 1
    notes = a.notes + d.warnings
    if notes:
        r += 1
        ws.cell(r, 1, "참고").font = BOLD
        for note in notes:
            r += 1
            ws.cell(r, 1, "※ " + note).font = NOTE
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 48


# ====================================================================== DCF
def _build_dcf(wb, res: DCFResult, hist: dict) -> dict:
    d, a = res.company, res.a
    n = len(a.years)
    ws = wb.create_sheet("DCF")
    hs, hr, HL, HA = f"'{hist['sheet']}'", hist["rows"], hist["last"], hist["avg"]
    first, last_col = 3, 2 + n
    F, L = C(first), C(last_col)
    note_col = last_col + 1

    ws["A1"] = f"{d.corp_name} FCFF DCF 모델 (단위: 백만원)"
    ws["A1"].font = TITLE
    ws.cell(3, 1, "연도")
    ws.cell(3, 2, f"{d.latest.year}A")
    for i, y in enumerate(a.years):
        ws.cell(3, first + i, f"{y}E")
    ws.cell(3, note_col, "산출 근거")
    for col in range(1, note_col + 1):
        ws.cell(3, col).font, ws.cell(3, col).fill = BOLD, HEAD_FILL

    _section(ws, 4, "① 추정 가정 (연도별로 덮어쓸 수 있음)", note_col)
    der = a.derived
    drivers = [
        (5, "매출 성장률", a.revenue_growth, "revenue_growth", "growth",
         "과거 매출 CAGR → 마지막 해에 영구성장률로 선형 수렴"),
        (6, "영업이익률", a.ebit_margin, "ebit_margin", "ebit_margin", "최근 평균 영업이익률"),
        (7, "D&A / 매출", a.da_pct, "da_pct", "da_pct", "최근 평균 D&A/매출"),
        (8, "CAPEX / 매출", a.capex_pct, "capex_pct", "capex_pct", "최근 평균 CAPEX/매출"),
        (9, "순운전자본 / 매출", a.nwc_pct, "nwc_pct", "nwc_pct", "최근 평균 NWC/매출"),
    ]
    for r, label, vals, dkey, hkey, why in drivers:
        ws.cell(r, 1, label)
        ws.cell(r, 2, f"={hs}!{HL}{hr[hkey]}").number_format = PCT
        ws.cell(r, 2).font = LINK
        linked = der.get(dkey) == "history"
        for i, v in enumerate(vals):
            col = first + i
            if linked and dkey == "revenue_growth":
                f = f"={hs}!{HA}{hr['cagr']}" if i == 0 else \
                    f"=${F}$5+(tg-${F}$5)*{i}/{max(n - 1, 1)}"
                c = ws.cell(r, col, f)
                c.font = LINK if i == 0 else Font()
            elif linked:
                c = ws.cell(r, col, f"={hs}!${HA}${hr[hkey]}")
                c.font = LINK
            else:
                c = ws.cell(r, col, v)
                c.font = INPUT
            c.number_format = PCT
        ws.cell(r, note_col, (why + " (Historical 연결)") if linked else "직접 입력 / 기본값").font = NOTE

    _section(ws, 10, "② 손익 및 잉여현금흐름", note_col)
    lines = {
        11: ("매출액", "전년 매출 × (1 + 성장률)"),
        12: ("영업이익 (EBIT)", "매출 × 영업이익률"),
        13: ("(−) 법인세", "EBIT × 법인세율"),
        14: ("NOPAT", "EBIT − 법인세"),
        15: ("(+) 감가상각비", "매출 × D&A 비율"),
        16: ("(−) CAPEX", "매출 × CAPEX 비율"),
        17: ("순운전자본", "매출 × NWC 비율 (기준연도도 같은 비율로 정규화)"),
        18: ("(−) 순운전자본 증감", "당기 NWC − 전기 NWC"),
        19: ("FCFF", "NOPAT + D&A − CAPEX − ΔNWC"),
        20: ("할인기간 (년)", "t − 0.5 (기중 할인) 또는 t"),
        21: ("할인계수", "1 / (1 + WACC)^할인기간"),
        22: ("FCFF 현재가치", "FCFF × 할인계수"),
        24: ("EBITDA", "EBIT + D&A"),
    }
    for r, (label, why) in lines.items():
        ws.cell(r, 1, label)
        ws.cell(r, note_col, why).font = NOTE
    link = lambda key: f"={hs}!{HL}{hr[key]}"  # noqa: E731
    ws.cell(11, 2, link("revenue")).font = LINK
    ws.cell(12, 2, link("operating_income")).font = LINK
    ws.cell(15, 2, link("depreciation_amortization")).font = LINK
    ws.cell(16, 2, link("capex")).font = LINK
    ws.cell(17, 2, f"=B11*{F}9")
    ws.cell(24, 2, link("ebitda")).font = LINK
    for i in range(n):
        col = first + i
        x, p = C(col), C(col - 1)
        f = {
            11: f"={p}11*(1+{x}5)", 12: f"={x}11*{x}6", 13: f"={x}12*tax_rate", 14: f"={x}12-{x}13",
            15: f"={x}11*{x}7", 16: f"={x}11*{x}8", 17: f"={x}11*{x}9", 18: f"={x}17-{p}17",
            19: f"={x}14+{x}15-{x}16-{x}18", 20: f"={i + 1}-IF(mid_year=1,0.5,0)",
            21: f"=1/(1+wacc)^{x}20", 22: f"={x}19*{x}21", 24: f"={x}12+{x}15",
        }
        for r, formula in f.items():
            ws.cell(r, col, formula)
    for r in range(11, 25):
        for col in range(2, last_col + 1):
            ws.cell(r, col).number_format = "0.00" if r in (20, 21) else NUM
    for r in (14, 19):
        for col in range(1, last_col + 1):
            ws.cell(r, col).font = BOLD
            ws.cell(r, col).border = THIN

    _section(ws, 26, "③ 잔존가치 → 기업가치 → 주당가치", note_col)
    val_rows = [
        ("FCFF 현재가치 합계", f"=SUM({F}22:{L}22)", NUM, "Σ FCFF 현재가치"),
        ("잔존가치 산출용 FCFF (N+1년)",
         f"={L}11*(1+tg)*{L}6*(1-tax_rate)+{L}11*(1+tg)*{L}7-{L}11*(1+tg)*{L}8-({L}11*(1+tg)*{L}9-{L}17)",
         NUM, "N년 매출 × (1+g) 로 정규화한 FCFF"),
        ("잔존가치 (TV)", f"=IF(exit_multiple>0,{L}24*exit_multiple,B28/(wacc-tg))", NUM,
         "FCFF(N+1) / (WACC − g)  또는  EBITDA(N) × 배수"),
        ("TV 할인기간 (년)", f"=IF(exit_multiple>0,{n},{n}-IF(mid_year=1,0.5,0))", "0.00",
         "영구성장법+기중할인: N − 0.5, 배수법: N"),
        ("TV 현재가치", "=B29/(1+wacc)^B30", NUM, "TV / (1 + WACC)^기간"),
        ("기업가치 (EV)", "=B27+B31", NUM, "FCFF 현재가치 합계 + TV 현재가치"),
        ("(+) 현금 및 단기금융자산", "=cash", NUM, "Assumptions ④"),
        ("(−) 차입금·사채·리스부채", "=debt", NUM, "Assumptions ④"),
        ("(−) 비지배지분", "=nci", NUM, "Assumptions ④"),
        ("주주가치", "=B32+B33-B34-B35", NUM, "EV + 현금 − 차입금 − 비지배지분"),
        ("유통주식수 (주)", "=shares", NUM, "Assumptions ④"),
        ("주당 내재가치 (원)", "=B36*1000000/B37", NUM, "주주가치(백만원) × 10⁶ / 주식수"),
        ("현재 주가 (원)", "=price", NUM, "Assumptions ④"),
        ("상승여력", '=IF(B39>0,B38/B39-1,"")', PCT, "내재가치 / 주가 − 1"),
        ("EV 중 TV 비중", "=B31/B32", PCT, "70~80% 이상이면 영구성장 가정에 민감"),
        ("WACC", "=wacc", PCT, "Assumptions ②"),
        ("암묵적 Exit EV/EBITDA", f'=IF({L}24>0,B29/{L}24,"")', '0.0"x"', "TV / EBITDA(N) — 영구성장법 검증용"),
        ("암묵적 영구성장률", f"=IF(exit_multiple>0,(B29*wacc-{L}19)/(B29+{L}19),tg)", PCT,
         "배수법 사용 시 TV 에 내포된 g"),
    ]
    for i, (label, formula, fmt, why) in enumerate(val_rows):
        r = 27 + i
        ws.cell(r, 1, label)
        c = ws.cell(r, 2, formula)
        c.number_format = fmt
        ws.cell(r, note_col, why).font = NOTE
        if label.startswith(("기업가치", "주주가치", "주당")):
            ws.cell(r, 1).font = BOLD
            c.font, c.fill, c.border = BOLD, KEY_FILL, DOUBLE
    ws.column_dimensions["A"].width = 30
    for col in range(2, last_col + 1):
        ws.column_dimensions[C(col)].width = 13
    ws.column_dimensions[C(note_col)].width = 50
    ws.freeze_panes = "B4"
    return {"sheet": ws.title, "first": F, "last": L, "n": n}


# ====================================================================== Sensitivity
def _build_sensitivity(wb, res: DCFResult, dcf: dict):
    ws = wb.create_sheet("Sensitivity")
    s = res.sensitivity
    a = res.a
    D, F, L = f"'{dcf['sheet']}'", dcf["first"], dcf["last"]
    ws["A1"] = "주당 내재가치 민감도 (원) — WACC × 영구성장률"
    ws["A1"].font = TITLE
    ws["A2"] = "모든 셀이 DCF 시트의 FCFF 를 참조하는 수식 (가정을 바꾸면 함께 갱신)"
    ws["A2"].font = NOTE
    ws.cell(4, 1, "WACC ↓ / g →").font = BOLD
    w_steps = [w - a.wacc for w in s["wacc"]]
    g_steps = [g - a.terminal_growth for g in s["growth"]]
    for j, st in enumerate(g_steps):
        c = ws.cell(4, 2 + j, f"=tg+({st:.6f})")
        c.number_format, c.font, c.fill = PCT, BOLD, HEAD_FILL
    for i, st in enumerate(w_steps):
        r = 5 + i
        c = ws.cell(r, 1, f"=wacc+({st:.6f})")
        c.number_format, c.font, c.fill = PCT, BOLD, HEAD_FILL
        for j in range(len(g_steps)):
            w, g = f"$A{r}", f"{C(2 + j)}$4"
            pv = f"SUMPRODUCT({D}!${F}$19:${L}$19,1/(1+{w})^{D}!${F}$20:${L}$20)"
            tv_fcff = (f"({D}!${L}$11*(1+{g})*{D}!${L}$6*(1-tax_rate)+{D}!${L}$11*(1+{g})*{D}!${L}$7"
                       f"-{D}!${L}$11*(1+{g})*{D}!${L}$8-({D}!${L}$11*(1+{g})*{D}!${L}$9-{D}!${L}$17))")
            tv = f"IF(exit_multiple>0,{D}!${L}$24*exit_multiple,{tv_fcff}/({w}-{g}))"
            ev = f"({pv}+{tv}/(1+{w})^{D}!$B$30)"
            f = f'=IF(OR(exit_multiple>0,{w}>{g}),({ev}+cash-debt-nci)*1000000/shares,"n/a")'
            cell = ws.cell(r, 2 + j, f)
            cell.number_format = NUM
            if i == len(w_steps) // 2 and j == len(g_steps) // 2:
                cell.fill, cell.font = KEY_FILL, BOLD
    ws.column_dimensions["A"].width = 16
    for j in range(len(g_steps)):
        ws.column_dimensions[C(2 + j)].width = 13


# ====================================================================== Sources
def _build_sources(wb, res: DCFResult):
    d = res.company
    last = d.latest
    ws = wb.create_sheet("Sources")
    ws["A1"] = f"계정 매핑 내역 — {last.year}년 DART fnlttSinglAcntAll ({d.corp_code})"
    ws["A1"].font = TITLE
    ws.append([])
    ws.append(["항목", "값 (백만원)", "사용된 공시 계정 [재무제표] 계정명 (XBRL ID)"])
    for c in ws[3]:
        c.font, c.fill = BOLD, HEAD_FILL
    for spec in SPECS:
        ws.append([spec.label, _m(last.values.get(spec.key)), last.sources.get(spec.key, "— 찾지 못함")])
        ws.cell(ws.max_row, 2).number_format = NUM
    ws.append([])
    ws.append([f"유통주식수: 주식의 총수 현황(stockTotqySttus) 발행주식총수 − 자기주식 = {res.a.shares:,.0f} 주"])
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 110
    for row in ws.iter_rows(min_row=4):
        row[2].alignment = Alignment(wrap_text=False)


# ====================================================================== Summary
def _build_summary(wb, res: DCFResult, dcf: dict):
    d, a = res.company, res.a
    ws = wb.create_sheet("Summary", 0)
    D = f"'{dcf['sheet']}'"
    ws["A1"] = f"{d.corp_name} ({d.stock_code}) DCF 밸류에이션 요약"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = (f"기준: {d.latest.year}년 사업보고서 {'연결' if d.fs_div == 'CFS' else '별도'}재무제표 "
                f"(DART) | 추정 {a.years[0]}~{a.years[-1]} | 단위: 백만원")
    ws["A2"].font = NOTE
    items = [
        ("주당 내재가치 (원)", f"={D}!B38", NUM, True),
        ("현재 주가 (원)", f"={D}!B39", NUM, False),
        ("상승여력", f"={D}!B40", PCT, False),
        (None, None, None, None),
        ("기업가치 (EV)", f"={D}!B32", NUM, False),
        ("  FCFF 현재가치 합계", f"={D}!B27", NUM, False),
        ("  잔존가치 현재가치", f"={D}!B31", NUM, False),
        ("(+) 현금 및 단기금융자산", "=cash", NUM, False),
        ("(−) 차입금·사채·리스부채", "=debt", NUM, False),
        ("(−) 비지배지분", "=nci", NUM, False),
        ("주주가치", f"={D}!B36", NUM, True),
        (None, None, None, None),
        ("WACC", "=wacc", PCT, False),
        ("영구성장률", "=tg", PCT, False),
        ("자기자본비용 (Ke)", "=ke", PCT, False),
        ("세전 타인자본비용 (Kd)", "=kd", PCT, False),
        ("법인세율", "=tax_rate", PCT, False),
        ("EV 중 TV 비중", f"={D}!B41", PCT, False),
    ]
    r = 3
    for label, f, fmt, key in items:
        r += 1
        if label is None:
            continue
        ws.cell(r, 1, label)
        c = ws.cell(r, 2, f)
        c.number_format = fmt
        if key:
            ws.cell(r, 1).font = BOLD
            c.font, c.fill = Font(bold=True, size=12), KEY_FILL
    r += 2
    ws.cell(r, 1, "계산 흐름").font = BOLD
    steps = [
        "1. Historical  : DART 공시 원본(회색) → 파생 지표 → 최근 평균·CAGR (수식)",
        "2. Assumptions : CAPM 자기자본비용, 타인자본비용, 세율, D/(D+E) → WACC / 현금·차입금 조정",
        "3. DCF         : 매출 → EBIT → NOPAT → + D&A − CAPEX − ΔNWC = FCFF → 할인 → TV → EV → 주당가치",
        "4. Sensitivity : WACC × 영구성장률 조합별 주당가치 (수식)",
        "5. Sources     : 각 항목이 어느 공시 계정에서 왔는지",
        "",
        "색상: 파란 글씨 = 입력값 / 초록 글씨 = 다른 시트 연결 / 검정 = 계산식 / 노란 칸 = 핵심 결과",
    ]
    for s in steps:
        r += 1
        ws.cell(r, 1, s).font = NOTE if s.startswith("색상") else Font()
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
