"""DCF 결과를 수식이 살아있는 Excel 파일로 출력.

- 파란 글씨 셀은 입력값: 바꾸면 DCF 시트가 자동 재계산된다.
- 금액 단위: 백만원 (주당가치는 원)
- Sensitivity 시트는 Python 에서 계산한 정적 값이다.
"""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

from .accounts import SPECS
from .model import DCFResult

UNIT = 1_000_000  # 백만원
NUM = '#,##0;[Red]-#,##0'
PCT = '0.0%;[Red]-0.0%'
INPUT = Font(color="0000CC")
BOLD = Font(bold=True)
HEAD_FILL = PatternFill("solid", fgColor="DDE5F0")
KEY_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Border(top=Side(style="thin"))


def _name(wb, name, sheet, cell):
    col, row = "".join(c for c in cell if c.isalpha()), "".join(c for c in cell if c.isdigit())
    wb.defined_names[name] = DefinedName(name, attr_text=f"'{sheet}'!${col}${row}")


def _m(x):
    return None if x is None else x / UNIT


def export_excel(res: DCFResult, path: str) -> None:
    wb = Workbook()
    d, a = res.company, res.a
    n = len(a.years)

    # ------------------------------------------------------------ Assumptions
    ws = wb.active
    ws.title = "Assumptions"
    ws["A1"] = f"{d.corp_name} ({d.stock_code}) DCF 가정"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"기준연도 {d.latest.year} / {'연결' if d.fs_div == 'CFS' else '별도'}재무제표 / 금액 단위: 백만원"
    src = {"rf": a.risk_free_rate, "erp": a.equity_risk_premium, "beta": a.beta,
           "size": a.size_premium, "wacc_override": a.wacc_override}
    # (name, label, value or formula, fmt, is_input, note)
    rows = [
        ("rf", "무위험이자율 (Rf)", src.get("rf"), PCT, True, "국고채 10년"),
        ("erp", "시장위험프리미엄 (ERP)", src.get("erp"), PCT, True, ""),
        ("beta", "베타 (β)", src.get("beta"), "0.00", True, ""),
        ("size_premium", "규모 프리미엄", src.get("size"), PCT, True, ""),
        ("ke", "자기자본비용 (Ke)", "=rf+beta*erp+size_premium", PCT, False, "CAPM"),
        ("kd", "세전 타인자본비용 (Kd)", a.pre_tax_cost_of_debt, PCT, True, "이자지급액/평균차입금 (공시)"),
        ("tax_rate", "법인세율", a.tax_rate, PCT, True, "과거 유효세율 평균 (공시)"),
        ("wd", "부채 비중 D/(D+E)", a.debt_weight, PCT, True, ""),
        ("wacc", "WACC", "=(1-wd)*ke+wd*kd*(1-tax_rate)", PCT, False, ""),
        ("tg", "영구성장률 (g)", a.terminal_growth, PCT, True, ""),
        ("exit_multiple", "Exit EV/EBITDA (0=영구성장법)", a.exit_ev_ebitda or 0, "0.0x", True, ""),
        ("mid_year", "기중 할인 (1=예, 0=아니오)", 1 if a.mid_year else 0, "0", True, ""),
        ("shares", "유통주식수 (주)", a.shares, NUM, True, "주식총수 현황: 발행주식 − 자기주식 (공시)"),
        ("price", "현재 주가 (원)", res.share_price or 0, NUM, True, ""),
        ("cash", "현금 및 단기금융자산", _m(res.cash_and_investments), NUM, True, f"{d.latest.year}년말 BS (공시)"),
        ("debt", "차입금 + 사채 + 리스부채", _m(res.total_debt), NUM, True, f"{d.latest.year}년말 BS (공시)"),
        ("nci", "비지배지분", _m(res.nci), NUM, True, f"{d.latest.year}년말 BS (공시)"),
    ]
    if src.get("wacc_override") is not None:
        rows[8] = ("wacc", "WACC (직접 입력)", src["wacc_override"], PCT, True, "override")
    ws.append([])
    ws.append(["항목", "값", "비고"])
    for c in ws[4]:
        c.font, c.fill = BOLD, HEAD_FILL
    for name, label, val, fmt, is_input, note in rows:
        ws.append([label, val, note])
        r = ws.max_row
        cell = ws.cell(r, 2)
        cell.number_format = fmt
        if is_input:
            cell.font = INPUT
        _name(wb, name, ws.title, f"B{r}")
    ws.append([])
    for note in a.notes + d.warnings:
        ws.append(["※ " + note])
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 44

    # ------------------------------------------------------------ DCF
    ws = wb.create_sheet("DCF")
    ws["A1"] = f"{d.corp_name} FCFF DCF (단위: 백만원)"
    ws["A1"].font = Font(bold=True, size=14)
    first_col, last_col = 3, 2 + n
    C = get_column_letter
    L = C(last_col)
    ws.cell(3, 1, "연도").font = BOLD
    ws.cell(3, 2, f"{d.latest.year}A").font = BOLD
    for i, y in enumerate(a.years):
        ws.cell(3, first_col + i, f"{y}E").font = BOLD
    for c in ws[3]:
        c.fill = HEAD_FILL

    driver_rows = [(5, "매출 성장률", a.revenue_growth), (6, "영업이익률", a.ebit_margin),
                   (7, "D&A / 매출", a.da_pct), (8, "CAPEX / 매출", a.capex_pct),
                   (9, "순운전자본 / 매출", a.nwc_pct)]
    ws.cell(4, 1, "추정 가정 (입력)").font = BOLD
    for r, label, vals in driver_rows:
        ws.cell(r, 1, label)
        for i, v in enumerate(vals):
            cell = ws.cell(r, first_col + i, v)
            cell.number_format, cell.font = PCT, INPUT

    last = d.latest
    ws.cell(5, 2, _growth_actual(d)).number_format = PCT
    ws.cell(6, 2, (last.ebit / last.revenue) if last.ebit is not None else None).number_format = PCT

    lines = {11: "매출액", 12: "영업이익 (EBIT)", 13: "(−) 법인세", 14: "NOPAT", 15: "(+) 감가상각비",
             16: "(−) CAPEX", 17: "순운전자본", 18: "(−) 순운전자본 증감", 19: "FCFF",
             20: "할인기간", 21: "할인계수", 22: "FCFF 현재가치", 24: "EBITDA"}
    for r, label in lines.items():
        ws.cell(r, 1, label)
    ws.cell(10, 1, "추정 손익 및 현금흐름").font = BOLD
    ws.cell(11, 2, _m(last.revenue))
    ws.cell(12, 2, _m(last.ebit))
    ws.cell(15, 2, _m(last.da))
    ws.cell(16, 2, _m(last.capex))
    ws.cell(17, 2, f"=B11*{C(first_col)}9")
    ws.cell(24, 2, "=B12+B15" if last.da is not None and last.ebit is not None else None)
    for i in range(n):
        col = first_col + i
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
    for col in range(1, last_col + 1):
        ws.cell(19, col).font = BOLD
        ws.cell(19, col).border = THIN

    ws.cell(26, 1, "밸류에이션").font = BOLD
    val_rows = [
        ("FCFF 현재가치 합계", f"=SUM({C(first_col)}22:{L}22)", NUM),
        ("잔존가치 산출용 FCFF (N+1)",
         f"={L}11*(1+tg)*{L}6*(1-tax_rate)+{L}11*(1+tg)*{L}7-{L}11*(1+tg)*{L}8-({L}11*(1+tg)*{L}9-{L}17)", NUM),
        ("잔존가치 (TV)", f"=IF(exit_multiple>0,{L}24*exit_multiple,B28/(wacc-tg))", NUM),
        ("TV 할인기간", f"=IF(exit_multiple>0,{n},{n}-IF(mid_year=1,0.5,0))", "0.00"),
        ("TV 현재가치", "=B29/(1+wacc)^B30", NUM),
        ("기업가치 (EV)", "=B27+B31", NUM),
        ("(+) 현금 및 단기금융자산", "=cash", NUM),
        ("(−) 차입금·사채·리스부채", "=debt", NUM),
        ("(−) 비지배지분", "=nci", NUM),
        ("주주가치", "=B32+B33-B34-B35", NUM),
        ("유통주식수 (주)", "=shares", NUM),
        ("주당 내재가치 (원)", "=B36*1000000/B37", NUM),
        ("현재 주가 (원)", "=price", NUM),
        ("상승여력", '=IF(B39>0,B38/B39-1,"")', PCT),
        ("EV 중 TV 비중", "=B31/B32", PCT),
        ("WACC", "=wacc", PCT),
    ]
    for i, (label, formula, fmt) in enumerate(val_rows):
        r = 27 + i
        ws.cell(r, 1, label)
        cell = ws.cell(r, 2, formula)
        cell.number_format = fmt
        if label.startswith(("기업가치", "주주가치", "주당")):
            ws.cell(r, 1).font = BOLD
            cell.font = BOLD
            cell.fill = KEY_FILL
    ws.column_dimensions["A"].width = 30
    for col in range(2, last_col + 1):
        ws.column_dimensions[C(col)].width = 14
    ws.freeze_panes = "B4"

    # ------------------------------------------------------------ Historical
    ws = wb.create_sheet("Historical")
    ws["A1"] = f"{d.corp_name} 과거 실적 (공시 추출, 단위: 백만원)"
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    ws.append(["항목"] + [y.year for y in d.years])
    for c in ws[3]:
        c.font, c.fill = BOLD, HEAD_FILL
    for spec in SPECS:
        ws.append([spec.label] + [_m(y.values.get(spec.key)) for y in d.years])
    ws.append([])
    derived = [
        ("EBITDA", lambda y: _m(y.ebitda), NUM),
        ("CAPEX 합계", lambda y: _m(y.capex), NUM),
        ("순운전자본", lambda y: _m(y.nwc), NUM),
        ("총차입금 (리스 포함)", lambda y: _m(y.total_debt), NUM),
        ("영업이익률", lambda y: y.ebit / y.revenue if y.ebit is not None and y.revenue else None, PCT),
        ("D&A / 매출", lambda y: y.da / y.revenue if y.da is not None and y.revenue else None, PCT),
        ("CAPEX / 매출", lambda y: y.capex / y.revenue if y.capex is not None and y.revenue else None, PCT),
        ("NWC / 매출", lambda y: y.nwc / y.revenue if y.nwc is not None and y.revenue else None, PCT),
        ("유효세율", lambda y: y.tax_rate, PCT),
    ]
    start = ws.max_row + 1
    for label, fn, _ in derived:
        ws.append([label] + [fn(y) for y in d.years])
    for r in range(4, ws.max_row + 1):
        fmt = PCT if r >= start + 4 else NUM
        for col in range(2, len(d.years) + 2):
            ws.cell(r, col).number_format = fmt
    ws.column_dimensions["A"].width = 26
    for col in range(2, len(d.years) + 2):
        ws.column_dimensions[C(col)].width = 14

    # ------------------------------------------------------------ Sensitivity
    ws = wb.create_sheet("Sensitivity")
    ws["A1"] = "주당 내재가치 민감도 (원) — WACC × 영구성장률 (Python 계산 정적 값)"
    ws["A1"].font = BOLD
    s = res.sensitivity
    ws.cell(3, 1, "WACC \\ g").font = BOLD
    for j, g in enumerate(s["growth"]):
        c = ws.cell(3, 2 + j, g)
        c.number_format, c.font, c.fill = PCT, BOLD, HEAD_FILL
    for i, w in enumerate(s["wacc"]):
        c = ws.cell(4 + i, 1, w)
        c.number_format, c.font, c.fill = PCT, BOLD, HEAD_FILL
        for j, v in enumerate(s["per_share"][i]):
            cell = ws.cell(4 + i, 2 + j, v)
            cell.number_format = NUM
            if i == len(s["wacc"]) // 2 and j == len(s["growth"]) // 2:
                cell.fill, cell.font = KEY_FILL, BOLD
    ws.column_dimensions["A"].width = 12

    # ------------------------------------------------------------ Sources
    ws = wb.create_sheet("Sources")
    ws["A1"] = f"계정 매핑 내역 ({d.latest.year}년, DART fnlttSinglAcntAll)"
    ws["A1"].font = BOLD
    ws.append([])
    ws.append(["항목", "값 (백만원)", "사용된 공시 계정"])
    for c in ws[3]:
        c.font, c.fill = BOLD, HEAD_FILL
    for spec in SPECS:
        v = last.values.get(spec.key)
        ws.append([spec.label, _m(v), last.sources.get(spec.key, "— 찾지 못함")])
        ws.cell(ws.max_row, 2).number_format = NUM
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 110
    for row in ws.iter_rows(min_row=4):
        row[2].alignment = Alignment(wrap_text=False)

    wb.move_sheet("DCF", offset=-1)
    wb.active = 0
    wb.save(path)


def _growth_actual(d):
    if len(d.years) >= 2 and d.years[-2].revenue and d.latest.revenue:
        return d.latest.revenue / d.years[-2].revenue - 1
    return None
