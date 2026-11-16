"""가상의 회사 '샘플전자' 의 OpenDART 응답(JSON)을 캐시 형식으로 생성한다.

실제 회사 데이터가 아니며, 오프라인 데모/테스트용이다.
    python examples/build_sample_cache.py  →  examples/sample_cache/
"""
from __future__ import annotations

import json
from pathlib import Path

OUT = Path(__file__).parent / "sample_cache"
CORP = {"corp_code": "00999999", "corp_name": "샘플전자", "stock_code": "999999", "modify_date": "20250301"}

# 단위: 원.  (연도: 값)  - 2020~2024
B = 100_000_000  # 1억
DATA = {
    2020: dict(rev=8000, op=720, pt=700, tax=170, ni=530, dep=380, amort=40, capex_ppe=520, capex_int=50,
               int_paid=45, ca=4200, cash=900, stfi=300, cl=2600, stb=600, cpd=200, bonds=800, ltb=500,
               lease_c=40, lease_nc=120, eq=6000, owners=5700, nci=300),
    2021: dict(rev=8900, op=830, pt=810, tax=195, ni=615, dep=400, amort=42, capex_ppe=560, capex_int=55,
               int_paid=44, ca=4600, cash=1000, stfi=320, cl=2800, stb=550, cpd=220, bonds=800, ltb=480,
               lease_c=42, lease_nc=125, eq=6550, owners=6230, nci=320),
    2022: dict(rev=9700, op=870, pt=840, tax=205, ni=635, dep=425, amort=45, capex_ppe=610, capex_int=60,
               int_paid=52, ca=5000, cash=1050, stfi=350, cl=3050, stb=600, cpd=250, bonds=900, ltb=450,
               lease_c=45, lease_nc=130, eq=7100, owners=6760, nci=340),
    2023: dict(rev=10300, op=950, pt=920, tax=225, ni=695, dep=450, amort=48, capex_ppe=640, capex_int=62,
               int_paid=58, ca=5300, cash=1150, stfi=380, cl=3200, stb=580, cpd=260, bonds=900, ltb=430,
               lease_c=47, lease_nc=135, eq=7700, owners=7340, nci=360),
    2024: dict(rev=11200, op=1060, pt=1030, tax=250, ni=780, dep=470, amort=50, capex_ppe=680, capex_int=65,
               int_paid=60, ca=5750, cash=1250, stfi=400, cl=3400, stb=560, cpd=270, bonds=1000, ltb=400,
               lease_c=50, lease_nc=140, eq=8400, owners=8020, nci=380),
}


def fmt(x):
    return f"{int(round(x * B)):,}" if x is not None else ""


def rows_for(year: int) -> list[dict]:
    cur, prev = DATA[year], DATA.get(year - 1)
    out = []

    def add(sj, sj_nm, acc_id, name, key=None, val=None, neg=False):
        v = cur[key] if key else val
        pv = (prev[key] if prev and key else None)
        s = -1 if neg else 1
        out.append({
            "rcept_no": f"{year + 1}0315000999", "reprt_code": "11011", "bsns_year": str(year),
            "corp_code": CORP["corp_code"], "sj_div": sj, "sj_nm": sj_nm,
            "account_id": acc_id, "account_nm": name, "account_detail": "-",
            "thstrm_nm": f"제 {year - 1990} 기", "thstrm_amount": fmt(v * s if v is not None else None),
            "frmtrm_nm": f"제 {year - 1991} 기", "frmtrm_amount": fmt(pv * s if pv is not None else None),
            "ord": str(len(out) + 1), "currency": "KRW",
        })

    BSN, ISN, CFN = "재무상태표", "손익계산서", "현금흐름표"
    c = cur
    add("BS", BSN, "ifrs-full_CurrentAssets", "유동자산", "ca")
    add("BS", BSN, "ifrs-full_CashAndCashEquivalents", "현금및현금성자산", "cash")
    add("BS", BSN, "dart_ShortTermDepositsNotClassifiedAsCashEquivalents", "단기금융상품", "stfi")
    add("BS", BSN, "ifrs-full_Inventories", "재고자산", val=c["ca"] * 0.35)
    add("BS", BSN, "ifrs-full_NoncurrentAssets", "비유동자산", val=c["eq"] + c["cl"] + 2000 - c["ca"])
    add("BS", BSN, "ifrs-full_Assets", "자산총계", val=c["eq"] + c["cl"] + 2000)
    add("BS", BSN, "ifrs-full_CurrentLiabilities", "유동부채", "cl")
    add("BS", BSN, "ifrs-full_ShorttermBorrowings", "단기차입금", "stb")
    add("BS", BSN, "-표준계정코드 미사용-", "유동성장기부채", "cpd")
    add("BS", BSN, "ifrs-full_CurrentLeaseLiabilities", "유동리스부채", "lease_c")
    add("BS", BSN, "ifrs-full_NoncurrentLiabilities", "비유동부채", val=2000)
    add("BS", BSN, "dart_BondsIssued", "사채", "bonds")
    add("BS", BSN, "ifrs-full_LongtermBorrowings", "장기차입금", "ltb")
    add("BS", BSN, "ifrs-full_NoncurrentLeaseLiabilities", "비유동리스부채", "lease_nc")
    add("BS", BSN, "ifrs-full_Liabilities", "부채총계", val=c["cl"] + 2000)
    add("BS", BSN, "ifrs-full_EquityAttributableToOwnersOfParent", "지배기업의 소유주에게 귀속되는 자본", "owners")
    add("BS", BSN, "ifrs-full_NoncontrollingInterests", "비지배지분", "nci")
    add("BS", BSN, "ifrs-full_Equity", "자본총계", "eq")

    add("IS", ISN, "ifrs-full_Revenue", "수익(매출액)", "rev")
    add("IS", ISN, "ifrs-full_CostOfSales", "매출원가", val=c["rev"] * 0.7)
    add("IS", ISN, "ifrs-full_GrossProfit", "매출총이익", val=c["rev"] * 0.3)
    add("IS", ISN, "dart_OperatingIncomeLoss", "영업이익(손실)", "op")
    add("IS", ISN, "ifrs-full_FinanceCosts", "금융비용", val=c["int_paid"] * 1.1)
    add("IS", ISN, "ifrs-full_ProfitLossBeforeTax", "법인세비용차감전순이익(손실)", "pt")
    add("IS", ISN, "ifrs-full_IncomeTaxExpenseContinuingOperations", "법인세비용", "tax")
    add("IS", ISN, "ifrs-full_ProfitLoss", "당기순이익(손실)", "ni")

    add("CF", CFN, "ifrs-full_CashFlowsFromUsedInOperatingActivities", "영업활동현금흐름",
        val=c["ni"] + c["dep"] + c["amort"])
    add("CF", CFN, "-표준계정코드 미사용-", "감가상각비", "dep")
    add("CF", CFN, "-표준계정코드 미사용-", "무형자산상각비", "amort")
    add("CF", CFN, "-표준계정코드 미사용-", "대손상각비", val=5)        # 매칭에서 제외되어야 함
    add("CF", CFN, "ifrs-full_InterestPaidClassifiedAsOperatingActivities", "이자의 지급", "int_paid", neg=True)
    add("CF", CFN, "ifrs-full_PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities",
        "유형자산의 취득", "capex_ppe", neg=True)
    add("CF", CFN, "ifrs-full_PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities",
        "무형자산의 취득", "capex_int", neg=True)
    return out


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "corpCode.json").write_text(json.dumps([CORP], ensure_ascii=False, indent=1), encoding="utf-8")
    for year in DATA:
        for fs_div in ("CFS",):
            payload = {"status": "000", "message": "정상", "list": rows_for(year)}
            name = (f"fnlttSinglAcntAll__bsns_year-{year}_corp_code-{CORP['corp_code']}"
                    f"_fs_div-{fs_div}_reprt_code-11011.json")
            (OUT / name).write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")
        shares = {"status": "000", "message": "정상", "list": [
            {"se": "보통주", "istc_totqy": "50,000,000", "tesstk_co": "2,000,000", "distb_stock_co": "48,000,000"},
            {"se": "우선주", "istc_totqy": "5,000,000", "tesstk_co": "-", "distb_stock_co": "5,000,000"},
            {"se": "합계", "istc_totqy": "55,000,000", "tesstk_co": "2,000,000", "distb_stock_co": "53,000,000"},
        ]}
        name = f"stockTotqySttus__bsns_year-{year}_corp_code-{CORP['corp_code']}_reprt_code-11011.json"
        (OUT / name).write_text(json.dumps(shares, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"wrote sample cache → {OUT}")


if __name__ == "__main__":
    main()
