"""피어(유사기업) 베타 언레버링 → 리레버링 (Hamada 식).

    언레버드 β_U = β_L / (1 + (1 − t) × D/E)          ← 피어별 자본구조 효과 제거
    대상 β_L    = 대표 β_U × (1 + (1 − t_대상) × D/E_대상)   ← 대상회사 자본구조 반영

피어의 언레버드 베타를 이미 알고 있으면(블룸버그 등) unlevered_beta 로 넣으면 D/E 없이 그대로 쓴다.
피어의 D/E 는 직접 입력하거나, DART 고유번호/종목코드를 주면 공시(재무상태표·주식수)에서 자동 산출한다.
  - D = 차입금 + 사채 + 리스부채 (대상회사와 같은 정의)
  - E = 주가 × 유통주식수 (주가를 주면 시가총액), 없으면 장부가 자본
"""
from __future__ import annotations

from dataclasses import dataclass, field
from statistics import mean, median

PEER_DEFAULT_TAX = 0.264   # 피어 세율 미지정 시: 법인세 최고세율 24% + 지방소득세


@dataclass
class Peer:
    name: str
    beta: float | None = None         # 관측(레버드) 베타
    de_ratio: float | None = None     # D/E. 없으면 debt / equity
    debt: float | None = None         # 원
    equity: float | None = None       # 원 (시가총액 권장)
    tax_rate: float | None = None
    company: str | None = None        # DART 조회용 회사명/종목코드/고유번호
    price: float | None = None        # 시가총액 계산용 주가 (DART 조회 시)
    source: str = ""                  # D/E 출처 메모
    unlevered_beta: float | None = None   # 직접 입력한 언레버드 베타 (있으면 D/E 없이 그대로 사용)

    @classmethod
    def from_dict(cls, d: dict) -> "Peer":
        known = {k: v for k, v in d.items() if k in cls.__dataclass_fields__}
        if "name" not in known:
            known["name"] = str(known.get("company", "peer"))
        return cls(**known)

    @property
    def de(self) -> float | None:
        if self.de_ratio is not None:
            return self.de_ratio
        if self.debt is not None and self.equity:
            return self.debt / self.equity
        return None


def unlever(beta_l: float, de: float, tax: float) -> float:
    return beta_l / (1 + (1 - tax) * de)


def relever(beta_u: float, de: float, tax: float) -> float:
    return beta_u * (1 + (1 - tax) * de)


@dataclass
class PeerBetaResult:
    rows: list[dict]                  # name, beta_l, de, tax, beta_u, source
    method: str
    unlevered: float                  # 대표 언레버드 베타
    target_de: float
    target_tax: float
    relevered: float
    excluded: list[str] = field(default_factory=list)
    default_tax: float = PEER_DEFAULT_TAX


def peer_beta(peers: list[Peer], target_de: float, target_tax: float, method: str = "median",
              default_tax: float = PEER_DEFAULT_TAX) -> PeerBetaResult:
    rows, excluded = [], []
    for p in peers:
        de = p.de
        tax = p.tax_rate if p.tax_rate is not None else default_tax
        if p.unlevered_beta is not None:          # 언레버드 베타를 직접 받은 경우
            rows.append(dict(name=p.name, beta_l=p.beta, de=de, tax=tax, debt=p.debt, equity=p.equity,
                             beta_u=p.unlevered_beta, given_u=True,
                             source=p.source or "언레버드 β 직접 입력"))
            continue
        if de is None or de < 0 or p.beta is None:
            excluded.append(f"{p.name}: 언레버드 β 도, (베타 + D/E) 도 없어 제외")
            continue
        rows.append(dict(name=p.name, beta_l=p.beta, de=de, tax=tax, debt=p.debt, equity=p.equity,
                         beta_u=unlever(p.beta, de, tax), given_u=False, source=p.source))
    if not rows:
        raise ValueError("사용할 수 있는 피어가 없습니다 (각 피어에 언레버드 β, 또는 베타 + D/E 가 필요).")
    us = [r["beta_u"] for r in rows]
    if method == "median":
        bu = median(us)
    elif method == "mean":
        bu = mean(us)
    else:
        raise ValueError(f"알 수 없는 집계 방식: {method} (median | mean)")
    return PeerBetaResult(rows=rows, method=method, unlevered=bu, target_de=target_de,
                          target_tax=target_tax, relevered=relever(bu, target_de, target_tax),
                          excluded=excluded, default_tax=default_tax)


def fill_peers_from_dart(client, peers: list[Peer], year: int) -> list[str]:
    """company 가 있고 D/E 가 비어 있는 피어의 부채·자본을 DART 에서 채운다. 경고 메시지 목록 반환."""
    from .dart import REPRT_CODES, DartError
    from .extract import extract_year, shares_outstanding

    notes = []
    reprt = REPRT_CODES["annual"]
    for p in peers:
        if p.de is not None or p.unlevered_beta is not None or not p.company:
            continue
        try:
            corp = client.find_corp(p.company)
            code = corp["corp_code"]
            div = "CFS"
            rows = client.financial_statements(code, year, reprt, "CFS")
            if not rows:
                div = "OFS"
                rows = client.financial_statements(code, year, reprt, "OFS")
            if not rows:
                notes.append(f"피어 {p.name}: {year}년 재무제표 없음 → 제외")
                continue
            y = extract_year(rows, year)
            if p.debt is None:
                p.debt = y.total_debt
            if p.equity is None:
                if p.price:
                    shares = shares_outstanding(client.share_status(code, year, reprt))
                    if shares:
                        p.equity = p.price * shares
                        basis = "시가총액"
                if p.equity is None:
                    p.equity = y.get("owners_equity") or y.get("total_equity")
                    basis = "장부가 자본"
            else:
                basis = "입력 자본"
            if p.name == p.company:
                p.name = corp["corp_name"]
            p.source = f"DART {year} {div} 차입금+리스 / {basis}"
        except DartError as e:
            notes.append(f"피어 {p.name}: DART 조회 실패 ({e}) → 제외")
    return notes
