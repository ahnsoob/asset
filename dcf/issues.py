"""데이터 점검 항목 — 공시에서 바로 가져오지 못해 가정·대체·보완한 값.

DCF 결과를 보기 전에 가장 먼저 확인해야 할 목록이다 (콘솔 맨 위, Excel 첫 시트 'Check').
"""
from __future__ import annotations

from dataclasses import dataclass, field

# 표시 순서 = 중요도 순
KINDS = {
    "assumed": ("가정", "⚠", "공시에서 추출 실패 → 대체 가정값 사용"),
    "default": ("기본값", "⚠", "공시에 없는 시장 데이터 → 입력하지 않아 기본값 사용"),
    "adjusted": ("조정", "△", "공시 기반 값이 범위를 벗어나 제한하거나 일부 연도만 사용"),
    "filled": ("보완", "✚", "본문 재무제표에 없어 주석 등 다른 공시에서 찾아 채움"),
    "zero": ("0 처리", "·", "해당 계정이 공시에 없어 0 으로 계산"),
}
ORDER = list(KINDS)


@dataclass
class DataIssue:
    kind: str                       # KINDS 의 키
    item: str                       # 항목 (예: "D&A / 매출")
    value: str                      # 실제로 사용한 값 (사람이 읽는 문자열)
    reason: str                     # 왜 이렇게 됐는지
    targets: list[str] = field(default_factory=list)
    # targets: Excel 강조 위치. 예) "driver:da_pct", "assump:tax_rate", "hist:depreciation_amortization:2025"

    @property
    def label(self) -> str:
        return KINDS[self.kind][0]

    @property
    def icon(self) -> str:
        return KINDS[self.kind][1]

    def to_dict(self) -> dict:
        return {"kind": self.kind, "label": self.label, "item": self.item, "value": self.value,
                "reason": self.reason, "targets": list(self.targets)}


def sort_issues(issues: list[DataIssue]) -> list[DataIssue]:
    return sorted(issues, key=lambda i: ORDER.index(i.kind))


def needs_attention(issues: list[DataIssue]) -> list[DataIssue]:
    """꼭 확인해야 할 항목 (가정·기본값·조정)."""
    return [i for i in issues if i.kind in ("assumed", "default", "adjusted")]
