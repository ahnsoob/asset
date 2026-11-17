"""OpenDART(금융감독원 전자공시) API 클라이언트.

- 응답은 cache_dir 에 JSON 으로 저장되어 같은 요청은 다시 호출하지 않는다.
- offline=True 이면 캐시만 사용한다 (네트워크/API 키 불필요).

API 키 발급: https://opendart.fss.or.kr  →  환경변수 DART_API_KEY 로 전달.
"""
from __future__ import annotations

import gzip
import io
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import requests

BASE_URL = "https://opendart.fss.or.kr/api"

# 보고서 코드
REPRT_CODES = {
    "annual": "11011",  # 사업보고서
    "half": "11012",    # 반기보고서
    "q1": "11013",      # 1분기보고서
    "q3": "11014",      # 3분기보고서
}

STATUS_OK = "000"
STATUS_NO_DATA = "013"


class DartError(RuntimeError):
    pass


def _decode(raw: bytes) -> str:
    """원문 XML 은 선언이 utf-8 이어도 실제로는 EUC-KR(cp949) 인 경우가 있다 (2022년 이전 공시 등)."""
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("cp949", errors="replace")


class DartClient:
    def __init__(self, api_key: str | None = None, cache_dir: str | Path = ".dart_cache",
                 offline: bool = False, session: requests.Session | None = None,
                 timeout: float = 30.0):
        self.api_key = api_key or os.environ.get("DART_API_KEY")
        self.cache_dir = Path(cache_dir)
        self.offline = offline
        self.session = session or requests.Session()
        self.timeout = timeout

    # ------------------------------------------------------------------ cache
    def _cache_path(self, endpoint: str, params: dict) -> Path:
        parts = [f"{k}-{params[k]}" for k in sorted(params) if k != "crtfc_key"]
        name = endpoint + ("__" + "_".join(parts) if parts else "")
        name = re.sub(r"[^\w\-.]", "_", name)
        return self.cache_dir / f"{name}.json"

    def _read_cache(self, path: Path):
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
        return None

    def _write_cache(self, path: Path, data) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")

    def _require_online(self, what: str) -> None:
        if self.offline:
            raise DartError(f"오프라인 모드인데 캐시에 없음: {what}")
        if not self.api_key:
            raise DartError("DART API 키가 없습니다. 환경변수 DART_API_KEY 를 설정하거나 --api-key 를 사용하세요.")

    # ------------------------------------------------------------------- core
    def get_json(self, endpoint: str, params: dict) -> dict:
        path = self._cache_path(endpoint, params)
        cached = self._read_cache(path)
        if cached is not None:
            return cached
        self._require_online(path.name)

        resp = self.session.get(f"{BASE_URL}/{endpoint}.json",
                                params={**params, "crtfc_key": self.api_key},
                                timeout=self.timeout)
        resp.raise_for_status()
        data = resp.json()
        status = data.get("status")
        if status == STATUS_OK:
            self._write_cache(path, data)
            return data
        if status == STATUS_NO_DATA:
            data = {"status": STATUS_NO_DATA, "message": data.get("message", ""), "list": []}
            self._write_cache(path, data)
            return data
        raise DartError(f"DART 오류 [{status}] {data.get('message')} (endpoint={endpoint}, params={params})")

    # ------------------------------------------------------------- corp code
    def corp_codes(self) -> list[dict]:
        """전체 회사 고유번호 목록 (corp_code, corp_name, stock_code, modify_date)."""
        path = self.cache_dir / "corpCode.json"
        cached = self._read_cache(path)
        if cached is not None:
            return cached
        self._require_online("corpCode.json")

        resp = self.session.get(f"{BASE_URL}/corpCode.xml",
                                params={"crtfc_key": self.api_key}, timeout=self.timeout * 4)
        resp.raise_for_status()
        if not zipfile.is_zipfile(io.BytesIO(resp.content)):
            raise DartError(f"corpCode 다운로드 실패: {resp.content[:300]!r}")
        with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
            xml_bytes = zf.read(zf.namelist()[0])
        root = ET.fromstring(xml_bytes)
        corps = []
        for el in root.iter("list"):
            corps.append({
                "corp_code": (el.findtext("corp_code") or "").strip(),
                "corp_name": (el.findtext("corp_name") or "").strip(),
                "stock_code": (el.findtext("stock_code") or "").strip(),
                "modify_date": (el.findtext("modify_date") or "").strip(),
            })
        self._write_cache(path, corps)
        return corps

    def find_corp(self, query: str) -> dict:
        """종목코드(6자리) / 고유번호(8자리) / 회사명으로 회사 검색. 상장사 우선."""
        q = query.strip()
        corps = self.corp_codes()
        if re.fullmatch(r"\d{6}", q):
            hits = [c for c in corps if c["stock_code"] == q]
        elif re.fullmatch(r"\d{8}", q):
            hits = [c for c in corps if c["corp_code"] == q]
        else:
            hits = [c for c in corps if c["corp_name"] == q]
            if not hits:
                hits = [c for c in corps if q in c["corp_name"]]
        if not hits:
            raise DartError(f"회사를 찾을 수 없습니다: {query}")
        listed = [c for c in hits if c["stock_code"]]
        pool = listed or hits
        if len(pool) > 1 and not any(c["corp_name"] == q for c in pool):
            names = ", ".join(f"{c['corp_name']}({c['stock_code'] or c['corp_code']})" for c in pool[:10])
            raise DartError(f"여러 회사가 검색됩니다. 정확한 이름이나 종목코드를 쓰세요: {names}")
        exact = [c for c in pool if c["corp_name"] == q]
        return (exact or pool)[0]

    # ------------------------------------------------------------ endpoints
    def financial_statements(self, corp_code: str, year: int, reprt_code: str = "11011",
                             fs_div: str = "CFS") -> list[dict]:
        """단일회사 전체 재무제표 (fnlttSinglAcntAll). fs_div: CFS(연결) / OFS(별도)."""
        data = self.get_json("fnlttSinglAcntAll", {
            "corp_code": corp_code, "bsns_year": str(year),
            "reprt_code": reprt_code, "fs_div": fs_div,
        })
        return data.get("list", [])

    def annual_reports(self, corp_code: str, year: int) -> list[str]:
        """해당 사업연도 사업보고서 접수번호들 (정정 포함, 최근 접수 순)."""
        data = self.get_json("list", {
            "corp_code": corp_code, "bgn_de": f"{year + 1}0101", "end_de": f"{year + 2}0630",
            "pblntf_ty": "A", "page_count": "100",
        })
        hits = [r for r in data.get("list", [])
                if re.search(rf"사업보고서\s*\({year}\.", r.get("report_nm", ""))]
        hits.sort(key=lambda r: r.get("rcept_no", ""), reverse=True)
        return [r["rcept_no"] for r in hits]

    def document_xml(self, rcept_no: str) -> str:
        """공시서류 원문 (document.xml) — 사업보고서 전체(주석 포함) XML 텍스트.

        수 MB 라서 캐시에는 gzip 으로 저장한다.
        """
        path = self.cache_dir / f"document__rcept_no-{re.sub(r'[^0-9]', '', rcept_no)}.xml.gz"
        if path.exists():
            return gzip.decompress(path.read_bytes()).decode("utf-8", errors="replace")
        self._require_online(path.name)

        resp = self.session.get(f"{BASE_URL}/document.xml",
                                params={"crtfc_key": self.api_key, "rcept_no": rcept_no},
                                timeout=self.timeout * 4)
        resp.raise_for_status()
        if not zipfile.is_zipfile(io.BytesIO(resp.content)):
            try:
                err = resp.json()
                msg = f"[{err.get('status')}] {err.get('message')}"
            except ValueError:
                msg = repr(resp.content[:200])
            raise DartError(f"공시 원문 다운로드 실패 (rcept_no={rcept_no}): {msg}")
        with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
            # 본문 XML 이 가장 크다 (첨부 감사보고서 등은 별도 파일)
            names = sorted((n for n in zf.namelist() if n.lower().endswith(".xml")),
                           key=lambda n: zf.getinfo(n).file_size, reverse=True)
            if not names:
                raise DartError(f"공시 원문에 XML 이 없습니다 (rcept_no={rcept_no})")
            raw = zf.read(names[0])
        text = _decode(raw)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(gzip.compress(text.encode("utf-8")))
        return text

    def share_status(self, corp_code: str, year: int, reprt_code: str = "11011") -> list[dict]:
        """주식의 총수 현황 (stockTotqySttus)."""
        data = self.get_json("stockTotqySttus", {
            "corp_code": corp_code, "bsns_year": str(year), "reprt_code": reprt_code,
        })
        return data.get("list", [])
