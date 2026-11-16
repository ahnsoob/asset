# DART 공시 기반 DCF 밸류에이션

금융감독원 전자공시(OpenDART) API에서 재무제표와 주식 수를 자동으로 가져와 **FCFF DCF** 로 주당 내재가치를 계산합니다.
결과는 콘솔 리포트, **수식이 살아있는 Excel**, JSON 으로 받을 수 있습니다.

```
OpenDART API ──► dart.py (호출 + 캐시) ──► extract.py + accounts.py (계정 매핑)
                                                │
                     연도별 표준 재무데이터 (매출, EBIT, D&A, CAPEX, NWC, 차입금, 현금, 주식수 …)
                                                │
                    model.py (가정 자동 산출 → WACC → FCFF 추정 → TV → EV → 주당가치)
                                                │
                       report.py (콘솔)   excel.py (xlsx, 수식)   cli.py (--json)
```

## 빠른 시작

```bash
pip install -r requirements.txt

# 1) API 키 없이 가상 회사 '샘플전자' 로 돌려보기
python -m dcf 샘플전자 --year 2024 --offline --cache-dir examples/sample_cache --excel sample.xlsx

# 2) 실제 회사 (https://opendart.fss.or.kr 에서 API 키 발급)
export DART_API_KEY=발급받은키
python -m dcf 005930 --year 2024 --beta 1.0 --price 55000 --excel samsung.xlsx
python -m dcf 삼성전자 --year 2024 --growth 8,7,6,4,3 --margin 12 --tg 2 --wacc 8.5
python -m dcf 삼성전자 --config examples/assumptions.example.json --json result.json
```

회사는 **회사명 / 종목코드(6자리) / DART 고유번호(8자리)** 로 지정합니다. 퍼센트는 `8`, `8%`, `0.08` 모두 같은 값입니다.

## 공시에서 자동 추출되는 항목

| 항목 | DART 출처 | 용도 |
|---|---|---|
| 매출액, 영업이익, 법인세차감전이익, 법인세비용 | 손익계산서 (`fnlttSinglAcntAll`) | 성장률, 마진, 유효세율 |
| 감가상각비·무형자산상각비 | 현금흐름표 조정 항목 | D&A/매출 |
| 유형·무형자산의 취득 | 현금흐름표 투자활동 | CAPEX/매출 |
| 이자의 지급 | 현금흐름표 | 세전 타인자본비용 |
| 유동자산·유동부채·현금·단기금융상품·단기차입금·유동성장기부채 | 재무상태표 | 영업 순운전자본 |
| 단기/장기차입금, 사채, 리스부채, 비지배지분, 자본 | 재무상태표 | 순차입금, 자본구조 |
| 발행주식총수 − 자기주식 | 주식의 총수 현황 (`stockTotqySttus`) | 주당가치 분모 |

- 계정은 XBRL `account_id` 로 먼저 찾고, 없으면 계정명(`Ⅰ.` 등 번호, 공백 제거 후) 정규식으로 찾습니다. 규칙은 `dcf/accounts.py` 에 있으니 회사별로 빠지는 계정이 있으면 패턴만 추가하면 됩니다.
- 어떤 공시 계정이 쓰였는지는 Excel **Sources** 시트에서 확인할 수 있습니다.
- 연결재무제표(CFS)가 없으면 별도(OFS)로 자동 전환합니다 (`--fs` 로 고정 가능).
- API 응답은 `.dart_cache/` 에 저장되어 같은 요청은 다시 호출하지 않습니다 (API 키는 캐시에 저장되지 않음).

## 가정 (지정하지 않으면 공시 데이터에서 자동 산출)

| 가정 | 기본 산출 방식 | CLI |
|---|---|---|
| 매출 성장률 | 과거 매출 CAGR(−10%~30% 로 제한) → 마지막 해에 영구성장률로 선형 수렴 | `--growth` |
| 영업이익률, D&A/매출, CAPEX/매출, NWC/매출 | 최근 3개년 평균 | `--margin --da --capex --nwc` |
| 법인세율 | 최근 3개년 유효세율 평균, 이상치면 26.4% (법인세 24% + 지방세) | `--tax` |
| 자기자본비용 | CAPM: Rf 3.0% + β 1.0 × ERP 6.0% (+ 규모 프리미엄) | `--rf --beta --erp --size-premium` |
| 타인자본비용 | 이자지급액 / 평균 차입금 (Rf ~ Rf+8% 로 제한), 없으면 Rf+1.5% | `--kd` |
| 자본구조 | 주가 입력 시 시가총액, 아니면 장부가 기준 D/(D+E) | `--debt-weight` |
| 잔존가치 | 영구성장 2.0% (또는 Exit EV/EBITDA 배수) | `--tg`, `--exit-multiple` |
| 할인 | 기중(mid-year) 할인 | `--no-mid-year` |

연도별 값은 `--growth 10,8,6,4,3` 처럼 쉼표로 넣거나 JSON 설정 파일(`--config`)에 리스트로 넣습니다.
**베타, 무위험이자율, 시장위험프리미엄, 주가는 공시에 없는 시장 데이터**이므로 직접 넣는 것을 권장합니다.

## Excel 출력

- **DCF**: 추정 손익·FCFF·현재가치·EV→주주가치→주당가치. 파란 글씨(성장률, 마진 등)를 바꾸면 자동 재계산됩니다.
- **Assumptions**: WACC 구성요소, 세율, 순차입금, 주식수 (이름 정의된 셀, 예: `wacc`, `tg`)
- **Historical**: 공시에서 추출한 과거 실적과 비율
- **Sensitivity**: WACC × 영구성장률 민감도 (Python 계산 정적 값)
- **Sources**: 항목별로 매칭된 공시 계정

## 계산 방식

```
FCFF   = EBIT × (1 − t) + D&A − CAPEX − ΔNWC
TV     = FCFF(N+1) / (WACC − g)            ← N+1년은 매출 × (1+g) 로 정규화해 다시 계산
EV     = Σ FCFF_t / (1+WACC)^(t−0.5) + TV / (1+WACC)^(N−0.5)     (기중 할인 시)
주주가치 = EV + 현금·단기금융자산 − (차입금 + 사채 + 리스부채) − 비지배지분
주당가치 = 주주가치 / (발행주식총수 − 자기주식)
```

## 한계 / 주의

- 평가 기준일은 기준연도 말입니다 (분기 데이터 미반영).
- 리스부채를 차입금에 포함하는 IFRS 16 방식이지만, 사용권자산 취득은 CAPEX 에 포함되지 않습니다.
- 금융업(은행·보험·증권)은 FCFF DCF 에 맞지 않습니다 (매출/운전자본 개념이 다름).
- 관계기업투자·투자부동산 같은 비영업자산은 EV 에 따로 더하지 않습니다. 비중이 크면 결과에 수동으로 반영하세요.
- 현금흐름표에 감가상각비가 따로 나오지 않는 회사는 D&A = CAPEX 로 가정하고 경고를 표시합니다.

## 테스트

```bash
pip install pytest
python -m pytest
python examples/build_sample_cache.py   # 샘플 캐시 재생성
```
