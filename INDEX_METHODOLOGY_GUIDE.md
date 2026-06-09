# 지수 방법론 분석 가이드

> 작성일: 2026-06-09  
> 분석 대상: 100개 지수 방법론 PDF (Google Drive)  
> 산출물: `index_methodology_analysis_v2_2026.xlsx`

---

## 1. Google Drive PDF 접근 방법

### 1-1. 환경 구성

Claude Code on the web(원격 실행 환경)에서 Google Drive MCP 서버를 통해 PDF를 읽었다.  
**외부 인터넷은 차단**되어 있으나, 허용된 MCP 서버(Google Drive, GitHub)는 사용 가능하다.

```
환경: Claude Code on the web (격리된 원격 컨테이너)
MCP 서버 ID: mcp__ed91be2c-7ea9-43ef-ac23-8441787dbd0d
허용 도구: mcp__ed91be2c-...__read_file_content
차단 항목: 외부 HTTP(motherduck, duckdb extensions, google.com 등)
```

### 1-2. 파일 목록 확보

먼저 Drive의 파일 ID 목록을 TSV로 확보해야 한다.

```
형식: 순번 \t 파일명 \t Google_Drive_file_id
예시:
1   1_KRX_KOSPI200.pdf          1abc...xyz
2   2_KRX_KOSDAQ150.pdf         2def...uvw
```

파일 목록 위치: `/tmp/all_files_numbered.tsv`

### 1-3. PDF 읽기 (MCP 도구)

```python
# Claude Code 내부에서 MCP 도구로 읽는 방법 (에이전트 프롬프트로 지시)
# 도구명: mcp__ed91be2c-7ea9-43ef-ac23-8441787dbd0d__read_file_content
# 파라미터: {"file_id": "<Google Drive file_id>"}
```

**주의**: 한 에이전트당 20개씩 병렬 처리(5개 에이전트 × 20파일)가 효율적이다.  
100개를 직렬로 처리하면 컨텍스트 한도 초과 위험이 있다.

### 1-4. 로컬 환경에서 재현하는 방법

오프라인/로컬에서 동일 작업을 하려면:

```python
# 방법 A: Google Drive API (service account)
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

creds = Credentials.from_service_account_file('service_account.json',
    scopes=['https://www.googleapis.com/auth/drive.readonly'])
service = build('drive', 'v3', credentials=creds)

# 파일 다운로드
file_id = '1abc...xyz'
request = service.files().get_media(fileId=file_id)
content = request.execute()  # bytes

# 방법 B: PyDrive2
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

gauth = GoogleAuth()
drive = GoogleDrive(gauth)
f = drive.CreateFile({'id': file_id})
f.GetContentFile('local_copy.pdf')
```

```python
# PDF 텍스트 추출 (pdfminer 또는 pypdf)
from pdfminer.high_level import extract_text
text = extract_text('local_copy.pdf')
```

---

## 2. 지수 방법론 → 편입/편출 반영 방식

### 2-1. 분석 결과 데이터 구조

100개 지수에 대해 아래 필드를 추출했다.

| 필드명 | 설명 | 예시 |
|--------|------|------|
| `inclusion_rule` | 편출입 결정 산식 | "1차: 시총 누적 85% + 거래대금 85%; 2차: 기존 N×110% 이내..." |
| `weight_formula` | 비중 결정 산식 | "w_i = FMC_i / Σ FMC_j, CAP 30%" |
| `sector_standard` | 산업분류기준 | GICS / FICS / WICS / KRX업종분류 / 없음 |
| `cap_threshold` | 시총 누적 컷오프 | 85% |
| `liq_pct` | 유동성 컷오프 | 85% |
| `keep_buffer` | 기존구성 유지 버퍼 | 110% (기존 종목은 N×1.1 이내면 유지) |
| `new_buffer` | 신규편입 버퍼 | 90% (신규 종목은 N×0.9 이내여야 편입) |
| `target_n` | 목표 구성종목 수 | 200 |
| `ceiling_pct` | 단일종목 상한 CAP | 30% |

### 2-2. KRX 8단계 Rule Engine (KRX·파생 지수 공통 구조)

대부분의 KRX 계열 지수(KOSPI200, KOSDAQ150, KRX300 등)와 이를 기반으로 한 섹터지수는 아래 8단계를 따른다.

```
Step 1. Universe Filter
        └─ 상장기간, 관리종목·투자경고 제외, 기업지배구조 필터 등

Step 2. Sector Classification
        └─ GICS 기준으로 10개 산업군 분류 (또는 지수별 자체 기준)

Step 3. 1차 선정 (Primary Selection)
        ├─ cap_threshold: 산업군 내 시총 누적 상위 X% 이내
        └─ liq_pct:       산업군 내 거래대금 순위 상위 X% 이내 (교집합)

Step 4. 2차 선정 (Buffer Rule)
        ├─ 기존 구성종목: 산업군 배분 N × keep_buffer(110%) 이내면 유지
        └─ 신규 편입후보: 산업군 배분 N × new_buffer(90%) 이내여야 편입

Step 5. 3차 선정 (Tertiary)
        └─ tertiary_mode: 목표 N 미달 시 전산업 시총순 보충 / 또는 탈락

Step 6. Merger Whitelist
        └─ 합병·분할 중인 종목 특례 처리

Step 7. Matching (산업군 배분 조정)
        └─ 산업군별 편입 수 = floor(N × 산업군시총비중)

Step 8. Large-cap Special
        └─ large_cap_special_rank: 시총 상위 K위 이내 종목 자동편입 특례
```

#### KRX 8단계 파라미터 예시 (KOSPI 200)

```python
params = {
    "target_n": 200,
    "cap_threshold": 0.85,      # 시총 누적 85% 이내
    "liq_pct": 0.85,            # 거래대금 85% 이내
    "keep_buffer": 1.10,        # 기존 구성종목: N×110%
    "new_buffer": 0.90,         # 신규 후보: N×90%
    "tertiary_mode": "전산업_시총순_보충",
    "large_cap_special_rank": 50,   # 시총 상위 50위 자동편입
    "ceiling_pct": None,            # KOSPI200은 CAP 없음
    "sector_standard": "GICS",
}
```

### 2-3. 비중 방식 분류 (100개 지수)

| 비중 방식 | 지수 수 | 산식 |
|-----------|---------|------|
| 유동시총가중 + CAP | 40개 | `w_i = FMC_i / Σ FMC_j`, CAP 적용 후 초과분 재배분 |
| 유동시총가중 (CAP 없음) | 16개 | `w_i = FMC_i / Σ FMC_j` |
| 동일가중 | 7개 | `w_i = 1/N` |
| 고정비중 | 7개 | 종목별 고정 % (TOP3: 25% 고정 등) |
| 배당가중 | 3개 | `w_i = DIV_i / Σ DIV_j` |
| 기타(스코어 틸팅 등) | 27개 | 팩터 스코어 × FMC 혼합 등 |

### 2-4. 편입/편출 자동 추정 로직 (Python 구현 설계)

나중에 종목 데이터를 받으면 아래 구조로 구현한다.

```python
import pandas as pd
import numpy as np

def estimate_rebalancing(index_params: dict, stock_data: pd.DataFrame) -> dict:
    """
    index_params: all_v2_data_v4.json에서 해당 지수 레코드
    stock_data: 종목별 시가총액, 거래대금, 섹터, 현재 구성여부 등
    
    returns: {
        'inclusions': [...],   # 신규 편입 추정 종목
        'exclusions': [...],   # 편출 추정 종목
        'weights': {...},      # 종목별 추정 비중
    }
    """
    target_n = index_params['target_n']
    cap_thr   = index_params['cap_threshold']   # 예: 0.85
    liq_pct   = index_params['liq_pct']         # 예: 0.85
    keep_buf  = index_params['keep_buffer']      # 예: 1.10
    new_buf   = index_params['new_buffer']       # 예: 0.90
    ceiling   = index_params['ceiling_pct']      # 예: 0.30 또는 None
    sector_col = 'gics_sector'                   # stock_data 컬럼명

    results_by_sector = {}

    for sector, grp in stock_data.groupby(sector_col):
        # Step 1: 시총 누적 필터
        grp = grp.sort_values('market_cap', ascending=False)
        grp['cumcap_pct'] = grp['market_cap'].cumsum() / grp['market_cap'].sum()
        cap_pass = grp[grp['cumcap_pct'] <= cap_thr]

        # Step 2: 유동성 필터
        grp_liq = grp.sort_values('avg_turnover', ascending=False)
        liq_cutoff = int(len(grp_liq) * liq_pct)
        liq_pass = set(grp_liq.iloc[:liq_cutoff]['ticker'])

        # Step 3: 교집합 (1차 선정)
        primary = cap_pass[cap_pass['ticker'].isin(liq_pass)]

        # Step 4: 버퍼룰 (2차 선정)
        sector_alloc = _get_sector_alloc(sector, stock_data, target_n)
        current_members = set(grp[grp['is_current_member']]['ticker'])
        keep_limit = int(sector_alloc * keep_buf)
        new_limit  = int(sector_alloc * new_buf)

        kept = [t for t in primary['ticker'] if t in current_members][:keep_limit]
        new_candidates = [t for t in primary['ticker']
                          if t not in current_members][:new_limit]
        selected = list(dict.fromkeys(kept + new_candidates))  # 순서 유지
        results_by_sector[sector] = selected

    # Step 5: 집계 → 비중 계산
    all_selected = [t for tickers in results_by_sector.values() for t in tickers]
    weights = _calc_weights(all_selected, stock_data, ceiling)

    current_set = set(stock_data[stock_data['is_current_member']]['ticker'])
    selected_set = set(all_selected)

    return {
        'inclusions': sorted(selected_set - current_set),
        'exclusions': sorted(current_set - selected_set),
        'weights': weights,
        'total_selected': len(all_selected),
    }


def _get_sector_alloc(sector: str, stock_data: pd.DataFrame, target_n: int) -> int:
    """산업군별 시총비중 기반 배분 수 계산"""
    total_cap = stock_data['market_cap'].sum()
    sector_cap = stock_data[stock_data['gics_sector'] == sector]['market_cap'].sum()
    return max(1, int(target_n * sector_cap / total_cap))


def _calc_weights(tickers: list, stock_data: pd.DataFrame,
                  ceiling: float | None) -> dict:
    """유동시총가중 + CAP 재배분"""
    df = stock_data[stock_data['ticker'].isin(tickers)].copy()
    df['raw_w'] = df['float_market_cap'] / df['float_market_cap'].sum()

    if ceiling is None:
        return dict(zip(df['ticker'], df['raw_w']))

    # CAP 반복 재배분
    for _ in range(20):  # 최대 20회 반복
        capped = df['raw_w'].clip(upper=ceiling)
        excess = (df['raw_w'] - capped).sum()
        if excess < 1e-9:
            break
        under_cap = capped < ceiling
        capped[under_cap] += excess * (capped[under_cap] / capped[under_cap].sum())
        df['raw_w'] = capped

    return dict(zip(df['ticker'], df['raw_w']))
```

### 2-5. 지수 유형별 편입 로직 요약

#### A. KRX 8단계형 (KOSPI200·KOSDAQ150·KRX300·섹터지수)

```
편입 조건: 1차(시총+유동성 교집합) ∩ 2차(버퍼룰) → 목표 N종목
편출 조건: 기존 구성종목이 1차 선정 풀에서 탈락 AND 버퍼 초과
비중: 유동시총가중, 일부 CAP(20~30%)
```

#### B. 테마·FICS형 (FnGuide 계열)

```
편입 조건: FICS 특정 업종 코드 소속 + 시총/거래대금 기준 상위 N
편출 조건: FICS 업종 이탈 OR 시총·유동성 기준 미달
비중: 유동시총가중 + CAP (20% 일반, 특례 다름)
특이사항: 일부 지수는 키워드 스코어(딥서치 NLP) 활용
```

#### C. 고정구성형 (FnGuide 조선TOP3 등)

```
편입 조건: FICS 업종 내 매출액/시총 순위 상위 N (완전 고정)
편출 조건: 없음 (정기 리뷰 시 순위 재산정)
비중: TOP3 각 25% 고정, 나머지 유동시총가중
```

#### D. 파생형 (KOSPI200 섹터지수 등)

```
편입 조건: 모지수(KOSPI200·KOSDAQ150) 구성종목 중 해당 GICS 섹터
편출 조건: 모지수 편출 또는 섹터 재분류
비중: 유동시총가중, 단일종목 CAP 20%
수시 변경: 모지수 변경 시 자동 연동
```

#### E. 팩터·배당형

```
편입 조건: 배당수익률 기준 상위 N 또는 팩터 스코어 상위 N
편출 조건: 순위 하락 + 버퍼(±20%) 초과
비중: 배당가중 또는 팩터스코어 × 유동시총가중 혼합
```

#### F. MSCI형

```
편입 조건: FIF(외국인포함비율) ≥ 15% + ATVR 유동성 + 시가총액
           Large Cap 70%±5% / Standard 85%±5% / IMI 99% 커버리지
편출 조건: 커버리지 하한 이탈 + 120%/80% 버퍼 초과
비중: w_i = FIF_i × FMC_i / Σ(FIF_j × FMC_j)
```

---

## 3. 산출물 파일 구조

### Excel 파일: `index_methodology_analysis_v2_2026.xlsx`

| 시트명 | 내용 | 행 수 |
|--------|------|--------|
| `June2026_리밸런싱_요약` | 6월 리밸런싱 대상 전체 | 100 |
| `산업분류기준_분석` | 산업분류 기준별 정렬 | 100 |
| `June2026_Y_리밸런싱` | 리밸런싱 확정(Y) 지수만 | 57 |
| `편출입_비중_산식` | 편출입 결정 산식 + 비중 산식 | 100 |
| `KRX_Rule_Engine_파라미터` | KRX 8단계 수치 파라미터 | 13 |
| `전체_지수_기본정보` | 마스터 데이터 전체 | 100 |
| `제공사별_통계` | 제공사 집계 | ~12 |

### JSON 파일: `/tmp/all_v2_data_v4.json`

```json
{
  "file": "1_KRX_KOSPI200.pdf",
  "index_name_ko": "코스피 200",
  "index_name_en": "KOSPI 200",
  "provider": "한국거래소(KRX)",
  "provider_type": "거래소",
  "index_type": "시장대표",
  "target_n": 200,
  "market_scope": "KOSPI",
  "weight_method": "유동시총가중",
  "frequency": "반기",
  "review_date_rule": "5월·11월 두 번째 금요일",
  "effective_date_rule": "6월·12월 두 번째 월요일",
  "cap_threshold": 0.85,
  "liq_pct": 0.85,
  "keep_buffer": 1.10,
  "new_buffer": 0.90,
  "tertiary_mode": "전산업_시총순_보충",
  "ceiling_pct": null,
  "large_cap_special_rank": 50,
  "sector_standard": "GICS",
  "sector_count": 10,
  "june_2026_rebalancing": "Y - 6월 두 번째 월요일 발효",
  "inclusion_rule": "1차: 산업군(GICS 10개)별 일평균시총 누적 85% 이내 + 거래대금 순위 산업군 내 85% 이내 교집합; 2차: 기존구성 산업군 N×110% 이내 유지·신규 N×90% 이내; 3차: 200종목 미달 시 전산업 시총순 보충; 대형주특례: 시총 상위 50위 이내 자동편입",
  "weight_formula": "w_i = FMC_i / Σ FMC_j (유동시총가중, CAP 없음)"
}
```

---

## 4. 다음 단계: 자동 편출입 추정 엔진 구성

### 필요 입력 데이터

```
stock_data.csv 컬럼 (최소 요구사항):
  ticker          - 종목코드
  market_cap      - 시가총액 (원)
  float_market_cap - 유동시총 (원)
  avg_turnover    - 일평균거래대금 (원, 60영업일 기준)
  gics_sector     - GICS 섹터명 (또는 FICS, WICS)
  is_current_member - 현재 구성종목 여부 (bool)
  listing_months  - 상장 경과 개월 수
  is_admin        - 관리종목 여부 (bool)
```

### 실행 흐름

```
1. all_v2_data_v4.json + stock_data.csv 로드
2. 대상 지수 선택 (june_2026_rebalancing == 'Y')
3. 각 지수별 estimate_rebalancing() 실행
4. 결과: 신규 편입 / 편출 종목 목록 + 추정 비중
5. Excel 출력 또는 비교 리포트 생성
```

### MotherDuck 활용 (로컬 환경 전용)

```python
import duckdb

# 로컬 또는 MotherDuck 연결
conn = duckdb.connect('md:?motherduck_token=<TOKEN>')

# 종목 데이터 로드
conn.execute("""
    CREATE TABLE IF NOT EXISTS stock_master AS
    SELECT * FROM read_csv_auto('stock_data.csv')
""")

# 지수 방법론 JSON 로드
conn.execute("""
    CREATE TABLE IF NOT EXISTS index_params AS
    SELECT * FROM read_json_auto('all_v2_data_v4.json')
""")

# 편출입 추정 결과 저장
conn.execute("""
    CREATE TABLE IF NOT EXISTS rebalancing_estimate (
        index_name VARCHAR,
        ticker     VARCHAR,
        action     VARCHAR,  -- 'INCLUDE' / 'EXCLUDE' / 'KEEP'
        est_weight DOUBLE,
        run_date   DATE
    )
""")
```

---

## 5. 주의사항 및 한계

| 항목 | 내용 |
|------|------|
| 데이터 기준일 | PDF 발행 시점 기준 — 방법론 개정 시 재분석 필요 |
| 편입 추정 정확도 | 방법론 PDF 해석 기반 — 실제 산출기관 계산과 차이 가능 |
| 수시 변경 | 파생형 지수는 모지수 변경 시 즉시 반영 필요 |
| MSCI | FIF(외국인포함비율) 데이터 별도 필요 |
| 섹터 분류 | GICS/FICS/WICS 코드 데이터 별도 필요 |
| 버퍼 계산 | 산업군별 배분 N이 가변적 — 시총비중 기반 동적 계산 필요 |
