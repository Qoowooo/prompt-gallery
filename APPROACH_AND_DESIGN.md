# 지수 방법론 분석 프로젝트 — 접근 방식 및 설계 노트

> 작성일: 2026-06-09  
> 대상: 100개 지수 방법론 PDF 분석 전체 과정  
> 브랜치: `claude/drive-folder-access-Qtlsp`

---

## 1. 프로젝트 목표

| 목표 | 세부 내용 |
|------|----------|
| **단기** | 100개 지수의 2026년 6월 리밸런싱 일정 자동 추출 |
| **중기** | 지수 편입/편출 결정 산식 + 비중 산식 체계화 |
| **장기** | 종목 데이터 입력 시 편출입 종목 및 비중 자동 추정 |

---

## 2. 전체 작업 흐름

```
[1] Google Drive PDF 100개
         │
         ▼
[2] MCP 도구로 PDF 텍스트 추출 (병렬 에이전트 5개 × 20파일)
         │
         ▼
[3] LLM 파싱 → 구조화 JSON (v1 → v2 → v3 → v4 스키마 진화)
         │
         ├─── [3a] KRX 8단계 파라미터 추출
         ├─── [3b] 산업분류기준(sector_standard) 추출
         └─── [3c] 편출입 산식 + 비중 산식 추출
         │
         ▼
[4] 100개 레코드 병합 → all_v2_data_v4.json
         │
         ▼
[5] Excel 생성 (7개 시트) + MD 가이드 작성
         │
         ▼
[6] GitHub 푸시 → 오프라인 환경에서 활용
```

---

## 3. 핵심 설계 결정

### 3-1. 병렬 에이전트 처리

**문제**: 100개 PDF를 직렬로 처리하면 컨텍스트 한도 초과 + 응답 지연.

**결정**: 5개 백그라운드 에이전트를 동시에 실행, 각 20파일 담당.

```
Agent 1: 파일 1~20   → /tmp/batch1_results.json
Agent 2: 파일 21~40  → /tmp/batch2_results.json
Agent 3: 파일 41~60  → /tmp/batch3_results.json
Agent 4: 파일 61~80  → /tmp/batch4_results.json
Agent 5: 파일 81~100 → /tmp/batch5_results.json
```

**효과**: 직렬 대비 ~5배 속도 향상, 메인 컨텍스트 보호.

**주의점**: 에이전트마다 파일 ID를 명시적으로 주입해야 한다.  
에이전트는 이전 컨텍스트를 공유하지 않으므로, 프롬프트에 파일 ID 목록을 직접 포함시킬 것.

---

### 3-2. 파일 ID 관리 (가장 중요한 교훈)

**문제**: 초기 배치에서 파일명만 전달했더니, 에이전트가 Drive에서 파일을 못 찾아 전부 "없음" 반환.

**원인**: Google Drive MCP는 파일명이 아닌 `file_id`로만 파일 접근 가능.

**해결**: 파일 목록을 `(순번, 파일명, file_id)` 3열 TSV로 관리.

```
# /tmp/all_files_numbered.tsv 형식
순번  파일명                          file_id
1     1_KRX_KOSPI200.pdf              1abc...xyz
2     2_KRX_KOSDAQ150.pdf             2def...uvw
...
```

**추가 주의**: 동일한 물리 파일이 Google Drive에서 다른 이름으로 노출되는 경우가 있음.  
파일명 기준 매핑이 아니라 `file_id` 기준 매핑을 항상 우선시할 것.

```python
# 잘못된 방법
formula_lookup[rec['file_name']]  # 파일명 불일치 위험

# 올바른 방법
formula_lookup[rec['file_id']]    # file_id는 변하지 않음
```

---

### 3-3. 스키마 진화 (v1 → v4)

분석을 반복하면서 스키마를 점진적으로 확장했다. 한 번에 모든 필드를 뽑으려 하지 말고, **라운드별로 하나의 관심사만 추출**하는 것이 LLM 파싱 품질에 유리하다.

| 버전 | 추가 필드 | 목적 |
|------|----------|------|
| v1 | 기본 정보 (이름, 제공사, 비중, 리밸런싱) | 6월 리밸런싱 일정 파악 |
| v2 | KRX 8단계 파라미터 (cap_threshold, liq_pct, keep_buffer, new_buffer 등) | Rule Engine 수치 추출 |
| v3 | sector_standard (GICS/FICS/WICS 등) | 산업분류 기준 파악 |
| v4 | inclusion_rule, weight_formula | 자동 추정 엔진용 산식 |

---

### 3-4. LLM 파싱 품질 제어

PDF → 구조화 JSON 변환 시 LLM에 다음을 명시했다:

1. **"모른다"를 허용**: 정보가 없으면 `null` 또는 `"없음"` 반환, 추정 금지
2. **단위 통일**: 버퍼는 소수(`1.10`), 시총 컷오프는 소수(`0.85`)
3. **명시적 필드만 추출**: 문서에 없는 값은 채우지 않음
4. **파생지수 구분**: 모지수 편출입 연동인 경우 `"파생"` 플래그

```python
# 파싱 프롬프트 핵심 원칙 (에이전트 지시문)
"""
- 문서에 명시된 값만 추출하라. 추론하지 마라.
- 버퍼는 1.0 기준 소수(기존 구성종목 110% → 1.10)
- 정보 없으면 null, 해당 없으면 "없음" 또는 "해당없음"
- 모지수 연동 지수는 inclusion_rule에 "파생형" 명시
"""
```

---

### 3-5. 데이터 검증 패턴

배치 결과가 돌아올 때마다 아래 세 가지를 반드시 확인했다:

```python
# 1. 커버리지 확인 (누락 파일 없는지)
matched = set(r['file'] for r in results)
expected = set(row['파일명'] for row in file_list)
missing = expected - matched
print(f"누락: {missing}")

# 2. 핵심 필드 비율 확인 (파싱 실패율)
valid = [r for r in results if r.get('inclusion_rule') not in (None, '', '없음')]
print(f"유효 파싱률: {len(valid)}/{len(results)}")

# 3. 이상치 확인 (범위 벗어난 수치)
for r in results:
    cap = r.get('cap_threshold')
    if cap and not (0.5 <= cap <= 1.0):
        print(f"이상치: {r['file']} cap_threshold={cap}")
```

---

## 4. 분석 결과 요약

### 4-1. 데이터셋 개요

| 항목 | 수치 |
|------|------|
| 총 분석 지수 | 100개 |
| 제공사 수 | 6개 (KRX, FnGuide계열, iSelect, KEDI, MSCI, 기타) |
| 6월 리밸런싱 Y | 57개 |
| 6월 리밸런싱 N | 20개 |
| 조건부/불명확 | 23개 |

### 4-2. 제공사별 분포

| 제공사 | 지수 수 | 특징 |
|--------|---------|------|
| FnGuide 계열 | 38개 | FICS 기준, 테마형 다수 |
| KRX | 30개 | GICS 기준, 8단계 Rule Engine |
| iSelect (NH투자증권) | 14개 | FICS 기반 테마형 |
| MSCI | 8개 | FIF 기반, 글로벌 커버리지 |
| KEDI (한국경제신문) | 5개 | KICS 자체 분류 |
| 기타 (S&P, Bloomberg 등) | 5개 | 글로벌 기준 |

### 4-3. 리밸런싱 주기

| 주기 | 지수 수 |
|------|---------|
| 연 2회 (반기) | 48개 |
| 분기 4회 | 18개 |
| 연 1회 | 14개 |
| 없음 (파생형) | 13개 |
| 기타/불명 | 7개 |

### 4-4. 산업분류 기준

| 기준 | 지수 수 | 주요 사용처 |
|------|---------|------------|
| 없음 (단일 테마) | 56개 | 테마지수, 파생지수 |
| GICS | 19개 | KRX 계열, MSCI |
| FICS | 17개 | FnGuide, iSelect |
| 기타 (KICS, WICS 등) | 8개 | KEDI, NICE, Akros |

### 4-5. 비중 방식

| 비중 방식 | 지수 수 |
|-----------|---------|
| 유동시총가중 + CAP | 40개 |
| 유동시총가중 (CAP 없음) | 16개 |
| 고정비중 | 10개 |
| 시총가중 (유동비율 미적용) | 5개 |
| 동일가중 | 1개 |
| 기타 (팩터 스코어 혼합 등) | 28개 |

---

## 5. 실패 사례 및 교훈

### 실패 1: 배치 파일 ID 오염

**상황**: sector_standard 추출 배치 3~5에서 이전 v2 배치의 파일명/ID를 그대로 복사해 사용.  
**증상**: 19/20 파일에서 "Entity not found" 반환.  
**원인**: 에이전트 프롬프트에 하드코딩된 파일 ID가 실제 Drive와 다른 구버전 ID였음.  
**해결**: `/tmp/all_files_numbered.tsv`를 단일 진실 원본(Source of Truth)으로 확정. 모든 배치는 이 파일의 file_id만 사용.

```
핵심 교훈: 파일 ID는 절대로 프롬프트에 하드코딩하지 않는다.
           항상 마스터 TSV를 읽어서 주입한다.
```

### 실패 2: 파일명 불일치

**상황**: Google Drive 노출 파일명과 v2 데이터 파일명이 3개 파일에서 불일치.

```
v2 데이터 파일명                        Drive 실제 파일명
30_KRX_SECU.pdf                   → 30_KRX_Securities_Sector.pdf  (같은 file_id)
46_ISELECT_KDEF_TOP10.pdf         → 46_iSelect_KDEF_TOP10.pdf
MSCI_0_INFORMATION_DOCUMENT.pdf   → MSCI_0_INFORMATION_DOCUMENT_Description_...pdf
```

**해결**: `manual_mappings` 딕셔너리로 예외 처리. 장기적으로는 file_id 기준 매핑 필수.

### 실패 3: MotherDuck 연결 불가

**상황**: 원격 실행 환경(Claude Code on the web)에서 외부 인터넷 전체 차단.  
**증상**: `extensions.duckdb.org`, `api.motherduck.com`, `google.com` 모두 HTTP 403.  
**결론**: 네트워크 허용 정책을 설정한 환경에서만 MotherDuck 연결 가능.  
**대안**: 로컬 환경(Claude Code CLI/데스크탑)에서 MotherDuck 연동.

---

## 6. 자동 편출입 추정 엔진 — 다음 단계 설계

### 6-1. 입력 → 처리 → 출력

```
입력 1: all_v2_data_v4.json (100개 지수 방법론 파라미터)
입력 2: stock_data.csv      (종목별 시가총액, 거래대금, 섹터 등)
입력 3: current_members.csv (현재 각 지수 구성종목 목록)
         │
         ▼
    estimate_rebalancing()
    ├─ Step 1: Universe Filter (상장기간, 관리종목 제외)
    ├─ Step 2: Sector 분류 (GICS/FICS/WICS)
    ├─ Step 3: 1차 선정 (cap_threshold ∩ liq_pct)
    ├─ Step 4: 2차 선정 (keep_buffer, new_buffer)
    ├─ Step 5: 3차 선정 (tertiary_mode)
    ├─ Step 6: Large-cap 특례
    └─ Step 7: 비중 계산 (_calc_weights)
         │
         ▼
출력: rebalancing_report.xlsx
    ├─ 신규 편입 추정 종목
    ├─ 편출 추정 종목
    └─ 종목별 추정 비중
```

### 6-2. 필요한 stock_data.csv 스펙

```csv
ticker, name_ko, market_cap, float_market_cap, avg_turnover_60d,
gics_sector, fics_sector, wics_sector,
listing_months, is_admin_stock, is_warning_stock,
dividend_yield, foreign_ownership_ratio
```

### 6-3. 지수별 처리 분기

```python
def route_index(params: dict) -> str:
    """지수 유형 판별 → 처리 함수 라우팅"""
    if params.get('cap_threshold'):
        return 'krx_8step'        # KRX 8단계 Rule Engine
    elif '파생' in params.get('index_type', ''):
        return 'derivative'       # 모지수 연동, 자체 선정 없음
    elif params.get('inclusion_rule', '').startswith('고정'):
        return 'fixed'            # 고정 구성, 순위만 재산정
    elif 'FICS' in params.get('sector_standard', ''):
        return 'fics_theme'       # FICS 업종 기반 테마
    elif 'FIF' in params.get('inclusion_rule', ''):
        return 'msci'             # MSCI FIF 기반
    else:
        return 'generic'          # 기본 시총 상위 N
```

### 6-4. 검증 방법

- **과거 리밸런싱 데이터와 비교**: 과거 발표 편출입 내역 vs. 추정 결과 일치율
- **비중 오차 측정**: `|추정비중 - 실제비중|` 평균 절대 오차(MAE)
- **버퍼 경계 종목 검토**: 버퍼 ±2%p 이내 종목은 불확실 구간으로 별도 표시

---

## 7. 파일 목록

| 파일 | 위치 | 설명 |
|------|------|------|
| `index_methodology_analysis_v2_2026.xlsx` | `/home/user/prompt-gallery/` | 최종 산출물 (7 시트) |
| `INDEX_METHODOLOGY_GUIDE.md` | `/home/user/prompt-gallery/` | 기술 참조 가이드 |
| `APPROACH_AND_DESIGN.md` | `/home/user/prompt-gallery/` | 이 문서 |
| `all_v2_data_v4.json` | `/tmp/` | 100개 지수 마스터 데이터 (세션 한정) |
| `all_files_numbered.tsv` | `/tmp/` | 파일명 ↔ Drive file_id 매핑 (세션 한정) |

> `/tmp/` 파일은 세션 종료 시 소멸. 오프라인 환경에서 재사용하려면  
> `all_v2_data_v4.json`을 repo에 커밋하거나 MotherDuck에 저장할 것.

---

## 8. 재현 체크리스트

오프라인/로컬 환경에서 이 작업을 다시 하거나 확장할 때:

- [ ] Google Drive API 인증 설정 (service account JSON 또는 OAuth)
- [ ] `all_files_numbered.tsv` 확보 (파일명 + file_id 3열)
- [ ] DuckDB + MotherDuck 토큰 설정
- [ ] 배치 처리 시 에이전트당 파일 수 ≤ 20개 유지
- [ ] 각 배치 결과 저장 후 커버리지 검증 (누락 파일 없는지)
- [ ] 파일명 불일치 예외처리 (`manual_mappings`) 확인
- [ ] `all_v2_data_v4.json` → `stock_data.csv` 연동 후 추정 엔진 실행
