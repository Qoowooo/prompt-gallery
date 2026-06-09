"""
ETF June 2026 Rebalancing Analysis
====================================
June 2026 정기 리밸런싱 대상 ETF의 편입/편출 종목 및 비중 변화 산출

Usage:
    export MOTHERDUCK_TOKEN=<token>
    python scripts/etf_rebalancing_analysis.py

Output:
    output/etf_rebalancing_2026_06.xlsx
"""

import os
import sys
import re
import logging
from pathlib import Path
from datetime import datetime

import duckdb
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# ─── Logging ────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

# ─── Paths ──────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

METHODOLOGY_XLSX = DATA_DIR / "index_methodology.xlsx"
ETF_MATCHING_XLSX = DATA_DIR / "etf_index_matching.xlsx"
OUTPUT_XLSX = OUTPUT_DIR / "etf_rebalancing_2026_06.xlsx"

# ─── Style Constants ─────────────────────────────────────────
CLR_HEADER = "1F4E79"       # 진한 파란색
CLR_BUY    = "C6EFCE"       # 연초록 (매수)
CLR_SELL   = "FFC7CE"       # 연빨강 (매도)
CLR_HOLD   = "FFFFFF"       # 흰색 (유지)
CLR_NEW    = "FFEB9C"       # 연노란색 (신규 편입)
CLR_SUBHDR = "BDD7EE"       # 연파란색 (소헤더)
CLR_SUMHDR = "2F75B6"       # 중간 파란 (요약 헤더)


# ─────────────────────────────────────────────────────────────
#  INDEX NAME NORMALISATION
# ─────────────────────────────────────────────────────────────

# 방법론 Excel(영문 표기) → DB etf_index(한글 표기) 수동 매핑
_KO_MAP = {
    "KOSPI 200":              "코스피 200",
    "KOSPI 100":              "코스피 100",
    "KOSPI 50":               "코스피 50",
    "KOSPI 200 정보기술":      "코스피 200 정보기술",
    "KOSPI 200 중공업 섹터지수": "코스피 200 중공업",
    "코리아 밸류업 지수":      "코리아 밸류업 지수",
    "KRX 300 지수":           "KRX 300",
    "KRX 부동산리츠인프라 지수": "KRX 부동산리츠인프라 지수",
}

# DB etf_index값이 방법론 지수명(한)에 포함하는 약식 매핑
_DB_NORM = {
    "FnGuide 리츠부동산인프라":       "FnGuide 리츠부동산인프라 지수",
    "FnGuide TOP 5 Plus Total Return 지수": "FnGuide TOP 5 Plus 지수",
    "FnGuide AI반도체 TOP3+ 지수":    "FnGuide AI 반도체 TOP3+ 지수",
    "FnGuide 고배당주 플러스 자사주 지수 (PR)": "FnGuide 고배당주 지수",
    "FnGuide 성장 지수(PR)":           "FnGuide 성장 지수",
    "FnGuide 기업가치 지수(시장가격)":  "FnGuide 기업가치 지수",
    "DeepSearch 원자력 TOP 10지수":    "DeepSearch 원자력 Top10 지수",
    "WISE 삼성그룹 밸류 인덱스":       "WISE삼성그룹밸류인덱스",
    "코리아 밸류업 TR 지수":           "코리아 밸류업 지수",
    "이Select 원자력SMR 지수(Price Return)": "iSelect K원자력SMR 지수",
    "iSelect 원자력SMR 지수(Price Return)": "iSelect K원자력SMR 지수",
    "iSelect K방산&우주":              "iSelect K방산&우주 지수",
    "iSelect 방산 TOP10 지수(Price Return)": "iSelect K방산TOP10 지수",
    "iSelect 조선TOP10 Index(PR)":    "iSelect 조선TOP10 지수",
    "iSelect 비메모리반도체 지수(시장가격지수)": "iSelect 비메모리반도체 지수",
    "iSelect 2차전지 지수 (시장가격지수)": "iSelect 2차전지 지수",
    "iSelect AI&로봇 지수(시장가격)":  "iSelect AI&로봇 지수",
    "iSelect AI 전력핵심설비 지수(Price Return)": "iSelect AI전력핵심설비 지수",
    "iSelect AI 반도체핵심장비 지수":  "iSelect AI반도체핵심장비 지수",
    "iSelect 바이오헬스케어 PR 지수":  "iSelect 바이오헬스케어 PR 지수",
    "iSelect 코리아 원자력 지수 (Price Return)": "iSelect 코리아 원자력 지수",
    "KEDI 메가테크지수(PR)":           "KEDI 메가테크 지수",
    "KEDI 코리아 휴머노이드로봇산업 지수 (PR)": "KEDI 코리아휴머노이드로봇산업 지수",
    "MKF 현대차그룹+ FW":             "MKF 현대차그룹 지수",
    "Dow Jones Korea Dividend 30 지수 (Price Return)": "다우존스 한국 배당 30 지수",
    "NICE K반도체 TOP2 MAX+ 지수":    "NICE K 반도체 TOP2 MAX+ 지수",
    "MSCI Korea Index":               "MSCI 글로벌 투자가능시장 지수 방법론",
    "FnGuide 은행고배당플러스TOP10 지수(시장가격 지수)": "FnGuide 은행 고배당 플러스 TOP 10 지수",
    "FnGuide 자동차TOP3플러스지수(PR)": "FnGuide 자동차 TOP3 플러스 지수",
}

# GICS 섹터 → KRX 섹터지수명 매핑
_GICS_TO_SECTOR_IDX = {
    "Information Technology": "코스피 200 정보기술",
    "Industrials":             "코스피 200 중공업",
    "Financials":              "코스피 200 금융",
    "Health Care":             "코스피 200 헬스케어",
    "Energy":                  "코스피 200 에너지/화학",
    "Materials":               "코스피 200 철강/소재",
    "Consumer Discretionary":  "코스피 200 경기소비재",
    "Consumer Staples":        "코스피 200 생활소비재",
    "Communication Services":  "코스피 200 커뮤니케이션서비스",
    "Real Estate":             "코스피 200 헬스케어",  # 없으면 건너뜀
    "Utilities":               "코스피 200 에너지/화학",
}


def normalize_idx_name(name: str) -> str:
    """방법론 지수명 → DB etf_index 정규화"""
    if name in _KO_MAP:
        return _KO_MAP[name]
    return name


def reverse_normalize(db_name: str) -> str:
    """DB etf_index → 방법론 지수명 정규화"""
    if db_name in _DB_NORM:
        return _DB_NORM[db_name]
    return db_name


def fuzzy_match(name: str, candidates: pd.Index, threshold: int = 3) -> str | None:
    """간단한 fuzzy matching (공통 토큰 기반)"""
    name_tokens = set(re.split(r'\s+|_|-', name.lower()))
    best, best_score = None, 0
    for c in candidates:
        c_tokens = set(re.split(r'\s+|_|-', c.lower()))
        score = len(name_tokens & c_tokens)
        if score > best_score:
            best_score = score
            best = c
    return best if best_score >= threshold else None


# ─────────────────────────────────────────────────────────────
#  DATABASE CONNECTION & DISCOVERY
# ─────────────────────────────────────────────────────────────

def connect_motherduck(token: str) -> duckdb.DuckDBPyConnection:
    log.info("MotherDuck 연결 중...")
    conn = duckdb.connect(f"md:?motherduck_token={token}")
    log.info("연결 성공")
    return conn


def discover_db(conn: duckdb.DuckDBPyConnection) -> str:
    """사용 가능한 DB/테이블 탐색 후 사용할 DB명 반환"""
    dbs = conn.execute("SHOW DATABASES").df()
    log.info(f"DB 목록: {dbs['database_name'].tolist()}")

    # 'dim_etf' 테이블이 있는 DB 찾기
    for db in dbs["database_name"]:
        if db in ("memory", "system", "temp"):
            continue
        try:
            tables = conn.execute(f"SHOW TABLES IN {db}").df()
            tbl_names = tables.iloc[:, 0].tolist()
            log.info(f"  {db}: {tbl_names}")
            if "dim_etf" in tbl_names:
                log.info(f"  → 사용 DB: {db}")
                return db
        except Exception as e:
            log.debug(f"  {db} 접근 불가: {e}")

    raise RuntimeError("dim_etf 테이블을 찾을 수 없습니다. DB 이름을 확인하세요.")


def get_table_columns(conn: duckdb.DuckDBPyConnection, db: str, table: str) -> list:
    """테이블 컬럼 목록 조회"""
    try:
        df = conn.execute(f"DESCRIBE {db}.{table}").df()
        return df.iloc[:, 0].tolist()
    except Exception:
        return []


# ─────────────────────────────────────────────────────────────
#  DATA LOADING FROM MOTHERDUCK
# ─────────────────────────────────────────────────────────────

def load_etf_index_map(conn: duckdb.DuckDBPyConnection, db: str) -> pd.DataFrame:
    """dim_etf에서 ETF-지수 매핑 로드"""
    return conn.execute(f"""
        SELECT etf_ticker_ks, etf_index
        FROM {db}.dim_etf
        WHERE type = 'ETF'
          AND etf_index IS NOT NULL
    """).df()


def load_current_holdings(
    conn: duckdb.DuckDBPyConnection, db: str, etf_tickers: list
) -> pd.DataFrame:
    """fact_etf_holdings_daily 최신 일자 편입 비중"""
    if not etf_tickers:
        return pd.DataFrame()
    tickers_sql = ", ".join(f"'{t}'" for t in etf_tickers)
    return conn.execute(f"""
        WITH latest AS (
            SELECT etf_ticker_ks, MAX(date) AS max_date
            FROM {db}.fact_etf_holdings_daily
            WHERE etf_ticker_ks IN ({tickers_sql})
            GROUP BY etf_ticker_ks
        )
        SELECT h.etf_ticker_ks, h.stock_ticker_ks, h.value_per, h.date
        FROM {db}.fact_etf_holdings_daily h
        INNER JOIN latest l
               ON h.etf_ticker_ks = l.etf_ticker_ks
              AND h.date          = l.max_date
    """).df()


def load_etf_aum(
    conn: duckdb.DuckDBPyConnection, db: str, etf_tickers: list
) -> pd.DataFrame:
    """fact_etf_nav 최신 순자산총액"""
    if not etf_tickers:
        return pd.DataFrame()
    tickers_sql = ", ".join(f"'{t}'" for t in etf_tickers)
    return conn.execute(f"""
        WITH latest AS (
            SELECT etf_ticker_ks, MAX(date) AS max_date
            FROM {db}.fact_etf_nav
            WHERE etf_ticker_ks IN ({tickers_sql})
            GROUP BY etf_ticker_ks
        )
        SELECT n.etf_ticker_ks, n.etf_asset
        FROM {db}.fact_etf_nav n
        INNER JOIN latest l
               ON n.etf_ticker_ks = l.etf_ticker_ks
              AND n.date          = l.max_date
    """).df()


def load_stock_data(conn: duckdb.DuckDBPyConnection, db: str) -> pd.DataFrame:
    """fact_stock_data + dim_gics + dim_fics 최신 데이터"""
    # 컬럼 존재 여부 확인
    stock_cols  = get_table_columns(conn, db, "fact_stock_data")
    gics_cols   = get_table_columns(conn, db, "dim_gics")
    fics_cols   = get_table_columns(conn, db, "dim_fics")

    log.info(f"fact_stock_data 컬럼: {stock_cols}")
    log.info(f"dim_gics 컬럼: {gics_cols}")
    log.info(f"dim_fics 컬럼: {fics_cols}")

    # 공통 join 키 결정
    gics_key  = "stock_ticker_ks" if "stock_ticker_ks" in gics_cols else gics_cols[0]
    fics_key  = "stock_ticker_ks" if "stock_ticker_ks" in fics_cols else fics_cols[0]

    # 평균 거래대금 컬럼 있으면 사용
    turnover_col = next(
        (c for c in stock_cols if "turnover" in c.lower() or "거래" in c.lower()), None
    )
    extra_select = f", s.{turnover_col} AS avg_turnover" if turnover_col else ""

    # 종목명 컬럼
    name_col = next(
        (c for c in stock_cols if "name" in c.lower() or "종목명" in c.lower()), None
    )
    name_select = f", s.{name_col} AS stock_name" if name_col else ""

    return conn.execute(f"""
        WITH latest AS (
            SELECT MAX(date) AS max_date FROM {db}.fact_stock_data
        )
        SELECT
            s.stock_ticker_ks,
            s.stock_market,
            s.stock_industry,
            s.member_kospi200,
            s.member_kosdaq150,
            s.stock_float_rate,
            s.stock_cap,
            s.float_stock_cap
            {extra_select}
            {name_select},
            g.stock_gics,
            f.fics_l,
            f.fics_m,
            f.fics_s
        FROM {db}.fact_stock_data s
        CROSS JOIN latest l
        LEFT JOIN {db}.dim_gics g ON g.{gics_key} = s.stock_ticker_ks
        LEFT JOIN {db}.dim_fics f ON f.{fics_key} = s.stock_ticker_ks
        WHERE s.date = l.max_date
    """).df()


def load_stock_names(conn: duckdb.DuckDBPyConnection, db: str) -> dict:
    """종목코드 → 종목명 딕셔너리"""
    # dim_stock 또는 fact_stock_data에서 종목명 조회
    for tbl in ["dim_stock", "fact_stock_data"]:
        try:
            cols = get_table_columns(conn, db, tbl)
            name_col = next(
                (c for c in cols if "name" in c.lower() or "종목명" in c.lower()), None
            )
            if name_col:
                df = conn.execute(f"""
                    SELECT DISTINCT stock_ticker_ks, {name_col} AS name
                    FROM {db}.{tbl}
                    WHERE {name_col} IS NOT NULL
                """).df()
                return dict(zip(df["stock_ticker_ks"], df["name"]))
        except Exception:
            pass
    return {}


# ─────────────────────────────────────────────────────────────
#  METHODOLOGY LOADING
# ─────────────────────────────────────────────────────────────

def load_june2026_methodology() -> list:
    """index_methodology.xlsx > June2026_Y_리밸런싱 탭 로드"""
    wb = openpyxl.load_workbook(METHODOLOGY_XLSX)
    ws = wb["June2026_Y_리밸런싱"]
    headers = [ws.cell(row=2, column=c).value for c in range(1, ws.max_column + 1)]
    data = []
    for r in range(3, ws.max_row + 1):
        row = {h: ws.cell(row=r, column=i).value for i, h in enumerate(headers, 1)}
        if row.get("순번"):
            data.append(row)
    log.info(f"방법론 로드 완료: {len(data)}개 지수")
    return data


def load_etf_matching_from_excel() -> pd.DataFrame:
    """etf_index_matching.xlsx에서 ETF-지수 매핑 로드"""
    wb = openpyxl.load_workbook(ETF_MATCHING_XLSX)
    ws = wb["index"]
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(2, ws.max_row + 1):
        row = {h: ws.cell(row=r, column=i).value for i, h in enumerate(headers, 1)}
        if row.get("etf_ticker_ks"):
            rows.append(row)
    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────
#  ETF ↔ INDEX MATCHING
# ─────────────────────────────────────────────────────────────

def build_index_etf_map(
    methodology: list,
    etf_db_map: pd.DataFrame,
    etf_excel_map: pd.DataFrame,
) -> dict:
    """
    각 지수명(한) → [etf_ticker_ks] 매핑 딕셔너리 생성
    우선순위: DB dim_etf > Excel etf_matching
    """
    # DB 매핑: etf_index → tickers
    db_idx2etf: dict = {}
    for _, row in etf_db_map.iterrows():
        idx = row["etf_index"]
        db_idx2etf.setdefault(idx, []).append(row["etf_ticker_ks"])

    # Excel 매핑 (DB가 없을 때 보완)
    excel_idx2etf: dict = {}
    for _, row in etf_excel_map.iterrows():
        idx = row["etf_index"]
        excel_idx2etf.setdefault(idx, []).append(row["etf_ticker_ks"])

    all_db_indices = set(db_idx2etf.keys())

    result = {}
    for row in methodology:
        idx_name = row["지수명(한)"]
        if not idx_name:
            continue

        # 1. 방법론명 → DB명 변환 후 직접 조회
        db_name = normalize_idx_name(idx_name)
        if db_name in db_idx2etf:
            result[idx_name] = db_idx2etf[db_name]
            continue

        # 2. 직접 조회 (변환 없이)
        if idx_name in db_idx2etf:
            result[idx_name] = db_idx2etf[idx_name]
            continue

        # 3. _DB_NORM 역방향 조회
        rev = reverse_normalize(idx_name)
        if rev in db_idx2etf:
            result[idx_name] = db_idx2etf[rev]
            continue

        # 4. Excel 매핑에서 fuzzy
        matched = fuzzy_match(idx_name, pd.Index(list(excel_idx2etf.keys())), threshold=3)
        if matched and matched in excel_idx2etf:
            result[idx_name] = excel_idx2etf[matched]
            log.debug(f"  fuzzy match: {idx_name!r} → {matched!r}")
            continue

        # 5. 매칭 실패
        log.warning(f"  ETF 매칭 실패: {idx_name!r}")
        result[idx_name] = []

    return result


# ─────────────────────────────────────────────────────────────
#  WEIGHT COMPUTATION UTILITIES
# ─────────────────────────────────────────────────────────────

def apply_cap(weights: pd.Series, cap: float, max_iter: int = 30) -> pd.Series:
    """유동시총가중 CAP 반복 재배분"""
    w = weights.copy()
    for _ in range(max_iter):
        over = (w > cap)
        excess = (w - cap).clip(lower=0).sum()
        if excess < 1e-12:
            break
        w = w.clip(upper=cap)
        under = w < cap
        if under.sum() == 0:
            break
        w.loc[under] += excess * w.loc[under] / w.loc[under].sum()
    return w


def calc_weights_float(
    stocks: pd.DataFrame,
    cap_pct: float | None = None,
    fmc_col: str = "float_stock_cap",
) -> pd.Series:
    """유동시총가중 비중 계산 (CAP 옵션)"""
    fmc = stocks[fmc_col].fillna(0).clip(lower=0)
    total = fmc.sum()
    if total == 0:
        return pd.Series(0.0, index=stocks.index)
    w = fmc / total
    if cap_pct:
        w = apply_cap(w, cap_pct)
    return w


# ─────────────────────────────────────────────────────────────
#  KRX 8-STEP RULE ENGINE
# ─────────────────────────────────────────────────────────────

def krx_sector_alloc(
    sector_cap: float, total_cap: float, target_n: int
) -> int:
    return max(1, int(target_n * sector_cap / total_cap))


def apply_krx_8step(
    stocks: pd.DataFrame,
    market: str,
    target_n: int,
    cap_threshold: float,
    keep_buffer: float,
    new_buffer: float,
    large_cap_rank: int | None,
    ceiling_pct: float | None,
    sector_col: str = "stock_gics",
    liq_col: str | None = None,       # 거래대금 컬럼 (없으면 시총만 사용)
    liq_pct: float = 0.85,
) -> pd.DataFrame:
    """
    KRX 8단계 Rule Engine 적용
    Returns: 선정된 종목 DataFrame (ticker, weight 컬럼 포함)
    """
    universe = stocks[stocks["stock_market"] == market].copy()
    universe = universe.dropna(subset=["stock_cap", "float_stock_cap"])
    universe = universe[universe["stock_cap"] > 0]

    current_set = set(
        universe.loc[universe.get("member_kospi200" if market == "KOSPI"
                                  else "member_kosdaq150", pd.Series(False)), "stock_ticker_ks"]
    )

    # Large-cap special (Step 8)
    special_set: set = set()
    if large_cap_rank:
        special_set = set(
            universe.nlargest(large_cap_rank, "stock_cap")["stock_ticker_ks"]
        )

    total_cap = universe["stock_cap"].sum()
    selected_tickers: list = []

    # Per-sector selection
    sectors = universe[sector_col].dropna().unique()
    for sector in sectors:
        sec_df = universe[universe[sector_col] == sector].copy()
        sec_cap = sec_df["stock_cap"].sum()
        alloc   = krx_sector_alloc(sec_cap, total_cap, target_n)

        # 1차 선정: 시총 누적 cap_threshold
        sec_df = sec_df.sort_values("stock_cap", ascending=False)
        sec_df["_cum_pct"] = sec_df["stock_cap"].cumsum() / sec_cap
        primary_pool = sec_df[sec_df["_cum_pct"] <= cap_threshold]["stock_ticker_ks"].tolist()

        # 유동성 필터 (거래대금 데이터 있을 때)
        if liq_col and liq_col in sec_df.columns:
            liq_cutoff = max(1, int(len(sec_df) * liq_pct))
            liq_pass = set(
                sec_df.nlargest(liq_cutoff, liq_col)["stock_ticker_ks"]
            )
            primary_pool = [t for t in primary_pool if t in liq_pass]

        if not primary_pool:
            primary_pool = sec_df.head(1)["stock_ticker_ks"].tolist()

        # 2차 선정: 버퍼룰
        keep_lim = int(alloc * keep_buffer)
        new_lim  = int(alloc * new_buffer)

        kept      = [t for t in primary_pool if t in current_set][:keep_lim]
        new_cands = [t for t in primary_pool if t not in current_set][:new_lim]
        sector_sel = list(dict.fromkeys(kept + new_cands))[:alloc]
        selected_tickers.extend(sector_sel)

    # 3차 선정: 전체 시총순 보충
    selected_set = set(selected_tickers)
    if len(selected_set) < target_n:
        gap = target_n - len(selected_set)
        fallback = (
            universe[~universe["stock_ticker_ks"].isin(selected_set)]
            .sort_values("stock_cap", ascending=False)
            .head(gap)["stock_ticker_ks"]
            .tolist()
        )
        selected_set.update(fallback)

    # Large-cap special 합산
    selected_set |= special_set

    result = universe[universe["stock_ticker_ks"].isin(selected_set)].copy()
    result["weight_new"] = calc_weights_float(result, ceiling_pct).values
    return result


def apply_sector_derivative(
    new_parent: pd.DataFrame,
    gics_sector: str,
    ceiling_pct: float = 0.20,
) -> pd.DataFrame:
    """KOSPI200 파생 섹터지수: 모지수에서 GICS 섹터 필터"""
    sec = new_parent[new_parent["stock_gics"] == gics_sector].copy()
    if sec.empty:
        return sec
    sec["weight_new"] = calc_weights_float(sec, ceiling_pct).values
    return sec


# ─────────────────────────────────────────────────────────────
#  PER-INDEX METHODOLOGY ROUTING
# ─────────────────────────────────────────────────────────────

def route_and_compute(
    idx_name: str,
    idx_params: dict,
    stocks: pd.DataFrame,
    krx_results: dict,          # pre-computed KRX parent results
) -> pd.DataFrame | None:
    """
    지수 유형 판별 후 신규 구성종목 + 비중 계산
    Returns: DataFrame(stock_ticker_ks, weight_new) 또는 None (산출 불가)
    """
    rule = idx_params.get("편출입 결정 산식") or ""
    sector_std = idx_params.get("산업분류기준") or ""
    provider = idx_params.get("제공사") or ""
    target_n = idx_params.get("목표종목수")

    # ── KRX 직접형 ──────────────────────────────────────────
    if idx_name == "KOSPI 200":
        return krx_results.get("KOSPI200")

    if idx_name in ("KOSPI 100 지수", "KOSPI 100"):
        p = krx_results.get("KOSPI200")
        if p is None:
            return None
        top = p.nlargest(100, "stock_cap").copy()
        top["weight_new"] = calc_weights_float(top, cap_pct=None).values
        return top

    if idx_name in ("KOSPI 50 지수", "KOSPI 50"):
        p = krx_results.get("KOSPI200")
        if p is None:
            return None
        top = p.nlargest(50, "stock_cap").copy()
        top["weight_new"] = calc_weights_float(top, cap_pct=None).values
        return top

    if idx_name == "KRX 300 지수":
        return krx_results.get("KRX300")

    # ── KRX 파생 섹터지수 ─────────────────────────────────
    if "KOSPI200 구성종목 중" in rule or "기초지수(KOSPI200)" in rule or "KOSPI200/KOSDAQ150/KRX300 구성종목" in rule:
        p = krx_results.get("KOSPI200")
        if p is None:
            return None
        # 섹터 결정
        for gics_sector, mapped_name in _GICS_TO_SECTOR_IDX.items():
            if mapped_name in idx_name or gics_sector.lower() in idx_name.lower():
                return apply_sector_derivative(p, gics_sector)
        # 직접 GICS 섹터명 매핑
        sector_map = {
            "정보기술": "Information Technology",
            "중공업":   "Industrials",
            "금융":     "Financials",
            "헬스케어": "Health Care",
            "에너지":   "Energy",
            "소재":     "Materials",
            "경기소비재": "Consumer Discretionary",
            "생활소비재": "Consumer Staples",
            "커뮤니케이션": "Communication Services",
            "부동산":   "Real Estate",
        }
        for ko, gics in sector_map.items():
            if ko in idx_name:
                return apply_sector_derivative(p, gics)
        return None

    # ── KOSPI 200 정보기술 (직접 명시) ──────────────────────
    if idx_name == "KOSPI 200 정보기술":
        p = krx_results.get("KOSPI200")
        return apply_sector_derivative(p, "Information Technology") if p is not None else None

    if "KOSPI 200 중공업" in idx_name:
        p = krx_results.get("KOSPI200")
        return apply_sector_derivative(p, "Industrials") if p is not None else None

    # ── KRX 부동산리츠인프라 ─────────────────────────────
    if "부동산리츠인프라" in idx_name and "KRX" in provider:
        # GICS Real Estate 종목 + 상장 리츠
        sec = stocks[
            (stocks["stock_gics"] == "Real Estate") |
            (stocks["fics_m"].str.contains("리츠", na=False))
        ].copy()
        sec = sec[sec["stock_cap"] > 0]
        if not sec.empty:
            sec["weight_new"] = calc_weights_float(sec, cap_pct=0.20).values
        return sec if not sec.empty else None

    # ── 코리아 밸류업 지수 ────────────────────────────────
    if "코리아 밸류업" in idx_name:
        # 1차: 시총 상위 400위 + (거래대금 상위 80% - 데이터 있을 때)
        universe = stocks[
            (stocks["stock_market"].isin(["KOSPI", "KOSDAQ"])) &
            (stocks["stock_cap"] > 0) &
            (stocks["stock_float_rate"].fillna(0) >= 10)
        ].copy()
        if universe.empty:
            return None
        top400 = universe.nlargest(400, "stock_cap")
        # 시총 상위 100종목 선택 (단순화)
        if target_n and isinstance(target_n, (int, float)):
            top = top400.nlargest(int(target_n), "float_stock_cap").copy()
        else:
            top = top400.nlargest(100, "float_stock_cap").copy()
        top["weight_new"] = calc_weights_float(top, cap_pct=None).values
        return top

    # ── FnGuide 은행 고배당 플러스 TOP 10 ────────────────
    if "은행 고배당" in idx_name and "FnGuide" in provider:
        sec = stocks[
            stocks["fics_m"].str.contains("은행|상업은행", na=False)
        ].copy()
        if not sec.empty and target_n:
            sec = sec.nlargest(int(target_n), "stock_cap")
            sec["weight_new"] = calc_weights_float(sec, cap_pct=0.35).values
        return sec if not sec.empty else None

    # ── FnGuide K-반도체 지수 ────────────────────────────
    if ("K-반도체" in idx_name or "K반도체" in idx_name) and "FnGuide" in provider:
        sec = stocks[
            stocks["fics_m"].str.contains("반도체", na=False) &
            (stocks["float_stock_cap"].fillna(0) > 100_000_000_000)  # 1000억 이상
        ].copy()
        if not sec.empty and target_n:
            sec = sec.nlargest(int(target_n), "float_stock_cap")
            sec["weight_new"] = calc_weights_float(sec, cap_pct=0.35).values
        return sec if not sec.empty else None

    # ── FnGuide 반도체 밸류체인 ──────────────────────────
    if "반도체 밸류체인" in idx_name and "FnGuide" in provider:
        sec = stocks[
            (stocks["fics_l"] == "IT") &
            stocks["fics_m"].str.contains("반도체|하드웨어", na=False)
        ].copy()
        if not sec.empty and target_n:
            n = int(target_n) if target_n else 70
            sec = sec.nlargest(n, "float_stock_cap")
            sec["weight_new"] = calc_weights_float(sec, cap_pct=0.30).values
        return sec if not sec.empty else None

    # ── FnGuide 리츠부동산인프라 ─────────────────────────
    if "리츠부동산인프라" in idx_name and "FnGuide" in provider:
        sec = stocks[
            stocks["fics_m"].str.contains("리츠|부동산|인프라", na=False) &
            (stocks["stock_market"] == "KOSPI")
        ].copy()
        if not sec.empty and target_n:
            sec = sec.nlargest(int(target_n), "stock_cap")
            sec["weight_new"] = calc_weights_float(sec, cap_pct=0.30).values
        return sec if not sec.empty else None

    # ── FnGuide IT플러스 ──────────────────────────────────
    if "IT플러스" in idx_name and "FnGuide" in provider:
        sec = stocks[
            (stocks["fics_l"] == "IT") &
            stocks["fics_m"].str.contains("소프트웨어|하드웨어|반도체|디스플레이", na=False) &
            (stocks["stock_cap"].fillna(0) > 300_000_000_000)
        ].copy()
        if not sec.empty:
            sec["weight_new"] = calc_weights_float(sec, cap_pct=0.25).values
        return sec if not sec.empty else None

    # ── FnGuide K-신재생에너지 플러스 ────────────────────
    if "신재생에너지" in idx_name and "FnGuide" in provider:
        sec = stocks[
            stocks["fics_m"].str.contains("신재생|태양광|풍력|에너지", na=False)
        ].copy()
        if not sec.empty and target_n:
            sec = sec.nlargest(int(target_n), "float_stock_cap")
            sec["weight_new"] = calc_weights_float(sec, cap_pct=0.20).values
        return sec if not sec.empty else None

    # ── FnGuide 자동차 TOP3 플러스 ───────────────────────
    if "자동차 TOP3" in idx_name and "FnGuide" in provider:
        sec = stocks[
            stocks["fics_m"].str.contains("자동차|부품", na=False) &
            (stocks["stock_cap"].fillna(0) > 200_000_000_000)
        ].copy()
        if not sec.empty and target_n:
            sec = sec.nlargest(int(target_n), "stock_cap")
            sec["weight_new"] = calc_weights_float(sec, cap_pct=0.30).values
        return sec if not sec.empty else None

    # ── FnGuide 기업가치 지수 ────────────────────────────
    if "기업가치" in idx_name and "FnGuide" in provider:
        universe = stocks[
            (stocks["stock_cap"].fillna(0) > 1_000_000_000_000) &  # 1조 이상
            (stocks["stock_float_rate"].fillna(0) >= 10)
        ].copy()
        if not universe.empty and target_n:
            top = universe.nlargest(int(target_n), "float_stock_cap")
            top["weight_new"] = calc_weights_float(top, cap_pct=0.25).values
            return top
        return None

    # ── WISE 삼성그룹 밸류인덱스 ─────────────────────────
    if "삼성그룹" in idx_name:
        # 종목명에 '삼성' 포함 종목 (근사값)
        if "stock_name" in stocks.columns:
            sec = stocks[stocks["stock_name"].str.contains("삼성", na=False)].copy()
        else:
            # 삼성 계열사는 약 20개 → 시총 상위로 근사
            sec = stocks[stocks["stock_cap"].fillna(0) > 0].nlargest(20, "stock_cap").copy()
        if not sec.empty:
            sec["weight_new"] = calc_weights_float(sec, cap_pct=0.30).values
        return sec if not sec.empty else None

    # ── FnGuide 네트워크 인프라 ──────────────────────────
    if "네트워크 인프라" in idx_name:
        sec = stocks[
            (stocks["fics_l"] == "IT") &
            (stocks["stock_cap"].fillna(0) > 0)
        ].copy()
        if not sec.empty and target_n:
            sec = sec.nlargest(int(target_n), "stock_cap")
            sec["weight_new"] = calc_weights_float(sec, cap_pct=0.15).values
        return sec if not sec.empty else None

    # 나머지: 산출 불가 → None 반환 (현재 보유 종목 그대로 표시)
    return None


# ─────────────────────────────────────────────────────────────
#  COMPUTE PRE-REQS (KOSPI200 etc.)
# ─────────────────────────────────────────────────────────────

def compute_krx_parents(stocks: pd.DataFrame) -> dict:
    """KOSPI200, KRX300 사전 계산"""
    results = {}

    current_k200 = set(
        stocks.loc[stocks["member_kospi200"].fillna(False).astype(bool), "stock_ticker_ks"]
    )

    # KOSPI 200
    log.info("KOSPI 200 재구성 중...")
    kospi200 = apply_krx_8step(
        stocks,
        market="KOSPI",
        target_n=200,
        cap_threshold=0.85,
        keep_buffer=1.10,
        new_buffer=0.90,
        large_cap_rank=50,
        ceiling_pct=None,
        sector_col="stock_gics",
    )
    # member_kospi200 override
    kospi200["member_kospi200"] = True
    results["KOSPI200"] = kospi200
    log.info(f"  → {len(kospi200)}종목 선정")

    # KRX 300
    log.info("KRX 300 재구성 중...")
    krx300 = apply_krx_8step(
        stocks,
        market="KOSPI",  # KRX300은 KOSPI+KOSDAQ이지만 단순화
        target_n=300,
        cap_threshold=0.80,
        keep_buffer=1.10,
        new_buffer=0.90,
        large_cap_rank=None,
        ceiling_pct=None,
        sector_col="stock_gics",
    )
    results["KRX300"] = krx300
    log.info(f"  → {len(krx300)}종목 선정")

    return results


# ─────────────────────────────────────────────────────────────
#  MAIN ANALYSIS
# ─────────────────────────────────────────────────────────────

def run_analysis(
    methodology: list,
    index_etf_map: dict,
    holdings_df: pd.DataFrame,
    aum_df: pd.DataFrame,
    stocks: pd.DataFrame,
    stock_names: dict,
) -> dict:
    """
    모든 지수에 대해 변경 전/후 비중 + 금액 산출
    Returns: {idx_name: DataFrame(…)}
    """
    # Pre-compute KRX parent indices
    krx_parents = compute_krx_parents(stocks)

    # AUM 딕셔너리
    aum_dict = (
        dict(zip(aum_df["etf_ticker_ks"], aum_df["etf_asset"]))
        if not aum_df.empty else {}
    )

    results = {}

    for row in methodology:
        idx_name = row["지수명(한)"]
        if not idx_name:
            continue

        log.info(f"처리 중: {idx_name}")
        etf_tickers = index_etf_map.get(idx_name, [])

        # 현재 보유 종목 (모든 ETF 합산 → 가중 평균 비중)
        current_holdings = {}  # stock_ticker → {value_per, etf_ticker, aum}
        if etf_tickers:
            etf_hold = holdings_df[holdings_df["etf_ticker_ks"].isin(etf_tickers)]
            for _, h in etf_hold.iterrows():
                ticker = h["stock_ticker_ks"]
                etf    = h["etf_ticker_ks"]
                w_pct  = h["value_per"] or 0.0
                aum    = aum_dict.get(etf, 0.0) or 0.0
                if ticker not in current_holdings:
                    current_holdings[ticker] = {"weight_before": 0.0, "aum": 0.0, "etf_ticker": etf}
                current_holdings[ticker]["weight_before"] += w_pct
                current_holdings[ticker]["aum"] = max(current_holdings[ticker]["aum"], aum)

        # 신규 비중 계산
        new_comp = route_and_compute(idx_name, row, stocks, krx_parents)

        new_weights = {}
        if new_comp is not None and not new_comp.empty:
            for _, r in new_comp.iterrows():
                t = r["stock_ticker_ks"]
                w = r.get("weight_new", 0.0) or 0.0
                new_weights[t] = w * 100  # decimal → %

        # 전체 종목 집합
        all_tickers = set(current_holdings.keys()) | set(new_weights.keys())

        # ETF AUM (첫 번째 ETF 기준)
        total_aum = 0.0
        if etf_tickers:
            total_aum = sum(aum_dict.get(t, 0.0) for t in etf_tickers)

        rows_out = []
        for ticker in sorted(all_tickers):
            w_before = current_holdings.get(ticker, {}).get("weight_before", 0.0)
            w_after  = new_weights.get(ticker, 0.0) if new_comp is not None else w_before

            delta = w_after - w_before
            # 금액 = 변경 비중(%p) × etf_asset / 100
            amount = delta * total_aum / 100.0

            if abs(delta) < 0.001 and new_comp is not None:
                action = "유지"
            elif new_comp is None:
                action = "방법론 미구현"
            elif w_before == 0 and w_after > 0:
                action = "편입"
            elif w_before > 0 and w_after == 0:
                action = "편출"
            else:
                action = "비중변경"

            sname = stock_names.get(ticker, "")
            cap = stocks.set_index("stock_ticker_ks")["stock_cap"].get(ticker, 0.0)
            fmc = stocks.set_index("stock_ticker_ks")["float_stock_cap"].get(ticker, 0.0)

            rows_out.append({
                "종목코드":      ticker,
                "종목명":        sname,
                "현재_편입비중": round(w_before, 4),
                "신규_편입비중": round(w_after, 4),
                "비중변화_pct":  round(delta, 4),
                "구분":          action,
                "ETF_AUM":       round(total_aum, 2),
                "추정금액":      round(amount, 2),
                "시가총액":      cap,
                "유동시가총액":  fmc,
            })

        df_out = pd.DataFrame(rows_out)
        df_out["_sort"] = df_out["비중변화_pct"].abs()
        df_out = df_out.sort_values("_sort", ascending=False).drop(columns="_sort")
        df_out.reset_index(drop=True, inplace=True)

        results[idx_name] = {
            "df":          df_out,
            "etf_tickers": etf_tickers,
            "methodology_row": row,
            "computable":  new_comp is not None,
        }

    return results


# ─────────────────────────────────────────────────────────────
#  EXCEL GENERATION
# ─────────────────────────────────────────────────────────────

def _hdr_fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def _thin_border() -> Border:
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)


def _font(bold: bool = False, color: str = "000000", size: int = 10) -> Font:
    return Font(bold=bold, color=color, size=size, name="맑은 고딕")


def _align(h: str = "center", v: str = "center", wrap: bool = False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def _apply_header_row(ws, row: int, cols: list, fill_color: str):
    for i, col_name in enumerate(cols, 1):
        cell = ws.cell(row=row, column=i, value=col_name)
        cell.fill = _hdr_fill(fill_color)
        cell.font = _font(bold=True, color="FFFFFF")
        cell.alignment = _align("center")
        cell.border = _thin_border()


def _set_col_widths(ws, col_widths: dict):
    for col, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


def clean_sheet_name(name: str, max_len: int = 31) -> str:
    """Excel 시트명 정제 (31자 제한, 특수문자 제거)"""
    name = re.sub(r'[\\/:?*\[\]]', '', name)
    if len(name) > max_len:
        name = name[:max_len - 3] + "..."
    return name


def fill_index_sheet(ws, idx_name: str, data: dict):
    """개별 지수 탭 작성"""
    df        = data["df"]
    etf_tickers = data.get("etf_tickers", [])
    row_meta  = data.get("methodology_row", {})
    computable = data.get("computable", False)

    # ── 메타 정보 헤더 ────────────────────────────────────
    ws.cell(row=1, column=1, value="지수명").font = _font(bold=True)
    ws.cell(row=1, column=2, value=idx_name).font  = _font()
    ws.cell(row=2, column=1, value="제공사").font  = _font(bold=True)
    ws.cell(row=2, column=2, value=row_meta.get("제공사", "")).font = _font()
    ws.cell(row=3, column=1, value="대상 ETF").font = _font(bold=True)
    ws.cell(row=3, column=2, value=", ".join(etf_tickers) or "매칭 없음").font = _font()
    ws.cell(row=4, column=1, value="방법론 구현").font = _font(bold=True)
    ws.cell(row=4, column=2, value="O (자동 산출)" if computable else "△ (현재 보유 기준)").font = _font()
    ws.cell(row=5, column=1, value="편출입 산식").font = _font(bold=True)
    ws.cell(row=5, column=2, value=row_meta.get("편출입 결정 산식", "")).font = _font()
    ws.cell(row=5, column=2).alignment = _align("left", "center", wrap=True)
    ws.row_dimensions[5].height = 40
    ws.merge_cells(start_row=5, start_column=2, end_row=5, end_column=8)

    # ── 컬럼 헤더 ────────────────────────────────────────
    HDR_ROW = 7
    headers = [
        "종목코드", "종목명",
        "현재 편입 비중 (%)",
        "신규 편입 비중 (%)",
        "비중 변화 (%p)",
        "구분",
        "ETF AUM",
        "추정 금액\n(AUM × 비중변화/100)",
    ]
    _apply_header_row(ws, HDR_ROW, headers, CLR_HEADER)

    col_w = {1: 14, 2: 22, 3: 16, 4: 16, 5: 14, 6: 14, 7: 14, 8: 22}
    _set_col_widths(ws, col_w)

    # ── 데이터 행 ─────────────────────────────────────────
    ACTION_COLORS = {
        "편입":     CLR_BUY,
        "편출":     CLR_SELL,
        "비중변경": CLR_NEW,
        "유지":     CLR_HOLD,
        "방법론 미구현": "F2F2F2",
    }

    for ri, (_, row) in enumerate(df.iterrows(), start=HDR_ROW + 1):
        action = row["구분"]
        bg = ACTION_COLORS.get(action, CLR_HOLD)
        vals = [
            row["종목코드"],
            row["종목명"],
            row["현재_편입비중"],
            row["신규_편입비중"],
            row["비중변화_pct"],
            action,
            row["ETF_AUM"],
            row["추정금액"],
        ]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.fill = _hdr_fill(bg)
            c.border = _thin_border()
            c.font   = _font()
            if ci in (3, 4, 5):
                c.number_format = "0.0000"
                c.alignment = _align("right")
            elif ci in (7, 8):
                c.number_format = "#,##0.00"
                c.alignment = _align("right")
            else:
                c.alignment = _align("left")

    ws.freeze_panes = f"A{HDR_ROW + 1}"


def fill_summary_sheet(ws, results: dict, stocks: pd.DataFrame):
    """개별 종목 종합 시트 작성"""
    ws.cell(row=1, column=1, value="ETF 리밸런싱 개별 종목 종합 (2026년 6월)").font = Font(
        bold=True, size=14, color=CLR_HEADER, name="맑은 고딕"
    )
    ws.cell(row=2, column=1, value=f"산출일: {datetime.today().strftime('%Y-%m-%d')}").font = _font()

    HDR_ROW = 4
    headers = [
        "종목코드", "종목명",
        "편입 추정금액 합계",
        "편출 추정금액 합계",
        "순매수 금액",
        "현재 시가총액",
        "유동 시가총액",
        "관련 지수 수",
    ]
    _apply_header_row(ws, HDR_ROW, headers, CLR_SUMHDR)

    # 집계
    agg: dict = {}
    for idx_name, data in results.items():
        df = data["df"]
        for _, row in df.iterrows():
            t = row["종목코드"]
            if not t:
                continue
            if t not in agg:
                agg[t] = {
                    "종목명": row["종목명"],
                    "buy": 0.0, "sell": 0.0, "count": 0,
                }
            amt = row["추정금액"]
            if amt > 0:
                agg[t]["buy"]  += amt
            elif amt < 0:
                agg[t]["sell"] += amt
            agg[t]["count"] += 1

    # 시가총액 매핑
    cap_map  = stocks.set_index("stock_ticker_ks")["stock_cap"].to_dict()
    fmc_map  = stocks.set_index("stock_ticker_ks")["float_stock_cap"].to_dict()

    rows = []
    for t, v in agg.items():
        net = v["buy"] + v["sell"]
        rows.append({
            "종목코드":         t,
            "종목명":          v["종목명"],
            "편입금액합계":     round(v["buy"], 2),
            "편출금액합계":     round(v["sell"], 2),
            "순매수금액":       round(net, 2),
            "현재시가총액":     cap_map.get(t, 0.0),
            "유동시가총액":     fmc_map.get(t, 0.0),
            "관련지수수":       v["count"],
        })

    df_sum = pd.DataFrame(rows).sort_values("순매수금액", ascending=False)

    COL_FMT = {3: "#,##0.00", 4: "#,##0.00", 5: "#,##0.00", 6: "#,##0", 7: "#,##0"}
    for ri, (_, row) in enumerate(df_sum.iterrows(), start=HDR_ROW + 1):
        net = row["순매수금액"]
        bg  = CLR_BUY if net > 0 else (CLR_SELL if net < 0 else CLR_HOLD)
        vals = [
            row["종목코드"], row["종목명"],
            row["편입금액합계"], row["편출금액합계"], row["순매수금액"],
            row["현재시가총액"], row["유동시가총액"], row["관련지수수"],
        ]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.fill  = _hdr_fill(bg)
            c.border = _thin_border()
            c.font  = _font()
            if ci in COL_FMT:
                c.number_format = COL_FMT[ci]
                c.alignment = _align("right")
            else:
                c.alignment = _align("left")

    col_w = {1: 14, 2: 22, 3: 18, 4: 18, 5: 16, 6: 18, 7: 18, 8: 10}
    _set_col_widths(ws, col_w)
    ws.freeze_panes = f"A{HDR_ROW + 1}"


def generate_excel(results: dict, stocks: pd.DataFrame):
    wb = Workbook()

    # 1. 요약 시트
    ws_sum = wb.active
    ws_sum.title = "개별종목_종합"
    fill_summary_sheet(ws_sum, results, stocks)

    # 2. 개별 지수 시트
    seen_names: set = set()
    for idx_name, data in results.items():
        sname = clean_sheet_name(idx_name)
        # 중복 탭명 처리
        base = sname
        cnt  = 1
        while sname in seen_names:
            sname = f"{base[:28]}_{cnt}"
            cnt  += 1
        seen_names.add(sname)

        ws = wb.create_sheet(title=sname)
        fill_index_sheet(ws, idx_name, data)
        log.info(f"  시트 생성: {sname}")

    wb.save(OUTPUT_XLSX)
    log.info(f"Excel 저장 완료: {OUTPUT_XLSX}")


# ─────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────

def main():
    token = os.environ.get("MOTHERDUCK_TOKEN")
    if not token:
        log.error("MOTHERDUCK_TOKEN 환경변수가 설정되지 않았습니다.")
        sys.exit(1)

    # 1. MotherDuck 연결
    conn = connect_motherduck(token)
    db   = discover_db(conn)

    # 2. 방법론 로드
    methodology     = load_june2026_methodology()
    etf_excel_map   = load_etf_matching_from_excel()

    # 3. DB 데이터 로드
    etf_db_map  = load_etf_index_map(conn, db)
    stocks      = load_stock_data(conn, db)
    stock_names = load_stock_names(conn, db)
    log.info(f"종목 데이터: {len(stocks)}행, 컬럼: {list(stocks.columns)}")

    # 4. ETF 매핑 구성
    index_etf_map = build_index_etf_map(methodology, etf_db_map, etf_excel_map)
    all_etf_tickers = sorted({t for tl in index_etf_map.values() for t in tl})
    log.info(f"대상 ETF: {len(all_etf_tickers)}개")

    # 5. 현재 보유 및 AUM 로드
    holdings_df = load_current_holdings(conn, db, all_etf_tickers)
    aum_df      = load_etf_aum(conn, db, all_etf_tickers)
    log.info(f"보유 데이터: {len(holdings_df)}행, AUM: {len(aum_df)}행")

    # 6. 종목명 보완 (stock_data에서)
    if "stock_name" in stocks.columns:
        for _, r in stocks.iterrows():
            t = r["stock_ticker_ks"]
            n = r.get("stock_name", "")
            if t and n and t not in stock_names:
                stock_names[t] = n

    # 7. 분석 실행
    results = run_analysis(
        methodology, index_etf_map,
        holdings_df, aum_df,
        stocks, stock_names,
    )

    # 8. Excel 생성
    generate_excel(results, stocks)
    log.info("완료!")


if __name__ == "__main__":
    main()
