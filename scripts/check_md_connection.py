"""
MotherDuck 연결 테스트 + DB 스키마 확인 스크립트
Usage: MOTHERDUCK_TOKEN=<token> python scripts/check_md_connection.py
"""
import os, sys
import duckdb

token = os.environ.get("MOTHERDUCK_TOKEN")
if not token:
    print("❌ MOTHERDUCK_TOKEN 환경변수 없음")
    sys.exit(1)

print("=== MotherDuck 연결 테스트 ===")
try:
    conn = duckdb.connect(f"md:?motherduck_token={token}")
    print("✅ 연결 성공!")
except Exception as e:
    print(f"❌ 연결 실패: {e}")
    sys.exit(1)

# DB 목록
print("\n=== 데이터베이스 목록 ===")
dbs = conn.execute("SHOW DATABASES").df()
print(dbs.to_string(index=False))

# 각 DB의 테이블 탐색
target_tables = {"dim_etf", "fact_etf_holdings_daily", "fact_etf_nav",
                 "fact_stock_data", "dim_gics", "dim_fics"}

for db in dbs["database_name"]:
    if db in ("memory", "system", "temp"):
        continue
    try:
        tables = conn.execute(f"SHOW TABLES IN {db}").df()
        tbl_list = tables.iloc[:, 0].tolist()
        found = [t for t in tbl_list if t in target_tables]
        print(f"\n=== DB: {db} ({len(tbl_list)}개 테이블) ===")
        for t in tbl_list:
            mark = "✅" if t in target_tables else "  "
            print(f"  {mark} {t}")

        # 발견된 테이블 상세 확인
        for t in found:
            print(f"\n  --- {db}.{t} 컬럼 ---")
            try:
                cols = conn.execute(f"DESCRIBE {db}.{t}").df()
                print(cols[["column_name","column_type"]].to_string(index=False))
            except Exception as e:
                print(f"    컬럼 조회 실패: {e}")

            print(f"\n  --- {db}.{t} 행 수 + 최신 날짜 ---")
            try:
                cnt = conn.execute(f"SELECT COUNT(*) as cnt FROM {db}.{t}").fetchone()[0]
                print(f"    행 수: {cnt:,}")
                # date 컬럼 확인
                cols2 = conn.execute(f"DESCRIBE {db}.{t}").df()
                date_cols = cols2[cols2["column_name"].str.contains("date|Date", na=False)]["column_name"].tolist()
                for dc in date_cols[:1]:
                    mx = conn.execute(f"SELECT MAX({dc}) FROM {db}.{t}").fetchone()[0]
                    print(f"    최신 {dc}: {mx}")
            except Exception as e:
                print(f"    통계 조회 실패: {e}")

    except Exception as e:
        print(f"\n  {db}: 테이블 목록 조회 실패 ({e})")

print("\n=== 완료 ===")
