from pathlib import Path

import pandas as pd

# SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame.
# Try using .loc[row_indexer,col_indexer] = value instead
# See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
pd.options.mode.chained_assignment = None


def run(order_file: Path, product_file: Path, result_file: Path):
    """무신사 글로벌

    Args:
        order_file (Path): 정산내역
        product_file (Path): 상품내역
        result_file (Path): 최종파일
    """

    CHANNEL_NAME = "무신사"
    SUB_CHANNEL_NAME = "통합계정_글로벌"

    order_df = pd.read_excel(order_file)
    product_df = pd.read_excel(product_file)

    df = order_df

    # ---------- 컬럼 정리 ----------

    # 첫 번째 행 제거
    # df = df.drop(index=0)

    def remove_rows(df, start, end):
        # Drop the specified range of rows by their indices
        df = df.drop(df.index[start : end + 1])
        # Reset the index to keep the DataFrame tidy
        df = df.reset_index(drop=True)
        return df

    df = remove_rows(df, 0, 1)

    # ---------- 주문일자 ----------
    # '주문일' 컬럼으로부터 '주문일자' 컬럼을 생성
    # 주문일: 20240401
    # 주문일자: 2024-04-01
    def extract_order_date(x):
        x = str(x)
        # return f"{x[0:4]}-{x[4:6]}-{x[6:8]}"
        return x[0:10]

    df['주문일자'] = df['주문일'].map(extract_order_date)
    # df['주문일자'] = df['주문일'].apply(lambda x: f"{x[0:4]}-{x[4:6]}-{x[6:8]}")

    # ---------- 상품명(한글) ----------
    # TODO: 상품내역 파일에서 스타일넘버를 기준으로 한글 상품명 가져오기.
    product_df = product_df.rename(columns={'상품명': '상품명(한글)'})
    df_right = product_df[['스타일넘버', '상품명(한글)']]
    df_right = df_right.drop_duplicates()
    df = pd.merge(df, df_right, how='left', on='스타일넘버')

    # ---------- 매출자료 양식에 맞게 변경 ----------
    COLUMNS = [
        "채널명",
        "차수",
        "주문일(결제일)",
        "주문번호",
        "상품주문번호",
        "상품명",
        "옵션",
        "컬러",
        "사이즈",
        "수량",
        "(판매처) 정상판매가",
        "할인 (플랫폼)",
        "할인 (브랜드)",
        "고객결제",
        "매출(v+)",
        "수수료",
        "정산가",
        "매출구분",
    ]
    df["채널명"] = CHANNEL_NAME
    df['차수'] = SUB_CHANNEL_NAME
    df["주문일(결제일)"] = df["주문일자"]
    # df["주문번호"] = df["주문번호"]
    df["상품주문번호"] = df['주문일련번호']
    df["상품명"] = df["상품명(한글)"]  #
    # df["옵션"] = df['옵션']
    df["컬러"] = "-"
    df["사이즈"] = "-"
    # df["수량"] = df["수량"]
    df["(판매처) 정상판매가"] = df['판매금액소계']
    df["할인 (플랫폼)"] = df['PROMO 할인 금액(원화)']
    df["할인 (브랜드)"] = df['할인금액']
    df["고객결제"] = df['매출액-배송비-관세']
    df["매출(v+)"] = df['고객결제']
    df["수수료"] = df['판매수수료']
    df["정산가"] = df["매출(v+)"] - df["수수료"]
    df["매출구분"] = "제품"

    df_result = df[COLUMNS]

    # # ---------- 저장 ----------
    with pd.ExcelWriter(result_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="전체")
        df_result.to_excel(writer, index=False, sheet_name="매출자료")


if __name__ == '__main__':
    TOP_DIR = Path('/Users/jwseo/Movies/9월정산')
    # ROZLEY_DIR = TOP_DIR / "로즐리"
    # CITY_DIR = TOP_DIR / "시티브리즈"
    # ARTID_DIR = TOP_DIR / "아티드"

    order_file = TOP_DIR / '시티_무신사_글로벌_정산내역_2408.xlsx'
    product_file = TOP_DIR / '시티_무신사글로벌_상품내역_2408.xlsx'
    result_file = TOP_DIR / '시티_글로벌_무신사_통합✅_2408.xlsx'
    run(order_file, product_file, result_file)
