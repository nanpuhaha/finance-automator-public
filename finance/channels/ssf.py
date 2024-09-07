from pathlib import Path

import pandas as pd

# SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame.
# Try using .loc[row_indexer,col_indexer] = value instead
# See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
pd.options.mode.chained_assignment = None


def run(order_file: Path, delivery_file: Path, result_file: Path):
    """SSG

    Args:
        order_file (Path): 정산상세내역
        delivery_file (Path): 상세 배송비내역
        result_file (Path): 최종파일
    """

    CHANNEL_NAME = "SSF"

    # 데이터를 pandas DataFrame으로 읽기
    order_df = pd.read_excel(order_file)
    delivery_df = pd.read_excel(delivery_file)

    df = order_df.copy(deep=True)

    # ---------- 합계 행 제거 ----------
    # '정산기간' 값이 '합계'가 아닌 행만 남기기
    df = df.loc[df['정산기간'] != '합계']

    # ---------- 주문일자 ----------
    # 주문번호로부터 주문일자를 추출하여 '주문일자' 컬럼을 추가
    # 주문번호 OD202404042502341 -> 주문일자 2024-04-04
    df['주문일자'] = df['주문번호'].apply(lambda x: f"{x[2:6]}-{x[6:8]}-{x[8:10]}")

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
    df['차수'] = '-'
    df["주문일(결제일)"] = df["주문일자"]
    df["주문번호"] = df["주문번호"]
    df["상품주문번호"] = '-'
    df["상품명"] = df["상품명"]
    # df["옵션"] = df["옵션"]
    df["컬러"] = "-"
    df["사이즈"] = "-"
    # df["수량"] = df["수량"]
    df["(판매처) 정상판매가"] = df['정소가']
    df["할인 (플랫폼)"] = df['물산부담(C)']
    df["할인 (브랜드)"] = df['정소가'] - df['판매가(A)']
    df["고객결제"] = df['결제금액(A-B)']
    df["매출(v+)"] = df['결제금액(A-B)']
    df["수수료"] = df['입점수수료(F) (물산부담(C)제외)']
    df["정산가"] = df['업체지급금액(G)']
    df["매출구분"] = "제품"
    df_result = df[COLUMNS]

    # '브랜드명'이 "시티"이면서 '구분'이 '판매'인 데이터만 추출하여 새로운 데이터프레임 생성
    df_city_sale = df.query('브랜드명 == "Citybreeze"')
    df_artid_sale = df.query('브랜드명 == "ARTID"')

    df_city_sale_data = df_city_sale[COLUMNS]
    df_artid_sale_data = df_artid_sale[COLUMNS]

    # ---------- 저장 ----------
    with pd.ExcelWriter(result_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="전체")
        df_city_sale.to_excel(writer, index=False, sheet_name="판매_시티")
        df_artid_sale.to_excel(writer, index=False, sheet_name="판매_아티드")
        df_city_sale_data.to_excel(writer, index=False, sheet_name="매출자료_시티")
        df_artid_sale_data.to_excel(writer, index=False, sheet_name="매출자료_아티드")


if __name__ == '__main__':
    TOP_DIR = Path('/Users/jwseo/Movies/9월정산')
    ROZLEY_DIR = TOP_DIR / "로즐리"
    CITY_DIR = TOP_DIR / "시티브리즈"
    ARTID_DIR = TOP_DIR / "아티드"

    # order_file = TOP_DIR / '아티드_SSF_정산내역_2408.xlsx'
    order_file = TOP_DIR / '아티드_SSF_입점사 정산상세내역_2408.xlsx'

    # delivery_file = TOP_DIR / '아티드_SSF_배송내역_2408.xlsx'
    delivery_file = TOP_DIR / '아티드_SSF_상세 배송비내역_2408.xlsx'

    result_file = TOP_DIR / '아티드_SSF_통합✅_2408.xlsx'
    run(order_file, delivery_file, result_file)
