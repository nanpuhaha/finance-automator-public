from pathlib import Path

import pandas as pd

# SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame.
# Try using .loc[row_indexer,col_indexer] = value instead
# See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
pd.options.mode.chained_assignment = None


def run(order_file: Path, deduction_file: Path, result_file: Path):
    """에이블리

    Args:
        order_file (Path): 매출내역(구매확정)
        deduction_file (Path): 차감내역(업체)
        result_file (Path): 최종파일
    """

    CHANNEL_NAME = "롯데ON"

    # 데이터를 pandas DataFrame으로 읽기
    order_df = pd.read_excel(order_file)
    deduction_df = pd.read_excel(deduction_file)

    df = order_df.copy(deep=True)

    # ---------- 합계 행 삭제 ----------
    # '주문번호' 컬럼 값이 비어있는데, '판매금액' 컬럼 값이 비어있지 않은 행을 삭제
    df = df[~(df["주문번호"].isna() & df["판매금액"].notna())]
    # df = df.dropna(subset=["주문번호"]) # '주문번호'가 NaN인 행 삭제

    # ---------- 브랜드명 ----------
    # '하위거래처명' 컬럼이 NaN인 경우, '로즐리'로 채우기
    df["하위거래처명"] = df["하위거래처명"].fillna("로즐리")
    df["브랜드명"] = df["하위거래처명"]

    # ---------- 주문일자 ----------
    # '주문번호' 컬럼으로부터 '주문일자' 컬럼을 생성
    # 주문번호 : 2024041815378905
    # 주문일자 : 2024-04-18

    def extract_order_date(order_number):
        # 주문번호를 변환하여 주문일자를 추출하는 함수
        # 문자열이 아닌 경우 문자열로 변환
        order_number = str(order_number)
        return f"{order_number[0:4]}-{order_number[4:6]}-{order_number[6:8]}"

    # 주문번호를 변환하여 주문일자 컬럼 추가
    df["주문일자"] = df["주문번호"].apply(extract_order_date)
    # df['주문일자'] = df['주문번호'].apply(lambda x: f"{x[0:4]}-{x[4:6]}-{x[6:8]}")
    # TypeError: x = 2024050617567955.0, type(x) = <class 'float'>

    # ---------- 구분 ----------
    # '구분' 컬럼 생성
    # FIXME: 무엇을 기준으로 '구분'을 생성할 것인가?
    # FIXME: '판매자상품명' = '()' 인 경우를 '배송비'로 처리할 수도 있고,
    # FIXME: '표준카테고리' = NaN 인 경우를 '배송비'로 처리할 수도 있고,
    # FIXME: '판매수량' = 0 인 경우를 '배송비'로 처리할 수도 있고,
    # FIXME: '총배송비' > 0 인 경우를 '배송비'로 처리할 수도 있고,
    df["구분"] = df["판매수량"].apply(lambda x: "배송비" if x == 0 else "판매")

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
    CITY_COLUMNS = [*COLUMNS, "제품구분"]
    df["채널명"] = CHANNEL_NAME
    df["차수"] = "-"
    # 표준카테고리
    # '브랜드남성의류>반팔티셔츠>라운드넥반팔티'
    # '남성의류>트레이닝복>트레이닝하의'
    # 시티브리즈 맨
    # df['차수'] = df['표준카테고리'].apply(lambda x: '맨즈별도' if '남성의류' in x else "-")
    # df['제품구분'] = df['표준카테고리'].apply(lambda x: '맨' if '남성의류' in x else "우먼")

    df["주문일(결제일)"] = df["주문일자"]
    df["주문번호"] = df["주문번호"]
    df["상품주문번호"] = "-"
    df["상품명"] = df["판매자상품명"]
    df["옵션"] = "-"
    df["컬러"] = "-"
    df["사이즈"] = "-"
    df["수량"] = df["판매수량"]
    df["(판매처) 정상판매가"] = df["판매금액"]
    df["할인 (플랫폼)"] = df["상품할인(당사부담)"]
    df["할인 (브랜드)"] = df["셀러즉시할인금액"] + df["상품할인(셀러부담)"]
    df["고객결제"] = (
        df["(판매처) 정상판매가"] - df["할인 (플랫폼)"] - df["할인 (브랜드)"]
    )
    df["매출(v+)"] = df["정산대상판매금액"]
    df["수수료"] = df["총 수수료 합계"]
    df["정산가"] = df["지급대상금액"]
    df["매출구분"] = "제품"

    df_result = df[COLUMNS]

    # '브랜드명'이 "시티"이면서 '구분'이 '판매'인 데이터만 추출하여 새로운 데이터프레임 생성
    df_artid_sale = df.query('브랜드명 == "아티드" and 구분 == "판매"')
    df_city_sale = df.query(
        '"시티브리즈" in 브랜드명 and 구분 == "판매"'
    )  # '시티브리즈 맨'도 포함
    df_rozley_sale = df.query('브랜드명 == "로즐리" and 구분 == "판매"')

    def is_man_brand(brand_name: str) -> bool:
        return brand_name == "시티브리드 맨"

    def check_is_man(brand_name: str, man_value: str, woman_value: str) -> str:
        return man_value if is_man_brand(brand_name) else woman_value

    df_city_sale.loc[:, "차수"] = df_city_sale["브랜드명"].apply(
        # df_city_sale['차수'] = df_city_sale['브랜드명'].apply(
        lambda x: "맨즈별도" if is_man_brand(x) else "-"
    )
    df_city_sale.loc[:, "제품구분"] = df_city_sale["브랜드명"].apply(
        # df_city_sale['제품구분'] = df_city_sale['브랜드명'].apply(
        lambda x: "맨" if is_man_brand(x) else "우먼"
    )
    # ---------- SettingWithCopyWarning ----------
    # SettingWithCopyWarning:
    # A value is trying to be set on a copy of a slice from a DataFrame.
    # Try using .loc[row_indexer,col_indexer] = value instead

    # See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
    #   df_city_sale['차수'] = df_city_sale['브랜드명'].apply(

    # ---------- SettingWithCopyWarning ----------
    # SettingWithCopyWarning:
    # A value is trying to be set on a copy of a slice from a DataFrame.
    # Try using .loc[row_indexer,col_indexer] = value instead

    # See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
    #   df_city_sale.loc[:, '제품구분'] = df_city_sale['브랜드명'].apply(

    df_artid_sale_data = df_artid_sale[COLUMNS]
    df_city_sale_data = df_city_sale[CITY_COLUMNS]
    df_rozley_sale_data = df_rozley_sale[COLUMNS]

    GROUP_COLUMNS = ["지급대상금액"]
    df_grouped = df.groupby(by=["브랜드명", "구분"], group_keys=True)[
        GROUP_COLUMNS
    ].sum()
    df_grouped2 = df.groupby(by=["구분", "브랜드명"], group_keys=True)[
        GROUP_COLUMNS
    ].sum()

    df_delivery = df.query('구분 == "배송비"')
    # 총합계 계산하여 새로운 행 추가
    df_grouped3 = df_delivery.groupby(by=["브랜드명"], group_keys=True)[
        GROUP_COLUMNS
    ].sum()
    df_grouped3.loc["총합계"] = df_delivery[GROUP_COLUMNS].sum()

    # ---------- 저장 ----------
    with pd.ExcelWriter(result_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="전체")
        df_artid_sale.to_excel(writer, index=False, sheet_name="판매_아티드")
        df_city_sale.to_excel(writer, index=False, sheet_name="판매_시티")
        df_rozley_sale.to_excel(writer, index=False, sheet_name="판매_로즐리")
        df_artid_sale_data.to_excel(writer, index=False, sheet_name="매출자료_아티드")
        df_city_sale_data.to_excel(writer, index=False, sheet_name="매출자료_시티")
        df_rozley_sale_data.to_excel(writer, index=False, sheet_name="매출자료_로즐리")
        df_delivery.to_excel(writer, index=False, sheet_name="배송")
        df_grouped.to_excel(writer, sheet_name="피봇1")
        df_grouped2.to_excel(writer, sheet_name="피봇2")
        df_grouped3.to_excel(writer, sheet_name="배송비_피봇")


if __name__ == "__main__":
    TOP_DIR = Path("/Users/jwseo/Movies/9월정산")
    ROZLEY_DIR = TOP_DIR / "로즐리"
    CITY_DIR = TOP_DIR / "시티브리즈"
    ARTID_DIR = TOP_DIR / "아티드"

    order_file = (
        TOP_DIR / "롯데ON_중개_매출내역(구매확정)_2408.xlsx"
    )  # 매출내역(구매확정)
    deduction_file = (
        TOP_DIR / "롯데ON_중개_차감내역(업체)_2408.xlsx"
    )  # 차감내역(업체)
    result_file = TOP_DIR / "롯데ON_통합✅_2408.xlsx"
    run(order_file, deduction_file, result_file)
