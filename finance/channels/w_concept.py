from pathlib import Path

import pandas as pd


def run(order_file: Path, delivery_file: Path, return_file: Path, result_file: Path):
    """W컨셉

    Args:
        order_file (Path): 정산내역
        delivery_file (Path): 반품내역
        return_file (Path): 배송완료내역
        result_file (Path): 최종파일
    """
    CHANNEL_NAME = "W컨셉"

    # 데이터를 pandas DataFrame으로 읽기
    order_df = pd.read_excel(order_file)
    delivery_df = pd.read_excel(delivery_file)
    return_df = pd.read_excel(return_file)

    df = order_df.copy(deep=True)

    # ---------- 브랜드명 ----------
    # '브랜드명' 컬럼의 NaN 값을 이전 값으로 채우기
    df["브랜드명"] = df["브랜드명"].ffill()

    # ---------- 주문일자 ----------
    return_df = return_df[["주문상세번호", "주문일자"]]
    return_df = return_df.drop_duplicates()
    return_df.columns = ["주문상세번호", "주문일자(반품)"]
    df = pd.merge(df, return_df, on="주문상세번호", how="left")

    delivery_df = delivery_df[["주문상세번호", "주문일자"]]
    delivery_df = (
        delivery_df.drop_duplicates()
    )  # 중복 제거 필요. 동일한 '주문상세번호'로 '주문유형'이 '교환회수완료', '교환배송완료' 2개의 데이터가 있을 수 있음.
    delivery_df.columns = ["주문상세번호", "주문일자(배송완료)"]
    df = pd.merge(df, delivery_df, on="주문상세번호", how="left")

    def determine_order_date(row):
        return_date = row["주문일자(반품)"]
        delivery_date = row["주문일자(배송완료)"]

        if pd.isna(return_date) and pd.isna(delivery_date):
            return None
        elif pd.isna(return_date):
            return delivery_date
        elif pd.isna(delivery_date):
            return return_date
        elif return_date == delivery_date:
            return return_date
        else:
            return "WARNING"

    df["주문일자"] = df.apply(determine_order_date, axis=1)

    # ---------- 구분 ----------
    # '구분' 컬럼의 값을 변경
    # '정상마진' -> '판매' , '역마진' -> '판매', '배송료' -> '배송비'로 변경
    df.rename(columns={"구분": "구분_원본"}, inplace=True)
    df["구분"] = df["구분_원본"].map(lambda x: "배송비" if x == "배송료" else "판매")

    # ---------- 차수 ----------
    # 새로운 컬럼 '차수'를 추가
    # '해외판매 수수료' 값이 0보다 크면 '해외', 0이면 '국내'로 설정
    df["차수"] = df["해외판매 수수료"].apply(lambda x: "해외" if x > 0 else "국내")

    # ---------- 중복 데이터 정리 ----------
    # 579	2024년 5월 	(주)이스트엔드	배송완료	2024-05-04	Y01969339	35755304	윤*영	윤*영	ARTID	305695347	20064130	0	믹스 니트 가디건 셋업_BLUE		BLUE_FREE		1	143100	0	101601	0	18628	0	0	0	0	0	0	0	0	124472	124472	22871	0	101601	29		37108548	판매		2024-04-12	2024-04-12	국내	W컨셉	2024-04-12	35755304	BLUE_FREE	-	-	143100	18628	0	124472	124472	22871	101601	제품
    # 579	2024년 5월 	(주)이스트엔드	배송완료	2024-05-04	Y01969339	35755304	윤*영	윤*영	ARTID	305695347	20064130	0	믹스 니트 가디건 셋업_BLUE		BLUE_FREE		1	143100	0	101601	0	18628	0	0	0	0	0	0	0	0	124472	124472	22871	0	101601	29		37108548	판매		2024-04-12	2024-04-12	국내	W컨셉	2024-04-12	35755304	BLUE_FREE	-	-	143100	18628	0	124472	124472	22871	101601	제품

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
    # df['차수'] = df['차수']
    df["주문일(결제일)"] = df["주문일자"]
    df["주문번호"] = df["주문번호"]
    df["상품주문번호"] = df["주문상세번호"]
    df["상품명"] = df["상품명"]
    df["옵션"] = df["옵션1"]
    df["컬러"] = "-"
    df["사이즈"] = "-"
    # df["수량"] = df["수량"]
    df["(판매처) 정상판매가"] = df["주문금액합계(기본가)"] + df["주문금액합계(옵션가)"]
    df["할인 (플랫폼)"] = df["본사 쿠폰"] + df["사용적립금"] + df["제휴몰 할인"]
    df["할인 (브랜드)"] = df["업체 쿠폰"]
    df["고객결제"] = df["고객결제금액"]
    df["매출(v+)"] = df["순매출금액"]
    df["수수료"] = df["세금계산서"] + df["해외판매 수수료"]
    df["정산가"] = df["업체지급금액"]
    df["매출구분"] = "제품"
    df_result = df[COLUMNS]

    # '브랜드명'이 "시티"이면서 '구분'이 '판매'인 데이터만 추출하여 새로운 데이터프레임 생성
    df_city_sale = df.query('브랜드명 == "CITYBREEZE" and 구분 == "판매"')
    df_artid_sale = df.query('브랜드명 == "ARTID" and 구분 == "판매"')

    df_city_sale_data = df_city_sale[COLUMNS]
    df_artid_sale_data = df_artid_sale[COLUMNS]

    GROUP_COLUMNS = ["순매출금액", "세금계산서", "해외판매 수수료", "업체지급금액"]
    df_grouped = df.groupby(by=["브랜드명", "구분"], group_keys=True)[
        GROUP_COLUMNS
    ].sum()
    df_grouped2 = df.groupby(by=["구분", "브랜드명"], group_keys=True)[
        GROUP_COLUMNS
    ].sum()

    def set_autofit_and_tab_color(
        writer: pd.ExcelWriter, sheet_name: str, color: str | None = None
    ):
        sheet = writer.sheets[sheet_name]
        sheet.autofit()
        if color:
            sheet.set_tab_color(color)

    def df_to_excel(
        df: pd.DataFrame,
        writer: pd.ExcelWriter,
        sheet_name: str,
        color: str | None = None,
        index: bool = False,
    ):
        df.to_excel(writer, index=index, sheet_name=sheet_name)
        set_autofit_and_tab_color(writer, sheet_name, color)

    # Define the custom method
    def _df_to_excel(
        self: pd.ExcelWriter,
        df: pd.DataFrame,
        sheet_name: str,
        color: str | None = None,
        index: bool = False,
    ):
        df.to_excel(self, index=index, sheet_name=sheet_name)
        set_autofit_and_tab_color(writer, sheet_name, color)

    # Add the custom method to the pd.ExcelWriter class
    pd.ExcelWriter.df_to_excel = _df_to_excel

    # ---------- 저장 ----------
    # df_result.to_excel(result_file, index=False, sheet_name='통합'
    # with pd.ExcelWriter(result_file, engine="openpyxl") as writer:
    with pd.ExcelWriter(result_file, engine="xlsxwriter") as writer:
        # workbook: Workbook = writer.book

        # df_to_excel(df, writer, "전체", "red")
        # df_to_excel(df_city_sale, writer, "판매_시티", "orange")
        # df_to_excel(df_artid_sale, writer, "판매_아티드", "yellow")
        # df_to_excel(df_city_sale_data, writer, "매출자료_시티", "green")
        # df_to_excel(df_artid_sale_data, writer, "매출자료_아티드", "blue")
        # df_to_excel(df_grouped, writer, "피봇1", "navy")
        # df_to_excel(df_grouped2, writer, "피봇2", "purple")

        writer.df_to_excel(df, "전체", "red")
        writer.df_to_excel(df_city_sale, "판매_시티", "orange")
        writer.df_to_excel(df_artid_sale, "판매_아티드", "yellow")
        writer.df_to_excel(df_city_sale_data, "매출자료_시티", "green")
        writer.df_to_excel(df_artid_sale_data, "매출자료_아티드", "blue")
        writer.df_to_excel(df_grouped, "피봇1", "navy", index=True)
        writer.df_to_excel(df_grouped2, "피봇2", "purple", index=True)


if __name__ == '__main__':
    TOP_DIR = Path('/Users/jwseo/Movies/9월정산')
    ROZLEY_DIR = TOP_DIR / "로즐리"
    CITY_DIR = TOP_DIR / "시티브리즈"
    ARTID_DIR = TOP_DIR / "아티드"

    order_file = TOP_DIR / 'W컨셉_정산종합내역_2408.xlsx'
    delivery_file = TOP_DIR / 'W컨셉_배송완료내역_2408.xlsx'
    # return_file = TOP_DIR / 'W컨셉_반품완료내역_2408.xlsx'
    return_file = TOP_DIR / 'W컨셉_반품내역_2408.xlsx'
    result_file = TOP_DIR / 'W컨셉_통합✅_2408.xlsx'
    run(order_file, delivery_file, return_file, result_file)
