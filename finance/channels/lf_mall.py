from pathlib import Path

import pandas as pd

# SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame.
# Try using .loc[row_indexer,col_indexer] = value instead
# See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
pd.options.mode.chained_assignment = None


def run(order_file: Path, result_file: Path):
    """SSG

    Args:
        order_file (Path): 정산내역
        result_file (Path): 최종파일
    """

    CHANNEL_NAME = "LF몰"

    order_df = pd.read_excel(order_file, skiprows=3)

    df = order_df

    # ---------- 컬럼 정리 ----------

    columns = df.columns
    # ['매출일', '주문번호', '주문상세번호', '주문자', '수령자', '상품코드', '사이즈', '상품명', '패밀리여부',
    #     '결제형태', '정산구분', '판매수량', '정산금액', '판매금액', '수수료율', '판매수수료', '업체부담금액',
    #     'Unnamed: 17', 'LF부담금액', 'Unnamed: 19', 'Unnamed: 20', '배송비 금액',
    #     'Unnamed: 22', 'Unnamed: 23', '추가배송비', 'LF부담배송비', '동봉금액', '정산조정금액',
    #     '배송완료일', '정산월', '비고']

    first_row_values = df.iloc[0].to_dict()
    # {'매출일': nan, .., '판매수수료': nan, '업체부담금액': '쿠폰', 'Unnamed: 17': 'EGM', 'LF부담금액': '쿠폰', ...}

    first_row = [{"column": k, "value": v} for k, v in first_row_values.items()]
    # first_row = [
    #     {
    #         'column': '매출일',
    #         'value': nan,
    #     },
    #     ...
    #     {
    #         'column': '업체부담금액',
    #         'value': '쿠폰',
    #     },
    #     {
    #         'column': 'Unnamed: 17',
    #         'value': 'EGM',
    #     },
    #     ...
    # ]

    def transform_columns(first_row):
        temp = []
        last_valid_column = None
        for item in first_row:
            column = item['column']
            value = item['value']

            if "Unnamed" in column:
                if last_valid_column is not None:
                    temp.append({'column': last_valid_column, 'value': value})
            else:
                last_valid_column = column
                temp.append({'column': column, 'value': value})

        new_columns = []
        for item in temp:
            column = item['column']
            value = item['value']

            if pd.isna(value):
                new_columns.append(column)
            else:
                new_columns.append(f"{column}({value})")

        return new_columns

    new_columns = transform_columns(first_row)

    def remove_first_row(df):
        # Drop the first row by its index
        df = df.drop(df.index[0])
        # Reset the index to keep the DataFrame tidy
        df = df.reset_index(drop=True)
        return df

    df = remove_first_row(df)

    df = df.rename(columns=dict(zip(columns, new_columns)))

    # ---------- 컬럼 정리 ----------
    df['업체부담금액'] = df['업체부담금액(쿠폰)'] + df['업체부담금액(EGM)']
    df['LF부담금액'] = (
        df['LF부담금액(쿠폰)'] + df['LF부담금액(EGM)'] + df['LF부담금액(마일리지)']
    )
    df['배송비 금액'] = (
        df['배송비 금액(기본)'] + df['배송비 금액(반품)'] + df['배송비 금액(교환)']
    )

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
    df["주문일(결제일)"] = df["매출일"]
    # df["주문번호"] = df["주문번호"]
    df["상품주문번호"] = df['주문상세번호']
    # df["상품명"] = df["상품명"]
    df["옵션"] = df['사이즈']
    df["컬러"] = "-"
    df["사이즈"] = "-"
    df["수량"] = df["판매수량"]
    df["(판매처) 정상판매가"] = df['판매금액']
    df["할인 (플랫폼)"] = (
        df['LF부담금액(쿠폰)'] + df['LF부담금액(EGM)'] + df['LF부담금액(마일리지)']
    )
    df["할인 (브랜드)"] = df['업체부담금액(쿠폰)']
    df["고객결제"] = (
        df['(판매처) 정상판매가'] - df["할인 (플랫폼)"] - df["할인 (브랜드)"]
    )
    df["매출(v+)"] = df['고객결제']
    df["수수료"] = -1 * df['판매수수료'] - df['할인 (플랫폼)']
    df["정산가"] = df['정산금액']
    df["매출구분"] = "제품"

    # ---------- 구분 ----------
    # '정산구분' 컬럼 값이 '배송비' -> '배송비'
    # 그 외에는 '판매'
    # FIXME: 아래로 하면 그냥 전부 판매로 나옴. 배송비가 판매에 같이 들어가있는 경우.
    # 필요하면 SSG처럼 분리해야.
    df['구분'] = df['정산구분'].map(lambda x: "배송비" if x == "배송비" else "판매")

    df_sale = df.query("구분 == '판매'")
    df_delivery = df.query("구분 == '배송비'")

    df_result = df_sale[COLUMNS]

    GROUP_COLUMNS = [
        '정산금액',
        '판매금액',
        '배송비 금액',
        '추가배송비',
        '업체부담금액',
        'LF부담금액',
        '판매수수료',
        '동봉금액',
    ]
    df_grouped = df.groupby(by=["구분"], group_keys=True)[GROUP_COLUMNS].sum()
    df_grouped.loc["총합계"] = df[GROUP_COLUMNS].sum()

    # ---------- 저장 ----------
    with pd.ExcelWriter(result_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="전체")
        df_sale.to_excel(writer, index=False, sheet_name="판매")
        df_delivery.to_excel(writer, index=False, sheet_name="배송")
        df_result.to_excel(writer, index=False, sheet_name="매출자료")
        df_grouped.to_excel(writer, index=True, sheet_name="피봇")


if __name__ == '__main__':
    TOP_DIR = Path('/Users/jwseo/Movies/9월정산')
    ROZLEY_DIR = TOP_DIR / "로즐리"
    CITY_DIR = TOP_DIR / "시티브리즈"
    ARTID_DIR = TOP_DIR / "아티드"

    order_file = TOP_DIR / '시티_LF몰_정산내역_2408.xlsx'
    result_file = TOP_DIR / '시티_LF몰_통합✅_2408.xlsx'
    run(order_file, result_file)

    order_file = TOP_DIR / '아티드_LF몰_정산내역_2408.xlsx'
    result_file = TOP_DIR / '아티드_LF몰_통합✅_2408.xlsx'
    run(order_file, result_file)
