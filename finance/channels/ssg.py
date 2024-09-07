from pathlib import Path

import pandas as pd

# SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame.
# Try using .loc[row_indexer,col_indexer] = value instead
# See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
pd.options.mode.chained_assignment = None


def run(order_file: Path, result_file: Path):
    """SSG

    Args:
        order_file (Path): 위수탁 마감 업체별 상세내역
        result_file (Path): 최종파일
    """

    CHANNEL_NAME = "SSG"

    order_df = pd.read_excel(order_file)

    df = order_df

    # ===============================================================
    # 배송비 =V5-X5
    # V = 판매단가
    # X = 상품대총판매가
    # 계산할 필요 없이 "고객부담 배송비"라는 컬럼이 있음. 이거 사용하면 될 듯

    # ---------- 주문일자 ----------
    # '원주문ID' 컬럼으로부터 '주문일자' 컬럼을 생성
    # 원주문ID: 20240105777255
    # 주문일자: 2024-04-21
    # df['주문일자'] = df['원주문ID'].map(extract_order_date)
    df['주문일자'] = df['원주문ID'].apply(lambda x: f"{x[0:4]}-{x[4:6]}-{x[6:8]}")

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
    df['차수'] = '-'
    df["주문일(결제일)"] = df["주문일자"]
    df["주문번호"] = df["원주문ID"]
    df["상품주문번호"] = '-'
    # df["상품명"] = df["상품명"]
    df["옵션"] = df['단품']
    df["컬러"] = "-"
    df["사이즈"] = "-"
    # df["수량"] = df["수량"]
    df["(판매처) 정상판매가"] = df['상품대총판매가']
    df["할인 (플랫폼)"] = df['SSG부담 (할인부담금)']
    df["할인 (브랜드)"] = 0
    df["고객결제"] = (
        df['(판매처) 정상판매가'] - df["할인 (플랫폼)"] - df["할인 (브랜드)"]
    )
    df["매출(v+)"] = df['고객결제']
    df["수수료"] = df['판매수수료']
    df["정산가"] = df["매출(v+)"] - df["수수료"]
    df["매출구분"] = "제품"

    # ---------- 구분 ----------
    # '수량' 열 값이 0인 행 또는 '정산금액(VAT포함)' 값 = '판매단가' 값 = '고객부담 배송비' 값인 행은 '배송비'로 구분
    def classify(row):
        if row['수량'] == 0:
            return '배송비'
        elif row['정산금액(VAT포함)'] == row['판매단가'] == row['고객부담 배송비']:
            return '배송비'
        else:
            return '판매'

    df['구분'] = df.apply(classify, axis=1)

    # 수량이 0이 아닌데 고객부담 배송비가 있는 행은 배송 시트와 판매 시트에 둘 다 넣는데, 이때 판매 시트에 고객부담 배송비를 0원으로 변경.
    df_sale = df.query('수량 > 0').copy()
    df_sale.loc[df_sale['고객부담 배송비'] > 0, '고객부담 배송비'] = 0
    # df_delivery = df.query("'고객부담 배송비' > 0").copy()
    df_delivery = df[df['고객부담 배송비'] > 0].copy()

    df_result = df_sale[COLUMNS]

    GROUP_COLUMNS = [
        '상품대총판매가',
        '판매수수료',
        '협력사 배송비(VAT포함)',
        '배송비 (VAT포함)',
        '고객부담 배송비',
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

    order_file = (
        TOP_DIR
        / '아티드_SSG_위수탁 마감 업체별 상세내역_2408.xlsx'
    )
    # order_file = TOP_DIR / '아티드_SSG_정산내역_2408.xlsx'
    result_file = TOP_DIR / '아티드_SSG_통합✅_2408.xlsx'
    run(order_file, result_file)
