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

    CHANNEL_NAME = "카카오선물하기"

    # 데이터를 pandas DataFrame으로 읽기
    order_df = pd.read_excel(order_file)

    # df = order_df.copy(deep=True)
    df = order_df

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
    df["주문일(결제일)"] = df['주문 결제일']
    df["주문번호"] = df['주문번호']
    df["상품주문번호"] = '-'
    # df["상품명"] = df['상품명']
    df["옵션"] = df['옵션명']
    df["컬러"] = "-"
    df["사이즈"] = "-"
    # df["수량"] = df['수량']
    df["(판매처) 정상판매가"] = df['상품주문금액']
    df["할인 (플랫폼)"] = 0
    df["할인 (브랜드)"] = df['판매자할인금액합계']
    df["고객결제"] = (
        df['(판매처) 정상판매가'] - df['할인 (플랫폼)'] - df['할인 (브랜드)']
    )
    df["매출(v+)"] = df['정산기준금액']
    df["수수료"] = df['수수료합계']
    df["정산가"] = df['판매정산금액']
    df["매출구분"] = "제품"

    # df_sale = df[df['정산구분'] == "일반상품"]
    df_sale = df.query('정산구분 == "일반상품"')
    # df_sale = df.loc[df['정산구분'] == '일반상품']
    # df_sale = df[df['정산구분'].isin(['일반상품'])]

    df_result = df_sale[COLUMNS]

    GROUP_COLUMNS = ["결제금액", '수수료합계', '카카오할인금액합계', '판매정산금액']
    df_grouped = df.groupby(by=["정산구분"], group_keys=True)[GROUP_COLUMNS].sum()
    df_grouped.loc["총합계"] = df[GROUP_COLUMNS].sum()

    # ---------- 저장 ----------
    with pd.ExcelWriter(result_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="전체")
        df_sale.to_excel(writer, index=False, sheet_name="판매")
        df_result.to_excel(writer, index=False, sheet_name="매출자료")
        df_grouped.to_excel(writer, sheet_name="피봇")


if __name__ == '__main__':
    TOP_DIR = Path('/Users/jwseo/Movies/9월정산')
    ROZLEY_DIR = TOP_DIR / "로즐리"
    CITY_DIR = TOP_DIR / "시티브리즈"
    ARTID_DIR = TOP_DIR / "아티드"

    order_file = TOP_DIR / '시티_카카오선물하기_정산내역_2408.xlsx'
    result_file = TOP_DIR / '시티_카카오선물하기_통합✅_2408.xlsx'
    run(order_file, result_file)
