from pathlib import Path

import pandas as pd

# SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame.
# Try using .loc[row_indexer,col_indexer] = value instead
# See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
pd.options.mode.chained_assignment = None


def run(order_file1: Path, order_file2: Path, confirm_file: Path, result_file: Path):
    """에이블리

    Args:
        order_file1 (Path): 정산세부내역1차
        order_file2 (Path): 정산세부내역2차
        confirm_file (Path): 구매확정 내역
        result_file (Path): 최종파일
    """

    CHANNEL_NAME = "에이블리"

    # 데이터를 pandas DataFrame으로 읽기
    order_df1 = pd.read_excel(order_file1)
    order_df2 = pd.read_excel(order_file2)
    confirm_df = pd.read_excel(confirm_file)

    order_df1["차수"] = '1차'
    order_df2["차수"] = '2차'

    # ---------- 데이터 합치기 ----------
    # 1차와 2차 데이터를 합치기
    df = pd.concat([order_df1, order_df2], ignore_index=True)

    # ---------- 주문일자 ----------
    # '결제 완료일' 컬럼으로부터 '주문일자' 컬럼을 생성
    # 결제 완료일 : 2024-04-28 21:45:10
    # 주문일자 : 2024-04-28
    df['주문일자'] = df['결제 완료일'].apply(lambda x: x.strftime('%Y-%m-%d'))

    # ---------- 구분 ----------
    # '구분' 컬럼을 생성
    # '결제 금액' == 0 AND '배송비' > 0 --> 구분 = '배송비'
    # 그 외엔 '판매'
    df['구분'] = df.apply(
        lambda row: '배송비' if row['결제 금액'] == 0 and row['배송비'] > 0 else '판매',
        axis=1,
    )

    # ---------- 정산금에서 배송비 제외 ----------
    # 결제 금액도 있고 배송비도 있는 경우, 정산금에서 배송비를 뺌
    # 굳이?

    # ---------- 병합 ----------
    df['상품주문번호'] = df['상품 주문 번호']
    confirm_df = confirm_df[
        ["상품주문번호", '주문번호', '상품명', '판매가', '옵션 정보', '수량']
    ]
    confirm_df = confirm_df.drop_duplicates()
    confirm_df.columns = [
        "상품주문번호",
        '주문번호',
        '상품명',
        '정상판매가',
        '옵션',
        '수량',
    ]  # 컬럼명 변경
    df = pd.merge(df, confirm_df, on="상품주문번호", how="left")

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
    df["주문일(결제일)"] = df['주문일자']
    # df["주문번호"] = df['주문번호']
    # df["상품주문번호"] = df["상품주문번호"]
    # df["상품명"] = df['상품명']
    # df["옵션"] = df['옵션']
    df["컬러"] = "-"
    df["사이즈"] = "-"
    # df["수량"] = df['수량']
    df["(판매처) 정상판매가"] = df['정상판매가']
    df["할인 (플랫폼)"] = df['프로모션 지원금']
    df["할인 (브랜드)"] = df['정상판매가'] - df['결제 금액'] - df['프로모션 지원금']
    df["고객결제"] = df['결제 금액']
    df["매출(v+)"] = df['결제 금액']
    df["수수료"] = df['결제 수수료'] + df['플랫폼 수수료']
    df["정산가"] = df["매출(v+)"] - df["수수료"]
    df["매출구분"] = "제품"

    # df_sale = df[df['구분'] == "판매"]
    df_sale = df.query('구분 == "판매"')
    # df_sale = df.loc[df['구분'] == '판매']
    # df_sale = df[df['구분'].isin(['판매'])]

    df_result = df_sale[COLUMNS]

    GROUP_COLUMNS = [
        '결제 금액',
        '배송비',
        '결제 수수료',
        '프로모션 지원금',
        '플랫폼 수수료',
        '정산금',
    ]
    df_grouped = df.groupby(by=["차수"], group_keys=True)[GROUP_COLUMNS].sum()
    df_grouped.loc["총합계"] = df[GROUP_COLUMNS].sum()

    # ---------- 저장 ----------
    with pd.ExcelWriter(result_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="전체")
        df_sale.to_excel(writer, index=False, sheet_name="판매")
        df_result.to_excel(writer, index=False, sheet_name="매출자료")
        df_grouped.to_excel(writer, sheet_name="피봇")


if __name__ == '__main__':
    TOP_DIR = Path('/Users/jwseo/Movies/9월정산')

    # TODO: 10일에 할 수 있음

    order_file1 = (
        TOP_DIR / '로즐리_에이블리_정산내역_1차_2408.xlsx'
    )  # 정산세부내역1차
    order_file2 = (
        TOP_DIR / '로즐리_에이블리_정산내역_2차_.xlsx'
    )  # 정산세부내역2차
    confirm_file = TOP_DIR / '로즐리_에이블리_구매확정 내역_2408.xlsx'  # 구매확정 내역
    result_file = TOP_DIR / '로즐리_에이블리_통합❤️_2408.xlsx'  # 최종파일
    run(order_file1, order_file2, confirm_file, result_file)

    order_file1 = (
        TOP_DIR / '시티_에이블리_정산내역_1차_2408.xlsx'
    )  # 정산세부내역1차
    order_file2 = (
        TOP_DIR / '시티_에이블리_정산내역_2차.xlsx'
    )  # 정산세부내역2차
    confirm_file = (
        TOP_DIR / '시티_에이블리_구매확정내역_2408.xlsx'
    )  # 구매확정 내역
    result_file = TOP_DIR / '시티_에이블리_통합❤️_2408.xlsx'  # 최종파일
    run(order_file1, order_file2, confirm_file, result_file)
