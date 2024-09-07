from pathlib import Path

import pandas as pd

try:
    from finance.xl.password import unlock_if_locked
except ModuleNotFoundError:
    import sys

    sys.path.append(str(Path(__file__).parent.parent.parent))
    from finance.xl.password import unlock_if_locked

# SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame.
# Try using .loc[row_indexer,col_indexer] = value instead
# See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
pd.options.mode.chained_assignment = None


def run(order_file: Path, confirm_file: Path, settlement_file: Path, result_file: Path):
    """네이버페이

    Args:
        order_file (Path): 부가세신고내역
        confirm_file (Path): 구매확정내역
        settlement_file (Path): 건별정산내역
        result_file (Path): 최종파일
    """

    CHANNEL_NAME = "네이버페이"

    # 암호 잠금 해제
    unlock_if_locked(order_file)
    unlock_if_locked(confirm_file)
    unlock_if_locked(settlement_file)

    # 데이터를 pandas DataFrame으로 읽기
    order_df = pd.read_excel(order_file)
    confirm_df = pd.read_excel(confirm_file)
    settlement_df = pd.read_excel(settlement_file)

    order_df["주문번호"] = order_df["주문번호"].astype(str)
    order_df["상품주문번호"] = order_df["상품주문번호"].astype(str)
    confirm_df["상품주문번호"] = confirm_df["상품주문번호"].astype(str)
    settlement_df["상품주문번호"] = settlement_df["상품주문번호"].astype(str)

    df = order_df.copy(deep=True)
    df = df.drop(
        columns=["구분"]
    )  # 건별정산 시트의 '구분' 컬럼을 써야 해서, 부가세 시트의 '구분' 컬럼 삭제

    # # ---------- 건별정산내역 ----------
    # df['매출']
    # # =VLOOKUP($D3,'/Users/jwseo/Movies/다솜님_5월_TEST/After/2. 채널별정산서/아티드/자사/[아티드_네이버페이_건별정산내역_2405.xlsx]SettleCaseByCase'!$C$1:$Q$132,11,0)
    # # D열 = '상품주문번호'
    # # 11 = '정산기준금액' 열

    # df['수수료']
    # # =VLOOKUP($D3,'/Users/jwseo/Movies/다솜님_5월_TEST/After/2. 채널별정산서/아티드/자사/[아티드_네이버페이_건별정산내역_2405.xlsx]SettleCaseByCase'!$C$1:$Q$132,12,0)
    # # 12 = '네이버페이 주문관리 수수료' 열

    # df['정산']
    # # =VLOOKUP($D3,'/Users/jwseo/Movies/다솜님_5월_TEST/After/2. 채널별정산서/아티드/자사/[아티드_네이버페이_건별정산내역_2405.xlsx]SettleCaseByCase'!$C$1:$Q$132,15,0)
    # # 15 = '정산예정금액' 열

    # # ---------- 구매확정내역 ----------
    # df['정상판매가']
    # # =VLOOKUP($D3,'/Users/jwseo/Movies/다솜님_5월_TEST/After/2. 채널별정산서/아티드/자사/[아티드_네이버페이_구매확정내역_2405.xlsx]구매확정내역'!$A$1:$AP$119,22,0)
    # # 22 = '상품별 총 주문금액' 열

    # df['옵션']
    # # =VLOOKUP($D3,'/Users/jwseo/Movies/다솜님_5월_TEST/After/2. 채널별정산서/아티드/자사/[아티드_네이버페이_구매확정내역_2405.xlsx]구매확정내역'!$A$1:$AP$119,16,0)
    # # 16 = '옵션정보' 열

    # df['수량']
    # # =VLOOKUP($D3,'/Users/jwseo/Movies/다솜님_5월_TEST/After/2. 채널별정산서/아티드/자사/[아티드_네이버페이_구매확정내역_2405.xlsx]구매확정내역'!$A$1:$AP$119,17,0)
    # # 17 = '수량' 열

    # ---------- 건별정산내역 ----------
    # '매출', '수수료', '정산' 데이터를 order_df와 병합
    settlement_df = settlement_df[
        [
            "상품주문번호",
            "구분",
            "정산기준금액",
            "네이버페이 주문관리 수수료",
            "정산예정금액",
        ]
    ]
    settlement_df.columns = ["상품주문번호", "구분", "매출", "수수료", "정산"]
    df = pd.merge(df, settlement_df, on="상품주문번호", how="left")

    # ---------- 구매확정내역 ----------
    # '정상판매가', '옵션', '수량' 데이터를 order_df와 병합
    confirm_df = confirm_df[["상품주문번호", "상품별 총 주문금액", "옵션정보", "수량"]]
    confirm_df.columns = ["상품주문번호", "정상판매가", "옵션", "수량"]
    df = pd.merge(df, confirm_df, on="상품주문번호", how="left")

    # ---------- 주문일자 ----------

    # 주문번호에서 주문일자 추출
    # 주문번호 "2024042017461030" -> 주문일자 "2024-04-20"
    df["주문일자"] = df["주문번호"].apply(lambda x: f"{x[:4]}-{x[4:6]}-{x[6:8]}")

    # ---------- 판매만 ----------
    DIVISION_COLUMN = "구분"
    print(df[DIVISION_COLUMN].unique())

    # '구분' 컬럼의 '상품주문' 값을 '판매'로 변경
    # df.loc[df[DIVISION_COLUMN] == '상품주문', DIVISION_COLUMN] = '판매'
    df[DIVISION_COLUMN] = df[DIVISION_COLUMN].replace({"상품주문": "판매"})

    df_sales = df[df[DIVISION_COLUMN] == "판매"]
    df_deliv = df[df[DIVISION_COLUMN] == "배송비"]

    GROUP_COLUMNS = ["총 매출금액", "과세매출", "매출", "수수료", "정산"]
    # df_deliv2 = df_deliv.groupby(by=['구분'], group_keys=True)[COLUMNS].sum()
    df2 = df.groupby(by=["구분"], group_keys=True)[GROUP_COLUMNS].sum()

    # 총합계 계산하여 새로운 행 추가
    total_sum = df[GROUP_COLUMNS].sum()
    df2.loc["총합계"] = total_sum

    # 매출자료 양식에 맞게 변경
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
    df_data = df_sales.copy(deep=True)
    df_data["채널명"] = "자사"
    df_data["차수"] = CHANNEL_NAME
    df_data["주문일(결제일)"] = df_sales["주문일자"]
    df_data["주문번호"] = df_sales["주문번호"]
    df_data["상품주문번호"] = df_sales["상품주문번호"]
    df_data["상품명"] = df_sales["상품명"]
    df_data["옵션"] = df_sales["옵션"]
    df_data["컬러"] = "-"
    df_data["사이즈"] = "-"
    df_data["수량"] = df_sales["수량"]
    df_data["(판매처) 정상판매가"] = df_sales["정상판매가"]
    df_data["할인 (플랫폼)"] = 0
    df_data["할인 (브랜드)"] = df_sales["정상판매가"] - df_sales["총 매출금액"]
    df_data["고객결제"] = df_sales["매출"]
    df_data["매출(v+)"] = df_sales["매출"]
    df_data["수수료"] = df_sales["수수료"] * (-1)
    df_data["정산가"] = df_sales["정산"]
    df_data["매출구분"] = "제품"
    df_data = df_data[COLUMNS]

    # df.to_excel(result_file, index=False, sheet_name='통합')
    with pd.ExcelWriter(result_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="전체")
        df_sales.to_excel(writer, index=False, sheet_name="판매")
        df_deliv.to_excel(writer, index=False, sheet_name="배송")
        # df_deliv2.to_excel(writer, sheet_name='배송피벗')
        df2.to_excel(writer, sheet_name="전체피벗")
        df_data.to_excel(writer, sheet_name="매출자료")


if __name__ == "__main__":
    TOP_DIR = Path("/Users/jwseo/Movies/9월정산")
    ROZLEY_DIR = TOP_DIR / "로즐리"
    CITY_DIR = TOP_DIR / "시티브리즈"
    ARTID_DIR = TOP_DIR / "아티드"

    print("로즐리")
    order_file = TOP_DIR / "로즐리_네이버페이_부가세신고내역_2408.xlsx"
    confirm_file = TOP_DIR / "로즐리_네이버페이_구매확정내역_2408.xlsx"
    settlement_file = TOP_DIR / "로즐리_네이버페이_건별정산내역_2408.xlsx"
    result_file = TOP_DIR / "로즐리_네이버페이_통합✅_2408.xlsx"
    run(order_file, confirm_file, settlement_file, result_file)

    print("시티브리즈")
    order_file = TOP_DIR / "시티_네이버페이_부가세신고내역_2408.xlsx"
    confirm_file = TOP_DIR / "시티_네이버페이_구매확정내역_2408.xlsx"
    settlement_file = TOP_DIR / "시티_네이버페이_건별정산내역_2408.xlsx"
    result_file = TOP_DIR / "시티_네이버페이_통합✅_2408.xlsx"
    run(order_file, confirm_file, settlement_file, result_file)

    print("아티드")
    order_file = TOP_DIR / "아티드_네이버페이_부가세신고내역_2408.xlsx"
    confirm_file = TOP_DIR / "아티드_네이버페이_구매확정내역_2408.xlsx"
    settlement_file = TOP_DIR / "아티드_네이버페이_건별정산내역_2408.xlsx"
    result_file = TOP_DIR / "아티드_네이버페이_통합✅_2408.xlsx"
    run(order_file, confirm_file, settlement_file, result_file)
