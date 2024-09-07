from pathlib import Path

import pandas as pd

# SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame.
# Try using .loc[row_indexer,col_indexer] = value instead
# See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
pd.options.mode.chained_assignment = None


def run(
    order_file: Path,
    settlement_file: Path,
    quick_deliv_file: Path | None,
    result_file: Path,
):
    """SSG

    Args:
        order_file (Path): 구매확정관리
        settlement_file (Path): 정산기준내역
        quick_deliv_file (Path | None): 직진배송
        result_file (Path): 최종파일
    """

    CHANNEL_NAME = "지그재그"

    # 데이터를 pandas DataFrame으로 읽기
    # '상품주문번호' 컬럼을 문자열로 지정하여 엑셀 파일 읽기
    # 안 그러면 settlement_df['상품주문번호']는 float64로 불러와지면서 값이 달라져버림 (빈 값(NaN)이 있어서 float64로 되는 것으로 추정됨)
    # 달라지면 merge가 제대로 되지 않음.
    order_df = pd.read_excel(order_file, dtype={'주문번호': str, '상품주문번호': str})
    settlement_df = pd.read_excel(
        settlement_file, dtype={'주문번호': str, '상품주문번호': str}
    )
    if quick_deliv_file:
        quick_deliv_df = pd.read_excel(
            quick_deliv_file, dtype={'주문번호': str, '상품주문번호': str}
        )

    # df = order_df.copy(deep=True)
    df = settlement_df

    # ---------- 주문일자 ----------
    # 주문발생일 -> 주문일자
    # 20240507 -> 2024-05-07
    df['주문일자'] = df['주문발생일'].map(
        lambda x: pd.to_datetime(str(x), format='%Y%m%d').strftime('%Y-%m-%d')
    )

    # ---------- 구분 ----------
    # '상품명' == '일반 배송비' -> '구분' == '배송비'
    # 그 외에는 '판매'
    # start = time.perf_counter()
    # '일반 배송비'
    # '반품 배송비'
    df['구분'] = df['상품명'].map(lambda x: '배송비' if '배송비' in x else '판매')
    # 0.0011037910589948297

    # df.loc[df['상품명'] == '일반 배송비', "구분"] = "배송비"
    # df.loc[df['상품명'] != '일반 배송비', "구분"] = "판매"
    # 0.003560542012564838
    # end = time.perf_counter()
    # print(end - start)

    # ---------- 서비스 타입 ----------
    # '직진배송'만 채워져 있음. 비어있는 값은 '일반'으로 채우기
    df['서비스 타입'] = df['서비스 타입'].fillna('일반')

    # ---------- 옵션정보 ----------

    # df_sale = df.loc[df['구분'] == '판매']
    df_sale = df.query("구분 == '판매'")
    df_delivery = df.query("구분 == '배송비'")

    # 서비스 타입에 따라 다른 파일에서 옵션정보를 가져오기
    order_df = order_df[["상품주문번호", "옵션정보", "수량"]]
    order_df.columns = ["상품주문번호", "옵션(일반)", "수량(일반)"]
    df_sale = pd.merge(df_sale, order_df, on="상품주문번호", how="left")

    if quick_deliv_file:
        quick_deliv_df = quick_deliv_df[["상품주문번호", "옵션정보", "수량"]]
        quick_deliv_df.columns = ["상품주문번호", "옵션(직진배송)", "수량(직진배송)"]
        df_sale = pd.merge(df_sale, quick_deliv_df, on="상품주문번호", how="left")

    GROUP_COLUMNS = ['상품주문액', '수수료 금액', '정산금액']
    df_grouped = df.groupby(by=["서비스 타입", "구분"], group_keys=True)[
        GROUP_COLUMNS
    ].sum()
    total_sum = df_grouped.sum()
    total_sum.name = ('총합계', '-')
    df_grouped = pd.concat([df_grouped, total_sum.to_frame().T])

    df_grouped2 = df.groupby(by=["구분", "서비스 타입"], group_keys=True)[
        GROUP_COLUMNS
    ].sum()
    total_sum2 = df_grouped2.sum()
    total_sum2.name = ('총합계', '-')
    df_grouped2 = pd.concat([df_grouped2, total_sum2.to_frame().T])

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
    df['차수'] = df['서비스 타입']
    df["주문일(결제일)"] = df["주문일자"]
    df["주문번호"] = df["주문번호"]
    df["상품주문번호"] = df["상품주문번호"]
    df["상품명"] = df["상품코드"]
    # df["옵션"] = df["옵션1"]
    # df["컬러"] = "-"
    # df["사이즈"] = "-"
    # df["수량"] = df["수량"]
    # df["(판매처) 정상판매가"] = df['할인액'] + df['상품주문액']
    # df["할인 (플랫폼)"] = (
    #     df['주문할인_자사부담']
    #     + df['즉시할인_자사부담']
    #     + df['쿠폰할인_자사부담']
    #     + df['적립금']
    # )
    # df["할인 (브랜드)"] = df['할인액']
    # df["고객결제"] = df['결제금액(배송비제외)']
    # df["매출(v+)"] = df['매출금액']
    # # df["수수료"] = df['수수료']
    # df["정산가"] = df['업체입금예정액']
    # df["매출구분"] = "제품"
    # df_result = df[COLUMNS]

    # ---------- 저장 ----------
    # with pd.ExcelWriter(result_file, engine="openpyxl") as writer:
    with pd.ExcelWriter(result_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="전체")
        df_sale.to_excel(writer, index=False, sheet_name="판매")
        df_delivery.to_excel(writer, index=False, sheet_name="배송비")
        df_grouped.to_excel(writer, index=True, sheet_name="그룹핑1")
        df_grouped2.to_excel(writer, index=True, sheet_name="그룹핑2")


if __name__ == '__main__':
    TOP_DIR = Path('/Users/jwseo/Movies/9월정산')
    ROZLEY_DIR = TOP_DIR / "로즐리"
    CITY_DIR = TOP_DIR / "시티브리즈"
    ARTID_DIR = TOP_DIR / "아티드"

    order_file = ROZLEY_DIR / '로즐리_지그재그_구매확정관리_4-5-6월.xlsx'
    # order_file = ROZLEY_DIR / '로즐리_지그재그_구매확정관리_56월.xlsx'
    # order_file = (
    #     ROZLEY_DIR / '로즐리_지그재그_구매확정관리_240506.xlsx'
    # )
    settlement_file = ROZLEY_DIR / '로즐리_지그재그_정산기준내역_6월.xlsx'
    quick_deliv_file = ROZLEY_DIR / '로즐리_지그재그_직진배송_6월.xlsx'
    result_file = ROZLEY_DIR / '로즐리_지그재그_통합✅_2408.xlsx'
    run(order_file, settlement_file, quick_deliv_file, result_file)

    # order_file = CITY_DIR / '시티_지그재그_구매확정관리_6월.xlsx'
    # settlement_file = CITY_DIR / '시티_지그재그_정산기준내역_6월.xlsx'
    # quick_deliv_file = None
    # result_file = CITY_DIR / '시티_지그재그_통합✅_2408.xlsx'
    # run(order_file, settlement_file, quick_deliv_file, result_file)
