from pathlib import Path

import pandas as pd

# SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame.
# Try using .loc[row_indexer,col_indexer] = value instead
# See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
pd.options.mode.chained_assignment = None


def run(order_file: Path, result_file: Path):
    CHANNEL_NAME = "29CM"

    order_df = pd.read_excel(order_file)
    df = order_df

    # ---------- 구분 ----------
    # 기존 '구분' 컬럼은 '구분_원본' 컬럼으로 이름 변경하고,
    # '구분_원본' 컬럼 값이 '배송비'이면 '구분' 컬럼 값을 '배송비'로, 그 외에는 '판매'로 지정
    df = df.rename(columns={'구분': '구분_원본'})
    df['구분'] = df['구분_원본'].map(lambda x: '배송비' if x == '배송비' else '판매')

    # ---------- 주문일자 ----------
    # '주문번호' 컬럼으로부터 '주문일자' 컬럼을 생성
    # 주문번호: ORD20240421-3020148
    # 주문일자: 2024-04-21
    # 주문번호의 형식 “ORDYYYYMMDD-XXXXXXX”에서 “YYYYMMDD”를 추출하여 “YYYY-MM-DD” 형식으로 변환하여 ‘주문일자’ 열에 추가합니다.
    def extract_order_date(order_number):
        # 주문번호를 변환하여 주문일자를 추출하는 함수
        return f"{order_number[3:7]}-{order_number[7:9]}-{order_number[9:11]}"

    # 주문번호를 변환하여 주문일자 컬럼 추가
    # df['주문일자'] = df['주문번호'].map(extract_order_date)
    df['주문일자'] = df['주문번호'].apply(lambda x: f"{x[3:7]}-{x[7:9]}-{x[9:11]}")
    # TypeError: x = 2024050617567955.0, type(x) = <class 'float'>

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
    df["옵션"] = df['옵션명']
    df["컬러"] = "-"
    df["사이즈"] = "-"
    df["수량"] = df["수량"]
    df["(판매처) 정상판매가"] = df['판매가']
    df["할인 (플랫폼)"] = df['쿠폰할인액(29CM부담)']
    df["할인 (브랜드)"] = df['쿠폰할인액(파트너부담)']
    df["고객결제"] = df['실판매가']
    df["매출(v+)"] = df['실판매가']
    df["수수료"] = df['수수료(판매수수료-29CM부담쿠폰할인액)']
    df["정산가"] = df["매출(v+)"] - df["수수료"]
    df["매출구분"] = "제품"

    df_sale = df.query('구분 == "판매"')
    df_delivery = df.query('구분 == "배송비"')
    df_result = df_sale[COLUMNS]

    GROUP_COLUMNS = [
        '실판매가',
        '수수료(판매수수료-29CM부담쿠폰할인액)',
        '판매지원금',
        '배송비',
        '정산금액',
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
    # ROZLEY_DIR = TOP_DIR / "로즐리"
    # CITY_DIR = TOP_DIR / "시티브리즈"
    # ARTID_DIR = TOP_DIR / "아티드"

    # 정산내역, 정산상세내역

    order_file = TOP_DIR / '시티_29CM_정산상세내역_2408.xlsx'
    result_file = TOP_DIR / '시티_29CM_통합✅_2408.xlsx'
    run(order_file, result_file)

    order_file = TOP_DIR / '시티_29CM_맨_정산상세내역_2408.xlsx'
    result_file = TOP_DIR / '시티_맨_29CM_통합✅_2408.xlsx'
    run(order_file, result_file)

    order_file = TOP_DIR / '아티드_29CM_정산상세내역_2408.xlsx'
    result_file = TOP_DIR / '아티드_29CM_통합✅_2408.xlsx'
    run(order_file, result_file)
