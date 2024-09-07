from pathlib import Path

import pandas as pd

# SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame.
# Try using .loc[row_indexer,col_indexer] = value instead
# See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
pd.options.mode.chained_assignment = None


def run(settle_file1: Path, settle_file2: Path, order_file: Path, result_file: Path):
    """SSG

    Args:
        settle_file1 (Path): 정산관리1차
        settle_file2 (Path): 정산관리2차
        order_file (Path): 정산내역
        result_file (Path): 최종파일
    """

    CHANNEL_NAME = "서울스토어"

    # 데이터를 pandas DataFrame으로 읽기
    settle_df1 = pd.read_excel(settle_file1)
    settle_df2 = pd.read_excel(settle_file2)
    order_df = pd.read_excel(order_file)

    # ---------- 차수 ----------
    # 데이터를 합치기 전에 정산관리1차와 2차 데이터에 '차수' 컬럼을 추가
    # 정산관리1차 -> 차수 '1차'
    # 정산관리2차 -> 차수 '2차'
    settle_df1["차수"] = '1차'
    settle_df2["차수"] = '2차'

    # ---------- 데이터 합치기 ----------
    # 정산관리1차와 2차 데이터를 합치기
    df = pd.concat([settle_df1, settle_df2])

    # ---------- 주문일자 ----------
    # '주문일시' 컬럼으로부터 '주문일자' 컬럼을 생성
    # 주문일시 : 2024-04-17 14:28:57
    # 주문일자 : 2024-04-17
    df['주문일자'] = df['주문일시'].apply(lambda x: x[:10])

    # ---------- 구분 ----------
    # '구분' 컬럼을 생성
    # FIXME: '구분' 컬럼을 새로 추가할 때, 어떤 걸 기준으로 판단하는지?
    df['구분'] = '판매'

    # ---------- 옵션 ----------
    # 옵션은 구매확정파일에서 가져옴
    # 주문상세번호 기준으로
    order_df = order_df[["주문상세번호", "옵션정보"]]
    order_df = order_df.drop_duplicates()
    order_df.columns = ["주문상세번호", "옵션"]  # 컬럼명 변경
    df = pd.merge(df, order_df, on="주문상세번호", how="left")

    # ---------- 정수형 변환 ----------
    # 문자열 금액 데이터를 정수형으로 변환하는 함수 정의
    def convert_to_int(df, columns):
        for col in columns:
            if df[col].dtype == 'object':  # 문자열 타입인 경우에만 처리
                df[col] = df[col].str.replace(',', '')  # 콤마 제거
                df[col] = df[col].astype(int)  # 정수형으로 변환
        return df

    # 변환할 컬럼 리스트
    columns_to_convert = [
        '판매금액',
        '옵션금액',
        '수량',
        '쿠폰금액',
        '서울스토어지원금',
        '쿠폰할인초과금',
        '셀러부담쿠폰',
        '결산금액',
        '판매수수료',
        '판매수수료(vat)',
        '판매수수료 합계',
        '셀러정산금',
        '반품배송비',
        '반품배송비수수료(공급가액)',
        '반품배송비수수료(vat)',
        '반품배송비수수료합계',
    ]

    # 함수 적용하여 변환
    df = convert_to_int(df, columns_to_convert)

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
    # df["주문번호"] = df["주문번호"]
    df["상품주문번호"] = df["주문상세번호"]
    # df["상품명"] = df["상품명"]
    # df["옵션"] = df["옵션"]
    df["컬러"] = "-"
    df["사이즈"] = "-"
    # df["수량"] = df["수량"]
    df["(판매처) 정상판매가"] = df['결산금액']
    df["할인 (플랫폼)"] = -1 * df['서울스토어지원금']
    df["할인 (브랜드)"] = -1 * df['셀러부담쿠폰']
    df["고객결제"] = (
        df['(판매처) 정상판매가'] - df['할인 (플랫폼)'] - df['할인 (브랜드)']
    )
    df["매출(v+)"] = df["고객결제"]
    df["수수료"] = df['판매수수료 합계'] + df['서울스토어지원금']
    df["정산가"] = df['매출(v+)'] - df['수수료']
    df["매출구분"] = "제품"
    df_result = df[COLUMNS]

    GROUP_COLUMNS = [
        '판매금액',
        '판매수수료 합계',
        '셀러정산금',
        '반품배송비',
        '반품배송비수수료(공급가액)',
        '반품배송비수수료(vat)',
        '반품배송비수수료합계',
    ]
    df_grouped = df.groupby(by=["구분"], group_keys=True)[GROUP_COLUMNS].sum()

    # FIXME: 피봇 테이블에 적힌 내용이 뭔지?
    #     1) 배송비 1차, 2차 구분!
    #     2) 부가세이용료 1차 2차 각각 33000원씩 차감

    # ---------- 저장 ----------
    with pd.ExcelWriter(result_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="전체")
        df_result.to_excel(writer, index=False, sheet_name="매출자료")
        df_grouped.to_excel(writer, sheet_name="피봇")


if __name__ == '__main__':
    TOP_DIR = Path('/Users/jwseo/Movies/9월정산')
    ROZLEY_DIR = TOP_DIR / "로즐리"
    CITY_DIR = TOP_DIR / "시티브리즈"
    ARTID_DIR = TOP_DIR / "아티드"

    # 시티_무신사_정산내역.xlsx
    # 시티_맨_무신사_정산내역.xlsx
    # 시티_글로벌_무신사_정산내역.xlsx

    settle_file1 = CITY_DIR / '시티_서울스토어_정산관리1차_2408.xlsx'
    settle_file2 = CITY_DIR / '시티_서울스토어_정산관리2차_2408.xlsx'
    order_file = CITY_DIR / '시티_서울스토어_구매확정_2408.xlsx'
    result_file = CITY_DIR / '시티_서울스토어_통합✅_2408.xlsx'
    run(settle_file1, settle_file2, order_file, result_file)
