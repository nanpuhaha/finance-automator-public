# FIXME: 피봇 테이블 값이 이상함

from pathlib import Path

import pandas as pd

# SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame.
# Try using .loc[row_indexer,col_indexer] = value instead
# See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
pd.options.mode.chained_assignment = None


def run(order_file: Path, result_file: Path):
    """무신사

    Args:
        order_file (Path): 정산내역
        result_file (Path): 최종파일
    """

    CHANNEL_NAME = "무신사"

    order_df = pd.read_excel(order_file)

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
    # first_row: [
    #     {
    #         'column': '-',
    #         'value': '구분',
    #     },
    #     {
    #         'column': 'Unnamed: 1',
    #         'value': '일자',
    #     },
    #     {
    #         'column': 'Unnamed: 2',
    #         'value': '주문번호',
    #     },
    #     {
    #         'column': 'Unnamed: 3',
    #         'value': '주문일련번호',
    #     },
    #     {
    #         'column': 'Unnamed: 4',
    #         'value': '복수',
    #     },
    #     {
    #         'column': 'Unnamed: 5',
    #         'value': '쿠폰',
    #     },
    #     {
    #         'column': 'Unnamed: 6',
    #         'value': '상품',
    #     },
    #     {
    #         'column': 'Unnamed: 7',
    #         'value': '옵션',
    #     },
    #     {
    #         'column': 'Unnamed: 8',
    #         'value': '스타일넘버',
    #     },
    #     {
    #         'column': 'Unnamed: 9',
    #         'value': '상품코드',
    #     },
    #     {
    #         'column': 'Unnamed: 10',
    #         'value': '출고형태',
    #     },
    #     {
    #         'column': 'Unnamed: 11',
    #         'value': '매장픽업',
    #     },
    #     {
    #         'column': 'Unnamed: 12',
    #         'value': '주문자',
    #     },
    #     {
    #         'column': 'Unnamed: 13',
    #         'value': '수령자',
    #     },
    #     {
    #         'column': 'Unnamed: 14',
    #         'value': '결제방법',
    #     },
    #     {
    #         'column': 'Unnamed: 15',
    #         'value': '과세',
    #     },
    #     {
    #         'column': 'Unnamed: 16',
    #         'value': '수량',
    #     },
    #     {
    #         'column': '판매금액(판매수수료 계산 기준 금액)(1)',
    #         'value': '판매금액',
    #     },
    #     {
    #         'column': 'Unnamed: 18',
    #         'value': '클레임금액',
    #     },
    #     {
    #         'column': 'Unnamed: 19',
    #         'value': '소계',
    #     },
    #     {
    #         'column': '할인금액(2)',
    #         'value': '할인',
    #     },
    #     {
    #         'column': 'Unnamed: 21',
    #         'value': '무신사쿠폰',
    #     },
    #     {
    #         'column': 'Unnamed: 22',
    #         'value': '적립금',
    #     },
    #     {
    #         'column': 'Unnamed: 23',
    #         'value': '장바구니할인',
    #     },
    #     {
    #         'column': 'Unnamed: 24',
    #         'value': '장바구니쿠폰',
    #     },
    #     {
    #         'column': 'Unnamed: 25',
    #         'value': '그룹할인',
    #     },
    #     {
    #         'column': 'Unnamed: 26',
    #         'value': '소계',
    #     },
    #     {
    #         'column': '브랜드 할인(3)',
    #         'value': '업체쿠폰',
    #     },
    #     {
    #         'column': '추가 금액(4)',
    #         'value': '배송비',
    #     },
    #     {
    #         'column': 'Unnamed: 29',
    #         'value': '반품배송비',
    #     },
    #     {
    #         'column': 'Unnamed: 30',
    #         'value': '저단가배송비지원액',
    #     },
    #     {
    #         'column': 'Unnamed: 31',
    #         'value': '기타정산액',
    #     },
    #     {
    #         'column': 'Unnamed: 32',
    #         'value': '소계',
    #     },
    #     {
    #         'column': '매출액(5)=(1)+(2)+(3)+(4)',
    #         'value': '과세',
    #     },
    #     {
    #         'column': 'Unnamed: 34',
    #         'value': '면세',
    #     },
    #     {
    #         'column': 'Unnamed: 35',
    #         'value': '소계',
    #     },
    #     {
    #         'column': '수수료(판매)(7)',
    #         'value': '수수료율(%)',
    #     },
    #     {
    #         'column': 'Unnamed: 37',
    #         'value': '판매수수료',
    #     },
    #     {
    #         'column': 'Unnamed: 38',
    #         'value': '할인금액',
    #     },
    #     {
    #         'column': 'Unnamed: 39',
    #         'value': '소계',
    #     },
    #     {
    #         'column': 'Unnamed: 40',
    #         'value': '판매지원금(6)',
    #     },
    #     {
    #         'column': '수수료(패널티)(8)',
    #         'value': '반영금액',
    #     },
    #     {
    #         'column': '수수료(청구반품비)(9)',
    #         'value': '반영금액',
    #     },
    #     {
    #         'column': '수수료(10)=(7)+(8)+(9)',
    #         'value': '반영금액',
    #     },
    #     {
    #         'column': '정산(11)=(5)+(6)-(10)',
    #         'value': '합계금액',
    #     },
    #     {
    #         'column': '-.1',
    #         'value': '주문상태',
    #     },
    #     {
    #         'column': 'Unnamed: 46',
    #         'value': '클레임상태',
    #     },
    #     {
    #         'column': 'Unnamed: 47',
    #         'value': '주문일',
    #     },
    #     {
    #         'column': 'Unnamed: 48',
    #         'value': '배송완료일',
    #     },
    #     {
    #         'column': 'Unnamed: 49',
    #         'value': '클레임완료일',
    #     },
    #     {
    #         'column': 'Unnamed: 50',
    #         'value': '구매확정일',
    #     },
    #     {
    #         'column': 'Unnamed: 51',
    #         'value': '비고',
    #     },
    # ]

    last_valid_column = None
    for i, item in enumerate(first_row):
        col = item['column']
        val = item['value']

        if "Unnamed" not in col:
            last_valid_column = col
        else:
            first_row[i]['column'] = last_valid_column

    #  '-.1' -> '-'
    for i, item in enumerate(first_row):
        if item['column'] == '-.1':
            first_row[i]['column'] = '-'

    new_columns = [
        # -
        '구분',
        '일자',
        '주문번호',
        '주문일련번호',
        '복수',
        '쿠폰',
        '상품',
        '옵션',
        '스타일넘버',
        '상품코드',
        '출고형태',
        '매장픽업',
        '주문자',
        '수령자',
        '결제방법',
        '과세',
        '수량',
        # 판매금액(판매수수료 계산 기준 금액)(1)
        '판매금액',
        '클레임금액',
        '판매금액 소계',
        # 할인금액(2)
        '할인',
        '무신사쿠폰',
        '적립금',
        '장바구니할인',
        '장바구니쿠폰',
        '그룹할인',
        '할인금액 소계',
        # 브랜드 할인(3)
        '업체쿠폰',
        # 추가 금액(4)
        '배송비',
        '반품배송비',
        '저단가배송비지원액',
        '기타정산액',
        '추가 금액 소계',
        # 매출액(5)=(1)+(2)+(3)+(4)
        '과세',
        '면세',
        '매출액 소계',
        # 수수료(판매)(7)
        '수수료율(%)',
        '판매수수료',
        '할인금액',
        '수수료(판매) 소계',
        '판매지원금(6)',
        # 수수료(패널티)(8)
        '수수료(패널티)',
        # 수수료(청구반품비)(9)
        '수수료(청구반품비)',
        # 수수료(10)=(7)+(8)+(9)
        '수수료',
        # 정산(11)=(5)+(6)-(10)
        '정산액',
        # -
        '주문상태',
        '클레임상태',
        '주문일',
        '배송완료일',
        '클레임완료일',
        '구매확정일',
        '비고',
    ]

    df = df.rename(columns=dict(zip(columns, new_columns)))

    def remove_rows(df, start, end):
        # Drop the specified range of rows by their indices
        df = df.drop(df.index[start : end + 1])
        # Reset the index to keep the DataFrame tidy
        df = df.reset_index(drop=True)
        return df

    df = remove_rows(df, 0, 5)

    # ---------- 주문일자 ----------
    # '주문일' 컬럼으로부터 '주문일자' 컬럼을 생성
    # 주문일: 20240401
    # 주문일자: 2024-04-01
    def extract_order_date(x):
        x = str(x)
        return f"{x[0:4]}-{x[4:6]}-{x[6:8]}"

    df['주문일자'] = df['주문일'].map(extract_order_date)
    # df['주문일자'] = df['주문일'].apply(lambda x: f"{x[0:4]}-{x[4:6]}-{x[6:8]}")

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
    # df["주문번호"] = df["주문번호"]
    df["상품주문번호"] = df['주문일련번호']
    df["상품명"] = df["상품"]
    # df["옵션"] = df['옵션']
    df["컬러"] = "-"
    df["사이즈"] = "-"
    # df["수량"] = df["수량"]
    df["(판매처) 정상판매가"] = df['판매금액 소계']
    df["할인 (플랫폼)"] = -1 * df['할인금액 소계']
    df["할인 (브랜드)"] = -1 * df['업체쿠폰']
    df["고객결제"] = df['매출액 소계'] - df['추가 금액 소계']
    df["매출(v+)"] = df['고객결제']
    # df["수수료"] = df['수수료']
    df["정산가"] = df["매출(v+)"] - df["수수료"]
    df["매출구분"] = "제품"

    df_result = df[COLUMNS]

    GROUP_COLUMNS = [
        '배송비',
        '반품배송비',
        '저단가배송비지원액',
        '기타정산액',
        '추가 금액 소계',
        '매출액 소계',
        '판매지원금(6)',
        '수수료',
        '정산액',
    ]
    df_grouped = df.groupby(by=["구분"], group_keys=True)[GROUP_COLUMNS].sum()
    df_grouped.loc["총합계"] = df[GROUP_COLUMNS].sum()

    # FIXME: 피봇 테이블 값이 이상함

    # # ---------- 저장 ----------
    with pd.ExcelWriter(result_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="전체")
        df_result.to_excel(writer, index=False, sheet_name="매출자료")
        df_grouped.to_excel(writer, index=True, sheet_name="피봇")


if __name__ == '__main__':
    TOP_DIR = Path('/Users/jwseo/Movies/9월정산')
    ROZLEY_DIR = TOP_DIR / "로즐리"
    CITY_DIR = TOP_DIR / "시티브리즈"
    ARTID_DIR = TOP_DIR / "아티드"

    # 시티_무신사_정산내역.xlsx
    # 시티_맨_무신사_정산내역.xlsx
    # 시티_글로벌_무신사_정산내역.xlsx

    order_file = TOP_DIR / '아티드_무신사_정산내역_2408.xlsx'
    result_file = TOP_DIR / '아티드_무신사_통합✅_2408.xlsx'
    print("아티드_무신사")
    run(order_file, result_file)

    order_file = TOP_DIR / '시티_무신사_정산내역_2408.xlsx'
    result_file = TOP_DIR / '시티_무신사_통합✅_2408.xlsx'
    print("시티_무신사")
    run(order_file, result_file)

    order_file = TOP_DIR / '시티_무신사_맨_정산내역_2408.xlsx'
    result_file = TOP_DIR / '시티_맨_무신사_통합✅_2408.xlsx'
    print("시티_맨_무신사")
    run(order_file, result_file)
