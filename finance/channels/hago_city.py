# TODO: xls -> xlsx
from pathlib import Path

import pandas as pd

# SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame.
# Try using .loc[row_indexer,col_indexer] = value instead
# See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
pd.options.mode.chained_assignment = None


def run(order_file: Path, delivery_file: Path, product_file: Path, result_file: Path):
    """SSG

    Args:
        order_file (Path): 판매분_정산_집계_내역
        delivery_file (Path): 배송비_정산_집계_내역
        product_file (Path): 셀메이트상품목록
        result_file (Path): 최종파일
    """

    CHANNEL_NAME = "하고"
    BRAND_NAME = "시티"

    # 데이터를 pandas DataFrame으로 읽기
    order_df = pd.read_excel(order_file)
    delivery_df = pd.read_excel(delivery_file)

    # product_df = pd.read_excel(product_file, sheet_name=BRAND_NAME, skiprows=1) # 셀메이트
    product_df = pd.read_excel(product_file)

    df = order_df.copy(deep=True)

    # ---------- 주문일자 ----------
    # '주문일시' 컬럼으로부터 '주문일자' 컬럼을 생성
    # 주문일시 : 2024-04-28 21:45:10
    # 주문일자 : 2024-04-28
    df['주문일자'] = df['주문일시'].apply(lambda x: x[:10])

    # ---------- 제품구분 ---------
    # '린넨 브이넥 가디건_IVORY' -> '린넨 브이넥 가디건'
    # '[예약배송할인] [1차 4/17,2차 5/8 예약배송] 루즈핏 사선 절개 싱글 자켓_GREY' -> '루즈핏 사선 절개 싱글 자켓'
    def clean_product_name(product_name: str) -> str:
        # 마지막 ]의 위치를 찾습니다.
        last_bracket_index = product_name.rfind(']')
        if last_bracket_index != -1:
            # 마지막 ] 뒤의 문자열을 제거합니다.
            product_name = product_name[last_bracket_index + 1 :].strip()

        # _의 위치를 찾습니다.
        underscore_index = product_name.find('_')
        if underscore_index != -1:
            # _ 앞의 문자열만 남깁니다.
            product_name = product_name[:underscore_index].strip()

        return product_name

    df['상품명_정리'] = df['상품명'].map(clean_product_name)

    # '시티_하고_상품 목록_2408.xlsx' 컬럼명들
    product_list_columns = [
        '분류',
        '상품코드',
        '상품 번호',
        '제휴사 상품 코드',
        '상품 유형',
        '상품명',
        '진열 상품명',
        '총 재고',
        '색상 / 사이즈 / 옵션재고',
        '진열 상태',
        '승인 상태',
        '주문 제작 상품',
        '해외 배송 상품',
        '브랜드 관리자',
        '브랜드',
        '정가',
        '판매가',
        '1차 할인가',
        '마진',
        '조회수',
        '결제수',
        '구매율',
        '등록일시',
        '수정일시',
    ]

    product_df = product_df.rename(columns={'상품명': '상품목록_상품명'})
    product_df['구분'] = product_df['분류'].map(
        lambda x: '우먼' if '여성' in x else '맨'
    )

    df_right = product_df[
        ['상품코드', '상품 번호', '진열 상품명', '구분', '상품목록_상품명']
    ]
    df_right = df_right.drop_duplicates()

    df = df.rename(columns={"상품 코드": "상품코드"})
    df = pd.merge(df, df_right, how='left', on='상품코드')

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
        "제품구분",
    ]
    df["채널명"] = CHANNEL_NAME
    df['차수'] = '-'
    df["주문일(결제일)"] = df['주문일자']
    # df["주문번호"] = df['주문번호']
    df["상품주문번호"] = '-'
    # df["상품명"] = df['상품명']
    # df["옵션"] = df['옵션']
    df["컬러"] = "-"
    df["사이즈"] = "-"
    # df["수량"] = df['수량']
    df["(판매처) 정상판매가"] = df['총 판매가']
    df["할인 (플랫폼)"] = df['총 본사 부담 할인 금액']
    df["할인 (브랜드)"] = df['총 브랜드 관리자 부담 할인 금액']
    df["고객결제"] = df['총 판매분 부가세 신고 매출액']
    df["매출(v+)"] = df['총 판매분 부가세 신고 매출액']
    df["수수료"] = df['총 판매 수수료']
    df["정산가"] = df['총 정산 금액']
    df["매출구분"] = "제품"
    df['제품구분'] = df['구분']

    df_result = df[COLUMNS]

    # ---------- 저장 ----------
    with pd.ExcelWriter(result_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="전체")
        df_result.to_excel(writer, index=False, sheet_name="매출자료")
        df_right.to_excel(writer, index=False, sheet_name="제품구분")
        order_df.to_excel(writer, index=False, sheet_name="raw판매분")
        delivery_df.to_excel(writer, index=False, sheet_name="raw배송비")


if __name__ == '__main__':
    TOP_DIR = Path('/Users/jwseo/Movies/9월정산')
    # ROZLEY_DIR = TOP_DIR / "로즐리"
    # CITY_DIR = TOP_DIR / "시티브리즈"
    # ARTID_DIR = TOP_DIR / "아티드"

    order_file = (
        TOP_DIR / '시티_하고_판매분_정산_집계_내역_2408.xlsx'
    )  # 판매분_정산_집계_내역
    delivery_file = (
        TOP_DIR / '시티_하고_배송비_정산_집계_내역_2408.xlsx'
    )  # 배송비_정산_집계_내역
    product_file = (
        TOP_DIR / '시티_하고_상품 목록_2408.xlsx'
    )  # 셀메이트상품목록
    result_file = TOP_DIR / '시티_하고_통합✅_2408.xlsx'  # 최종파일
    run(order_file, delivery_file, product_file, result_file)
