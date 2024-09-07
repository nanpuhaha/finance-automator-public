import streamlit as st

st.set_page_config(layout="wide")

st.info(
    "1️⃣ xls 파일은 xlsx로 변환 후 업로드해주세요.\n\n2️⃣ 업로드 파일명은 `브랜드명_(중략)_연월.xlsx` 형식으로 작성해주세요.\n\n    (예시: 시티-맨_ㅁㅁㅁ_ㅁㅁㅁ_ㅁㅁㅁ_2408.xlsx)\n\n    시티_맨❌ 시티-맨✅"
)

st.warning("지그재그는 해결되지 않은 문제가 있습니다.")
st.error("카카오페이, 카페24는 미완성입니다.")
