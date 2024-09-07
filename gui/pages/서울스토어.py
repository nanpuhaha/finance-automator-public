from pathlib import Path

import streamlit as st
from streamlit.runtime.uploaded_file_manager import UploadedFile

try:
    from finance.channels.seoul_store import run
    from finance.xl.password import unlock_if_locked
    from utils.temp import TEMP_DIR, clear_temp_dir
except ModuleNotFoundError:
    import sys

    sys.path.append(str(Path(__file__).parent.parent.parent))
    from finance.channels.seoul_store import run
    from finance.xl.password import unlock_if_locked
    from utils.temp import TEMP_DIR, clear_temp_dir


TITLE = "서울스토어"

st.set_page_config(layout="wide")

# 파일 업로드
st.title(TITLE)

st.info(
    "1️⃣ xls 파일은 xlsx로 변환 후 업로드해주세요.\n\n2️⃣ 업로드 파일명은 `브랜드명_(중략)_연월.xlsx` 형식으로 작성해주세요.\n\n    (예시: 로즐리_ㅁㅁㅁ_ㅁㅁㅁ_ㅁㅁㅁ_2408.xlsx)"
)

col1, col2, col3 = st.columns(3)

with col1:
    file1: UploadedFile | None = st.file_uploader(
        "정산관리1차 파일 업로드",
        type=["xlsx"],
        help="xls 파일은 xlsx로 변환 후 업로드해주세요.",
    )
with col2:
    file2: UploadedFile | None = st.file_uploader(
        "정산관리2차 파일 업로드",
        type=["xlsx"],
        help="xls 파일은 xlsx로 변환 후 업로드해주세요.",
    )
with col3:
    file3: UploadedFile | None = st.file_uploader(
        "구매확정 파일 업로드",
        type=["xlsx"],
        help="xls 파일은 xlsx로 변환 후 업로드해주세요.",
    )

# 파일 업로드 상태 확인
files_uploaded = file1 is not None and file2 is not None and file3 is not None

button_col, warning_col = st.columns([1, 3])

# 사용자 정의 스타일을 추가
st.markdown(
    """
    <style>
    .stButton>button {
        margin-top: 5px;
        padding: 10px 20px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# 작업 실행 버튼 (비활성화 제어)
with button_col:
    with st.form(key="file_upload_form", border=False):
        submit_button = st.form_submit_button("작업 실행", disabled=not files_uploaded)

# 경고 메시지
with warning_col:
    if not files_uploaded:
        st.warning("모든 파일을 선택해야 작업을 실행할 수 있습니다.")

# 처리 버튼
if submit_button:
    if files_uploaded:

        # 파일을 로컬 임시 경로로 저장
        file1_path = TEMP_DIR / file1.name
        file2_path = TEMP_DIR / file2.name
        file3_path = TEMP_DIR / file3.name

        with open(file1_path, "wb") as f:
            f.write(file1.getbuffer())
        with open(file2_path, "wb") as f:
            f.write(file2.getbuffer())
        with open(file3_path, "wb") as f:
            f.write(file3.getbuffer())

        # 파일명에서 브랜드명과 연월 추출
        brand_name = file1_path.name.split("_")[0]
        date_info = file1_path.name.split("_")[-1].replace(".xlsx", "")
        result_file_name = f"{brand_name}_{TITLE}_통합✅_{date_info}.xlsx"
        result_file_path = TEMP_DIR / result_file_name

        # run 함수 실행
        run(file1_path, file2_path, file3_path, result_file_path)

        # 결과 파일 다운로드 버튼 제공
        with open(result_file_path, "rb") as f:
            result_data = f.read()

        file1_path.unlink()
        file2_path.unlink()
        file3_path.unlink()
        result_file_path.unlink()

        if st.download_button(
            label="결과 파일 다운로드",
            data=result_data,
            file_name=result_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ):
            # 다운로드 완료 후 TEMP 디렉토리 비우기
            clear_temp_dir(TEMP_DIR)

    else:
        st.warning("모든 파일을 업로드해주세요.")
