import base64
import io
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

import requests
import streamlit as st

import pipeline as pl

ASSET_CATEGORY_PATH = Path(__file__).parent / "카테고리번호_기본.xlsx"
ASSET_TEMPLATE_PATH = Path(__file__).parent / "스마트스토어기본_기본.xlsx"

GITHUB_OWNER = "jihun2da"
GITHUB_REPO = "product-update-generator"
GITHUB_BRANCH = "main"

st.set_page_config(page_title="신상품 업데이트 파일 생성기", layout="wide")


def save_default_to_github(github_path: str, new_bytes: bytes, commit_message: str):
    """업로드한 파일을 깃허브 저장소의 github_path에 커밋해서 앱의 기본값을 영구적으로 교체합니다.
    (성공하면 저장소가 자동으로 재배포되어, 이후에는 업로드 없이도 이 파일이 기본값으로 쓰입니다.)

    st.secrets['GITHUB_TOKEN']에 이 저장소에 대한 쓰기 권한(Contents: Read and write)이 있는
    GitHub Personal Access Token이 등록되어 있어야 합니다. (Streamlit Cloud 앱 관리 화면의
    Settings -> Secrets 에서 등록. README.txt 참고)
    """
    try:
        token = st.secrets.get("GITHUB_TOKEN")
    except Exception:
        token = None

    if not token:
        return False, (
            "GITHUB_TOKEN이 설정되어 있지 않아 기본값을 영구 저장할 수 없습니다. "
            "Streamlit Cloud 앱의 Settings → Secrets에 GITHUB_TOKEN을 등록해주세요 (README.txt 참고)."
        )

    api_url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{github_path}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github+json",
    }

    try:
        get_resp = requests.get(api_url, headers=headers, params={"ref": GITHUB_BRANCH}, timeout=15)
        sha = get_resp.json().get("sha") if get_resp.status_code == 200 else None

        payload = {
            "message": commit_message,
            "content": base64.b64encode(new_bytes).decode("utf-8"),
            "branch": GITHUB_BRANCH,
        }
        if sha:
            payload["sha"] = sha

        put_resp = requests.put(api_url, headers=headers, json=payload, timeout=20)
        if put_resp.status_code in (200, 201):
            return True, None

        try:
            detail = put_resp.json().get("message", put_resp.text)
        except Exception:
            detail = put_resp.text
        return False, f"깃허브 저장 실패 ({put_resp.status_code}): {detail}"
    except Exception as e:
        return False, f"깃허브 저장 중 오류: {e}"

st.title("신상품 업데이트 파일 생성기")
st.caption(
    "기린.xlsx / S파일 두 개만 올리면, 지금까지 손으로 하시던 열 복사·파일 저장·카테고리 갱신·"
    "옵션ver3·자동업데이트파일생성(2026) 과정을 한 번에 실행해 결과 파일을 만들어 드립니다."
)

with st.sidebar:
    st.header("1. 필수 파일 업로드")
    kirin_file = st.file_uploader("기린.xlsx (E,F,G / AJ 값을 가져올 소스 파일)", type=["xlsx"])
    sfile_file = st.file_uploader("S파일 (예: S20260820.xlsx)", type=["xlsx"])

    st.header("2. 선택 파일 (없으면 기본값 사용)")
    category_file = st.file_uploader(
        "카테고리번호.xlsx (업로드 안 하면 앱에 저장된 기본 파일 사용)", type=["xlsx"]
    )
    save_category_default = st.checkbox(
        "업로드한 카테고리번호.xlsx를 앱의 새 기본값으로 영구 저장(교체)",
        value=False,
        disabled=category_file is None,
        help="체크하고 실행하면, 이번 실행뿐 아니라 앞으로 업로드하지 않아도 이 파일이 기본값으로 쓰입니다.",
    )
    template_file = st.file_uploader(
        "스마트스토어기본.xlsx (업로드 안 하면 앱에 저장된 기본 템플릿 사용)", type=["xlsx"]
    )
    save_template_default = st.checkbox(
        "업로드한 스마트스토어기본.xlsx를 앱의 새 기본값으로 영구 저장(교체)",
        value=False,
        disabled=template_file is None,
        help="체크하고 실행하면, 이번 실행뿐 아니라 앞으로 업로드하지 않아도 이 파일이 기본값으로 쓰입니다.",
    )

    st.header("3. 옵션 설정")
    today = datetime.today()
    dated_date = st.date_input("날짜파일(옵션ver3 결과) 날짜", value=today)
    next_s_date = st.date_input("다음 라운드용 S파일 날짜", value=today + timedelta(days=1))
    category_row = st.number_input("카테고리번호.xlsx 갱신 행 번호", min_value=2, value=29, step=1)
    smartstore_col = st.text_input("스마트스토어(N파일) 카테고리 매핑 컬럼명", value="스마트오토")

    run_clicked = st.button("전체 파이프라인 실행", type="primary", use_container_width=True)


def build_zip(stage1: dict, stage2: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in stage1.items():
            zf.writestr(f"1단계_예제생성/{name}", data)
        for name, data in stage2.items():
            zf.writestr(f"2단계_판매처파일/{name}", data)
    return buf.getvalue()


if run_clicked:
    if not kirin_file or not sfile_file:
        st.error("기린.xlsx 와 S파일은 필수입니다. 왼쪽에서 두 파일을 모두 올려주세요.")
        st.stop()

    kirin_bytes = kirin_file.read()
    sfile_bytes = sfile_file.read()

    if category_file:
        category_bytes = category_file.read()
    elif ASSET_CATEGORY_PATH.exists():
        category_bytes = ASSET_CATEGORY_PATH.read_bytes()
        st.info("카테고리번호.xlsx를 별도로 올리지 않아 앱에 저장된 기본 파일을 사용했습니다.")
    else:
        st.error("카테고리번호.xlsx가 없습니다. 파일을 올려주세요.")
        st.stop()

    if template_file:
        template_bytes = template_file.read()
    elif ASSET_TEMPLATE_PATH.exists():
        template_bytes = ASSET_TEMPLATE_PATH.read_bytes()
        st.info("스마트스토어기본.xlsx를 별도로 올리지 않아 앱에 저장된 기본 템플릿을 사용했습니다.")
    else:
        template_bytes = None
        st.warning("스마트스토어기본.xlsx 템플릿이 없어 N파일(스마트스토어 변환) 생성을 건너뜁니다.")

    if category_file and save_category_default:
        with st.spinner("카테고리번호.xlsx를 앱의 새 기본값으로 저장하는 중..."):
            ok, err = save_default_to_github(
                "카테고리번호_기본.xlsx", category_bytes,
                "카테고리번호_기본.xlsx 갱신 (앱에서 업로드)",
            )
        if ok:
            st.success(
                "카테고리번호.xlsx를 앱의 새 기본값으로 저장했습니다. "
                "(재배포가 완료되기까지 1~2분 정도 걸릴 수 있습니다)"
            )
        else:
            st.error(f"카테고리번호.xlsx 기본값 저장 실패: {err}")

    if template_file and save_template_default:
        with st.spinner("스마트스토어기본.xlsx를 앱의 새 기본값으로 저장하는 중..."):
            ok, err = save_default_to_github(
                "스마트스토어기본_기본.xlsx", template_bytes,
                "스마트스토어기본_기본.xlsx 갱신 (앱에서 업로드)",
            )
        if ok:
            st.success(
                "스마트스토어기본.xlsx를 앱의 새 기본값으로 저장했습니다. "
                "(재배포가 완료되기까지 1~2분 정도 걸릴 수 있습니다)"
            )
        else:
            st.error(f"스마트스토어기본.xlsx 기본값 저장 실패: {err}")

    with st.spinner("파이프라인 실행 중..."):
        try:
            result = pl.run_pipeline(
                kirin_bytes=kirin_bytes,
                sfile_bytes=sfile_bytes,
                category_bytes=category_bytes,
                template_bytes=template_bytes,
                dated_filename_date=dated_date.strftime("%Y%m%d"),
                next_sfile_date=next_s_date.strftime("%Y%m%d"),
                category_row=int(category_row),
                smartstore_target_col=smartstore_col,
            )
        except Exception as e:
            st.exception(e)
            st.stop()

    if result["errors"]:
        st.error("아래 오류가 발생했습니다. 결과가 일부 누락됐을 수 있습니다.")
        for e in result["errors"]:
            st.write("❌", e)

    if result["logs"]:
        with st.expander("실행 로그 / 참고 사항 보기", expanded=True):
            for l in result["logs"]:
                st.write("•", l)

    st.success("파이프라인 실행 완료")

    zip_bytes = build_zip(result["stage1"], result["stage2"])
    st.download_button(
        "전체 결과 ZIP 다운로드",
        data=zip_bytes,
        file_name=f"업데이트파일_{datetime.today().strftime('%Y%m%d_%H%M')}.zip",
        mime="application/zip",
        type="primary",
    )

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1단계 결과 (예제 생성 과정)")
        for name, data in result["stage1"].items():
            st.download_button(name, data=data, file_name=name, key=f"s1_{name}")
    with col2:
        st.subheader("2단계 결과 (판매처별 최종 파일)")
        for name, data in result["stage2"].items():
            st.download_button(name, data=data, file_name=name, key=f"s2_{name}")
else:
    st.info("왼쪽에서 파일을 올리고 '전체 파이프라인 실행'을 눌러주세요.")
