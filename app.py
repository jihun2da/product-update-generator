import base64
import io
import json
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

import requests
import streamlit as st

import pipeline as pl

ASSET_CATEGORY_PATH = Path(__file__).parent / "카테고리번호_기본.xlsx"
ASSET_TEMPLATE_PATH = Path(__file__).parent / "스마트스토어기본_기본.xlsx"
ASSET_VENDOR_CONFIG_PATH = Path(__file__).parent / "판매처설정_기본.json"

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

def load_category_vendor_names(category_bytes) -> list:
    """카테고리번호.xlsx(bytes)의 1행 헤더 중 '원본'을 제외한 판매처 이름 목록을 반환합니다.
    새 판매처를 추가할 때 '이 판매처와 동일한 카테고리 값 복사' 선택지에 씁니다."""
    if not category_bytes:
        return []
    try:
        wb = pl.load_wb(category_bytes, data_only=False)
        ws = wb.active
        names = []
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=1, column=c).value
            if v is None:
                continue
            v = str(v).strip()
            if v and v != '원본':
                names.append(v)
        return names
    except Exception:
        return []


def load_vendor_config() -> dict:
    """앱에 저장된 '판매처 삭제/추가' 기본 설정(판매처설정_기본.json)을 읽어옵니다.
    파일이 없거나 읽기 실패하면 빈 설정(제외 없음, 추가 없음)을 반환합니다."""
    default = {"excluded_vendors": [], "custom_vendors": []}
    if not ASSET_VENDOR_CONFIG_PATH.exists():
        return default
    try:
        data = json.loads(ASSET_VENDOR_CONFIG_PATH.read_text(encoding="utf-8"))
        excluded = [v for v in (data.get("excluded_vendors") or []) if isinstance(v, str)]
        custom = [c for c in (data.get("custom_vendors") or []) if isinstance(c, dict) and c.get("name")]
        return {"excluded_vendors": excluded, "custom_vendors": custom}
    except Exception:
        return default


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

    st.header("4. 판매처 파일 관리 (삭제 / 추가)")

    # 앱에 저장된 기본 설정은 세션당 한 번만 불러오고, 이후에는 사용자가 화면에서
    # 직접 편집한 내용(세션 상태)을 그대로 유지합니다.
    if "vendor_config_loaded" not in st.session_state:
        _vc = load_vendor_config()
        st.session_state["vendor_config_loaded"] = True
        st.session_state["vendor_excluded_default"] = [
            v for v in _vc["excluded_vendors"] if v in pl.ALL_OPTIONAL_VENDOR_NAMES
        ]
        st.session_state["custom_vendor_list"] = _vc["custom_vendors"]

    with st.expander("삭제할 판매처 선택 (더 이상 생성하지 않을 파일)", expanded=False):
        st.caption(
            "예제생성 파일 중 곽충현/문정희, 판매처 파일 18개 중 X(날짜)파일을 제외한 "
            "나머지 중에서 골라 생성을 건너뛸 수 있습니다. X파일은 스마트스토어(N파일) "
            "변환에 내부적으로 쓰여 제외 목록에서 뺐습니다."
        )
        excluded_vendors = st.multiselect(
            "생성하지 않을 파일",
            options=pl.ALL_OPTIONAL_VENDOR_NAMES,
            default=st.session_state["vendor_excluded_default"],
            key="excluded_vendors_select",
            help="여기서 선택한 판매처는 이번 실행부터 결과물(다운로드 목록/ZIP)에서 빠집니다.",
        )

    # 새 판매처의 카테고리 매핑 방식 선택지 (내부 key -> 화면 표시 라벨)
    CATEGORY_MODE_LABELS = {
        "copy": "기존 판매처와 동일한 값 사용 (권장)",
        "value": "모든 행에 같은 값 하나 입력",
        "none": "카테고리 매핑 안 함 (파일만 생성)",
    }

    # 카테고리번호.xlsx는 판매처마다(컬럼마다) 원본 코드별로 값이 다릅니다(상품 종류별로
    # 분류가 다르기 때문) — 그래서 "복사할 기존 판매처 선택" 옵션에 쓸 이름 목록을,
    # 지금 화면에 올라와 있는(또는 앱 기본) 카테고리번호.xlsx에서 미리 읽어둡니다.
    if category_file is not None:
        _category_bytes_for_ui = category_file.getvalue()
    elif ASSET_CATEGORY_PATH.exists():
        _category_bytes_for_ui = ASSET_CATEGORY_PATH.read_bytes()
    else:
        _category_bytes_for_ui = None
    existing_vendor_names_for_copy = load_category_vendor_names(_category_bytes_for_ui)

    with st.expander("새 판매처 파일 추가", expanded=False):
        st.caption(
            "예제생성/판매처 파일과 같은 스타일(상품명 처리 방식)로 새 판매처 파일을 추가합니다. "
            "카테고리번호.xlsx는 판매처마다 상품 코드별로 값이 다르므로, 대부분은 이미 설정되어 "
            "있는 비슷한 판매처를 골라 '기존 판매처와 동일한 값 사용'을 쓰는 것이 정확합니다."
        )

        editing_idx = st.session_state.get("editing_custom_vendor_index")

        if st.session_state["custom_vendor_list"]:
            st.write("현재 추가된 판매처:")
            for idx, cv in enumerate(st.session_state["custom_vendor_list"]):
                style_label = pl.VENDOR_STYLE_INFO.get(
                    cv.get("style", pl.DEFAULT_VENDOR_STYLE), {}
                ).get("label", cv.get("style", ""))
                mode = cv.get("category_mode")
                if not mode:
                    mode = "value" if cv.get("category_value") else "none"
                if mode == "copy":
                    cat_desc = f"'{cv.get('category_source')}'와 동일한 값 사용"
                elif mode == "value":
                    cat_desc = f"모든 행 '{cv.get('category_value')}'"
                else:
                    cat_desc = "매핑 안 함"
                col_a, col_b, col_c = st.columns([5, 1, 1])
                with col_a:
                    prefix = "✏️ " if idx == editing_idx else "• "
                    st.write(f"{prefix}**{cv.get('name')}** — {style_label} · 카테고리: {cat_desc}")
                with col_b:
                    if st.button("수정", key=f"edit_custom_vendor_{idx}"):
                        st.session_state["editing_custom_vendor_index"] = idx
                        st.session_state["new_vendor_name"] = cv.get("name", "")
                        st.session_state["new_vendor_style"] = cv.get(
                            "style", pl.DEFAULT_VENDOR_STYLE
                        )
                        st.session_state["new_vendor_category_mode"] = mode
                        if mode == "copy" and cv.get("category_source") in existing_vendor_names_for_copy:
                            st.session_state["new_vendor_category_source"] = cv.get("category_source")
                        st.session_state["new_vendor_category_value"] = cv.get("category_value", "")
                        st.rerun()
                with col_c:
                    if st.button("삭제", key=f"remove_custom_vendor_{idx}"):
                        st.session_state["custom_vendor_list"].pop(idx)
                        if editing_idx == idx:
                            st.session_state.pop("editing_custom_vendor_index", None)
                        st.rerun()

        if editing_idx is not None:
            st.markdown(f"**판매처 수정 중: {st.session_state['custom_vendor_list'][editing_idx].get('name')}**")
        else:
            st.markdown("**새 판매처 추가**")

        new_vendor_name = st.text_input("판매처 이름 (파일명이 됩니다, 예: 신규거래처)", key="new_vendor_name")
        new_vendor_style = st.selectbox(
            "상품명 처리 스타일",
            options=list(pl.VENDOR_STYLE_INFO.keys()),
            format_func=lambda k: pl.VENDOR_STYLE_INFO[k]["label"],
            key="new_vendor_style",
        )
        new_vendor_category_mode = st.radio(
            "카테고리 매핑 방식",
            options=list(CATEGORY_MODE_LABELS.keys()),
            format_func=lambda k: CATEGORY_MODE_LABELS[k],
            key="new_vendor_category_mode",
        )
        if new_vendor_category_mode == "copy":
            if existing_vendor_names_for_copy:
                # 카테고리번호.xlsx가 이전 화면과 달라져 선택돼 있던 값이 더 이상 목록에
                # 없으면(예: 다른 카테고리번호.xlsx를 새로 업로드), selectbox가 오류 없이
                # 뜨도록 첫 번째 항목으로 되돌려둡니다.
                if st.session_state.get("new_vendor_category_source") not in existing_vendor_names_for_copy:
                    st.session_state["new_vendor_category_source"] = existing_vendor_names_for_copy[0]
                new_vendor_category_source = st.selectbox(
                    "복사할 기존 판매처 선택",
                    options=existing_vendor_names_for_copy,
                    key="new_vendor_category_source",
                    help="선택한 판매처의 카테고리번호.xlsx 값을 행(상품 코드)별로 그대로 복사해 옵니다.",
                )
            else:
                new_vendor_category_source = None
                st.warning("카테고리번호.xlsx에서 판매처 목록을 읽지 못했습니다. 먼저 카테고리번호.xlsx를 확인해주세요.")
            new_vendor_category_value = ""
        elif new_vendor_category_mode == "value":
            new_vendor_category_source = None
            new_vendor_category_value = st.text_input(
                "카테고리 값 (모든 행에 동일하게 채워집니다)",
                key="new_vendor_category_value",
            )
        else:
            new_vendor_category_source = None
            new_vendor_category_value = ""

        btn_col1, btn_col2 = st.columns([1, 1])
        with btn_col1:
            submit_label = "저장" if editing_idx is not None else "+ 판매처 추가"
            submit_clicked = st.button(submit_label, key="add_custom_vendor_btn", use_container_width=True)
        with btn_col2:
            cancel_clicked = False
            if editing_idx is not None:
                cancel_clicked = st.button("수정 취소", key="cancel_edit_custom_vendor_btn", use_container_width=True)

        if cancel_clicked:
            st.session_state.pop("editing_custom_vendor_index", None)
            st.rerun()

        if submit_clicked:
            name = new_vendor_name.strip()
            other_entries = [
                c for i, c in enumerate(st.session_state["custom_vendor_list"]) if i != editing_idx
            ]
            existing_names = {c.get("name") for c in other_entries}
            if not name:
                st.warning("판매처 이름을 입력해주세요.")
            elif name in pl.ALL_OPTIONAL_VENDOR_NAMES or name in existing_names:
                st.warning(f"'{name}'은 이미 사용 중인 판매처 이름입니다. 다른 이름을 입력해주세요.")
            elif new_vendor_category_mode == "copy" and not new_vendor_category_source:
                st.warning("복사할 기존 판매처를 선택해주세요.")
            else:
                new_entry = {
                    "name": name,
                    "style": new_vendor_style,
                    "category_mode": new_vendor_category_mode,
                    "category_source": new_vendor_category_source or "",
                    "category_value": new_vendor_category_value.strip(),
                }
                if editing_idx is not None:
                    st.session_state["custom_vendor_list"][editing_idx] = new_entry
                    st.session_state.pop("editing_custom_vendor_index", None)
                else:
                    st.session_state["custom_vendor_list"].append(new_entry)
                st.rerun()

    save_vendor_config_default = st.checkbox(
        "이 판매처 삭제/추가 설정을 앱의 새 기본값으로 영구 저장",
        value=False,
        help="체크하고 실행하면, 이번 실행뿐 아니라 앞으로 매번 다시 설정하지 않아도 "
        "이 삭제/추가 목록이 기본값으로 쓰입니다.",
    )

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

    custom_vendors = st.session_state.get("custom_vendor_list", [])

    if save_vendor_config_default:
        vendor_config_bytes = json.dumps(
            {"excluded_vendors": excluded_vendors, "custom_vendors": custom_vendors},
            ensure_ascii=False, indent=2,
        ).encode("utf-8")
        with st.spinner("판매처 삭제/추가 설정을 앱의 새 기본값으로 저장하는 중..."):
            ok, err = save_default_to_github(
                "판매처설정_기본.json", vendor_config_bytes,
                "판매처설정_기본.json 갱신 (앱에서 설정)",
            )
        if ok:
            st.success(
                "판매처 삭제/추가 설정을 앱의 새 기본값으로 저장했습니다. "
                "(재배포가 완료되기까지 1~2분 정도 걸릴 수 있습니다)"
            )
        else:
            st.error(f"판매처 설정 저장 실패: {err}")

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
                excluded_vendors=excluded_vendors,
                custom_vendors=custom_vendors,
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
