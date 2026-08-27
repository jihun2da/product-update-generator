"""
카페24 자동 업로드 (별도 탭 / 독립 기능)
========================================
기존 '신상품 업데이트 파일 생성기'(app.py) 로직과는 완전히 분리된 페이지입니다.
이 페이지의 코드가 어떻게 바뀌어도 app.py / pipeline.py 는 전혀 영향을 받지 않습니다.

실제 카페24 계정으로 로그인 → 업로드 화면 이동 → 파일 업로드 → 성공/실패 판별까지
전체 흐름을 사용자가 직접 테스트했고, 정상 동작하는 것을 확인했습니다(2026-08-26).
"""

import io
import re
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

import cafe24_uploader as up

# 파일명 끝이 8자리 날짜(YYYYMMDD)로 끝나는 경우, 그 날짜 부분을 떼어내고 비교하기 위한 패턴.
# 예: "X20260820" -> "X", "김하늘" -> "김하늘"(날짜가 없으면 그대로).
# 이렇게 하면 매일 파일명 뒤의 날짜가 바뀌어도(X20260820.xlsx, X20260821.xlsx, ...),
# Secrets에 등록된 계정 이름이 "X"(또는 날짜가 포함된 이름 아무거나)이기만 하면
# 앞부분("X")이 같은 파일과 자동으로 매칭됩니다.
_DATE_SUFFIX_RE = re.compile(r"\d{8}$")


def _normalize_for_match(name: str) -> str:
    return _DATE_SUFFIX_RE.sub("", name.strip())

st.set_page_config(page_title="카페24 자동 업로드", layout="wide")

st.title("카페24 자동 업로드 (베타 · 별도 기능)")

# 코드를 업데이트한 뒤 "정말 최신 코드로 테스트한 게 맞는지" 헷갈리는 경우가 많아서
# 추가한 버전 표시입니다. 로컬 파일을 새로 받은 뒤 앱을 재시작했는데도 이 값이 안
# 바뀌어 있다면, 아직 예전 코드로 실행 중이라는 뜻입니다.
st.caption(f"🔧 코드 버전: `{up.MODULE_VERSION}`")

st.success(
    "✅ 실제 카페24 계정으로 '로그인 → 업로드 화면 이동 → 파일 업로드 → 업로드 버튼 클릭'까지의 "
    "흐름을 테스트했고, 버튼 클릭 자체는 정상 동작합니다. 다만 클릭 이후 실제 서버 처리가 끝났는지는 "
    "화면 문구만으로 단정하기 어려워서, 실제 '업로드 처리' 요청과 그 응답을 직접 확인해서 "
    "성공/오류를 판별하도록 개선했습니다(아래 '③'의 대기 시간 설정 참고).\n\n"
    "- 실패(오류)로 판별된 계정은 재시도하지 않고 끝까지 진행한 뒤, 모아서 알려드립니다.\n"
    "- 계정별 스크린샷과 진단 정보는 '④ 업로드 결과'에서 계속 확인할 수 있으니, 새로운 계정이나 "
    "파일을 추가하실 때는 처음 한 번씩 결과를 확인해보시는 걸 권장드립니다.\n"
    "- 이 앱(app.py)의 다른 기능(파일 생성 등)에는 이 기능이 전혀 영향을 주지 않습니다."
)


# ---------------------------------------------------------------------------
# 1) Secrets에서 계정 목록 불러오기
# ---------------------------------------------------------------------------
def load_accounts():
    """Secrets에서 계정 목록을 읽어옵니다.

    각 계정은 Secrets 안에서 고유한 키(예: "모나마켓", 또는 같은 이름이 여러 개면
    "모나마켓_2", "모나마켓_3"...)를 가져야 하지만, 실제로 업로드할 때 어떤 파일과
    매칭시킬지는 별도의 "file" 값으로 관리합니다(생략하면 계정 키 자체를 사용).
    이렇게 분리해두면, 아이디/비밀번호는 다르지만 같은 파일을 올려야 하는 계정
    여러 개를 "이름이 같다"는 이유로 1개만 남기고 지우지 않고 전부 등록할 수 있습니다.
    """
    try:
        cafe24_conf = st.secrets.get("cafe24")
    except Exception:
        cafe24_conf = None
    if not cafe24_conf:
        return {}
    accounts_conf = None
    try:
        accounts_conf = cafe24_conf.get("accounts")
    except Exception:
        accounts_conf = None
    if not accounts_conf:
        return {}

    result = {}
    for name, info in accounts_conf.items():
        try:
            file_key = None
            try:
                file_key = info.get("file")
            except Exception:
                file_key = None
            file_key = str(file_key).strip() if file_key else ""
            if not file_key:
                file_key = name
            result[name] = {"id": info["id"], "password": info["password"], "file": file_key}
        except Exception:
            continue
    return result


accounts = load_accounts()

with st.expander("① Secrets 등록 도우미 (아이디비번.xlsx → Secrets 붙여넣기용 텍스트 생성)", expanded=(len(accounts) == 0)):
    st.write(
        "카페24 계정 아이디/비밀번호는 이 앱 화면이 아니라, Streamlit Cloud의 **Secrets**에만 "
        "등록합니다(저에게 알려주실 필요 없습니다). 아래에 '이름/아이디/비밀번호' 컬럼이 있는 "
        "엑셀을 올리면, Secrets에 붙여넣을 텍스트를 만들어서 **파일로만** 내려드립니다 "
        "(화면에 실제 아이디/비밀번호는 표시하지 않습니다)."
    )
    st.caption(
        "💡 서로 다른 아이디인데 **같은 파일을 업로드해야 하는 계정이 여러 개**라면 "
        "'이름' 컬럼에 같은 이름을 그대로 여러 번 써서 올려도 됩니다 — 더 이상 중복으로 "
        "취급해 1개만 남기지 않고, 전부 별도 계정으로 등록하면서 같은 파일과 매칭되도록 "
        "처리합니다. 필요하면 '파일명' 컬럼을 추가해서, 실제 업로드할 파일 이름을 "
        "'이름'과 다르게 직접 지정할 수도 있습니다(선택 사항)."
    )
    cred_file = st.file_uploader("이름/아이디/비밀번호 컬럼이 있는 엑셀 업로드", type=["xlsx"], key="cred_xlsx")

    def _find_col_map(columns):
        col_map = {}
        for col in columns:
            c = str(col).strip()
            if c in ("이름", "성함", "name", "Name"):
                col_map["name"] = col
            elif c in ("아이디", "ID", "id", "Id"):
                col_map["id"] = col
            elif c in ("비밀번호", "비번", "PW", "pw", "password", "Password"):
                col_map["pw"] = col
            elif c in ("파일명", "파일", "업로드파일명", "file", "File"):
                col_map["file"] = col
        return col_map

    if cred_file is not None:
        # 엑셀 파일에 시트가 여러 개일 수 있습니다(예: 상품 템플릿 시트 + 계정 목록 시트가
        # 한 파일 안에 같이 저장된 경우). pandas는 시트를 지정하지 않으면 첫 번째 시트만
        # 읽으므로, 이름/아이디/비밀번호 컬럼이 있는 시트를 찾을 때까지 모든 시트를
        # 순서대로 확인합니다.
        df = None
        col_map = {}
        checked_sheets = []
        try:
            xls = pd.ExcelFile(cred_file)
            for sheet_name in xls.sheet_names:
                sheet_df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str).fillna("")
                found = _find_col_map(sheet_df.columns)
                checked_sheets.append((sheet_name, list(sheet_df.columns)))
                if all(k in found for k in ("name", "id", "pw")):
                    df = sheet_df
                    col_map = found
                    if len(xls.sheet_names) > 1:
                        st.info(f"'{sheet_name}' 시트에서 이름/아이디/비밀번호 컬럼을 찾았습니다 (전체 시트: {xls.sheet_names}).")
                    break
        except Exception as e:
            st.error(f"엑셀을 읽는 중 오류가 발생했습니다: {e}")
            df = None

        if df is None:
            detail = "; ".join(f"[{name}] {cols}" for name, cols in checked_sheets)
            st.error(
                "엑셀의 어느 시트에서도 '이름', '아이디', '비밀번호' 컬럼을 찾지 못했습니다. "
                f"확인한 시트/컬럼: {detail}"
            )
        else:
                # 같은 '이름'이 여러 행에 있어도 더 이상 실수로 생긴 중복으로 취급해
                # 지우지 않습니다(예: 서로 다른 아이디인데 같은 파일을 올려야 하는 계정이
                # 여러 개인 경우). 대신 Secrets 안에서 키가 겹치면 안 되므로(TOML은 같은
                # 테이블을 두 번 선언하면 오류), 이름이 반복되는 행에는 뒤에 _2, _3 ...을
                # 붙여 고유한 "Secrets 키"를 자동으로 만들어줍니다. 실제 업로드할 파일과
                # 매칭할 때 쓰는 값("파일명")은 원래 이름(또는 '파일명' 컬럼이 있으면 그 값)을
                # 그대로 유지하므로, 이름이 같은 계정들은 모두 같은 파일과 매칭됩니다.
                name_seen_count = {}
                rows = []  # (secrets_key, id, pw, file_key, original_name)
                for _, r in df.iterrows():
                    name = str(r[col_map["name"]]).strip()
                    id_ = str(r[col_map["id"]]).strip()
                    pw = str(r[col_map["pw"]]).strip()
                    file_key = ""
                    if "file" in col_map:
                        file_key = str(r[col_map["file"]]).strip()
                    if not file_key:
                        file_key = name
                    if not name or not id_ or not pw:
                        continue
                    name_seen_count[name] = name_seen_count.get(name, 0) + 1
                    n = name_seen_count[name]
                    secrets_key = name if n == 1 else f"{name}_{n}"
                    rows.append((secrets_key, id_, pw, file_key, name))

                dup_names = sorted({name for name, cnt in name_seen_count.items() if cnt > 1})

                # 미리보기는 이름/키/매칭용 파일명만 보여주고 아이디/비밀번호는 마스킹 처리
                preview_df = pd.DataFrame(
                    {
                        "Secrets 키": [k for k, _, _, _, _ in rows],
                        "매칭용 파일명": [fk for _, _, _, fk, _ in rows],
                        "아이디": ["*" * len(i) for _, i, _, _, _ in rows],
                        "비밀번호": ["*" * len(p) for _, _, p, _, _ in rows],
                    }
                )
                st.dataframe(preview_df, use_container_width=True, hide_index=True)

                if dup_names:
                    st.info(
                        "ℹ️ 아래 이름은 엑셀에 여러 번 나와서, Secrets 키를 이름_2, 이름_3 ... "
                        "형태로 자동으로 구분해 **모두 별도 계정으로 등록**했습니다. 매칭용 "
                        "파일명은 원래 이름 그대로이므로, 업로드 시 같은 파일이 이 계정들 "
                        f"전체에 업로드됩니다(의도한 경우가 아니면 확인해주세요): {dup_names}"
                    )

                if not rows:
                    st.error("사용 가능한 행이 없습니다 (이름/아이디/비밀번호 중 빈 값이 있는 행은 제외됩니다).")
                else:
                    toml_lines = []
                    for secrets_key, id_, pw, file_key, _orig_name in rows:
                        esc_key = secrets_key.replace("\\", "\\\\").replace('"', '\\"')
                        esc_id = id_.replace("\\", "\\\\").replace('"', '\\"')
                        esc_pw = pw.replace("\\", "\\\\").replace('"', '\\"')
                        esc_file = file_key.replace("\\", "\\\\").replace('"', '\\"')
                        toml_lines.append(f'[cafe24.accounts."{esc_key}"]')
                        toml_lines.append(f'id = "{esc_id}"')
                        toml_lines.append(f'password = "{esc_pw}"')
                        toml_lines.append(f'file = "{esc_file}"')
                        toml_lines.append("")
                    toml_text = "\n".join(toml_lines)

                    st.download_button(
                        "Secrets에 붙여넣을 텍스트 파일 다운로드 (cafe24_secrets.toml)",
                        data=toml_text.encode("utf-8"),
                        file_name="cafe24_secrets.toml",
                        mime="text/plain",
                    )
                    st.caption(
                        "다운로드한 파일 내용을, Streamlit Cloud → 이 앱 관리 화면 → Settings → "
                        "Secrets 에 기존 내용 아래에 그대로 붙여넣고 저장하시면 됩니다. "
                        "저장 후 앱이 재시작되면 아래 '②' 목록에 계정이 나타납니다."
                    )

st.divider()

# ---------------------------------------------------------------------------
# 2) 업로드할 파일 선택
#    파일명이 계정의 "매칭용 파일명"(기본값 = 계정 이름, '①'에서 '파일명' 컬럼으로
#    직접 지정도 가능)과 같아야 매칭됩니다. 계정 이름은 서로 달라도(같은 이름을 여러
#    계정이 공유해도) 매칭용 파일명이 같으면, 그 계정 전부가 같은 파일로 매칭됩니다
#    (아이디는 다른데 같은 파일을 올려야 하는 경우). 또한 파일명이 "X20260820"처럼
#    뒤에 날짜(YYYYMMDD)가 붙어 매일 바뀌는 경우에는, 그 날짜를 뺀 나머지("X")가
#    매칭용 파일명과 같으면 자동으로 매칭됩니다.
# ---------------------------------------------------------------------------
st.subheader("② 업로드할 판매처 파일 선택")

if not accounts:
    st.info(
        "아직 Secrets에 등록된 카페24 계정이 없습니다. 위 '①' 단계에서 텍스트 파일을 만들어 "
        "Streamlit Secrets에 등록해주세요."
    )
else:
    st.caption(f"현재 Secrets에 등록된 계정 수: {len(accounts)}개 ({', '.join(sorted(accounts.keys()))})")

st.caption(
    "💡 파일명이 'X20260820.xlsx'처럼 뒤에 날짜(8자리, 매일 바뀜)가 붙는 경우에는, "
    "Secrets에 계정의 매칭용 파일명을 날짜 없이 'X'로 등록해두면 날짜가 몇 월 며칠로 "
    "바뀌어도 자동으로 매칭됩니다 (앞부분 'X'만 같으면 됨). 아이디는 다른데 같은 파일을 "
    "올려야 하는 계정이 여러 개라면, '①'에서 이름을 같게 등록해두면 그 계정 전부가 "
    "같은 파일과 자동으로 매칭됩니다."
)

uploaded_files = st.file_uploader(
    "판매처별 결과 파일들을 올려주세요 (파일명이 Secrets에 등록된 계정의 매칭용 파일명과 같아야 "
    "자동 매칭됩니다. 예: 김하늘.xlsx / 날짜가 붙는 파일은 X20260820.xlsx 같은 형태도 가능)",
    type=["xlsx"],
    accept_multiple_files=True,
)

matched = []
unmatched_names = []
pattern_matched_notes = []

if uploaded_files:
    # 매칭용 파일명(accounts[key]["file"]) -> [계정 키, ...] 매핑. 이름이 같은 계정이
    # 여러 개면(아이디는 다르지만 같은 파일을 올려야 하는 경우) 한 파일명 아래에 계정
    # 키가 여러 개 모입니다.
    file_key_to_accounts = {}
    for acc_key, info in accounts.items():
        file_key_to_accounts.setdefault(info["file"], []).append(acc_key)

    # 정규화된(날짜 제거) 파일명 -> {원본 파일명 -> [계정 키, ...]}. 서로 다른 원본
    # 파일명이 같은 정규화 값으로 겹치면(예: "X20260101"과 "X"가 둘 다 등록된 경우)
    # 어느 쪽 계정 목록을 써야 할지 알 수 없으므로 모호한 경우로 취급합니다.
    normalized_index = {}
    for file_key, acc_list in file_key_to_accounts.items():
        norm = _normalize_for_match(file_key)
        normalized_index.setdefault(norm, {})[file_key] = acc_list

    for f in uploaded_files:
        stem = Path(f.name).stem.strip()
        matched_account_keys = []
        via_pattern = False

        if stem in file_key_to_accounts:
            # 파일명이 계정의 매칭용 파일명과 정확히 같은 경우 (기존 방식)
            matched_account_keys = file_key_to_accounts[stem]
        else:
            # 정확히 일치하는 파일명이 없으면, 뒤에 붙은 8자리 날짜를 뗀 값으로 다시 시도
            norm_stem = _normalize_for_match(stem)
            group = normalized_index.get(norm_stem, {})
            if len(group) == 1:
                (only_file_key, acc_list), = group.items()
                matched_account_keys = acc_list
                via_pattern = True
            elif len(group) > 1:
                unmatched_names.append(
                    f"{f.name} (날짜를 뗀 이름 '{norm_stem}'이 서로 다른 파일명 여러 개와 동시에 "
                    f"일치: {sorted(group.keys())} — 매칭용 파일명을 더 구체적으로 등록해주세요)"
                )
                continue

        if matched_account_keys:
            file_bytes = f.read()
            for acc_key in matched_account_keys:
                matched.append(
                    {
                        "name": acc_key,
                        "id": accounts[acc_key]["id"],
                        "password": accounts[acc_key]["password"],
                        "file_bytes": file_bytes,
                        "file_name": f.name,
                    }
                )
            if via_pattern:
                pattern_matched_notes.append(f"{f.name} → {', '.join(matched_account_keys)} (날짜 패턴 매칭)")
        else:
            unmatched_names.append(f.name)

    if matched:
        st.success(f"매칭된 계정 {len(matched)}개: {', '.join(m['name'] for m in matched)}")
    if pattern_matched_notes:
        st.info("📅 날짜 패턴으로 매칭된 파일:\n" + "\n".join(f"- {n}" for n in pattern_matched_notes))

    # 같은 계정으로 파일이 2개 이상 매칭되면(예: X20260820.xlsx, X20260821.xlsx를 실수로
    # 같이 올린 경우) 같은 계정에 연속으로 두 번 업로드가 실행되니 미리 알려줍니다.
    # (이름이 같은 서로 다른 계정 여러 개가 같은 파일 하나에 매칭되는 것은 정상 동작이라
    # 여기 해당하지 않습니다 — 계정 키 자체가 반복될 때만 경고합니다.)
    account_counts = {}
    for m in matched:
        account_counts[m["name"]] = account_counts.get(m["name"], 0) + 1
    dup_account_matches = {name: cnt for name, cnt in account_counts.items() if cnt > 1}
    if dup_account_matches:
        st.warning(
            "⚠️ 아래 계정은 파일이 2개 이상 매칭되어, 같은 계정으로 여러 번 업로드가 "
            f"진행됩니다(의도한 경우가 아니면 파일을 확인해주세요): {dup_account_matches}"
        )

    if unmatched_names:
        st.warning(
            "⚠️ 아래 파일은 매칭용 파일명이 Secrets에 등록된 계정과 일치하지 않아 업로드 "
            f"대상에서 제외됩니다: {unmatched_names}"
        )

st.divider()

# ---------------------------------------------------------------------------
# 3) 실행 옵션 + 업로드 시작
# ---------------------------------------------------------------------------
st.subheader("③ 업로드 실행")

headless = st.checkbox(
    "브라우저 화면 없이 실행(headless)",
    value=True,
    help="Streamlit Cloud 등 서버 환경에서는 반드시 체크된 상태로 사용해야 합니다. "
    "내 컴퓨터에서 직접 실행하며 브라우저 동작을 눈으로 보고 싶을 때만 해제하세요.",
)
debug_pause = st.checkbox(
    "디버그: 업로드 화면 진입 직후 자동 진행 멈추기 (선택자 확인용)",
    value=False,
    disabled=headless,
    help="체크하면 업로드 화면에 들어간 직후(팝업 자동 정리·파일 첨부 전) 자동화가 "
    "멈추고 Playwright 인스펙터 창이 뜹니다. 그 상태에서 브라우저 화면의 팝업이나 "
    "버튼을 개발자도구(F12)로 검사해 정확한 선택자(class, id, xpath 등)를 확인한 뒤, "
    "인스펙터의 '재생(▶ Resume)' 버튼을 눌러야 이어서 진행됩니다. "
    "'브라우저 화면 없이 실행(headless)'이 체크된 상태에서는 볼 화면이 없어 "
    "사용할 수 없습니다(자동으로 비활성화됩니다).",
)
submit_wait_sec = st.number_input(
    "업로드 처리 완료(서버 응답)까지 최대 대기 시간(초)",
    min_value=10, value=180, step=10,
    help="업로드 버튼을 누른 뒤, 실제 '업로드 처리' 요청의 서버 응답이 올 때까지 기다리는 "
    "최대 시간입니다. 이 시간 안에 응답을 받지 못하면 '오류'로 표시되고 진단 정보에 그 "
    "사실이 남습니다. 상품 수가 많아 서버 처리가 평소보다 오래 걸리는 계정이 있다면 이 "
    "값을 늘려보세요(기본 180초).",
)
post_delay = st.number_input(
    "계정별 업로드 버튼 클릭 후 최종 대기 시간(초)",
    min_value=1, value=5, step=1,
    help="위 '업로드 처리 완료' 대기가 끝난(또는 응답을 받은) 뒤, 화면 반영 등을 위해 "
    "추가로 더 기다리는 시간입니다. 보통은 기본값 그대로 두셔도 됩니다.",
)

start_clicked = st.button(
    "선택된 계정 업로드 시작",
    type="primary",
    disabled=(len(matched) == 0),
    use_container_width=True,
)

if start_clicked:
    ok, err = up.ensure_chromium_installed(status_cb=lambda msg: st.info(msg))
    if not ok:
        st.error(f"브라우저 엔진 준비 실패: {err}")
        st.stop()

    progress_bar = st.progress(0)
    status_area = st.empty()

    def progress_cb(i, total, result):
        progress_bar.progress(i / total)
        status_area.write(f"[{i}/{total}] {result.name} → {result.status}")

    debug_pause_active = debug_pause and not headless
    spinner_msg = f"{len(matched)}개 계정 업로드 진행 중... (계정당 처리 시간이 있어 시간이 걸릴 수 있습니다)"
    if debug_pause_active:
        spinner_msg += " — 디버그 모드: 업로드 화면 진입 시 Playwright 인스펙터에서 '재생'을 누를 때까지 멈춰 있습니다."

    with st.spinner(spinner_msg):
        results = up.run_batch_upload(
            matched, headless=headless, post_delay_sec=int(post_delay), progress_cb=progress_cb,
            debug_pause=debug_pause_active, submit_wait_sec=int(submit_wait_sec),
        )

    st.session_state["cafe24_last_results"] = results

# ---------------------------------------------------------------------------
# 4) 결과 표시
# ---------------------------------------------------------------------------
results = st.session_state.get("cafe24_last_results")
if results:
    st.divider()
    st.subheader("④ 업로드 결과")

    status_label = {
        "success": "✅ 성공",
        "fail": "❌ 업로드 실패",
        "login_fail": "❌ 로그인 실패",
        "error": "⚠️ 오류",
    }

    result_df = pd.DataFrame(
        {
            "이름": [r.name for r in results],
            "상태": [status_label.get(r.status, r.status) for r in results],
            "메시지": [r.message for r in results],
            "진단": [getattr(r, "diagnosis", "") for r in results],
        }
    )
    st.dataframe(result_df, use_container_width=True, hide_index=True)

    failed = [r for r in results if r.status != "success"]
    if failed:
        st.error(
            "재시도 없이 아래 계정은 실패로 처리됐습니다 (재시도하지 않았습니다 — 요청하신 대로): "
            + ", ".join(r.name for r in failed)
        )
    else:
        st.success("모든 계정이 성공으로 판별됐습니다. (자동 판별이므로 스크린샷으로 실제 반영 여부를 꼭 확인해주세요)")

    with st.expander("🔍 진단 정보 (버튼 클릭이 실제로 서버에 전달됐는지 증거로 확인)", expanded=True):
        st.caption(
            "💡 '버튼을 클릭했다'는 것과 '클릭이 실제로 서버에 요청을 보냈다'는 것은 다를 수 "
            "있습니다. 아래는 추측이 아니라 실제로 관찰한 증거(클릭 후 발생한 네트워크 요청/응답, "
            "버튼 HTML 변화, 콘솔 로그)를 바탕으로 자동 판정한 결과입니다. 문제가 계속되면 이 "
            "내용을 그대로 복사해서 알려주시면 훨씬 정확하게 원인을 찾을 수 있습니다."
        )
        for r in results:
            diag = getattr(r, "diagnosis", "")
            st.write(f"**{r.name}**")
            if diag.startswith("🔴"):
                st.error(diag)
            elif diag.startswith("🟠"):
                st.warning(diag)
            elif diag.startswith("🟢"):
                st.success(diag)
            elif diag:
                st.caption(diag)
            else:
                st.caption("진단 정보 없음 (로그인 실패 등 버튼 클릭 이전 단계에서 종료됨)")

            with st.expander(f"{r.name} — 상세 로그 (네트워크 / 콘솔 / 버튼 HTML)", expanded=False):
                btn_before = getattr(r, "button_html_before", None)
                btn_after = getattr(r, "button_html_after", None)
                st.write("**업로드 버튼 HTML — 클릭 전:**")
                st.code(btn_before or "(캡처 안 됨)", language="html")
                st.write("**업로드 버튼 HTML — 클릭 후:**")
                st.code(btn_after or "(캡처 안 됨)", language="html")
                if btn_before and btn_after and btn_before == btn_after:
                    st.caption("⚠️ 클릭 전후 버튼 HTML이 완전히 동일합니다 (변화 없음).")

                net_log = getattr(r, "network_log", None) or []
                st.write(f"**네트워크 요청/응답 로그** ({len(net_log)}건, xhr/fetch/document만):")
                if net_log:
                    net_lines = []
                    for e in net_log:
                        if e.get("kind") == "요청":
                            net_lines.append(f"[{e['phase']}] 요청 {e.get('method','')} {e.get('url','')}")
                        else:
                            net_lines.append(f"[{e['phase']}] 응답 {e.get('status','')} {e.get('url','')}")
                    st.code("\n".join(net_lines), language="text")
                else:
                    st.caption("기록된 요청 없음")

                console_log = getattr(r, "console_log", None) or []
                st.write(f"**브라우저 콘솔 로그** ({len(console_log)}건):")
                if console_log:
                    st.code("\n".join(console_log), language="text")
                else:
                    st.caption("기록된 콘솔 로그 없음")
            st.divider()

    with st.expander("계정별 스크린샷 보기 (클릭 전 / 클릭 직후 / 최종 3장 비교)", expanded=False):
        st.caption(
            "💡 업로드 버튼을 눌렀는데도 실제 반영이 안 되는 문제를 진단하기 위해, "
            "버튼을 누르기 직전 / 누른 직후 / 대기 후 최종 화면 3장을 함께 보여줍니다. "
            "'클릭 직전'과 '클릭 직후' 화면이 서로 똑같아 보인다면(아무 변화가 없다면), "
            "버튼 클릭이 실제로는 아무 동작도 하지 않았을 가능성이 있다는 뜻이니 이 3장을 "
            "함께 캡처해서 알려주시면 추가로 진단하는 데 큰 도움이 됩니다."
        )
        for r in results:
            st.write(f"**{r.name}** — {status_label.get(r.status, r.status)}")
            cols = st.columns(3)
            with cols[0]:
                st.caption("① 클릭 직전")
                shot_before = getattr(r, "screenshot_before_click", None)
                if shot_before:
                    st.image(shot_before, use_container_width=True)
                else:
                    st.caption("없음")
            with cols[1]:
                st.caption("② 클릭 직후")
                shot_after = getattr(r, "screenshot_after_click", None)
                if shot_after:
                    st.image(shot_after, use_container_width=True)
                else:
                    st.caption("없음")
            with cols[2]:
                st.caption("③ 최종(대기 후)")
                if r.screenshot:
                    st.image(r.screenshot, use_container_width=True)
                else:
                    st.caption("없음")
            st.divider()

    shots = []
    for r in results:
        if getattr(r, "screenshot_before_click", None):
            shots.append((f"{r.name}_1_클릭직전", r.screenshot_before_click))
        if getattr(r, "screenshot_after_click", None):
            shots.append((f"{r.name}_2_클릭직후", r.screenshot_after_click))
        if r.screenshot:
            shots.append((f"{r.name}_3_최종", r.screenshot))
    if shots:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, shot in shots:
                zf.writestr(f"{name}.png", shot)
        st.download_button(
            "계정별 스크린샷 전체 ZIP 다운로드",
            data=buf.getvalue(),
            file_name="cafe24_upload_screenshots.zip",
            mime="application/zip",
        )
