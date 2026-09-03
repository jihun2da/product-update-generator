"""
카페24 자동 업로드 (별도 탭 / 독립 기능)
========================================
파일 생성 로직('신상품 업데이트 파일 생성기', app.py의 pipeline.py 호출 부분)과는
여전히 완전히 분리되어 있습니다 — 이 페이지의 코드가 어떻게 바뀌어도 pipeline.py의
파일 생성 결과에는 전혀 영향을 주지 않습니다.

다만 "계정 매칭 / 업로드 실행 옵션 / 결과 화면" 부분은 app.py의 "생성된 파일 바로
업로드" 섹션과 완전히 동일한 코드(cafe24_upload_ui.py)를 공유합니다 — 두 화면에서
매칭 규칙이나 결과 표시가 서로 다르게 관리되다가 어긋나는 일이 없도록 하기 위해서입니다.
이 페이지 고유의 부분(① Secrets 등록 도우미, ② 파일 업로더 위젯)은 그대로 유지됩니다.

실제 카페24 계정으로 로그인 → 업로드 화면 이동 → 파일 업로드 → 성공/실패 판별까지
전체 흐름을 사용자가 직접 테스트했고, 정상 동작하는 것을 확인했습니다(2026-08-26).
"""

import pandas as pd
import streamlit as st

import cafe24_uploader as up
import cafe24_upload_ui as up_ui

# 계정 로딩 / 파일-계정 매칭(날짜 접미사 정규화 포함) / 실행 옵션 / 결과 표시는
# app.py("생성된 파일 바로 업로드" 섹션)와 완전히 동일한 코드를 쓰기 위해
# cafe24_upload_ui.py 공용 모듈로 옮겼습니다. 이 페이지는 그 모듈을 그대로 가져다 씁니다.

st.set_page_config(page_title="카페24 자동 업로드", layout="wide")

st.title("카페24 자동 업로드 (별도 기능)")

# 코드를 업데이트한 뒤 "정말 최신 코드로 테스트한 게 맞는지" 헷갈리는 경우가 많아서
# 추가한 버전 표시입니다. 로컬 파일을 새로 받은 뒤 앱을 재시작했는데도 이 값이 안
# 바뀌어 있다면, 아직 예전 코드로 실행 중이라는 뜻입니다.
st.caption(f"🔧 코드 버전: `{up.MODULE_VERSION}`")

st.success(
    "✅ 실제 카페24 계정으로 '로그인 → 업로드 화면 이동 → 파일 업로드 → 업로드 버튼 클릭 → 서버 처리 "
    "완료'까지 전체 흐름이 정상 동작하는 것을 확인했습니다. 클릭 후에는 실제 '업로드 처리' 요청과 "
    "그 응답을 직접 확인해서 성공/오류를 판별합니다(아래 '③'의 대기 시간 설정 참고).\n\n"
    "- 실패(오류)로 판별된 계정은 재시도하지 않고 끝까지 진행한 뒤, 모아서 알려드립니다.\n"
    "- 계정별 스크린샷과 진단 정보는 '④ 업로드 결과'에서 계속 확인할 수 있으니, 새로운 계정이나 "
    "파일을 추가하실 때는 처음 한 번씩 결과를 확인해보시는 걸 권장드립니다.\n"
    "- 이 앱(app.py)의 다른 기능(파일 생성 등)에는 이 기능이 전혀 영향을 주지 않습니다."
)


# ---------------------------------------------------------------------------
# 1) Secrets에서 계정 목록 불러오기 (cafe24_upload_ui.load_accounts()로 이동)
# ---------------------------------------------------------------------------
accounts = up_ui.load_accounts()

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

files = {f.name: f.read() for f in uploaded_files} if uploaded_files else {}
matched, unmatched_names, pattern_matched_notes, dup_account_matches = up_ui.match_files_to_accounts(files, accounts)
up_ui.render_match_preview(matched, unmatched_names, pattern_matched_notes, dup_account_matches)

st.divider()

# ---------------------------------------------------------------------------
# 3) 실행 옵션 + 업로드 시작, 4) 결과 표시
#    (app.py "생성된 파일 바로 업로드" 섹션과 완전히 동일한 코드를 공유합니다.)
# ---------------------------------------------------------------------------
st.subheader("③ 업로드 실행")
up_ui.render_execution_and_results(matched, key_prefix="cafe24page_")
