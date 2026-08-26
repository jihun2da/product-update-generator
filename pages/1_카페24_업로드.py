"""
카페24 자동 업로드 (별도 탭 / 독립 기능)
========================================
기존 '신상품 업데이트 파일 생성기'(app.py) 로직과는 완전히 분리된 페이지입니다.
이 페이지의 코드가 어떻게 바뀌어도 app.py / pipeline.py 는 전혀 영향을 받지 않습니다.

⚠️ 아직 실제 계정으로 로그인→업로드까지 이어지는 전체 흐름을 사용자가 직접 테스트하지
   않았습니다. 반드시 계정 1개로 먼저 테스트해보시고, 문제가 없으면 전체 계정으로
   진행해주세요. (자세한 내용은 화면 상단 안내 문구를 확인해주세요.)
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

st.warning(
    "⚠️ 이 기능은 아직 실제 카페24 계정으로 '로그인 → 업로드 화면 이동 → 파일 업로드 → "
    "성공/실패 판별'까지 전체 흐름을 직접 테스트하지 않았습니다.\n\n"
    "- 로그인 아이디/비번 입력창, 로그인 버튼 위치는 실제 사이트에서 구조적으로 확인했습니다.\n"
    "- 업로드 성공/실패 판별은 화면에 뜨는 문구를 기준으로 자동 판단하는데, 정확한 문구를 "
    "실제로 보지 못한 상태라 오탐(성공인데 실패로 표시, 또는 그 반대)이 있을 수 있습니다.\n"
    "- **반드시 계정 1개로 먼저 테스트**해서 계정별 스크린샷으로 실제 반영 여부를 확인해주시고, "
    "문제 없으면 전체 계정으로 진행해주세요.\n"
    "- 이 앱(app.py)의 다른 기능(파일 생성 등)에는 이 기능이 전혀 영향을 주지 않습니다."
)


# ---------------------------------------------------------------------------
# 1) Secrets에서 계정 목록 불러오기
# ---------------------------------------------------------------------------
def load_accounts():
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
            result[name] = {"id": info["id"], "password": info["password"]}
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
                # 이름 기준으로 마지막 값이 이기도록 먼저 dict로 정리합니다.
                # (TOML은 같은 테이블을 두 번 선언하는 것 자체가 오류라서, 중복 이름을
                #  그대로 두면 Secrets 텍스트 전체가 파싱 실패로 깨집니다 — 반드시 이름당
                #  1개 블록만 생성해야 합니다.)
                order = []
                seen = {}
                dup_names = []
                for _, r in df.iterrows():
                    name = str(r[col_map["name"]]).strip()
                    id_ = str(r[col_map["id"]]).strip()
                    pw = str(r[col_map["pw"]]).strip()
                    if not name or not id_ or not pw:
                        continue
                    if name in seen:
                        dup_names.append(name)
                    else:
                        order.append(name)
                    seen[name] = (id_, pw)

                rows = [(name, seen[name][0], seen[name][1]) for name in order]

                # 미리보기는 이름만 보여주고 아이디/비밀번호는 마스킹 처리
                preview_df = pd.DataFrame(
                    {
                        "이름": [n for n, _, _ in rows],
                        "아이디": ["*" * len(i) for _, i, _ in rows],
                        "비밀번호": ["*" * len(p) for _, _, p in rows],
                    }
                )
                st.dataframe(preview_df, use_container_width=True, hide_index=True)

                if dup_names:
                    st.warning(
                        "⚠️ 아래 이름이 엑셀에 중복으로 있습니다. Secrets는 이름별로 1개씩만 "
                        "저장되므로, 같은 이름 중 **엑셀에서 가장 마지막에 나온 행의 아이디/비밀번호만** "
                        f"사용됩니다: {sorted(set(dup_names))}"
                    )

                if not rows:
                    st.error("사용 가능한 행이 없습니다 (이름/아이디/비밀번호 중 빈 값이 있는 행은 제외됩니다).")
                else:
                    toml_lines = []
                    for name, id_, pw in rows:
                        esc_name = name.replace("\\", "\\\\").replace('"', '\\"')
                        esc_id = id_.replace("\\", "\\\\").replace('"', '\\"')
                        esc_pw = pw.replace("\\", "\\\\").replace('"', '\\"')
                        toml_lines.append(f'[cafe24.accounts."{esc_name}"]')
                        toml_lines.append(f'id = "{esc_id}"')
                        toml_lines.append(f'password = "{esc_pw}"')
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
# 2) 업로드할 파일 선택 (파일명 = 계정 이름과 일치해야 매칭됨. 단, 파일명이
#    "X20260820"처럼 뒤에 날짜(YYYYMMDD)가 붙어 매일 바뀌는 경우에는, 그 날짜를
#    뺀 나머지("X")가 계정 이름과 같으면 자동으로 매칭됩니다.)
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
    "Secrets에 계정 이름을 날짜 없이 'X'로 등록해두면 날짜가 몇 월 며칠로 바뀌어도 "
    "자동으로 매칭됩니다 (앞부분 'X'만 같으면 됨)."
)

uploaded_files = st.file_uploader(
    "판매처별 결과 파일들을 올려주세요 (파일명이 Secrets에 등록된 계정 이름과 같아야 자동 매칭됩니다. "
    "예: 김하늘.xlsx / 날짜가 붙는 파일은 X20260820.xlsx 같은 형태도 가능)",
    type=["xlsx"],
    accept_multiple_files=True,
)

matched = []
unmatched_names = []
pattern_matched_notes = []

if uploaded_files:
    # 정규화된(날짜 제거) 이름 -> [계정 이름, ...] 매핑. 날짜가 없는 일반 계정 이름은
    # 정규화해도 그대로이므로, 이 맵에는 모든 계정이 자기 자신의 정규화 키로도 들어갑니다.
    normalized_accounts = {}
    for acc_name in accounts:
        key = _normalize_for_match(acc_name)
        normalized_accounts.setdefault(key, []).append(acc_name)

    for f in uploaded_files:
        stem = Path(f.name).stem.strip()
        matched_account_name = None

        if stem in accounts:
            # 파일명이 계정 이름과 정확히 같은 경우 (기존 방식)
            matched_account_name = stem
        else:
            # 정확히 일치하는 계정이 없으면, 뒤에 붙은 8자리 날짜를 뗀 값으로 다시 시도
            norm_stem = _normalize_for_match(stem)
            candidates = normalized_accounts.get(norm_stem, [])
            if len(candidates) == 1:
                matched_account_name = candidates[0]
                if norm_stem != stem:
                    pattern_matched_notes.append(f"{f.name} → '{matched_account_name}' 계정 (날짜 패턴 매칭)")
            elif len(candidates) > 1:
                unmatched_names.append(
                    f"{f.name} (날짜를 뗀 이름 '{norm_stem}'이 여러 계정과 동시에 일치: {candidates} — "
                    "계정 이름을 더 구체적으로 등록해주세요)"
                )
                continue

        if matched_account_name:
            matched.append(
                {
                    "name": matched_account_name,
                    "id": accounts[matched_account_name]["id"],
                    "password": accounts[matched_account_name]["password"],
                    "file_bytes": f.read(),
                    "file_name": f.name,
                }
            )
        else:
            unmatched_names.append(f.name)

    if matched:
        st.success(f"매칭된 계정 {len(matched)}개: {', '.join(m['name'] for m in matched)}")
    if pattern_matched_notes:
        st.info("📅 날짜 패턴으로 매칭된 파일:\n" + "\n".join(f"- {n}" for n in pattern_matched_notes))

    # 같은 계정으로 파일이 2개 이상 매칭되면(예: X20260820.xlsx, X20260821.xlsx를 실수로
    # 같이 올린 경우) 같은 계정에 연속으로 두 번 업로드가 실행되니 미리 알려줍니다.
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
            "⚠️ 아래 파일은 이름이 Secrets에 등록된 계정과 일치하지 않아 업로드 대상에서 "
            f"제외됩니다: {unmatched_names}"
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
post_delay = st.number_input("계정별 업로드 버튼 클릭 후 대기 시간(초)", min_value=1, value=3, step=1)

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

    with st.spinner(f"{len(matched)}개 계정 업로드 진행 중... (계정당 처리 시간이 있어 시간이 걸릴 수 있습니다)"):
        results = up.run_batch_upload(
            matched, headless=headless, post_delay_sec=int(post_delay), progress_cb=progress_cb
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

    with st.expander("계정별 스크린샷 보기", expanded=False):
        for r in results:
            st.write(f"**{r.name}** — {status_label.get(r.status, r.status)}")
            if r.screenshot:
                st.image(r.screenshot, use_container_width=True)
            else:
                st.caption("스크린샷 없음")
            st.divider()

    shots = [(r.name, r.screenshot) for r in results if r.screenshot]
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
