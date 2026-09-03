"""
카페24 업로드 화면에서 공통으로 쓰는 로직(계정 로딩 / 파일-계정 매칭 / 실행 옵션 / 결과 표시).

pages/1_카페24_업로드.py 와 app.py("생성된 파일 바로 업로드" 섹션) 양쪽에서 이 모듈의
함수를 그대로 가져다 씁니다 — 매칭 규칙이나 결과 화면이 두 곳에서 따로 관리되다가
서로 어긋나는 일이 없도록 하기 위해서입니다.

실제 브라우저 자동화 로직(cafe24_uploader.py, Playwright로 로그인·업로드하는 부분)에는
전혀 손대지 않았습니다 — 이 모듈은 그 위에 얹는 Streamlit 화면 조립 로직만 담당합니다.
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
_DATE_SUFFIX_RE = re.compile(r"\d{8}$")


def normalize_for_match(name: str) -> str:
    return _DATE_SUFFIX_RE.sub("", name.strip())


def load_accounts() -> dict:
    """Secrets에서 계정 목록을 읽어옵니다.

    각 계정은 Secrets 안에서 고유한 키(예: "모나마켓", 또는 같은 이름이 여러 개면
    "모나마켓_2", "모나마켓_3"...)를 가져야 하지만, 실제로 업로드할 때 어떤 파일과
    매칭시킬지는 별도의 "file" 값으로 관리합니다(생략하면 계정 키 자체를 사용).
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


def match_files_to_accounts(files: dict, accounts: dict):
    """files({파일명(확장자 포함): bytes, ...})를 accounts(load_accounts()의 반환값)와 매칭합니다.

    반환: (matched, unmatched_names, pattern_matched_notes, dup_account_matches)
      matched: [{'name':, 'id':, 'password':, 'file_bytes':, 'file_name':}, ...]
      unmatched_names: [str, ...] (매칭 실패한 파일명/사유)
      pattern_matched_notes: [str, ...] (날짜 패턴으로 매칭된 경우 안내 문구)
      dup_account_matches: {계정키: 매칭된 파일 개수, ...} (같은 계정에 파일이 2개 이상 매칭된 경우)
    """
    matched = []
    unmatched_names = []
    pattern_matched_notes = []

    if not files:
        return matched, unmatched_names, pattern_matched_notes, {}

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
        norm = normalize_for_match(file_key)
        normalized_index.setdefault(norm, {})[file_key] = acc_list

    for fname, fbytes in files.items():
        stem = Path(fname).stem.strip()
        matched_account_keys = []
        via_pattern = False

        if stem in file_key_to_accounts:
            # 파일명이 계정의 매칭용 파일명과 정확히 같은 경우 (기존 방식)
            matched_account_keys = file_key_to_accounts[stem]
        else:
            # 정확히 일치하는 파일명이 없으면, 뒤에 붙은 8자리 날짜를 뗀 값으로 다시 시도
            norm_stem = normalize_for_match(stem)
            group = normalized_index.get(norm_stem, {})
            if len(group) == 1:
                (only_file_key, acc_list), = group.items()
                matched_account_keys = acc_list
                via_pattern = True
            elif len(group) > 1:
                unmatched_names.append(
                    f"{fname} (날짜를 뗀 이름 '{norm_stem}'이 서로 다른 파일명 여러 개와 동시에 "
                    f"일치: {sorted(group.keys())} — 매칭용 파일명을 더 구체적으로 등록해주세요)"
                )
                continue

        if matched_account_keys:
            for acc_key in matched_account_keys:
                matched.append(
                    {
                        "name": acc_key,
                        "id": accounts[acc_key]["id"],
                        "password": accounts[acc_key]["password"],
                        "file_bytes": fbytes,
                        "file_name": fname,
                    }
                )
            if via_pattern:
                pattern_matched_notes.append(f"{fname} → {', '.join(matched_account_keys)} (날짜 패턴 매칭)")
        else:
            unmatched_names.append(fname)

    # 같은 계정으로 파일이 2개 이상 매칭되면(예: X20260820.xlsx, X20260821.xlsx를 실수로
    # 같이 선택한 경우) 같은 계정에 연속으로 두 번 업로드가 실행되니 미리 알려줍니다.
    account_counts = {}
    for m in matched:
        account_counts[m["name"]] = account_counts.get(m["name"], 0) + 1
    dup_account_matches = {name: cnt for name, cnt in account_counts.items() if cnt > 1}

    return matched, unmatched_names, pattern_matched_notes, dup_account_matches


def render_match_preview(matched, unmatched_names, pattern_matched_notes, dup_account_matches):
    """match_files_to_accounts()의 결과를 화면에 안내 메시지로 보여줍니다."""
    if matched:
        st.success(f"매칭된 계정 {len(matched)}개: {', '.join(m['name'] for m in matched)}")
    if pattern_matched_notes:
        st.info("📅 날짜 패턴으로 매칭된 파일:\n" + "\n".join(f"- {n}" for n in pattern_matched_notes))
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


def render_execution_and_results(matched: list, key_prefix: str):
    """③ 실행 옵션 + 업로드 시작 버튼, ④ 업로드 결과를 렌더링합니다.

    key_prefix: 이 함수를 여러 화면(카페24 업로드 페이지 / app.py 생성 직후 업로드)에서
        동시에 써도 Streamlit 위젯 key와 세션 상태 키가 서로 겹치지 않게 하는 접두어입니다.
    """
    results_state_key = f"{key_prefix}cafe24_last_results"

    headless = st.checkbox(
        "브라우저 화면 없이 실행(headless)",
        value=True,
        key=f"{key_prefix}headless",
        help="Streamlit Cloud 등 서버 환경에서는 반드시 체크된 상태로 사용해야 합니다. "
        "내 컴퓨터에서 직접 실행하며 브라우저 동작을 눈으로 보고 싶을 때만 해제하세요.",
    )
    debug_pause = st.checkbox(
        "디버그: 업로드 화면 진입 직후 자동 진행 멈추기 (선택자 확인용)",
        value=False,
        disabled=headless,
        key=f"{key_prefix}debug_pause",
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
        key=f"{key_prefix}submit_wait_sec",
        help="업로드 버튼을 누른 뒤, 실제 '업로드 처리' 요청의 서버 응답이 올 때까지 기다리는 "
        "최대 시간입니다. 이 시간 안에 응답을 받지 못하면 '오류'로 표시되고 진단 정보에 그 "
        "사실이 남습니다. 상품 수가 많아 서버 처리가 평소보다 오래 걸리는 계정이 있다면 이 "
        "값을 늘려보세요(기본 180초).",
    )
    post_delay = st.number_input(
        "계정별 업로드 버튼 클릭 후 최종 대기 시간(초)",
        min_value=1, value=5, step=1,
        key=f"{key_prefix}post_delay",
        help="위 '업로드 처리 완료' 대기가 끝난(또는 응답을 받은) 뒤, 화면 반영 등을 위해 "
        "추가로 더 기다리는 시간입니다. 보통은 기본값 그대로 두셔도 됩니다.",
    )

    start_clicked = st.button(
        "선택된 계정 업로드 시작",
        type="primary",
        disabled=(len(matched) == 0),
        use_container_width=True,
        key=f"{key_prefix}start_upload_btn",
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

        st.session_state[results_state_key] = results

    results = st.session_state.get(results_state_key)
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
                            elif e.get("kind") == "요청실패":
                                net_lines.append(f"[{e['phase']}] ❌ 요청 실패 {e.get('url','')} ({e.get('error','')})")
                            else:
                                net_lines.append(f"[{e['phase']}] 응답 {e.get('status','')} {e.get('url','')}")
                        st.code("\n".join(net_lines), language="text")
                    else:
                        st.caption("기록된 요청 없음")

                    popup_log = getattr(r, "popup_page_log", None) or []
                    st.write(f"**팝업 페이지 처리 로그** ({len(popup_log)}건):")
                    st.caption(
                        "💡 '엑셀 업로드' 버튼을 누른 뒤 새로 열리는 팝업 페이지(예: \"해당 파일을 "
                        "업로드 하시겠습니까?\" 확인창이 네이티브 대화상자가 아니라 진짜 팝업 페이지로 "
                        "뜨는 경우)마다, 그 안에서 '확인' 버튼을 자동으로 클릭했는지 / 못 찾아서 그냥 "
                        "닫았는지를 기록합니다. '확인 버튼을 찾지 못함'이 반복된다면 이 페이지의 실제 "
                        "버튼 문구가 '확인'이 아닐 수 있으니 알려주세요."
                    )
                    if popup_log:
                        st.code("\n".join(popup_log), language="text")
                    else:
                        st.caption("기록된 팝업 페이지 없음 (네이티브 확인창만 있었거나, 팝업 자체가 없었을 수 있습니다)")

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
                key=f"{key_prefix}download_screenshots_zip",
            )
