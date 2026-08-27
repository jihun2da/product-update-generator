"""
카페24 관리자(https://eclogin.cafe24.com/Shop/)에 로그인해서
상품 엑셀 일괄등록/수정(ProductExcelManage) 화면에 파일을 자동으로 업로드하는 모듈.

중요 — 반드시 읽어주세요
------------------------
- 아이디/비밀번호는 반드시 Streamlit Secrets(st.secrets)로만 전달해야 합니다.
  이 파일이나 저장소 어디에도 실제 아이디/비밀번호를 하드코딩하거나 커밋하면 안 됩니다.
- Claude(AI 어시스턴트)는 실제 계정으로 로그인하는 행위 자체를 직접 수행할 수 없습니다
  (안전 정책상, 사용자가 허락해도 마찬가지입니다). 로그인 페이지의 아이디/비밀번호
  입력창, 로그인 버튼 셀렉터, 더미(가짜) 값 제출 시 오류 문구 등은 구조적으로
  확인했고, 이후 사용자가 실제 계정으로 "로그인 → 업로드 화면 이동 → 파일 첨부 →
  업로드 성공/실패 판별"까지 전체 흐름을 직접 테스트해서 정상 동작하는 것을
  확인했습니다(2026-08-26). 계정별로 저장되는 스크린샷은 새로운 계정/파일을
  추가할 때 참고용으로 계속 활용하시면 됩니다.
"""

import asyncio
import os
import sys
import subprocess
import tempfile
import time
from dataclasses import dataclass
from typing import Optional, List, Callable

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

LOGIN_URL = "https://eclogin.cafe24.com/Shop/"
ID_XPATH = 'xpath=//*[@id="mall_id"]'
PW_XPATH = 'xpath=//*[@id="userpasswd"]'
LOGIN_BTN_XPATH = 'xpath=//*[@id="frm_user"]/div/div[2]/button'
UPLOAD_PATH = "/disp/admin/shop1/product/ProductExcelManage"
UPLOAD_BTN_XPATH = 'xpath=//*[@id="dropZone"]/div/div[2]/button'

# 실제 계정으로 진단 로그를 받아 확인한 결과, "엑셀 업로드" 버튼을 클릭하면 이 URL
# 조각을 포함한 요청(POST .../exec/admin/product/ProductExcelSet)이 실제로 발생하는
# 것이 확인됐습니다(2026-08-27, 클릭 버튼의 class="btnSubmit", data-get-id="dropZone",
# 링크 텍스트 "엑셀 업로드"까지 애널리틱스 이벤트로 함께 확인). 즉 지금까지 고친
# 클릭 자체는 정상적으로 서버에 전달되고 있습니다 — 남은 문제는 "이 요청의 응답을
# 받기 전에 브라우저를 닫아버리는 것"입니다. 이 요청의 응답이 실제로 올 때까지
# 직접 기다리기 위한 기준값으로 사용합니다.
UPLOAD_SUBMIT_URL_HINT = "ProductExcelSet"

# 업로드 결과 화면에서 "실패"로 볼 수 있는 문구 후보 (실제 계정으로 사용 테스트해서
# 정상 동작하는 것을 확인했습니다 — 이후 문구가 다른 경우가 생기면 알려주시면 바로
# 고치겠습니다)
FAIL_KEYWORDS = ["실패", "오류가 발생", "다시 시도", "에러가 발생", "업로드에 실패", "형식이 올바르지"]
LOGIN_FAIL_TEXT = "아이디 또는 비밀번호를 확인해주세요"


@dataclass
class AccountResult:
    name: str
    status: str  # 'success' | 'fail' | 'login_fail' | 'error'
    message: str = ""
    screenshot: Optional[bytes] = None
    # 업로드 버튼을 누르기 직전 / 직후 화면도 함께 남겨서, "버튼 클릭이 실제로
    # 무언가를 했는지"를 화면 3장(클릭 전 / 클릭 직후 / 최종)으로 비교 확인할 수
    # 있게 합니다. 실패 원인 진단용이며, 못 찍었을 때는 None입니다.
    screenshot_before_click: Optional[bytes] = None
    screenshot_after_click: Optional[bytes] = None
    # "버튼을 눌렀을 때 실제로 무슨 일이 일어났는지"를 추측이 아니라 직접 관찰한
    # 증거로 확인하기 위한 진단 정보입니다. 계속 같은 문제가 재현될 때, 이 정보를
    # 보면 (1) 클릭 후 서버로 요청이 실제로 발생했는지 (2) 발생했다면 성공/실패
    # 응답이었는지 (3) 버튼의 HTML(문구/활성화 상태)이 클릭 전후로 바뀌었는지 (4)
    # 페이지에 자바스크립트 오류가 있었는지를 바로 알 수 있습니다.
    diagnosis: str = ""
    network_log: Optional[List[dict]] = None
    console_log: Optional[List[str]] = None
    button_html_before: Optional[str] = None
    button_html_after: Optional[str] = None
    # 실제로 관찰된, "엑셀 업로드" 버튼이 호출하는 서버 요청(UPLOAD_SUBMIT_URL_HINT
    # 포함)에 대한 구체적인 판정 결과입니다. submit_request_sent=True인데
    # submit_response_status가 None이면 "요청은 보냈지만 응답을 못 받음"이라는
    # 뜻으로, 지금까지 반복된 문제의 가장 유력한 원인입니다.
    submit_request_sent: bool = False
    submit_response_status: Optional[int] = None


def ensure_chromium_installed(status_cb: Optional[Callable[[str], None]] = None):
    """Playwright의 Chromium 브라우저 바이너리가 없으면 설치합니다(최초 1회, 1~2분 소요될 수 있음).
    이미 설치돼 있으면 아무 것도 하지 않습니다."""
    marker = os.path.expanduser("~/.cache/ms-playwright/.chromium_ready")
    if os.path.exists(marker):
        return True, None

    if status_cb:
        status_cb("브라우저 엔진(Chromium)을 최초 설치하는 중입니다. 1~2분 정도 걸릴 수 있습니다...")

    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=True, capture_output=True, text=True, timeout=600,
        )
    except Exception as e:
        return False, f"Chromium 설치 실패: {e}"

    os.makedirs(os.path.dirname(marker), exist_ok=True)
    with open(marker, "w") as f:
        f.write("ok")
    return True, None


def _close_popups(context, page):
    """페이지 위에 남아있는 팝업(새 탭 또는 인페이지 모달)을 최대한 정리합니다.
    실패해도 전체 흐름에 영향 주지 않도록 예외를 삼킵니다.

    로그인 직후뿐 아니라, 업로드 화면 진입 직후 / 파일 첨부 후 업로드 버튼을 누르기
    직전에도 호출합니다 — 예를 들어 "전용 엑셀양식 다운로드" 같은 안내 팝업이 업로드
    화면에서 새로 뜨면서 업로드 버튼을 가려 클릭이 안 먹히는 경우가 실제로 있었기
    때문입니다(실사용 테스트 중 발견 — 클릭이 막혀도 조용히 넘어가면 실제로는 아무
    것도 업로드되지 않았는데 "성공"으로 잘못 표시되는 문제가 있어, 이제는 클릭 자체가
    끝까지 실패하면 성공으로 처리하지 않고 오류로 보고합니다. 아래 upload_one_account
    참고)."""
    time.sleep(1)
    for p in list(context.pages):
        if p is not page:
            try:
                p.close()
            except Exception:
                pass

    close_selectors = [
        # "상품등록용 엑셀 다운로드" 안내 walkthrough(Step1/3 등) 팝업의 닫기(x) 버튼 —
        # 실제 화면에서 사용자가 직접 확인한 정확한 경로입니다(2026-08-27 최초 확인 후,
        # 실사용 중 안 닫히는 것이 확인되어 2026-08-27 같은 날 아래 경로로 재확인/수정).
        # 다른 선택자보다 먼저 시도합니다.
        'xpath=//*[@id="QA_upload2"]/div[2]/div[2]/div[1]/div/button',
        "button:has-text('확인')",
        "button:has-text('닫기')",
        "[aria-label='닫기']",
        "[aria-label='close']",
        ".btn_close",
        ".pop_close",
        ".close",
    ]
    for sel in close_selectors:
        try:
            loc = page.locator(sel).first
            if loc.count() > 0 and loc.is_visible(timeout=500):
                loc.click(timeout=1000)
                time.sleep(0.3)
        except Exception:
            continue

    try:
        page.keyboard.press("Escape")
    except Exception:
        pass


def _visible_popup_text(page):
    """현재 화면에 남아있는 팝업/모달로 보이는 요소가 있으면 그 텍스트 일부를 반환합니다
    (진단 메시지용, 최선 노력 — 실패하거나 못 찾으면 빈 문자열)."""
    selectors = [
        "[role='dialog']",
        ".layer_popup",
        ".pop_layer",
        ".popup",
        ".modal",
    ]
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            if loc.count() > 0 and loc.is_visible(timeout=300):
                text = loc.inner_text(timeout=500)
                text = " ".join(text.split())
                if text:
                    return text[:200]
        except Exception:
            continue
    return ""


def upload_one_account(playwright, name, cafe24_id, cafe24_pw, file_bytes, file_name,
                        headless=True, post_delay_sec=3, nav_timeout_ms=25000,
                        debug_pause=False, submit_wait_sec=180):
    """단일 계정으로 로그인 -> 업로드 화면 이동 -> 파일 업로드 시도.
    실패해도 예외를 던지지 않고 AccountResult로 결과를 반환합니다(재시도하지 않음).

    submit_wait_sec: 업로드 버튼 클릭 후, 실제 업로드 처리 요청(UPLOAD_SUBMIT_URL_HINT)의
    응답이 올 때까지 최대 몇 초 기다릴지. 실제 계정으로 받은 진단 로그를 보면 클릭 자체는
    정상적으로 이 요청을 서버로 보내지만, 상품 수가 많은 파일은 서버 처리가 오래 걸려
    응답이 늦게 올 수 있습니다 — 이 시간 안에 응답이 오지 않으면 브라우저를 닫지 않고
    끝까지 기다립니다(무한 대기는 아니며, 이 값을 넘기면 포기하고 진단 정보에 명확히
    표시합니다).

    debug_pause=True이면(그리고 headless=False일 때만) 업로드 화면에 진입한 직후,
    자동 팝업 정리를 하기 전에 Playwright 인스펙터로 실행을 멈춥니다 — 화면에 남은
    팝업/버튼의 정확한 선택자를 직접 확인하고 싶을 때 사용합니다. headless=True인
    경우에는 볼 화면이 없어 무한 대기하게 되므로 이 옵션을 무시합니다."""
    browser = None
    page = None
    screenshot_before_click = None
    screenshot_after_click = None
    button_html_before = None
    button_html_after = None
    diagnosis = ""
    network_log = []
    console_log = []
    try:
        browser = playwright.chromium.launch(headless=headless, args=["--no-sandbox"])
        context = browser.new_context(accept_downloads=False)
        page = context.new_page()

        # 카페24는 업로드 버튼을 누르면 브라우저 네이티브 확인창(예: "해당 파일을
        # 업로드 하시겠습니까?")을 띄웁니다. Playwright는 이런 다이얼로그를 처리하는
        # 핸들러가 없으면 기본적으로 자동 "취소" 처리를 해버려서, 실제로는 업로드가
        # 진행되지 않았는데도 그대로 넘어가 버리는 문제가 있었습니다(실사용 테스트 중
        # 확인 — 팝업은 잘 닫혔지만 그 다음 뜨는 이 확인창 때문에 업로드가 취소되고
        # 있었습니다). 페이지 생성 직후부터 등록해두어, 어떤 시점에 다이얼로그가 뜨든
        # 항상 "확인"을 누른 것과 동일하게 자동으로 승인(accept)합니다.
        dialog_messages = []

        def _handle_dialog(dialog):
            try:
                dialog_messages.append(dialog.message)
            except Exception:
                pass
            try:
                dialog.accept()
            except Exception:
                pass

        page.on("dialog", _handle_dialog)

        # "업로드 버튼을 눌렀는데 실제로 반영이 안 된다"는 문제가 대기 시간을 늘려도
        # 계속 재현돼서, 더 이상 추측으로 고치지 않고 실제로 무슨 일이 일어나는지
        # 직접 관찰하기 위한 계측입니다. 콘솔 메시지(자바스크립트 오류 포함)와
        # 네트워크 요청/응답을 기록해두고, 버튼 클릭 "전"과 "후"를 구분해서 남깁니다
        # — 클릭 후에 서버로 요청이 아예 안 갔다면 클릭이 실제로는 아무 동작도
        # 안 했다는 뜻이고, 요청은 갔는데 오류 응답이면 서버 쪽 문제라는 뜻이므로
        # 완전히 다른 해결책이 필요합니다. 이 정보는 화면 "④ 업로드 결과"의
        # "진단 정보"에서 그대로 확인할 수 있습니다.
        click_state = {"clicked": False}

        def _handle_console(msg):
            try:
                console_log.append(f"[{msg.type}] {msg.text}")
            except Exception:
                pass

        def _handle_request(request):
            try:
                if request.resource_type in ("xhr", "fetch", "document"):
                    network_log.append({
                        "phase": "클릭 후" if click_state["clicked"] else "클릭 전",
                        "kind": "요청",
                        "method": request.method,
                        "url": request.url,
                    })
            except Exception:
                pass

        def _handle_response(response):
            try:
                if response.request.resource_type in ("xhr", "fetch", "document"):
                    network_log.append({
                        "phase": "클릭 후" if click_state["clicked"] else "클릭 전",
                        "kind": "응답",
                        "status": response.status,
                        "url": response.url,
                    })
            except Exception:
                pass

        page.on("console", _handle_console)
        page.on("request", _handle_request)
        page.on("response", _handle_response)

        page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=nav_timeout_ms)
        page.locator(ID_XPATH).fill(cafe24_id)
        page.locator(PW_XPATH).fill(cafe24_pw)
        page.locator(LOGIN_BTN_XPATH).click()

        try:
            page.wait_for_load_state("networkidle", timeout=nav_timeout_ms)
        except PlaywrightTimeoutError:
            pass

        if "eclogin.cafe24.com" in page.url:
            body_text = ""
            try:
                body_text = page.locator("body").inner_text(timeout=2000)
            except Exception:
                pass
            shot = _safe_screenshot(page)
            reason = LOGIN_FAIL_TEXT if LOGIN_FAIL_TEXT in body_text else "로그인 후에도 로그인 페이지에 머물러 있습니다."
            return AccountResult(
                name=name, status="login_fail", message=reason, screenshot=shot,
                network_log=network_log, console_log=console_log,
            )

        _close_popups(context, page)

        try:
            base_url = page.evaluate("window.location.origin")
        except Exception:
            base_url = page.url.split("/disp/")[0]

        upload_url = base_url.rstrip("/") + UPLOAD_PATH
        page.goto(upload_url, wait_until="domcontentloaded", timeout=nav_timeout_ms)

        if debug_pause and not headless:
            # 여기서 멈춥니다 — 업로드 화면에 막 진입한 상태 그대로(팝업이 떠 있다면
            # 그 팝업도 건드리지 않은 상태)이며, 아래 자동 팝업 정리(_close_popups)는
            # 아직 실행되지 않았습니다. Playwright 인스펙터(별도 창)의 '재생' 버튼을
            # 누르기 전까지는 그대로 대기합니다.
            try:
                page.pause()
            except Exception:
                pass

        # 업로드 화면에 진입한 직후에도 안내 팝업(예: "전용 엑셀양식 다운로드")이 새로
        # 뜨는 경우가 있어, 파일을 첨부하기 전에 한 번 더 정리합니다.
        _close_popups(context, page)

        # 리눅스(Streamlit Cloud)에는 항상 있는 /tmp가 Windows에는 존재하지 않아,
        # 하드코딩된 "/tmp/..." 경로로 파일을 쓰면 Windows 로컬 실행에서
        # "[Errno 2] No such file or directory" 오류가 났습니다(실사용 테스트 중
        # 확인). OS에 맞는 임시 폴더를 알아서 찾아주는 tempfile.gettempdir()를
        # 사용하도록 수정 — Windows에서는 사용자 임시 폴더, 리눅스에서는 기존과 동일하게
        # /tmp를 사용합니다. 계정 이름에 경로 구분자가 섞여 있어도 안전하도록 파일명도
        # 정리합니다.
        safe_name = "".join(c for c in name if c not in '\\/:*?"<>|') or "account"
        temp_path = os.path.join(
            tempfile.gettempdir(), f"_cafe24_upload_{safe_name}_{os.getpid()}.xlsx"
        )
        with open(temp_path, "wb") as f:
            f.write(file_bytes)

        attached = False
        try:
            file_input = page.locator("#dropZone input[type=file]")
            if file_input.count() > 0:
                file_input.set_input_files(temp_path, timeout=5000)
                attached = True
        except Exception:
            attached = False

        if not attached:
            try:
                with page.expect_file_chooser(timeout=5000) as fc_info:
                    page.locator(UPLOAD_BTN_XPATH).click()
                fc_info.value.set_files(temp_path)
                attached = True
            except Exception:
                attached = False

        try:
            os.remove(temp_path)
        except Exception:
            pass

        if not attached:
            shot = _safe_screenshot(page)
            return AccountResult(
                name=name, status="error",
                message="업로드 파일 입력창을 찾지 못했습니다 (페이지 구조가 예상과 다를 수 있습니다).",
                screenshot=shot,
                network_log=network_log, console_log=console_log,
            )

        # 파일을 방금 첨부한 직후에는, 화면이 파일을 검증/미리보기 처리하는 중이라
        # "#dropZone" 버튼이 아직 "파일찾기"(파일 선택) 상태에서 "업로드"(업로드 시작)
        # 상태로 완전히 바뀌지 않았을 수 있습니다. 이 상태에서 바로 버튼을 누르면 클릭
        # 자체는 예외 없이 "성공"하지만, 실제로는 파일 선택창을 다시 여는 것과 같은
        # 동작만 하고 서버 업로드는 전혀 시작되지 않을 수 있습니다 — 화면(진행중)에는
        # 파일이 100%로 붙어있는 것처럼 보이는데 실제 업로드 이력에는 아무것도 안 남는,
        # 실사용 중 반복 확인된 증상과 정확히 일치합니다. 로컬 재현 테스트로 확인한
        # 문제이며, 버튼을 누르기 전에 먼저 (1) 파일 검증 관련 네트워크 활동이 잠잠해질
        # 때까지 짧게 기다리고 (2) 화면이 완전히 전환될 여유를 위해 추가로 잠깐 더
        # 기다립니다(둘 다 실패해도 예외를 삼키고 그냥 진행 — 무한 대기 방지).
        try:
            page.wait_for_load_state("networkidle", timeout=8000)
        except PlaywrightTimeoutError:
            pass
        except Exception:
            pass
        # 여기서도 순수 time.sleep() 대신 page.wait_for_timeout()을 씁니다 — 클릭
        # 직전에 이 대기를 걸어두는 이유가 바로 "지금까지 밀린 네트워크 이벤트들이
        # 확실히 다 처리되게" 하기 위함인데, page.on() 콜백은 Playwright API 호출을
        # 통해서만 pump된다는 것을 확인했으므로(아래 submit_wait_sec 대기 로직 참고),
        # 여기서 순수 time.sleep()을 쓰면 밀린 이벤트가 이 클릭 시점(click_state를
        # "클릭 후"로 바꾸는 시점) 이후에야 뒤늦게 처리되면서, 사실은 클릭 "전"에 있었던
        # 활동이 "클릭 후"로 잘못 분류될 위험이 있습니다.
        page.wait_for_timeout(1500)

        screenshot_before_click = _safe_screenshot(page)
        try:
            button_html_before = page.locator(UPLOAD_BTN_XPATH).first.evaluate(
                "el => el.outerHTML"
            )
        except Exception:
            button_html_before = None

        # 파일 첨부 후 새로 뜬 팝업이 업로드 버튼을 가릴 수 있어 한 번 더 정리한 뒤
        # 클릭합니다. 클릭 자체가 끝까지 실패하면(예: 팝업에 가려져 클릭이 막히는 경우)
        # 예외를 조용히 무시하고 넘어가지 않습니다 — 그렇게 하면 실제로는 업로드 버튼이
        # 눌리지 않았는데도 "성공"으로 잘못 표시되는 문제가 실사용 중 확인됐습니다.
        # 대신 팝업 정리 후 한 번 더 재시도하고, 그래도 안 되면 오류로 명확히 보고합니다
        # (재시도는 이 클릭 단계 안에서만이며, 계정 자체를 다시 시도하지는 않습니다).
        #
        # click_state["clicked"]는 click() 호출 "직전"에 True로 바꿉니다 — click() 호출이
        # 반환되는 시점과, 클릭으로 촉발된 페이지 내부 로직(예: confirm() 대화상자를 거쳐
        # 실제 fetch가 나가는 것)이 Python 쪽에 통지되는 시점 사이에는 실제로 순서가
        # 보장되지 않는다는 것이 로컬 테스트로 확인됐습니다 — click() 호출 "이후"에
        # 플래그를 세팅했더니, 클릭이 유발한 실제 요청이 백그라운드에서 먼저 처리되면서
        # "클릭 전"으로 잘못 분류되는 레이스 컨디션이 있었습니다. 클릭을 시도하는 순간
        # 자체를 기준점으로 삼아야 이 레이스가 없습니다.
        click_error = None
        for attempt in range(2):
            try:
                click_state["clicked"] = True
                page.locator(UPLOAD_BTN_XPATH).click(timeout=5000)
                click_error = None
                break
            except Exception as e:
                click_error = e
                _close_popups(context, page)
                time.sleep(0.5)

        if click_error is not None:
            popup_text = _visible_popup_text(page)
            shot = _safe_screenshot(page)
            msg = "업로드 버튼을 클릭하지 못했습니다"
            if popup_text:
                msg += f" (화면에 남은 팝업으로 보임: \"{popup_text}\")"
            else:
                msg += f" ({click_error})"
            msg += " — 이 계정은 재시도하지 않았습니다. 스크린샷을 확인해주세요."
            return AccountResult(
                name=name, status="error", message=msg, screenshot=shot,
                screenshot_before_click=screenshot_before_click,
                button_html_before=button_html_before,
                network_log=network_log, console_log=console_log,
            )

        # 클릭 직후 버튼 자체의 HTML(문구/활성화 상태 등)도 함께 남겨서, 클릭 전과
        # 비교했을 때 실제로 뭔가 바뀌었는지 확인할 수 있게 합니다.
        try:
            button_html_after = page.locator(UPLOAD_BTN_XPATH).first.evaluate(
                "el => el.outerHTML"
            )
        except Exception:
            button_html_after = None
        screenshot_after_click = _safe_screenshot(page)

        # 업로드 버튼 클릭(및 확인창 승인) 직후, 실제 서버 처리가 끝나기 전에 브라우저를
        # 닫아버리면 파일이 화면에는 "진행중 100%"로 붙어 있는 것처럼 보여도 실제로는
        # 아무것도 반영되지 않는 문제가 실사용 중 반복 확인됐습니다. 처음에는 일반적인
        # "networkidle"(페이지의 모든 네트워크 활동이 잠잠해질 때까지)로 기다렸는데,
        # 실제 진단 로그를 받아보니 이 방식의 근본적인 허점이 드러났습니다 — 카페24
        # 관리자 화면은 알림/대시보드 등 계속 폴링하는 백그라운드 요청이 많아서
        # "완전히 idle" 상태가 늦게 오거나, 반대로 우리가 정말 기다려야 하는 실제 업로드
        # 요청(UPLOAD_SUBMIT_URL_HINT, 실제로 관찰: POST .../product/ProductExcelSet)의
        # 응답이 아직 안 왔는데도 다른 조건으로 networkidle이 만족되어 너무 일찍 넘어갈
        # 수 있었습니다. 이제는 애매한 "전체가 잠잠한지"가 아니라, 우리가 실제로 클릭한
        # 그 요청의 응답이 왔는지를 직접 지켜봅니다 — 최대 submit_wait_sec(기본 180초)
        # 까지 기다리고, 그동안은 절대 브라우저를 닫지 않습니다.
        # 주의: 이 대기 루프는 반드시 page.wait_for_timeout()으로 쉬어야 합니다 —
        # Playwright의 동기(sync) API는 등록해둔 이벤트 콜백(page.on("response") 등)을
        # 메인 스레드가 Playwright API를 실제로 호출할 때 함께 처리(pump)하는 구조라서,
        # 순수 파이썬 time.sleep()으로만 기다리면 그동안 도착한 응답 이벤트가 큐에 쌓인
        # 채로 전달되지 않는 것이 로컬 테스트로 직접 확인됐습니다(실제로 서버는 1.5초 만에
        # 응답했는데도, time.sleep()으로 기다리는 동안은 network_log에 전혀 기록되지 않다가
        # 그 다음 Playwright API 호출(예: 화면 텍스트 읽기) 시점에야 한꺼번에 기록됨).
        # page.wait_for_timeout()은 Playwright 자체 대기 함수라 이 문제가 없습니다.
        submit_deadline = time.monotonic() + max(submit_wait_sec, 0)
        submit_response_entry = None
        while time.monotonic() < submit_deadline:
            submit_response_entry = next(
                (
                    e for e in network_log
                    if e["phase"] == "클릭 후" and e["kind"] == "응답"
                    and UPLOAD_SUBMIT_URL_HINT in e.get("url", "")
                ),
                None,
            )
            if submit_response_entry is not None:
                break
            page.wait_for_timeout(1000)

        # 응답을 받았다면(또는 최대 시간을 다 기다렸다면) 화면이 최종 상태로 갱신될
        # 여유만 짧게 더 둡니다 — 더 이상 이 시간이 "성공/실패를 좌우하는 핵심 대기"가
        # 아니므로 길게 잡을 필요가 없습니다.
        time.sleep(min(post_delay_sec, 5))

        body_text = ""
        try:
            body_text = page.locator("body").inner_text(timeout=2000)
        except Exception:
            pass

        shot = _safe_screenshot(page)

        # 네이티브 확인창(예: "해당 파일을 업로드 하시겠습니까?")에 뜬 문구도 실패 판별
        # 대상에 함께 포함합니다 — 카페24가 오류를 이런 확인창으로 띄우는 경우 페이지
        # 본문(body_text)만 봐서는 놓칠 수 있기 때문입니다.
        combined_text = body_text + " " + " ".join(dialog_messages)

        # 추측이 아니라 실제로 관찰된 증거를 바탕으로 "버튼을 누른 게 진짜 뭔가를
        # 했는지", 더 나아가 "그 요청이 실제로 끝났는지"까지 판정합니다. 실제 계정으로
        # 받은 진단 로그에서, 클릭 후 진짜 업로드 요청(UPLOAD_SUBMIT_URL_HINT)이 서버로
        # 나가는 것까지는 확인됐지만 그 응답을 받은 기록이 없었습니다 — 이게 "요청 자체가
        # 없음"과는 다른, 훨씬 더 구체적인 문제입니다. 아래 판정은 이 세 가지를 순서대로
        # 구분합니다: (1) 진짜 업로드 요청 자체가 없었다(버튼이 아무 동작도 안 함),
        # (2) 요청은 나갔는데 응답을 못 받았다(=지금까지 반복된 문제의 가장 유력한
        # 원인 — 서버 처리가 오래 걸리거나 연결이 끊어짐), (3) 응답은 받았는데
        # 오류였다(서버 쪽 문제), (4) 응답을 정상적으로 받았다(성공).
        post_click_requests = [e for e in network_log if e["phase"] == "클릭 후" and e["kind"] == "요청"]
        submit_request_entries = [e for e in post_click_requests if UPLOAD_SUBMIT_URL_HINT in e.get("url", "")]
        submit_request_sent = bool(submit_request_entries)
        submit_response_status = submit_response_entry["status"] if submit_response_entry else None
        button_changed = (
            button_html_before is not None
            and button_html_after is not None
            and button_html_before != button_html_after
        )

        if not post_click_requests:
            diagnosis = (
                "🔴 클릭 후 서버로 어떤 요청도 발생하지 않았습니다 — 버튼 클릭이 실제로는 "
                "아무 동작도 하지 않았을 가능성이 매우 높습니다(예: 이미 다른 상태로 바뀐 "
                "버튼을 다시 누른 경우 등). 아래 '클릭 전/후 버튼 HTML'과 '클릭 전/후 화면'을 "
                "함께 확인해주세요."
            )
        elif submit_request_sent and submit_response_entry is None:
            diagnosis = (
                f"🔴 실제 업로드 처리 요청(서버로 \"{UPLOAD_SUBMIT_URL_HINT}\" 요청)은 정상적으로 "
                f"전송됐지만, {submit_wait_sec}초를 기다려도 응답을 받지 못했습니다 — 지금까지 "
                "반복된 문제의 가장 유력한 원인으로 보입니다. 상품 수가 아주 많아 서버 처리가 "
                "이 시간보다 더 오래 걸리거나, 서버와의 연결이 중간에 끊어졌을 가능성이 있습니다. "
                "'③'에서 대기 시간을 더 늘려서 다시 시도해보시거나, 그래도 안 되면 이 로그를 "
                "그대로 알려주세요 — 대기 시간 문제가 아니라는 뜻이므로 다른 원인을 찾아야 합니다."
            )
        elif submit_response_status is not None and submit_response_status >= 400:
            diagnosis = (
                f"🟠 실제 업로드 처리 요청은 서버로 전달됐지만 오류 응답({submit_response_status})을 "
                "받았습니다. 버튼 클릭과 대기 자체는 정상이며, 서버 쪽(세션 만료, 파일 형식/용량 "
                "등)을 확인해봐야 하는 문제로 보입니다."
            )
        elif submit_response_status is not None:
            diagnosis = (
                f"🟢 실제 업로드 처리 요청이 서버로 전달됐고, 정상 응답({submit_response_status})까지 "
                "받았습니다 — 버튼 클릭과 서버 처리가 끝까지 정상적으로 이어진 것으로 보입니다"
                "(카페24 화면의 실제 반영 여부는 최종적으로 직접 확인해주세요)."
            )
        else:
            # UPLOAD_SUBMIT_URL_HINT와 일치하는 요청은 못 찾았지만(카페24 화면 구조가 달라졌을
            # 수 있음), 클릭 후 다른 요청은 있었던 애매한 경우 — 오류 응답 유무로만 대략 판정.
            post_click_error_responses = [
                e for e in network_log
                if e["phase"] == "클릭 후" and e["kind"] == "응답"
                and isinstance(e.get("status"), int) and e["status"] >= 400
            ]
            if post_click_error_responses:
                urls = ", ".join(sorted({e["url"] for e in post_click_error_responses}))
                diagnosis = (
                    f"🟠 클릭 후 서버로 요청은 발생했지만(업로드 요청인지는 특정 못함) 오류 응답을 "
                    f"받았습니다 ({urls})."
                )
            else:
                diagnosis = (
                    "🟡 클릭 후 다른 요청은 있었지만, 실제 업로드 요청으로 보이는 요청을 특정하지 "
                    "못했습니다 — 아래 네트워크 로그를 직접 확인해주세요."
                )
        if button_changed:
            diagnosis += " (버튼의 HTML도 클릭 전후로 변화가 감지됐습니다.)"

        if any(kw in combined_text for kw in FAIL_KEYWORDS):
            return AccountResult(
                name=name, status="fail", message="업로드 실패 문구가 감지됐습니다.", screenshot=shot,
                screenshot_before_click=screenshot_before_click,
                screenshot_after_click=screenshot_after_click,
                diagnosis=diagnosis, network_log=network_log, console_log=console_log,
                button_html_before=button_html_before, button_html_after=button_html_after,
                submit_request_sent=submit_request_sent, submit_response_status=submit_response_status,
            )

        dialog_note = f" (자동 승인한 확인창: {' / '.join(dialog_messages)})" if dialog_messages else ""

        # 이전에는 이 지점까지 오면(=명시적인 실패 문구가 없으면) 무조건 "성공"으로
        # 표시했는데, 이게 바로 반복 신고된 "실제로는 아무것도 안 됐는데 성공으로
        # 잘못 나온다"는 문제의 핵심 원인이었습니다. 이제는 실제 업로드 처리 요청의
        # 응답을 확인한 확실한 증거(🟢)가 있을 때만 "성공"으로 표시하고, 응답을 못
        # 받았거나(🔴) 애매한 경우(🟠/🟡)는 "오류"로 정직하게 표시해서 진단 정보를
        # 반드시 확인하도록 합니다.
        if not diagnosis.startswith("🟢"):
            return AccountResult(
                name=name, status="error",
                message="업로드 버튼은 눌렀지만, 서버 처리가 끝났다는 확실한 증거를 확인하지 "
                "못했습니다 — 아래 '진단 정보'를 꼭 확인해주세요." + dialog_note,
                screenshot=shot,
                screenshot_before_click=screenshot_before_click,
                screenshot_after_click=screenshot_after_click,
                diagnosis=diagnosis, network_log=network_log, console_log=console_log,
                button_html_before=button_html_before, button_html_after=button_html_after,
                submit_request_sent=submit_request_sent, submit_response_status=submit_response_status,
            )

        return AccountResult(
            name=name, status="success",
            message="업로드 시도 완료 (실제 업로드 처리 요청의 정상 응답까지 확인했습니다 — 그래도 "
            "카페24 화면에서 실제 반영 여부를 최종 확인해주세요)"
            + dialog_note,
            screenshot=shot,
            screenshot_before_click=screenshot_before_click,
            screenshot_after_click=screenshot_after_click,
            diagnosis=diagnosis, network_log=network_log, console_log=console_log,
            button_html_before=button_html_before, button_html_after=button_html_after,
            submit_request_sent=submit_request_sent, submit_response_status=submit_response_status,
        )

    except Exception as e:
        shot = _safe_screenshot(page) if page is not None else None
        return AccountResult(
            name=name, status="error", message=str(e), screenshot=shot,
            screenshot_before_click=screenshot_before_click,
            screenshot_after_click=screenshot_after_click,
            diagnosis=diagnosis,
            network_log=network_log, console_log=console_log,
            button_html_before=button_html_before, button_html_after=button_html_after,
        )
    finally:
        try:
            if browser:
                browser.close()
        except Exception:
            pass


def _safe_screenshot(page):
    try:
        return page.screenshot(full_page=False)
    except Exception:
        return None


def _ensure_windows_subprocess_event_loop():
    """Windows에서 Streamlit(=Tornado) 서버가 asyncio 이벤트 루프 정책을
    WindowsSelectorEventLoopPolicy로 설정해두는 경우가 있는데, 이 정책은
    자식 프로세스(subprocess) 실행을 지원하지 않아 Playwright가 Chromium을
    실행하려 할 때 'NotImplementedError'가 발생합니다(로컬 Windows에서 실제로
    확인된 오류). Playwright 실행 직전에 subprocess를 지원하는
    WindowsProactorEventLoopPolicy로 되돌려줍니다.
    Windows가 아니거나(예: Streamlit Cloud의 리눅스 서버) 어떤 이유로 정책
    변경이 안 되더라도 예외를 삼켜 기존 동작에는 영향이 없도록 합니다."""
    if not sys.platform.startswith("win"):
        return
    try:
        policy_cls = getattr(asyncio, "WindowsProactorEventLoopPolicy", None)
        if policy_cls is not None and not isinstance(asyncio.get_event_loop_policy(), policy_cls):
            asyncio.set_event_loop_policy(policy_cls())
    except Exception:
        pass


def run_batch_upload(accounts, headless=True, post_delay_sec=3, progress_cb=None,
                      debug_pause=False, submit_wait_sec=180) -> List[AccountResult]:
    """
    accounts: [{'name':.., 'id':.., 'password':.., 'file_bytes':.., 'file_name':..}, ...]
    progress_cb: 선택. callable(index, total, AccountResult) — 진행 상황 콜백(Streamlit 갱신용)
    debug_pause: 선택. True면 각 계정의 업로드 화면 진입 직후 자동 진행을 멈춥니다
        (headless=False일 때만 의미가 있습니다 — upload_one_account 참고).
    submit_wait_sec: 선택. 업로드 버튼 클릭 후 실제 업로드 처리 응답을 최대 몇 초까지
        기다릴지 (upload_one_account 참고).
    반환: AccountResult 리스트 (accounts와 같은 순서)
    """
    _ensure_windows_subprocess_event_loop()
    results = []
    with sync_playwright() as p:
        total = len(accounts)
        for i, acc in enumerate(accounts, start=1):
            result = upload_one_account(
                p, acc["name"], acc["id"], acc["password"],
                acc["file_bytes"], acc["file_name"],
                headless=headless, post_delay_sec=post_delay_sec,
                debug_pause=debug_pause, submit_wait_sec=submit_wait_sec,
            )
            results.append(result)
            if progress_cb:
                progress_cb(i, total, result)
    return results
