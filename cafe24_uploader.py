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
                        headless=True, post_delay_sec=3, nav_timeout_ms=25000):
    """단일 계정으로 로그인 -> 업로드 화면 이동 -> 파일 업로드 시도.
    실패해도 예외를 던지지 않고 AccountResult로 결과를 반환합니다(재시도하지 않음)."""
    browser = None
    page = None
    try:
        browser = playwright.chromium.launch(headless=headless, args=["--no-sandbox"])
        context = browser.new_context(accept_downloads=False)
        page = context.new_page()

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
            return AccountResult(name=name, status="login_fail", message=reason, screenshot=shot)

        _close_popups(context, page)

        try:
            base_url = page.evaluate("window.location.origin")
        except Exception:
            base_url = page.url.split("/disp/")[0]

        upload_url = base_url.rstrip("/") + UPLOAD_PATH
        page.goto(upload_url, wait_until="domcontentloaded", timeout=nav_timeout_ms)

        # 업로드 화면에 진입한 직후에도 안내 팝업(예: "전용 엑셀양식 다운로드")이 새로
        # 뜨는 경우가 있어, 파일을 첨부하기 전에 한 번 더 정리합니다.
        _close_popups(context, page)

        temp_path = f"/tmp/_cafe24_upload_{name}_{os.getpid()}.xlsx"
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
            )

        # 파일 첨부 후 새로 뜬 팝업이 업로드 버튼을 가릴 수 있어 한 번 더 정리한 뒤
        # 클릭합니다. 클릭 자체가 끝까지 실패하면(예: 팝업에 가려져 클릭이 막히는 경우)
        # 예외를 조용히 무시하고 넘어가지 않습니다 — 그렇게 하면 실제로는 업로드 버튼이
        # 눌리지 않았는데도 "성공"으로 잘못 표시되는 문제가 실사용 중 확인됐습니다.
        # 대신 팝업 정리 후 한 번 더 재시도하고, 그래도 안 되면 오류로 명확히 보고합니다
        # (재시도는 이 클릭 단계 안에서만이며, 계정 자체를 다시 시도하지는 않습니다).
        click_error = None
        for attempt in range(2):
            try:
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
            return AccountResult(name=name, status="error", message=msg, screenshot=shot)

        time.sleep(post_delay_sec)

        body_text = ""
        try:
            body_text = page.locator("body").inner_text(timeout=2000)
        except Exception:
            pass

        shot = _safe_screenshot(page)

        if any(kw in body_text for kw in FAIL_KEYWORDS):
            return AccountResult(name=name, status="fail", message="업로드 실패 문구가 감지됐습니다.", screenshot=shot)

        return AccountResult(
            name=name, status="success",
            message="업로드 시도 완료 (자동 판별 결과이므로, 스크린샷으로 실제 반영 여부를 확인해주세요)",
            screenshot=shot,
        )

    except Exception as e:
        shot = _safe_screenshot(page) if page is not None else None
        return AccountResult(name=name, status="error", message=str(e), screenshot=shot)
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


def run_batch_upload(accounts, headless=True, post_delay_sec=3, progress_cb=None) -> List[AccountResult]:
    """
    accounts: [{'name':.., 'id':.., 'password':.., 'file_bytes':.., 'file_name':..}, ...]
    progress_cb: 선택. callable(index, total, AccountResult) — 진행 상황 콜백(Streamlit 갱신용)
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
            )
            results.append(result)
            if progress_cb:
                progress_cb(i, total, result)
    return results
