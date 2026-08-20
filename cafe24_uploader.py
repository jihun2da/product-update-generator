"""
카페24 관리자(https://eclogin.cafe24.com/Shop/)에 로그인해서
상품 엑셀 일괄등록/수정(ProductExcelManage) 화면에 파일을 자동으로 업로드하는 모듈.

중요 — 반드시 읽어주세요
------------------------
- 아이디/비밀번호는 반드시 Streamlit Secrets(st.secrets)로만 전달해야 합니다.
  이 파일이나 저장소 어디에도 실제 아이디/비밀번호를 하드코딩하거나 커밋하면 안 됩니다.
- Claude(AI 어시스턴트)는 실제 계정으로 로그인하는 행위 자체를 직접 수행할 수 없습니다
  (안전 정책상, 사용자가 허락해도 마찬가지입니다). 그래서 이 모듈은:
    * 로그인 페이지의 아이디/비밀번호 입력창, 로그인 버튼 셀렉터가 실제로 존재하는지,
      더미(가짜) 값으로 채우고 제출했을 때 정상적으로 "아이디 또는 비밀번호를 확인해주세요"
      오류가 뜨는지까지는 구조적으로 확인했습니다.
    * 하지만 "실제 계정으로 로그인 → 업로드 화면 이동 → 파일 첨부 → 업로드 성공/실패 판별"
      까지 이어지는 전체 흐름은 실제 계정으로 사용자가 직접 최초 1회 테스트해서 확인해야
      합니다. 특히 업로드 성공/실패 판별은 화면에 뜨는 문구를 기준으로 하는데, 정확한 문구를
      실제로 보지 못한 상태로 작성했기 때문에 오탐(성공인데 실패로, 또는 그 반대로 표시)이
      있을 수 있습니다. 계정별로 저장되는 스크린샷을 함께 확인해주세요.
"""

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

# 업로드 결과 화면에서 "실패"로 볼 수 있는 문구 후보 (정확한 문구를 실제로 확인 못해
# 여러 후보를 넓게 잡았습니다 — 실제 사용해보시고 다르면 알려주시면 바로 고치겠습니다)
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
    """로그인 직후 뜨는 팝업(새 탭 또는 인페이지 모달)을 최대한 정리합니다.
    실패해도 전체 흐름에 영향 주지 않도록 예외를 삼킵니다."""
    time.sleep(1)
    for p in list(context.pages):
        if p is not page:
            try:
                p.close()
            except Exception:
                pass

    close_selectors = [
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

        try:
            page.locator(UPLOAD_BTN_XPATH).click(timeout=5000)
        except Exception:
            pass

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


def run_batch_upload(accounts, headless=True, post_delay_sec=3, progress_cb=None) -> List[AccountResult]:
    """
    accounts: [{'name':.., 'id':.., 'password':.., 'file_bytes':.., 'file_name':..}, ...]
    progress_cb: 선택. callable(index, total, AccountResult) — 진행 상황 콜백(Streamlit 갱신용)
    반환: AccountResult 리스트 (accounts와 같은 순서)
    """
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
