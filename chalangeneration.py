import asyncio
import os
import re
import sys
import logging
from datetime import datetime
from pathlib import Path

import ollama
from openpyxl import load_workbook

from playwright.async_api import (
    async_playwright,
    TimeoutError as PlaywrightTimeoutError,
)


# ============================================================
# PATHS & CONFIGURATION
# ============================================================

def get_main_excel_file():
    candidates = [
        Path(r"E:\Supplementary Monthly Contributio\suplimentry_esic_users.xlsx"),
        Path(r"E:\Supplementary Monthly Contributio\esic_users.xlsx"),
    ]

    for c in candidates:
        if c.exists() and c.is_file():
            return c

    return candidates[0]


EXCEL_FILE = get_main_excel_file()

DATA_FOLDER = Path(
    r"E:\Supplementary Monthly Contributio\ESIC_AUTO\data"
)

CAPTCHA_DIR = DATA_FOLDER / "captcha"

CAPTCHA_DIR.mkdir(
    parents=True,
    exist_ok=True
)

OLLAMA_MODEL = "gemma3:4b"

MAX_RETRIES_PER_USER = 3

LOGIN_URL = (
    "https://portal.esic.gov.in/"
    "ESICInsurance1/ESICInsurancePortal/"
    "Portal_Loginnew.aspx"
)


# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

logger = logging.getLogger("esic_challan")


# ============================================================
# EXCEL HELPER
# ============================================================

def clean_excel_val(val):

    if val is None:
        return ""

    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val)

    if isinstance(val, int):
        return str(val)

    val_str = str(val).strip()

    if val_str.endswith(".0"):
        val_str = val_str[:-2]

    if re.fullmatch(
        r"\d+(\.\d+)?[eE][+-]?\d+",
        val_str
    ):
        try:
            return str(int(float(val_str)))
        except Exception:
            pass

    return val_str


def get_column_mapping(ws):

    mapping = {
        "loginid": 1,
        "password": 2,
        "response": 3
    }

    for col_idx in range(
        1,
        ws.max_column + 1
    ):

        val = ws.cell(
            1,
            col_idx
        ).value

        if not val:
            continue

        cleaned = re.sub(
            r"[^a-z0-9]",
            "",
            str(val).strip().lower()
        )

        if cleaned in (
            "loginid",
            "login",
            "username",
            "userid",
            "user"
        ):
            mapping["loginid"] = col_idx

        elif cleaned in (
            "password",
            "pass",
            "pwd"
        ):
            mapping["password"] = col_idx

        elif cleaned in (
            "response",
            "res",
            "result"
        ):
            mapping["response"] = col_idx

    return mapping


def read_distinct_users(
    excel_path=EXCEL_FILE
):

    if not excel_path.exists():
        raise FileNotFoundError(
            f"Excel file not found: {excel_path}"
        )

    wb = load_workbook(
        excel_path,
        data_only=True
    )

    ws = wb.active

    cols = get_column_mapping(ws)

    users = []

    seen = set()

    last_loginid = ""
    last_password = ""

    for r in range(
        2,
        ws.max_row + 1
    ):

        raw_login = clean_excel_val(
            ws.cell(
                r,
                cols["loginid"]
            ).value
        )

        raw_pwd = clean_excel_val(
            ws.cell(
                r,
                cols["password"]
            ).value
        )

        if raw_login:
            last_loginid = raw_login
        else:
            raw_login = last_loginid

        if raw_pwd:
            last_password = raw_pwd
        else:
            raw_pwd = last_password

        if not raw_login or not raw_pwd:
            continue

        key = (
            raw_login,
            raw_pwd
        )

        if key not in seen:

            seen.add(key)

            users.append(
                (
                    raw_login,
                    raw_pwd,
                    r
                )
            )

    wb.close()

    return users


# ============================================================
# CAPTCHA - OLLAMA
# ============================================================

def _get_captcha_sync(image_path):

    try:

        response = ollama.chat(
            model=OLLAMA_MODEL,
            messages=[
                {
                    "role": "user",

                    "content": (
                        "Read the CAPTCHA text from this image. "
                        "It contains a short 6-7 character "
                        "alphanumeric code. "
                        "Reply ONLY with these exact characters. "
                        "Do NOT add any extra words, spaces, "
                        "or punctuation."
                    ),

                    "images": [
                        str(image_path)
                    ],
                }
            ],
        )

        raw_text = response[
            "message"
        ][
            "content"
        ].strip()

        cleaned = re.sub(
            r"[^a-zA-Z0-9]",
            "",
            raw_text
        )

        logger.info(
            f"Ollama CAPTCHA extracted: {cleaned}"
        )

        return cleaned

    except Exception as e:

        logger.error(
            f"Ollama CAPTCHA extraction failed: {e}"
        )

        return ""


async def get_captcha_from_ollama(
    image_path
):

    return await asyncio.to_thread(
        _get_captcha_sync,
        image_path
    )


# ============================================================
# DISMISS POPUPS
# ============================================================

async def dismiss_all_popups(page):

    popup_selectors = [

        "input#div1_close",

        "input#btnpnlSanjeevanicolse",

        "input#btnpnlcolse",

        "input#btnpnlPwdcolse",

        'input[name="div1_close"]',

        'input[name="btnpnlSanjeevanicolse"]',

        'input[name="btnpnlcolse"]',

        'input[name="btnpnlPwdcolse"]',

        'input[value="I Agree"]',

        'input[value="Close"]',

        'input[value="X"]',

    ]

    for _ in range(4):

        clicked_any = False

        for selector in popup_selectors:

            try:

                el = page.locator(
                    selector
                ).first

                if (
                    await el.count() > 0
                    and await el.is_visible()
                ):

                    btn_id = (
                        await el.get_attribute("id")
                        or selector
                    )

                    logger.info(
                        f"Closing popup: {btn_id}"
                    )

                    await el.click(
                        force=True
                    )

                    await page.wait_for_timeout(
                        400
                    )

                    clicked_any = True

            except Exception:
                pass

        if not clicked_any:
            break


# ============================================================
# LOGIN FLOW
# ============================================================

async def fill_login_and_captcha(
    page,
    loginid,
    password,
    captcha_path,
    auto_mode=True
):

    # --------------------------------------------------------
    # USERNAME
    # --------------------------------------------------------

    user_input = page.locator(
        'input#txtUserName, '
        'input[name*="UserName"], '
        'input[name*="username"], '
        'input[placeholder*="User"]'
    ).first

    if await user_input.count() == 0:
        raise RuntimeError(
            "Username input field not found."
        )

    await user_input.fill(
        str(loginid)
    )

    # --------------------------------------------------------
    # PASSWORD
    # --------------------------------------------------------

    password_input = page.locator(
        'input#txtPassword, '
        'input[name*="Password"], '
        'input[type="password"]:visible'
    ).first

    if await password_input.count() == 0:
        raise RuntimeError(
            "Password field not found."
        )

    await password_input.fill(
        str(password)
    )

    # --------------------------------------------------------
    # CAPTCHA IMAGE
    # --------------------------------------------------------

    captcha_img = page.locator(
        'img#img1, '
        'img[src*="ChallanHandler"], '
        'img[id*="Captcha"], '
        'img[id*="captcha"]'
    ).first

    if await captcha_img.count() == 0:
        raise RuntimeError(
            "ESIC CAPTCHA image element not found."
        )

    await captcha_img.screenshot(
        path=str(captcha_path)
    )

    # --------------------------------------------------------
    # CAPTCHA OCR
    # --------------------------------------------------------

    captcha_text = await get_captcha_from_ollama(
        captcha_path
    )

    if (
        len(captcha_text) < 4
        or len(captcha_text) > 10
    ):

        if not auto_mode:
            captcha_text = input(
                "Enter CAPTCHA manually: "
            ).strip()

    # --------------------------------------------------------
    # CAPTCHA INPUT
    # --------------------------------------------------------

    captcha_input = page.locator(
        'input#txtChallanCaptcha, '
        'input[name="txtChallanCaptcha"], '
        'input[placeholder*="Captcha"]'
    ).first

    if await captcha_input.count() == 0:
        raise RuntimeError(
            "ESIC CAPTCHA input field not found."
        )

    await captcha_input.fill(
        captcha_text
    )

    logger.info(
        f"CAPTCHA filled: {captcha_text}"
    )


# ============================================================
# LOGIN
# ============================================================

async def login(
    page,
    loginid,
    password,
    safe_loginid,
    auto_mode=True
):

    page.on(
        "dialog",
        lambda dialog:
        asyncio.create_task(
            dialog.accept()
        )
    )

    # --------------------------------------------------------
    # OPEN LOGIN PAGE
    # --------------------------------------------------------

    await page.goto(
        LOGIN_URL,
        wait_until="domcontentloaded",
        timeout=60000
    )

    await page.wait_for_timeout(
        1000
    )

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    captcha_path = (
        CAPTCHA_DIR
        / f"captcha_{safe_loginid}_{timestamp}.png"
    )

    # --------------------------------------------------------
    # USERNAME + PASSWORD + CAPTCHA
    # --------------------------------------------------------

    await fill_login_and_captcha(
        page,
        loginid,
        password,
        captcha_path,
        auto_mode=auto_mode
    )

    # --------------------------------------------------------
    # WAIT BEFORE LOGIN
    # --------------------------------------------------------

    logger.info(
        "Waiting 2 seconds before clicking Login..."
    )

    await page.wait_for_timeout(
        2000
    )

    # --------------------------------------------------------
    # LOGIN BUTTON
    # --------------------------------------------------------

    login_button = page.locator(
        'input#btnLogin, '
        'input[name="btnLogin"], '
        'input[value="Login"], '
        'input.login100-form-btn-gou'
    ).first

    if await login_button.count() == 0:

        login_button = page.get_by_role(
            "button",
            name="Login",
            exact=False
        )

    if await login_button.count() == 0:
        raise RuntimeError(
            "LOGIN button not found."
        )

    await login_button.click()

    logger.info(
        "Login button clicked."
    )

    # --------------------------------------------------------
    # WAIT FOR PAGE
    # --------------------------------------------------------

    try:

        await page.wait_for_load_state(
            "domcontentloaded",
            timeout=30000
        )

    except PlaywrightTimeoutError:
        pass

    await page.wait_for_timeout(
        4000
    )


# ============================================================
# GENERATE CHALLAN → VIEW
# ============================================================

async def open_generate_challan_and_view(
    page,
    loginid
):

    # --------------------------------------------------------
    # LOGIN VALIDATION
    # --------------------------------------------------------

    if "Portal_Loginnew.aspx" in page.url:

        error_msg_el = page.locator(
            '#lblMsg, '
            '.error, '
            'span[style*="color:Red"], '
            'span[style*="color:red"]'
        ).first

        if await error_msg_el.count() > 0:

            err_text = await (
                error_msg_el.inner_text()
            )

            if err_text.strip():

                raise RuntimeError(
                    f"Login Failed on ESIC Portal: "
                    f"{err_text.strip()}"
                )

        raise RuntimeError(
            "Login Failed "
            "(Invalid CAPTCHA or Password). "
            "Webpage remained on Login screen."
        )

    logger.info(
        "Login successful."
    )

    # --------------------------------------------------------
    # CLOSE POPUPS
    # --------------------------------------------------------

    logger.info(
        "Canceling/dismissing dashboard popups..."
    )

    await page.wait_for_timeout(
        1000
    )

    await dismiss_all_popups(
        page
    )

    await page.wait_for_timeout(
        1000
    )

    # --------------------------------------------------------
    # GENERATE CHALLAN (OPENS POPUP WINDOW)
    # --------------------------------------------------------

    logger.info(
        "Searching for Generate Challan link..."
    )

    generate_challan = page.locator(
        "#lnkGenerateChallan"
    )

    await generate_challan.wait_for(
        state="visible",
        timeout=30000
    )

    logger.info(
        "Generate Challan button found. Clicking to open Challan window..."
    )

    challan_page = None
    try:
        async with page.expect_popup(timeout=15000) as popup_info:
            await generate_challan.click(force=True)
        challan_page = await popup_info.value
        logger.info(f"Captured Challan popup window: {challan_page.url}")
    except Exception as exc:
        logger.warning(f"Popup expect notice: {exc}. Checking open pages in context...")
        all_pages = page.context.pages
        logger.info(f"Total open pages: {len(all_pages)}")
        if len(all_pages) > 1:
            challan_page = all_pages[-1]
        else:
            challan_page = page

    # Handle dialogs automatically on the challan page
    challan_page.on(
        "dialog",
        lambda dialog: asyncio.create_task(dialog.accept())
    )

    await challan_page.bring_to_front()

    try:
        await challan_page.wait_for_load_state(
            "domcontentloaded",
            timeout=20000
        )
    except Exception:
        pass

    # --------------------------------------------------------
    # WAIT 3 SECONDS AFTER GENERATE CHALLAN
    # --------------------------------------------------------

    logger.info(
        "Waiting 3 seconds before locating View button..."
    )

    await challan_page.wait_for_timeout(
        3000
    )

    # --------------------------------------------------------
    # VIEW BUTTON (id=ctl00_HomePageContent_btnView, name=ctl00$HomePageContent$btnView)
    # --------------------------------------------------------

    logger.info(
        "Searching for View button on Challan window..."
    )

    view_button = challan_page.locator(
        '#ctl00_HomePageContent_btnView, '
        'input#ctl00_HomePageContent_btnView, '
        'input[name="ctl00$HomePageContent$btnView"], '
        'input[value="View"], '
        'button:has-text("View")'
    ).first

    # Fallback to main page if popup was not separate
    if await view_button.count() == 0 and challan_page != page:
        view_button = page.locator(
            '#ctl00_HomePageContent_btnView, '
            'input#ctl00_HomePageContent_btnView, '
            'input[name="ctl00$HomePageContent$btnView"], '
            'input[value="View"], '
            'button:has-text("View")'
        ).first
        if await view_button.count() > 0:
            challan_page = page

    await view_button.wait_for(
        state="visible",
        timeout=30000
    )

    logger.info(
        "View button found."
    )

    # --------------------------------------------------------
    # FOCUS + ENTER DIRECTLY ON VIEW BUTTON (AND CLICK)
    # --------------------------------------------------------

    try:
        logger.info(
            "Focusing View button..."
        )
        await view_button.focus()
        await challan_page.wait_for_timeout(
            500
        )

        logger.info(
            "Pressing ENTER on View button..."
        )
        await challan_page.keyboard.press(
            "Enter"
        )
        logger.info(
            "ENTER pressed on View button."
        )

        await challan_page.wait_for_timeout(
            1000
        )

        # Also force click to ensure postback triggers
        await view_button.click(
            force=True
        )
        logger.info(
            "View button clicked."
        )

    except Exception as e:
        logger.warning(
            f"Enter action notice: {e}. Using click fallback..."
        )
        await view_button.click(
            force=True
        )

    await challan_page.wait_for_timeout(
        5000
    )

    try:
        await challan_page.wait_for_load_state(
            "domcontentloaded",
            timeout=20000
        )
    except Exception:
        pass

    logger.info(
        "Generate Challan → View action completed successfully."
    )
    logger.info(
        f"Current URL: {challan_page.url}"
    )

    # Keep page open so user can review the results
    await challan_page.wait_for_timeout(
        8000
    )


# ============================================================
# PROCESS ONE USER
# ============================================================

async def process_user(
    browser,
    loginid,
    password,
    row_number
):

    context = await browser.new_context()

    page = await context.new_page()

    safe_loginid = re.sub(
        r"[^a-zA-Z0-9_-]",
        "_",
        str(loginid)
    )

    try:

        logger.info(
            "=" * 70
        )

        logger.info(
            f"Processing User: {loginid}"
        )

        logger.info(
            f"Excel Row: {row_number}"
        )

        logger.info(
            "=" * 70
        )

        # ----------------------------------------------------
        # LOGIN
        # ----------------------------------------------------

        await login(
            page,
            loginid,
            password,
            safe_loginid,
            auto_mode=True
        )

        # ----------------------------------------------------
        # GENERATE CHALLAN + VIEW
        # ----------------------------------------------------

        await open_generate_challan_and_view(
            page,
            loginid
        )

        logger.info(
            f"SUCCESS: {loginid}"
        )

        return True

    except Exception as e:

        logger.error(
            f"FAILED: {loginid}"
        )

        logger.error(
            f"Error: {e}"
        )

        return False

    finally:

        await context.close()


# ============================================================
# MAIN
# ============================================================

async def main():

    logger.info(
        "=" * 80
    )

    logger.info(
        "ESIC CHALLAN AUTOMATION STARTED"
    )

    logger.info(
        "=" * 80
    )

    # --------------------------------------------------------
    # READ USERS
    # --------------------------------------------------------

    users = read_distinct_users(
        EXCEL_FILE
    )

    logger.info(
        f"Total distinct users: {len(users)}"
    )

    if not users:

        logger.warning(
            "No users found in Excel."
        )

        return

    # --------------------------------------------------------
    # START PLAYWRIGHT
    # --------------------------------------------------------

    async with async_playwright() as p:

        browser = await p.chromium.launch(
            headless=False
        )

        try:

            # ------------------------------------------------
            # PROCESS EACH USER
            # ------------------------------------------------

            for (
                loginid,
                password,
                row_number
            ) in users:

                success = False

                for attempt in range(
                    1,
                    MAX_RETRIES_PER_USER + 1
                ):

                    logger.info(
                        f"Attempt {attempt}/"
                        f"{MAX_RETRIES_PER_USER} "
                        f"for {loginid}"
                    )

                    success = await process_user(
                        browser,
                        loginid,
                        password,
                        row_number
                    )

                    if success:
                        break

                    if attempt < MAX_RETRIES_PER_USER:

                        logger.info(
                            "Retrying user..."
                        )

                        await asyncio.sleep(
                            2
                        )

                if not success:

                    logger.error(
                        f"User failed after "
                        f"{MAX_RETRIES_PER_USER} attempts: "
                        f"{loginid}"
                    )

        finally:

            await browser.close()

    logger.info(
        "=" * 80
    )

    logger.info(
        "ESIC CHALLAN AUTOMATION COMPLETED"
    )

    logger.info(
        "=" * 80
    )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    try:

        asyncio.run(
            main()
        )

    except KeyboardInterrupt:

        logger.info(
            "Automation stopped by user."
        )

    except Exception as e:

        logger.exception(
            f"Fatal error: {e}"
        )