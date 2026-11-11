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
# PATHS
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

SHEET_NAME = None  # None = active sheet



OLLAMA_MODEL = "gemma3:4b"



MAX_RETRIES_PER_USER = 3

# Production flow: initial Submit -> payroll fields -> final Submit -> Excel status.




LOGIN_URL = (

    "https://portal.esic.gov.in/"

    "ESICInsurance1/ESICInsurancePortal/"

    "Portal_Loginnew.aspx"

)





# ============================================================

# DOWNLOAD FOLDERS

# ============================================================



DATE_FOLDER_NAME = datetime.now().strftime("%d%m%Y")



TARGET_DOWNLOAD_DIR = (

    DATA_FOLDER

    / DATE_FOLDER_NAME

    / "data_downloaded"

)



# ============================================================
# APPROVED PAYROLL EXCEL - FIXED PRODUCTION SOURCE
# ============================================================
# The approved payroll Excel is the source for ONLY these 3 portal fields:
#   Column C = ESIC Number
#   Column D = Paid Days
#   Column E = ESIC Gross Payable
#
# Exact production file name supplied by user.
APPROVED_PAYROLL_FILENAME = (
    "ESIS.MEDICAL SERVICES ESI_Aug_2026_Approved_Payroll.xlsx"
)

# User-specified payroll folder.
APPROVED_PAYROLL_DIR = Path(
    r"E:\Supplementary Monthly Contributio"
)

# Also support the existing project folder spelling as a fallback,
# without changing the main ESIC project paths.
APPROVED_PAYROLL_DIR_FALLBACK = Path(
    r"E:\Supplementary Monthly Contribution"
)

CAPTCHA_DIR = DATA_FOLDER / "captcha"





CAPTCHA_DIR.mkdir(

    parents=True,

    exist_ok=True,

)



TARGET_DOWNLOAD_DIR.mkdir(

    parents=True,

    exist_ok=True,

)





# ============================================================

# LOGGING

# ============================================================



logging.basicConfig(

    level=logging.INFO,

    format="%(asctime)s | %(levelname)s | %(message)s",

)



logger = logging.getLogger("esic_portal")





# ============================================================

# EXCEL VALUE CLEANER

# ============================================================



def clean_excel_val(val):

    """

    Converts Excel values cleanly to strings.

    Avoids scientific notation and trailing .0.

    """



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



    # Scientific notation

    if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", val_str):

        try:

            return str(int(float(val_str)))

        except Exception:

            pass



    return val_str





# ============================================================

# CAPTCHA - OLLAMA

# ============================================================



def _get_captcha_sync(image_path):

    """

    Synchronous call to Ollama to read ESIC CAPTCHA.

    """



    try:

        response = ollama.chat(

            model=OLLAMA_MODEL,

            messages=[

                {

                    "role": "user",

                    "content": (

                        "Read the CAPTCHA text from this image. "

                        "It contains a short 6-7 character alphanumeric code. "

                        "Reply ONLY with these exact characters. "

                        "Do NOT add any extra words, spaces, or punctuation."

                    ),

                    "images": [str(image_path)],

                }

            ],

        )



        raw_text = response["message"]["content"].strip()



        cleaned = re.sub(

            r"[^a-zA-Z0-9]",

            "",

            raw_text,

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





async def get_captcha_from_ollama(image_path):

    """

    Async wrapper around Ollama sync call.

    """



    return await asyncio.to_thread(

        _get_captcha_sync,

        image_path,

    )





# ============================================================

# DISMISS POPUPS

# ============================================================



async def dismiss_all_popups(page):

    """

    Dismiss ESIC popup dialogs.

    """



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

                el = page.locator(selector).first



                if (

                    await el.count() > 0

                    and await el.is_visible()

                ):

                    btn_id = (

                        await el.get_attribute("id")

                        or selector

                    )



                    await el.click(

                        timeout=3000

                    )



                    logger.info(

                        f"Dismissed Popup: {btn_id}"

                    )



                    await page.wait_for_timeout(1000)



                    clicked_any = True



            except Exception:

                pass



        if not clicked_any:

            break





# ============================================================

# LOGIN + CAPTCHA

# ============================================================



async def fill_login_and_captcha(

    page,

    loginid,

    password,

    captcha_path,

    auto_mode=True,

):



    # ----------------------------

    # Username

    # ----------------------------



    username_input = page.locator(

        'input#txtUserName, '

        'input[name*="UserName"], '

        'input[type="text"]:visible'

    ).first



    if await username_input.count() == 0:

        raise RuntimeError(

            "Login ID field not found."

        )



    await username_input.fill(

        str(loginid)

    )



    # ----------------------------

    # Password

    # ----------------------------



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



    # ----------------------------

    # CAPTCHA image

    # ----------------------------



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



    # ----------------------------

    # Read CAPTCHA using Ollama

    # ----------------------------



    captcha_text = (

        await get_captcha_from_ollama(

            captcha_path

        )

    )



    if len(captcha_text) < 4 or len(captcha_text) > 10:



        if auto_mode:



            logger.warning(

                "Ollama CAPTCHA result looks invalid: "

                f"'{captcha_text}'"

            )



        else:



            captcha_text = input(

                "Enter CAPTCHA manually: "

            ).strip()



    # ----------------------------

    # CAPTCHA input

    # ----------------------------



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

    auto_mode=True,

):



    page.on("dialog", lambda dialog: asyncio.create_task(dialog.accept()))

    await page.goto(
        LOGIN_URL,
        wait_until="domcontentloaded",
        timeout=60000,
    )



    await page.wait_for_timeout(1000)



    timestamp = datetime.now().strftime(

        "%Y%m%d_%H%M%S"

    )



    captcha_path = (

        CAPTCHA_DIR

        / f"captcha_{safe_loginid}_{timestamp}.png"

    )



    await fill_login_and_captcha(

        page,

        loginid,

        password,

        captcha_path,

        auto_mode=auto_mode,

    )



    logger.info(

        "Waiting 10 seconds before clicking Login..."

    )



    await page.wait_for_timeout(10000)



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

            exact=False,

        )



    if await login_button.count() == 0:

        raise RuntimeError(

            "LOGIN button not found."

        )



    await login_button.click()



    logger.info(

        "Login button clicked."

    )



    try:

        await page.wait_for_load_state(

            "domcontentloaded",

            timeout=30000,

        )

    except PlaywrightTimeoutError:

        pass



    await page.wait_for_timeout(4000)





# ============================================================

# SELECT MONTH + YEAR

# ============================================================



async def select_month_year(

    page,

    month,

    year,

):

    """

    Selects Month and Year from ESIC Monthly Contribution page.



    Month and Year are read from Excel:

        Column D = Month

        Column E = Year



    Example:

        D2 = Aug

        E2 = 2026

    """



    month = str(month).strip()

    year = str(year).strip()



    if not month:

        raise RuntimeError(

            "Month is empty in Excel."

        )



    if not year:

        raise RuntimeError(

            "Year is empty in Excel."

        )



    # -----------------------------------------

    # Normalize month names

    # -----------------------------------------



    month_map = {

        "jan": "Jan",

        "january": "Jan",



        "feb": "Feb",

        "february": "Feb",



        "mar": "Mar",

        "march": "Mar",



        "apr": "Apr",

        "april": "Apr",



        "may": "May",



        "jun": "Jun",

        "june": "Jun",



        "jul": "Jul",

        "july": "Jul",



        "aug": "Aug",

        "august": "Aug",



        "sep": "Sep",

        "september": "Sep",



        "oct": "Oct",

        "october": "Oct",



        "nov": "Nov",

        "november": "Nov",



        "dec": "Dec",

        "december": "Dec",

    }



    month_normalized = month_map.get(

        month.lower(),

        month,

    )



    year_normalized = year



    logger.info(

        "Excel Contribution Period: "

        f"Month={month_normalized}, "

        f"Year={year_normalized}"

    )



    # -----------------------------------------

    # Find visible SELECT elements

    # -----------------------------------------



    selects = page.locator(

        "select:visible"

    )



    select_count = await selects.count()



    logger.info(

        f"Visible dropdowns found: {select_count}"

    )



    if select_count == 0:

        raise RuntimeError(

            "No visible dropdown found on "

            "Monthly Contribution page."

        )



    month_select = None

    year_select = None



    # -----------------------------------------

    # Detect Month and Year dropdown

    # -----------------------------------------



    for i in range(select_count):



        select = selects.nth(i)



        try:



            options = await select.locator(

                "option"

            ).evaluate_all(

                """

                options => options.map(o => ({

                    text: (o.textContent || '').trim(),

                    value: o.value

                }))

                """

            )



            option_texts = [

                item["text"]

                for item in options

            ]



            logger.info(

                f"Dropdown {i} options: "

                f"{option_texts}"

            )



            # -----------------------------

            # Month detection

            # -----------------------------



            month_found = False



            for option in option_texts:



                option_clean = (

                    option.strip().lower()

                )



                if (

                    option_clean

                    == month_normalized.lower()

                    or option_clean.startswith(

                        month_normalized.lower()

                    )

                ):

                    month_found = True

                    break



            # -----------------------------

            # Year detection

            # -----------------------------



            year_found = False



            for option in option_texts:



                if (

                    option.strip()

                    == year_normalized

                ):

                    year_found = True

                    break



            if (

                month_found

                and month_select is None

            ):

                month_select = select



                logger.info(

                    f"Month dropdown detected "

                    f"at index {i}"

                )



            if (

                year_found

                and year_select is None

            ):

                year_select = select



                logger.info(

                    f"Year dropdown detected "

                    f"at index {i}"

                )



        except Exception as e:



            logger.debug(

                f"Could not inspect dropdown "

                f"{i}: {e}"

            )



    # -----------------------------------------

    # Validate

    # -----------------------------------------



    if month_select is None:

        raise RuntimeError(

            f"Month dropdown not found for "

            f"'{month_normalized}'."

        )



    if year_select is None:

        raise RuntimeError(

            f"Year dropdown not found for "

            f"'{year_normalized}'."

        )



    # =========================================

    # SELECT MONTH

    # =========================================



    month_options = await month_select.locator(

        "option"

    ).evaluate_all(

        """

        options => options.map(o => ({

            text: (o.textContent || '').trim(),

            value: o.value

        }))

        """

    )



    month_value = None



    for option in month_options:



        option_text = (

            option["text"]

            .strip()

            .lower()

        )



        if (

            option_text

            == month_normalized.lower()

            or option_text.startswith(

                month_normalized.lower()

            )

        ):

            month_value = option["value"]

            break



    if month_value is None:

        raise RuntimeError(

            f"Month '{month_normalized}' "

            "not available in dropdown."

        )



    await month_select.select_option(

        value=month_value

    )



    logger.info(

        f"Month selected: {month_normalized}"

    )



    await page.wait_for_timeout(1500)



    # =========================================

    # SELECT YEAR

    # =========================================



    year_options = await year_select.locator(

        "option"

    ).evaluate_all(

        """

        options => options.map(o => ({

            text: (o.textContent || '').trim(),

            value: o.value

        }))

        """

    )



    year_value = None



    for option in year_options:



        if (

            option["text"].strip()

            == year_normalized

        ):

            year_value = option["value"]

            break



    if year_value is None:

        raise RuntimeError(

            f"Year '{year_normalized}' "

            "not available in dropdown."

        )



    await year_select.select_option(

        value=year_value

    )



    logger.info(

        f"Year selected: {year_normalized}"

    )



    await page.wait_for_timeout(1500)



    logger.info(

        "ESIC Contribution Period selected "

        f"successfully: {month_normalized}-{year_normalized}"

    )





# ============================================================

# SELECT SUPPLEMENTARY CONTRIBUTION + SUBMIT

# ============================================================



async def select_supplementary_and_submit(
    page,
    month,
    year,
):
    """
    ESIC initial contribution-page flow.

    Month  -> Excel Column D
    Year   -> Excel Column E

    After Year is selected:
        1 Tab
        1 ArrowDown
        4 Tab
        1 ArrowDown
        1 Tab
        Mouse click Submit
        Wait 2 seconds
        1 Tab

    The final 1 Tab leaves focus on the Insurance Number field,
    which is then handled by fill_payroll_details_and_final_submit().
    """

    logger.info("Starting ESIC Monthly Contribution form...")

    await select_month_year(
        page,
        month,
        year,
    )

    await page.wait_for_timeout(1000)

    # IMPORTANT:
    # select_option() can leave focus on Month on this old ESIC page.
    # Explicitly focus Year before starting the keyboard sequence.
    selects = page.locator("select:visible")

    if await selects.count() < 2:
        raise RuntimeError("Month/Year dropdowns were not found.")

    year_select = selects.nth(1)
    await year_select.focus()

    logger.info("Year dropdown focused.")

    # 1 Tab
    await page.keyboard.press("Tab")
    await page.wait_for_timeout(300)
    logger.info("Keyboard step 1: TAB")

    # 1 Down
    await page.keyboard.press("ArrowDown")
    await page.wait_for_timeout(300)
    logger.info("Keyboard step 2: ARROW DOWN")

    # 4 Tabs
    for tab_number in range(1, 5):
        await page.keyboard.press("Tab")
        await page.wait_for_timeout(250)
        logger.info(f"Keyboard step: TAB {tab_number}/4")

    # 1 Down
    await page.keyboard.press("ArrowDown")
    await page.wait_for_timeout(3000)
    logger.info("Keyboard step: ARROW DOWN after 4 TABs; waited 3 seconds")

    # 1 Tab
    await page.keyboard.press("Tab")
    await page.wait_for_timeout(300)
    logger.info("Keyboard step: final TAB before initial Submit")

    # Ensure any additional dropdown (e.g. Type of Wages) is selected before Submit
    all_selects = page.locator("select:visible")
    sel_count = await all_selects.count()
    if sel_count > 2:
        for s_idx in range(2, sel_count):
            try:
                curr_sel = all_selects.nth(s_idx)
                val = await curr_sel.input_value()
                if not val or val in ("0", "-1", ""):
                    options = await curr_sel.locator("option").all()
                    if len(options) > 1:
                        opt_val = await options[1].get_attribute("value")
                        if opt_val:
                            await curr_sel.select_option(value=opt_val)
                            logger.info(f"Auto-selected dropdown {s_idx + 1} with option: {opt_val}")
                            await page.wait_for_timeout(1000)
            except Exception as e:
                logger.debug(f"Optional wage dropdown handling: {e}")

    # Initial Submit: actual ESIC input shown in DevTools.
    submit_button = page.locator(
        'input#ctl00_HomePageContent_CtrlBtnSubmit, '
        'input[name="ctl00$HomePageContent$CtrlBtnSubmit"], '
        'input[type="submit"][value="Submit"]'
    ).first

    if await submit_button.count() == 0:
        submit_button = page.get_by_role(
            "button",
            name="Submit",
            exact=True,
        ).first

    if await submit_button.count() == 0:
        raise RuntimeError("Initial Submit button not found.")

    # Mouse click as requested.
    await submit_button.click()
    logger.info("Initial Submit mouse click completed.")

    # Wait for ESIC postback/page update or navigation to Supplementary contribution grid
    for wait_cycle in range(6):
        await page.wait_for_timeout(1000)
        curr_url = page.url
        grid_count = await page.locator('[id*="CtrlTxtBxIpNo"]').count()
        if "suppliMonthlyContribution.aspx" in curr_url or grid_count > 0:
            logger.info("Successfully transitioned to Supplementary Contribution entry page.")
            break

        # If still on Home and dropdown 3 was missed, select it and click submit again
        wage_selects = page.locator("select:visible")
        if await wage_selects.count() > 2:
            try:
                last_sel = wage_selects.nth(2)
                l_val = await last_sel.input_value()
                if not l_val or l_val in ("0", "-1", ""):
                    opts = await last_sel.locator("option").all()
                    if len(opts) > 1:
                        await last_sel.select_option(index=1)
                        logger.info("Selected Wage Type dropdown option 1 and re-clicking submit...")
                        await page.wait_for_timeout(1000)
                        await submit_button.click()
                        await page.wait_for_timeout(2000)
            except Exception:
                pass

    try:
        await page.wait_for_load_state("domcontentloaded", timeout=10000)
    except Exception:
        pass

    # Wait for grid row 1 Insurance Number field to appear
    try:
        await page.locator('[id*="CtrlTxtBxIpNo"]').first.wait_for(state="visible", timeout=8000)
        logger.info("Supplementary Contribution grid is ready.")
    except Exception:
        pass

    logger.info("Initial ESIC contribution step completed.")


# ============================================================
# APPROVED PAYROLL EXCEL HELPERS
# ============================================================


def normalize_excel_header(value):
    if value is None:
        return ""

    return re.sub(
        r"[^a-z0-9]",
        "",
        str(value).strip().lower(),
    )


def find_approved_payroll_file(month, year):
    """
    Find the Approved Payroll Excel for the contribution period.
    Supports:
      1. Specific Month/Year files like 'MAY 2026.xlsx'
      2. Exact filename configured in APPROVED_PAYROLL_FILENAME
      3. Any XLSX in APPROVED_PAYROLL_DIR matching month/year
    """
    month_text = clean_excel_val(month).strip().lower()
    year_text = clean_excel_val(year).strip()

    month_map = {
        "1": ["jan", "january"], "01": ["jan", "january"],
        "2": ["feb", "february"], "02": ["feb", "february"],
        "3": ["mar", "march"], "03": ["mar", "march"],
        "4": ["apr", "april"], "04": ["apr", "april"],
        "5": ["may"], "05": ["may"],
        "6": ["jun", "june"], "06": ["jun", "june"],
        "7": ["jul", "july"], "07": ["jul", "july"],
        "8": ["aug", "august"], "08": ["aug", "august"],
        "9": ["sep", "september"], "09": ["sep", "september"],
        "10": ["oct", "october"], "11": ["nov", "november"],
        "12": ["dec", "december"],
    }
    month_tokens = month_map.get(month_text, [month_text])

    search_dirs = [
        APPROVED_PAYROLL_DIR,
        APPROVED_PAYROLL_DIR_FALLBACK,
        DATA_FOLDER,
    ]

    # Priority 1: Explicitly configured APPROVED_PAYROLL_FILENAME supplied by user
    if APPROVED_PAYROLL_FILENAME:
        for s_dir in search_dirs:
            exact = s_dir / APPROVED_PAYROLL_FILENAME
            if exact.exists() and exact.is_file():
                logger.info(f"SELECTED explicitly configured Approved Payroll Excel: {exact}")
                return exact

    candidates = []
    for s_dir in search_dirs:
        if not s_dir.exists():
            continue
        for f_path in s_dir.glob("*.xlsx"):
            f_name_lower = f_path.name.lower()
            if "esic_users" in f_name_lower or "~$" in f_name_lower:
                continue
            has_year = (not year_text) or (year_text in f_name_lower)
            has_month = any(tok in f_name_lower for tok in month_tokens)
            if has_year and has_month:
                candidates.append(f_path)

    # Fallback to any non-users xlsx in search_dirs
    if not candidates:
        for s_dir in search_dirs:
            if not s_dir.exists():
                continue
            for f_path in s_dir.glob("*.xlsx"):
                f_name_lower = f_path.name.lower()
                if "esic_users" in f_name_lower or "~$" in f_name_lower:
                    continue
                candidates.append(f_path)

    if not candidates:
        raise FileNotFoundError(
            f"No Approved Payroll Excel found for Month={month_text}, Year={year_text}."
        )

    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    selected = candidates[0]
    logger.info(f"SELECTED Approved Payroll Excel: {selected}")
    return selected


def _find_column(headers, priority_exact_list, contains_token_lists=None):
    """Find a column using normalized header names with priority."""
    for name in priority_exact_list:
        norm = normalize_excel_header(name)
        if norm in headers:
            return headers[norm]
    if contains_token_lists:
        for tokens in contains_token_lists:
            norm_tokens = [normalize_excel_header(t) for t in tokens]
            for h_norm, col in headers.items():
                if all(t in h_norm for t in norm_tokens):
                    return col
    return None


def get_approved_payroll_values(payroll_file, loginid=None, include_done=False):
    """
    Read the approved payroll Excel dynamically for the 3 portal fields:
      - Insurance Number (ESIC Number)
      - Paid Days
      - ESIC Gross Payable (or Total Monthly Wages)
    """
    wb = load_workbook(payroll_file, data_only=True)
    ws = wb.active

    header_row = 1
    headers = {
        normalize_excel_header(ws.cell(header_row, col).value): col
        for col in range(1, ws.max_column + 1)
        if ws.cell(header_row, col).value is not None
    }

    esic_col = _find_column(
        headers,
        ["ESIC Number", "Insurance Number", "IP Number", "ESIC No", "ESIC No."],
        [["esic", "number"], ["insurance", "number"]],
    )
    paid_days_col = _find_column(
        headers,
        ["Paid Days", "Paid Day", "ESIC Paid Days"],
        [["paid", "day"]],
    )
    gross_payable_col = _find_column(
        headers,
        ["ESIC Gross Payable", "Gross Payable", "Total Monthly Wages", "EE Gross Value", "Gross Value"],
        [["esic", "gross"], ["gross", "payable"], ["gross"]],
    )

    # Positional fallback if headers were missing or custom
    if esic_col is None:
        esic_col = 3
    if paid_days_col is None:
        paid_days_col = 12 if ws.max_column >= 12 else 4
    if gross_payable_col is None:
        gross_payable_col = 15 if ws.max_column >= 15 else 5

    status_col = get_status_column_index(ws)

    logger.info(
        f"Approved Payroll columns: ESIC={esic_col}, PaidDays={paid_days_col}, "
        f"Gross={gross_payable_col}, StatusCol={status_col}"
    )

    payroll_rows = []
    skipped_done = 0
    for row in range(header_row + 1, ws.max_row + 1):
        insurance_number = clean_excel_val(ws.cell(row, esic_col).value)
        if not insurance_number:
            continue

        status_val = clean_excel_val(ws.cell(row, status_col).value).strip()
        is_done = status_val.upper() == "DONE" or status_val.upper().startswith("DONE")

        if is_done and not include_done:
            skipped_done += 1
            continue

        paid_days = clean_excel_val(ws.cell(row, paid_days_col).value)
        gross_payable = clean_excel_val(ws.cell(row, gross_payable_col).value)

        payroll_rows.append({
            "insurance_number": insurance_number,
            "paid_days": paid_days or "0",
            "gross_payable": gross_payable or "0",
            "payroll_row": row,
            "status": status_val or "NOT DONE",
        })

    wb.close()

    if skipped_done > 0:
        logger.info(f"Skipped {skipped_done} employee row(s) already marked DONE in Excel.")

    logger.info(
        f"Pending Approved Payroll employee rows: {len(payroll_rows)}"
    )

    for item in payroll_rows[:5]:
        logger.info(
            f"Payroll row {item['payroll_row']} | "
            f"ESIC={item['insurance_number']} | "
            f"Days={item['paid_days']} | "
            f"Gross={item['gross_payable']} | "
            f"Status={item['status']}"
        )
    if len(payroll_rows) > 5:
        logger.info(f"... and {len(payroll_rows) - 5} more pending employees.")

    return payroll_rows


def get_status_column_index(ws):
    """Find or create Entry Status column in the payroll Excel."""
    for col in range(1, ws.max_column + 1):
        val = str(ws.cell(1, col).value or "").strip().lower()
        if val in ("entry status", "submission status", "portal status", "upload status", "done status"):
            return col
    new_col = ws.max_column + 1
    ws.cell(1, new_col).value = "Entry Status"
    return new_col


def save_workbook_safely(wb, file_path, max_retries=3):
    """
    Saves an openpyxl workbook safely. If Microsoft Excel has the file open (PermissionError),
    closes the locking Excel process to ensure changes are written without crashing.
    """
    import subprocess
    import time
    for attempt in range(max_retries):
        try:
            wb.save(file_path)
            return True
        except PermissionError:
            logger.warning(
                f"[FILE LOCKED] {Path(file_path).name} is open in Excel (PermissionError). "
                f"Closing Excel to save status updates (attempt {attempt + 1}/{max_retries})..."
            )
            try:
                subprocess.run(["taskkill", "/f", "/im", "excel.exe"], capture_output=True)
                time.sleep(1.5)
            except Exception:
                pass
        except Exception as e:
            logger.error(f"Error saving {file_path}: {e}")
            break
    try:
        wb.save(file_path)
        return True
    except Exception as e:
        logger.error(f"Final failure saving {file_path}: {e}")
        return False


def mark_payroll_rows_status(payroll_file, row_indices, status_value):
    """
    Updates the Entry Status column for specified row numbers in the payroll workbook.
    """
    if not row_indices:
        return
    try:
        wb = load_workbook(payroll_file)
        ws = wb.active
        status_col = get_status_column_index(ws)
        for r in row_indices:
            ws.cell(r, status_col).value = status_value
        success = save_workbook_safely(wb, payroll_file)
        wb.close()
        if success:
            logger.info(
                f"Updated {len(row_indices)} rows to '{status_value}' in {Path(payroll_file).name} (Col {status_col})."
            )
        else:
            logger.error(f"Could not save status in {payroll_file}.")
    except Exception as exc:
        logger.error(f"Failed to update status in {payroll_file}: {exc}")


# ============================================================
# SINGLE-SHEET (UNIFIED) MODE HELPERS
# ============================================================

def get_single_sheet_columns(ws):
    """
    Detects column numbers from header row (row 1) in single sheet format.
    Supports flexible header naming.
    """
    mapping = {
        "esic_number": None,
        "paid_days": None,
        "gross_payable": None,
        "loginid": None,
        "password": None,
        "response": None,
        "month": None,
        "year": None,
        "status": None,
    }
    for col_idx in range(1, ws.max_column + 1):
        val = ws.cell(1, col_idx).value
        if not val:
            continue
        cleaned = re.sub(r"[^a-z0-9]", "", str(val).strip().lower())
        if cleaned in ("esicnumber", "esicno", "ipnumber", "ipno", "insurance", "insurancenumber"):
            mapping["esic_number"] = col_idx
        elif cleaned in ("paiddays", "days", "workingdays", "noofdays"):
            mapping["paid_days"] = col_idx
        elif cleaned in ("basicpayable", "grosswages", "totalmonthlywages", "monthlywages", "wages", "gross", "basic"):
            mapping["gross_payable"] = col_idx
        elif cleaned in ("loginid", "login", "username", "userid", "user"):
            mapping["loginid"] = col_idx
        elif cleaned in ("password", "pass", "pwd"):
            mapping["password"] = col_idx
        elif cleaned in ("response", "res", "result"):
            mapping["response"] = col_idx
        elif cleaned in ("month", "mon"):
            mapping["month"] = col_idx
        elif cleaned in ("year", "yr"):
            mapping["year"] = col_idx
        elif cleaned in ("updatestatus", "status", "entrystatus"):
            mapping["status"] = col_idx

    # Positional defaults for standard 9-column format
    if not mapping["esic_number"] and ws.max_column >= 9:
        mapping["esic_number"] = 1
        mapping["paid_days"] = 2
        mapping["gross_payable"] = 3
        mapping["loginid"] = 4
        mapping["password"] = 5
        mapping["response"] = 6
        mapping["month"] = 7
        mapping["year"] = 8
        mapping["status"] = 9

    return mapping


def is_single_sheet_mode(excel_path=EXCEL_FILE, sheet_name=SHEET_NAME):
    """
    Checks if excel_path has employee columns (ESIC Number, Paid Days, Basic Payable).
    If so, it runs in single-sheet mode where credentials and employee data are unified.
    """
    if not Path(excel_path).exists():
        return False
    try:
        wb = load_workbook(excel_path, read_only=True)
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        cols = get_single_sheet_columns(ws)
        wb.close()
        return bool(cols.get("esic_number") and cols.get("paid_days") and cols.get("gross_payable"))
    except Exception:
        return False


def get_single_sheet_pending_rows(excel_path=EXCEL_FILE, loginid=None, month=None, year=None, sheet_name=SHEET_NAME):
    """
    Returns pending employee rows from the single sheet for a given loginid, month, year.
    A row is pending if:
      - loginid, month, year match
      - Update Status is not 'DONE' and not 'UPDATED'
      - response does not start with 'SUCCESS'
      - Insurance number is valid
    """
    wb = load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
    cols = get_single_sheet_columns(ws)

    pending = []
    target_login = str(loginid).strip() if loginid is not None else None
    target_month = str(month).strip().lower() if month is not None else None
    target_year = str(year).strip() if year is not None else None

    last_loginid = ""
    last_month = ""
    last_year = ""

    for r in range(2, ws.max_row + 1):
        r_loginid = clean_excel_val(ws.cell(r, cols["loginid"]).value).strip() if cols["loginid"] else ""
        r_month = clean_excel_val(ws.cell(r, cols["month"]).value).strip().lower() if cols["month"] else ""
        r_year = clean_excel_val(ws.cell(r, cols["year"]).value).strip() if cols["year"] else ""

        if r_loginid:
            last_loginid = r_loginid
        else:
            r_loginid = last_loginid

        if r_month:
            last_month = r_month
        else:
            r_month = last_month

        if r_year:
            last_year = r_year
        else:
            r_year = last_year

        if target_login and r_loginid != target_login:
            continue
        if target_month and r_month != target_month:
            continue
        if target_year and r_year != target_year:
            continue

        status_val = clean_excel_val(ws.cell(r, cols["status"]).value).strip().upper() if cols["status"] else ""
        resp_val = clean_excel_val(ws.cell(r, cols["response"]).value).strip().upper() if cols["response"] else ""

        if status_val in ("DONE", "UPDATED") or resp_val.startswith("SUCCESS"):
            continue

        ip = clean_excel_val(ws.cell(r, cols["esic_number"]).value).strip() if cols["esic_number"] else ""
        if not ip:
            continue

        days = clean_excel_val(ws.cell(r, cols["paid_days"]).value).strip() if cols["paid_days"] else "0"
        wages = clean_excel_val(ws.cell(r, cols["gross_payable"]).value).strip() if cols["gross_payable"] else "0"

        pending.append({
            "insurance_number": ip,
            "paid_days": days or "0",
            "gross_payable": wages or "0",
            "payroll_row": r,
            "status": status_val or "NOT DONE",
        })

    wb.close()
    return pending


def mark_single_sheet_rows_status(excel_path, row_indices, status_value="DONE", sheet_name=SHEET_NAME):
    """
    Updates the Update Status column for specified row indices in the single sheet.
    """
    if not row_indices:
        return
    try:
        wb = load_workbook(excel_path)
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        cols = get_single_sheet_columns(ws)
        status_col = cols.get("status") or 9
        for r in row_indices:
            ws.cell(r, status_col).value = status_value
        success = save_workbook_safely(wb, excel_path)
        wb.close()
        if success:
            logger.info(
                f"Updated {len(row_indices)} rows to '{status_value}' in {Path(excel_path).name} (Col {status_col})."
            )
        else:
            logger.error(f"Could not save status in {excel_path}.")
    except Exception as exc:
        logger.error(f"Failed to update single sheet status in {excel_path}: {exc}")


def mark_single_sheet_group_response(excel_path, row_indices, response_value, status_value="UPDATED", sheet_name=SHEET_NAME):
    """
    Updates both response column and Update Status column for a group of rows in the single sheet.
    """
    if not row_indices:
        return
    try:
        wb = load_workbook(excel_path)
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        cols = get_single_sheet_columns(ws)
        resp_col = cols.get("response") or 6
        status_col = cols.get("status") or 9
        for r in row_indices:
            ws.cell(r, resp_col).value = response_value
            if status_value:
                ws.cell(r, status_col).value = status_value
        success = save_workbook_safely(wb, excel_path)
        wb.close()
        if success:
            logger.info(
                f"Updated response to '{response_value}' for {len(row_indices)} rows in {Path(excel_path).name}."
            )
        else:
            logger.error(f"Could not save group response in {excel_path}.")
    except Exception as exc:
        logger.error(f"Failed to update group response in {excel_path}: {exc}")


def read_single_sheet_users(excel_path=EXCEL_FILE, sheet_name=SHEET_NAME):
    """
    Reads distinct (loginid, password, month, year) groups from the single sheet.
    Returns a list of tuples: (first_row, loginid, password, month, year, group_row_indices)
    """
    wb = load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
    cols = get_single_sheet_columns(ws)

    groups = {}  # key: (loginid, password, month, year) -> list of row indices
    last_loginid = ""
    last_password = ""
    last_month = ""
    last_year = ""

    for r in range(2, ws.max_row + 1):
        raw_login = clean_excel_val(ws.cell(r, cols["loginid"]).value).strip() if cols["loginid"] else ""
        raw_pwd = clean_excel_val(ws.cell(r, cols["password"]).value).strip() if cols["password"] else ""
        raw_mon = clean_excel_val(ws.cell(r, cols["month"]).value).strip().lower() if cols["month"] else ""
        raw_yr = clean_excel_val(ws.cell(r, cols["year"]).value).strip() if cols["year"] else ""

        if raw_login:
            last_loginid = raw_login
        else:
            raw_login = last_loginid

        if raw_pwd:
            last_password = raw_pwd
        else:
            raw_pwd = last_password

        if raw_mon:
            last_month = raw_mon
        else:
            raw_mon = last_month

        if raw_yr:
            last_year = raw_yr
        else:
            raw_yr = last_year

        if not raw_login or not raw_pwd or not raw_mon or not raw_yr:
            continue

        ip_val = clean_excel_val(ws.cell(r, cols["esic_number"]).value).strip() if cols["esic_number"] else ""
        if not ip_val:
            continue

        key = (raw_login, raw_pwd, raw_mon, raw_yr)
        if key not in groups:
            groups[key] = []
        groups[key].append(r)

    wb.close()

    result_users = []
    for (loginid, password, month, year), row_indices in groups.items():
        first_row = row_indices[0]
        result_users.append((first_row, loginid, password, month, year, row_indices))

    logger.info(f"Single-Sheet Mode: Found {len(result_users)} unique user/period group(s) across {sum(len(r) for r in groups.values())} employee rows.")
    return result_users


BATCH_SIZE = 8

# ============================================================
# ENTRY SPEED & TIMING CONFIGURATION (ENTRY SPEED TUNING)
# ============================================================
# ESIC portal allows a maximum of 10 rows per batch on Supplementary form.
# BATCH_SIZE = 10 ensures that the 10th record submits the batch without getting stuck on 'Add More'.
IP_WAIT_MS = 800         # Time (ms) after Insurance No (safe for ESIC AJAX IP name fetch)
DAYS_WAIT_MS = 200       # Time (ms) after Paid Days (instant local DOM input)
WAGES_WAIT_MS = 250      # Time (ms) after Monthly Wages (local event registration)
ADD_MORE_POLL_MS = 300   # Polling interval (ms) checking if new row appeared
ADD_MORE_MAX_WAIT_S = 8  # Max seconds to wait for new row after clicking 'Add More'


async def wait_for_aspnet_ajax(page, timeout_ms=4000):
    """
    Waits for ASP.NET AJAX PageRequestManager async postback to complete, if active.
    Prevents event collisions when ESIC fetches IP names or renders new rows.
    """
    try:
        start_t = asyncio.get_event_loop().time()
        await page.wait_for_timeout(200)
        while (asyncio.get_event_loop().time() - start_t) * 1000 < timeout_ms:
            is_busy = await page.evaluate("""
                () => {
                    try {
                        if (window.Sys && Sys.WebForms && Sys.WebForms.PageRequestManager) {
                            return Sys.WebForms.PageRequestManager.getInstance().get_isInAsyncPostBack();
                        }
                    } catch (e) {}
                    return false;
                }
            """)
            if not is_busy:
                break
            await page.wait_for_timeout(150)
    except Exception:
        pass


async def get_grid_row_count(page):
    """
    Accurately counts current data rows in Supplementary Contribution form.
    In ASP.NET GridView, each employee data row contains exactly one Insurance Number input.
    Header rows, Delete buttons, and footer submit rows are handled cleanly without discarding valid rows.
    """
    try:
        return await page.evaluate("""
            () => {
                // 1. Primary: Count all VISIBLE Insurance Number text inputs (exclude hidden fields like CtrlHdnFlsIpNo)
                const ipInputs = Array.from(document.querySelectorAll('input[id*="CtrlTxtBxIpNo"], input[id*="IpNo" i]'))
                    .filter(inp => inp.type !== 'hidden' && !inp.id.toLowerCase().includes('hdn'));
                if (ipInputs.length > 0) return ipInputs.length;

                // 2. Fallback: Count table rows in grid excluding header and footer
                const grid = document.querySelector('table[id*="CtrlGrdVwSuppMc"]') 
                          || document.querySelector('table#tblnew') 
                          || document.querySelector('.internalTable');
                if (!grid) return 0;

                const rows = Array.from(grid.querySelectorAll('tr')).filter(tr => {
                    if (tr.querySelector('th')) return false;
                    // Exclude footer row containing Add More or Submit
                    if (tr.querySelector('input[value*="Add More" i], input[value*="Submit" i], input[id*="ButtonAdd" i], input[id*="CtrlBtnSubmit" i]')) return false;
                    const inps = tr.querySelectorAll('input:not([type="hidden"]):not([type="checkbox"])');
                    return inps.length >= 1;
                });
                return rows.length;
            }
        """)
    except Exception:
        return 0


async def enter_employee_row(page, insurance_number, paid_days, gross_payable, row_idx):
    """
    Fills all 3 fields for an employee row on ESIC Supplementary Contribution page:
      1. Insurance Number -> wait for IP fetch postback & IP_WAIT_MS
      2. No. of Days for Which Wages Paid / Payable -> wait DAYS_WAIT_MS (200ms)
      3. Total Monthly Wages -> wait WAGES_WAIT_MS (250ms)
    Uses JavaScript DOM inspection to identify exact element IDs for row_idx.
    Guarantees row_idx is targeted accurately, preventing row overwrites.
    """
    logger.info(
        f"Row {row_idx} Entry | ESIC={insurance_number} | PaidDays={paid_days} | Wages={gross_payable}"
    )

    # Wait for target row to actually exist in the DOM
    row_info = {"success": False}
    for wait_i in range(25):
        try:
            row_info = await page.evaluate("""
                (rowIdx) => {
                    // 1. Primary: Locate all VISIBLE IP inputs across the form (strictly exclude CtrlHdnFlsIpNo)
                    const ipInputs = Array.from(document.querySelectorAll('input[id*="CtrlTxtBxIpNo"], input[id*="IpNo" i]'))
                        .filter(inp => inp.type !== 'hidden' && !inp.id.toLowerCase().includes('hdn'));
                    const totalRows = ipInputs.length;

                    if (totalRows < rowIdx) {
                        return { success: false, not_ready: true, totalRows: totalRows };
                    }

                    const targetIp = ipInputs[rowIdx - 1];
                    const targetRow = targetIp.closest('tr');

                    let daysId = null;
                    let wageId = null;

                    if (targetRow) {
                        // Priority 1: Match by ID/name pattern in target row (must not be hidden)
                        const daysInp = targetRow.querySelector(
                            'input:not([type="hidden"])[id*="CtrlTxtBxNoDaysWokd"], ' +
                            'input:not([type="hidden"])[id*="Days" i], ' +
                            'input:not([type="hidden"])[id*="Wokd" i]'
                        );
                        if (daysInp && !daysInp.id.toLowerCase().includes('hdn')) daysId = daysInp.id;

                        const wageInp = targetRow.querySelector(
                            'input:not([type="hidden"])[id*="CtrlTxtBxTotMnthlyWgs"], ' +
                            'input:not([type="hidden"])[id*="Wage" i], ' +
                            'input:not([type="hidden"])[id*="Wgs" i], ' +
                            'input:not([type="hidden"])[id*="Tot" i]'
                        );
                        if (wageInp && !wageInp.id.toLowerCase().includes('hdn')) wageId = wageInp.id;

                        // Priority 2: Fallback to non-button editable inputs in this row
                        if (!daysId || !wageId) {
                            const rowInps = Array.from(targetRow.querySelectorAll('input:not([type="hidden"]):not([type="checkbox"]):not([type="submit"]):not([type="button"]):not([type="image"])'))
                                .filter(inp => !inp.id.toLowerCase().includes('hdn'));
                            const ipIdx = rowInps.indexOf(targetIp);
                            if (!daysId && rowInps.length > ipIdx + 1) {
                                daysId = rowInps[ipIdx + 1].id;
                            }
                            if (!wageId && rowInps.length > ipIdx + 2) {
                                wageId = rowInps[ipIdx + 2].id;
                            }
                        }
                    }

                    return {
                        success: true,
                        totalRows: totalRows,
                        ipId: targetIp.id || null,
                        daysId: daysId,
                        wageId: wageId,
                    };
                }
            """, row_idx)
            if row_info.get("success"):
                break
        except Exception:
            pass
        await page.wait_for_timeout(300)

    if not row_info.get("success"):
        logger.warning(f"Row {row_idx} inspection warning: {row_info.get('error', 'Not ready')}. Using fallback selectors.")
        ip_id = None
        days_id = None
        wage_id = None
    else:
        ip_id = row_info.get("ipId")
        days_id = row_info.get("daysId")
        wage_id = row_info.get("wageId")
        logger.info(f"Row {row_idx} element IDs: IP={ip_id}, Days={days_id}, Wages={wage_id}")

    # ----------------------------------------------------
    # 1. INSURANCE NUMBER
    # ----------------------------------------------------
    if ip_id:
        ip_el = page.locator(f"#{ip_id}").first
    else:
        ip_locator = page.locator('input:not([type="hidden"]):not([id*="Hdn" i])[id*="CtrlTxtBxIpNo"], input:not([type="hidden"]):not([id*="Hdn" i])[id*="IpNo" i]')
        if await ip_locator.count() >= row_idx:
            ip_el = ip_locator.nth(row_idx - 1)
        else:
            ip_el = ip_locator.last

    try:
        await ip_el.scroll_into_view_if_needed(timeout=3000)
    except Exception:
        pass
    await ip_el.fill(str(insurance_number))
    await ip_el.press("Tab")
    try:
        await page.evaluate("""
            ([id, val]) => {
                let el = id ? document.getElementById(id) : null;
                if (!el || el.type === 'hidden' || (el.id||'').toLowerCase().includes('hdn')) {
                    const inputs = Array.from(document.querySelectorAll('input:not([type="hidden"]):not([type="checkbox"]):not([type="submit"]):not([type="button"])'))
                        .filter(inp => !(inp.id||'').toLowerCase().includes('hdn'));
                    const ipInputs = inputs.filter(inp => (inp.id||'').toLowerCase().includes('ip') || (inp.name||'').toLowerCase().includes('ip'));
                    if (ipInputs.length > 0) el = ipInputs[ipInputs.length - 1];
                }
                if (el) {
                    el.value = String(val);
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    el.dispatchEvent(new Event('blur', { bubbles: true }));
                }
            }
        """, [ip_id or "", str(insurance_number)])
    except Exception:
        pass
    logger.info(f"Row {row_idx} | Insurance Number entered: {insurance_number}")

    # Wait for ESIC AJAX IP name fetch postback to complete
    await wait_for_aspnet_ajax(page, timeout_ms=3000)
    await page.wait_for_timeout(IP_WAIT_MS)

    # ----------------------------------------------------
    # 2. PAID DAYS
    # ----------------------------------------------------
    if days_id:
        days_el = page.locator(f"#{days_id}").first
    else:
        days_locator = page.locator('input:not([type="hidden"]):not([id*="Hdn" i])[id*="CtrlTxtBxNoDaysWokd"], input:not([type="hidden"]):not([id*="Hdn" i])[id*="Days" i], input:not([type="hidden"]):not([id*="Hdn" i])[id*="Wokd" i]')
        if await days_locator.count() >= row_idx:
            days_el = days_locator.nth(row_idx - 1)
        else:
            days_el = days_locator.last

    try:
        await days_el.scroll_into_view_if_needed(timeout=3000)
    except Exception:
        pass
    await days_el.fill(str(paid_days))
    await days_el.press("Tab")
    try:
        await page.evaluate("""
            ([id, val]) => {
                let el = id ? document.getElementById(id) : null;
                if (!el || el.type === 'hidden' || (el.id||'').toLowerCase().includes('hdn')) {
                    const inputs = Array.from(document.querySelectorAll('input:not([type="hidden"]):not([type="checkbox"]):not([type="submit"]):not([type="button"])'))
                        .filter(inp => !(inp.id||'').toLowerCase().includes('hdn'));
                    const daysInputs = inputs.filter(inp => (inp.id||'').toLowerCase().includes('days') || (inp.name||'').toLowerCase().includes('days') || (inp.id||'').toLowerCase().includes('wokd'));
                    if (daysInputs.length > 0) el = daysInputs[daysInputs.length - 1];
                }
                if (el) {
                    el.value = String(val);
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    el.dispatchEvent(new Event('blur', { bubbles: true }));
                }
            }
        """, [days_id or "", str(paid_days)])
    except Exception:
        pass
    logger.info(f"Row {row_idx} | Paid Days entered: {paid_days}")

    # Fast wait (local input)
    await page.wait_for_timeout(DAYS_WAIT_MS)

    # ----------------------------------------------------
    # 3. TOTAL MONTHLY WAGES
    # ----------------------------------------------------
    if wage_id:
        wage_el = page.locator(f"#{wage_id}").first
    else:
        wage_locator = page.locator('input:not([type="hidden"]):not([id*="Hdn" i])[id*="CtrlTxtBxTotMnthlyWgs"], input:not([type="hidden"]):not([id*="Hdn" i])[id*="Wage" i], input:not([type="hidden"]):not([id*="Hdn" i])[id*="Wgs" i], input:not([type="hidden"]):not([id*="Hdn" i])[id*="Tot" i]')
        if await wage_locator.count() >= row_idx:
            wage_el = wage_locator.nth(row_idx - 1)
        else:
            wage_el = wage_locator.last

    try:
        await wage_el.scroll_into_view_if_needed(timeout=3000)
    except Exception:
        pass
    await wage_el.fill(str(gross_payable))
    await wage_el.press("Tab")
    
    try:
        await page.evaluate("""
            ([id, val]) => {
                let el = id ? document.getElementById(id) : null;
                if (!el || el.type === 'hidden' || (el.id||'').toLowerCase().includes('hdn')) {
                    const inputs = Array.from(document.querySelectorAll('input:not([type="hidden"]):not([type="checkbox"]):not([type="submit"]):not([type="button"])'))
                        .filter(inp => !(inp.id||'').toLowerCase().includes('hdn'));
                    const wageInputs = inputs.filter(inp => (inp.id||'').toLowerCase().includes('wage') || (inp.name||'').toLowerCase().includes('wage') || (inp.id||'').toLowerCase().includes('tot') || (inp.id||'').toLowerCase().includes('wgs'));
                    if (wageInputs.length > 0) el = wageInputs[wageInputs.length - 1];
                }
                if (el) {
                    el.value = String(val);
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    el.dispatchEvent(new Event('blur', { bubbles: true }));
                }
            }
        """, [wage_id or "", str(gross_payable)])
    except Exception:
        pass
    logger.info(f"Row {row_idx} | Total Monthly Wages entered: {gross_payable}")

    # Fast wait (local validation)
    await page.wait_for_timeout(WAGES_WAIT_MS)


async def click_add_more_and_wait(page, expected_min_rows=None):
    # Ensure any pending postback has completed before clicking Add More
    await wait_for_aspnet_ajax(page, timeout_ms=3000)

    # Get current row count before click safely
    prev_rows = await get_grid_row_count(page)
    target_rows = expected_min_rows if expected_min_rows else (prev_rows + 1)

    async def _do_click_add_more():
        selectors = [
            'input#ctl00_HomePageContent_ButtonAdd',
            'input[name="ctl00$HomePageContent$ButtonAdd"]',
            'input[type="submit"][value="Add More"]',
            'input[value="Add More"]',
            'input[value*="Add More" i]',
            'button:has-text("Add More")',
            'a:has-text("Add More")',
            'a#ctl00_HomePageContent_ButtonAdd',
        ]
        for sel in selectors:
            cand = page.locator(sel).first
            try:
                if await cand.count() > 0 and await cand.is_visible() and await cand.is_enabled():
                    await cand.scroll_into_view_if_needed(timeout=2000)
                    await cand.click(timeout=3000)
                    logger.info(f"Clicked 'Add More' button using selector: {sel}")
                    return True
            except Exception:
                continue

        # Fallback click via JS
        try:
            clicked_js = await page.evaluate("""
                () => {
                    const btn = document.getElementById('ctl00_HomePageContent_ButtonAdd') 
                             || Array.from(document.querySelectorAll('input, button, a')).find(el => (el.value || el.innerText || '').toLowerCase().includes('add more'));
                    if (btn) {
                        btn.click();
                        return true;
                    }
                    return false;
                }
            """)
            if clicked_js:
                logger.info("Clicked 'Add More' button via JavaScript.")
                return True
            else:
                logger.warning("Add More button not found, pressing Enter.")
                await page.keyboard.press("Enter")
                return True
        except Exception as e:
            logger.warning(f"JS click warning: {e}")
            return False

    await _do_click_add_more()

    # Dynamically poll for ASP.NET postback to complete and new row to appear
    max_polls = int(ADD_MORE_MAX_WAIT_S * 1000 / ADD_MORE_POLL_MS)
    retry_poll = int(2500 / ADD_MORE_POLL_MS)

    for poll_idx in range(max_polls):
        await page.wait_for_timeout(ADD_MORE_POLL_MS)
        try:
            curr_rows = await get_grid_row_count(page)
            if curr_rows >= target_rows:
                logger.info(f"New row opened in {(poll_idx + 1) * ADD_MORE_POLL_MS}ms! Grid rows: {prev_rows} -> {curr_rows}.")
                return True

            # If portal dropped the click or was busy, re-trigger click once at ~2.5s mark
            if poll_idx == retry_poll:
                logger.info(f"Row count ({curr_rows}) < {target_rows} after 2.5s. Re-triggering 'Add More' click...")
                await wait_for_aspnet_ajax(page, timeout_ms=2000)
                await _do_click_add_more()
        except Exception:
            continue

    logger.warning(f"Row count did not reach {target_rows} within {ADD_MORE_MAX_WAIT_S}s. Continuing...")
    return False


async def click_batch_submit(page):
    selectors = [
        'input#ctl00_HomePageContent_CtrlBtnSubmit',
        'input[name="ctl00$HomePageContent$CtrlBtnSubmit"]',
        'input[type="submit"][value="Submit"]',
        'input[value="Submit"]',
    ]
    clicked = False
    for sel in selectors:
        cand = page.locator(sel).first
        if await cand.count() > 0 and await cand.is_visible() and await cand.is_enabled():
            await cand.click()
            clicked = True
            logger.info("Clicked 'Submit' button.")
            break
    if not clicked:
        raise RuntimeError("Submit button ('ctl00_HomePageContent_CtrlBtnSubmit') not found.")

    # Press Enter after Submit as requested: "then submit button then enter"
    await page.wait_for_timeout(300)
    await page.keyboard.press("Enter")
    logger.info("Pressed Enter after Submit button.")

    await page.wait_for_timeout(2000)
    try:
        await page.wait_for_load_state("domcontentloaded", timeout=15000)
    except Exception:
        pass


async def open_supplementary_form(dashboard_page, existing_popup, month, year):
    """
    Ensures a fresh, clean Supplementary Monthly Contribution page is opened.
    Closes any existing popup first so submissions don't collide.
    """
    if existing_popup:
        try:
            if not existing_popup.is_closed():
                await existing_popup.close()
        except Exception:
            pass

    logger.info("Opening File Monthly Contributions window from dashboard...")
    await dashboard_page.bring_to_front()
    await dashboard_page.wait_for_timeout(1000)
    await dismiss_all_popups(dashboard_page)

    target = dashboard_page.locator("a#lnkMonthlyContribution").first
    if await target.count() == 0:
        target = dashboard_page.locator('a:has-text("Monthly Contribution"), a:has-text("File Monthly")').first
    await target.wait_for(state="visible", timeout=15000)

    async with dashboard_page.expect_popup(timeout=30000) as popup_info:
        await target.click()

    popup = await popup_info.value

    async def _on_popup_dialog(dlg):
        logger.warning(f"ESIC Portal Alert Dialog: '{dlg.message}' (auto-accepting)")
        try:
            await dlg.accept()
        except Exception:
            pass

    popup.on("dialog", lambda d: asyncio.create_task(_on_popup_dialog(d)))

    try:
        await popup.wait_for_load_state("domcontentloaded", timeout=30000)
    except Exception:
        pass
    await popup.wait_for_timeout(800)

    # Initial submit sequence (month & year)
    await select_supplementary_and_submit(popup, month, year)
    await popup.wait_for_timeout(800)
    return popup


async def fill_payroll_details_in_batches(dashboard_page, month, year, loginid):
    """
    Fills ESIC portal from Excel in batches of records.
    Supports Single-Sheet mode (suplimentry_esic_users.xlsx) or Approved Payroll Excel fallback.
    Continuous loop:
      - Reads pending records from Excel (blank or not DONE).
      - Opens fresh form from dashboard for up to BATCH_SIZE records.
      - Record 1..N-1 click 'Add More', Record N clicks 'Submit'.
      - After Submit, immediately marks rows as 'DONE' in Excel and saves safely.
      - Closes popup, pauses, and immediately starts next batch until 100% DONE!
    """
    single_mode = is_single_sheet_mode(EXCEL_FILE)
    if single_mode:
        payroll_file = None
        source_name = Path(EXCEL_FILE).name
    else:
        payroll_file = find_approved_payroll_file(month, year)
        source_name = Path(payroll_file).name

    batch_counter = 0
    total_submitted = 0

    while True:
        if single_mode:
            pending_rows = get_single_sheet_pending_rows(EXCEL_FILE, loginid=loginid, month=month, year=year)
        else:
            pending_rows = get_approved_payroll_values(payroll_file, loginid=None, include_done=False)

        if not pending_rows:
            logger.info(f"All employee records in {source_name} for {loginid} ({month}-{year}) are marked DONE!")
            return {
                "success": True,
                "total_submitted": total_submitted,
                "source_file": source_name,
            }

        total_pending = len(pending_rows)
        current_batch = pending_rows[:BATCH_SIZE]
        batch_counter += 1

        logger.info("=" * 70)
        logger.info(
            f"BATCH {batch_counter} | Entering {len(current_batch)} records "
            f"({total_pending} pending in {source_name})"
        )
        logger.info("=" * 70)

        popup_page = await open_supplementary_form(dashboard_page, None, month, year)
        await popup_page.wait_for_timeout(500)

        batch_row_indices = []
        try:
            for item_idx, item in enumerate(current_batch, start=1):
                insurance_number = item["insurance_number"]
                paid_days = item["paid_days"]
                gross_payable = item["gross_payable"]
                current_excel_row = item["payroll_row"]

                logger.info(
                    f"Batch {batch_counter} | Record {item_idx}/{len(current_batch)} (Row {current_excel_row}) | "
                    f"ESIC={insurance_number} | Days={paid_days} | Wages={gross_payable}"
                )

                # Enter all 3 fields (Insurance Number, Paid Days, Total Monthly Wages)
                await enter_employee_row(
                    popup_page,
                    insurance_number,
                    paid_days,
                    gross_payable,
                    item_idx,
                )

                batch_row_indices.append(current_excel_row)

                # Add More if not last record of batch
                if item_idx < len(current_batch):
                    await click_add_more_and_wait(popup_page, expected_min_rows=item_idx + 1)

            # Last record of batch: Submit
            logger.info(f"All {len(current_batch)} records of Batch {batch_counter} entered. Submitting batch...")
            await click_batch_submit(popup_page)

            # Mark all rows of this batch as DONE in Excel immediately!
            if single_mode:
                mark_single_sheet_rows_status(EXCEL_FILE, batch_row_indices, "DONE")
            else:
                mark_payroll_rows_status(payroll_file, batch_row_indices, "DONE")

            total_submitted += len(batch_row_indices)
            logger.info(
                f"SUCCESS: Batch {batch_counter} submitted and marked DONE in Excel! "
                f"({total_submitted} submitted so far)"
            )

            # Close popup to prepare cleanly for next batch
            await popup_page.wait_for_timeout(800)
            try:
                await popup_page.close()
            except Exception:
                pass
            popup_page = None

            # Small pause on dashboard before opening next batch
            await dashboard_page.wait_for_timeout(1000)

        except Exception as batch_exc:
            if popup_page:
                try:
                    if not popup_page.is_closed():
                        await popup_page.close()
                except Exception:
                    pass
            err_text = f"{type(batch_exc).__name__}: {batch_exc}"
            logger.error(f"Error occurred during Batch {batch_counter}: {err_text}")
            raise


async def open_target_and_download(
    page,
    loginid,
    target_dir,
    month,
    year,
):
    """
    Open Monthly Contribution and process all batches of 15 records continuously.
    """
    # --------------------------------------------------------
    # Login validation
    # --------------------------------------------------------
    if "Portal_Loginnew.aspx" in page.url:
        error_msg_el = page.locator(
            '#lblMsg, '
            '.error, '
            'span[style*="color:Red"], '
            'span[style*="color:red"]'
        ).first

        if await error_msg_el.count() > 0:
            err_text = await error_msg_el.inner_text()
            if err_text.strip():
                raise RuntimeError(
                    "Login Failed on ESIC Portal: "
                    f"{err_text.strip()}"
                )

        raise RuntimeError(
            "Login Failed (Invalid CAPTCHA or Password). "
            "Webpage remained on Login screen."
        )

    await page.wait_for_timeout(2000)
    await dismiss_all_popups(page)

    await page.keyboard.press("Enter")
    await page.wait_for_timeout(1500)

    # Process all batches of 15 records and submit
    final_result = await fill_payroll_details_in_batches(
        page,
        month,
        year,
        loginid,
    )

    logger.info(
        "Full ESIC contribution completed: "
        f"{final_result}"
    )

    return final_result
# READ USERS FROM EXCEL

# ============================================================



def read_users():

    """

    Excel structure:



    Column A = loginid

    Column B = password

    Column C = response

    Column D = Month

    Column E = year

    """



    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel file not found: "
            f"{EXCEL_FILE}"
        )

    if is_single_sheet_mode(EXCEL_FILE):
        return read_single_sheet_users(EXCEL_FILE, SHEET_NAME)

    wb = load_workbook(
        EXCEL_FILE,
        data_only=True,
    )



    ws = (

        wb[SHEET_NAME]

        if SHEET_NAME

        else wb.active

    )



    users = []



    for row in range(

        2,

        ws.max_row + 1,

    ):



        # -----------------------------------------

        # Read Excel columns

        # -----------------------------------------



        raw_loginid = ws.cell(

            row,

            1,

        ).value



        raw_password = ws.cell(

            row,

            2,

        ).value



        raw_month = ws.cell(

            row,

            4,

        ).value



        raw_year = ws.cell(

            row,

            5,

        ).value



        loginid = clean_excel_val(

            raw_loginid

        )



        password = clean_excel_val(

            raw_password

        )



        month = clean_excel_val(

            raw_month

        )



        year = clean_excel_val(

            raw_year

        )



        # -----------------------------------------

        # Login ID

        # -----------------------------------------



        if not loginid:

            continue



        # -----------------------------------------

        # Password

        # -----------------------------------------



        if not password:



            ws.cell(

                row,

                3,

            ).value = "PASSWORD_MISSING"



            continue



        # -----------------------------------------

        # Month

        # -----------------------------------------



        if not month:



            ws.cell(

                row,

                3,

            ).value = "MONTH_MISSING"



            logger.warning(

                f"Row {row}: Month missing "

                f"for Login ID {loginid}"

            )



            continue



        # -----------------------------------------

        # Year

        # -----------------------------------------



        if not year:



            ws.cell(

                row,

                3,

            ).value = "YEAR_MISSING"



            logger.warning(

                f"Row {row}: Year missing "

                f"for Login ID {loginid}"

            )



            continue



        # -----------------------------------------

        # Store user

        # -----------------------------------------



        users.append(

            (

                row,

                loginid,

                password,

                month,

                year,

            )

        )



    try:
        save_workbook_safely(wb, EXCEL_FILE)
        wb.close()
    except Exception as e:
        logger.warning(
            "Could not save Excel changes: "
            f"{e}"
        )



        logger.warning(

            "Could not save Excel changes: "

            f"{e}"

        )



    return users





# ============================================================

# WRITE RESPONSE TO EXCEL

# ============================================================



def write_response(row, response):
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[SHEET_NAME] if SHEET_NAME else wb.active
        ws.cell(row, 3).value = response
        save_workbook_safely(wb, EXCEL_FILE)
        wb.close()
    except Exception as e:
        logger.warning(f"Could not write status to Excel: {e}")


def write_update_status(row, status, col=6):
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[SHEET_NAME] if SHEET_NAME else wb.active
        ws.cell(row, col).value = status
        save_workbook_safely(wb, EXCEL_FILE)
        wb.close()
    except Exception as e:
        logger.warning(f"Could not write update status to Excel: {e}")




# ============================================================

# MAIN PORTAL DOWNLOAD

# ============================================================



async def run_portal_download(

    headless=False,

    auto_mode=True,

):



    logger.info(

        "=================================================="

    )



    logger.info(

        "Starting ESIC Portal Download Automation"

    )



    logger.info(

        f"Excel File: {EXCEL_FILE}"

    )



    logger.info(

        f"Download Dir: {TARGET_DOWNLOAD_DIR}"

    )



    logger.info(

        "=================================================="

    )



    # -----------------------------------------

    # Ensure directories

    # -----------------------------------------



    CAPTCHA_DIR.mkdir(

        parents=True,

        exist_ok=True,

    )



    TARGET_DOWNLOAD_DIR.mkdir(

        parents=True,

        exist_ok=True,

    )



    # -----------------------------------------

    # Read Excel

    # -----------------------------------------



    users = read_users()



    if not users:



        logger.warning(

            "No valid users found in Excel."

        )



        return True



    logger.info(

        f"Total Users Found: {len(users)}"

    )



    success_count = 0

    failed_count = 0



    # =========================================

    # PLAYWRIGHT

    # =========================================



    async with async_playwright() as p:



        browser = await p.chromium.launch(

            headless=headless,

            slow_mo=100 if not headless else 0,

        )



        # -------------------------------------

        # Process every user

        # -------------------------------------



        for index, user_entry in enumerate(users, start=1):
            row = user_entry[0]
            loginid = user_entry[1]
            password = user_entry[2]
            month = user_entry[3]
            year = user_entry[4]
            group_row_indices = user_entry[5] if len(user_entry) > 5 else [row]

            logger.info("--------------------------------------------------")
            logger.info(f"Processing User/Period {index}/{len(users)}")
            logger.info(f"Login ID : {loginid}")
            logger.info(f"Month    : {month}")
            logger.info(f"Year     : {year}")
            logger.info("--------------------------------------------------")

            single_mode = is_single_sheet_mode(EXCEL_FILE)
            if not single_mode:
                payroll_file = find_approved_payroll_file(month, year)
                source_name = Path(payroll_file).name
            else:
                payroll_file = None
                source_name = Path(EXCEL_FILE).name

            MAX_RESUME_ATTEMPTS = 10
            user_success = False

            for attempt in range(1, MAX_RESUME_ATTEMPTS + 1):
                # Check remaining pending/blank records in Excel
                if single_mode:
                    pending_employees = get_single_sheet_pending_rows(EXCEL_FILE, loginid=loginid, month=month, year=year)
                else:
                    pending_employees = get_approved_payroll_values(payroll_file, loginid=None, include_done=False)

                if not pending_employees:
                    logger.info(f"ALL employee records for {loginid} in {source_name} ({month}-{year}) are marked DONE!")
                    if single_mode:
                        mark_single_sheet_group_response(
                            EXCEL_FILE,
                            group_row_indices,
                            f"SUCCESS - Final Submitted | {month}-{year}",
                            "UPDATED",
                        )
                    else:
                        write_response(
                            row,
                            f"SUCCESS - Final Submitted | {month}-{year}",
                        )
                        write_update_status(
                            row,
                            "UPDATED",
                        )
                    user_success = True
                    success_count += 1
                    break

                if attempt > 1:
                    logger.info(
                        f"Auto-login Resume Attempt {attempt}/{MAX_RESUME_ATTEMPTS} for {loginid} | "
                        f"{len(pending_employees)} unuploaded/blank records remaining."
                    )

                context = None
                try:
                    # Browser context
                    context = await browser.new_context(
                        accept_downloads=True,
                        viewport={
                            "width": 1600,
                            "height": 900,
                        },
                    )
                    page = await context.new_page()

                    safe_loginid = "".join(
                        ch
                        if ch.isalnum() or ch in "-_"
                        else "_"
                        for ch in str(loginid)
                    )

                    # Auto-login to ESIC portal
                    await login(
                        page,
                        loginid,
                        password,
                        safe_loginid,
                        auto_mode=auto_mode,
                    )

                    # Open form and process batches
                    submit_result = await open_target_and_download(
                        page,
                        loginid,
                        TARGET_DOWNLOAD_DIR,
                        month,
                        year,
                    )

                    # Verify if all records from Excel are now DONE
                    if single_mode:
                        remaining_employees = get_single_sheet_pending_rows(EXCEL_FILE, loginid=loginid, month=month, year=year)
                    else:
                        remaining_employees = get_approved_payroll_values(payroll_file, loginid=None, include_done=False)

                    if not remaining_employees:
                        if single_mode:
                            mark_single_sheet_group_response(
                                EXCEL_FILE,
                                group_row_indices,
                                f"SUCCESS - Final Submitted | {month}-{year}",
                                "UPDATED",
                            )
                        else:
                            write_response(
                                row,
                                f"SUCCESS - Final Submitted | {month}-{year}",
                            )
                            write_update_status(
                                row,
                                "UPDATED",
                            )
                        logger.info(
                            f"User {loginid} SUCCESS: All records submitted and marked DONE in Excel | {month}-{year}"
                        )
                        user_success = True
                        success_count += 1
                        break
                    else:
                        logger.info(f"{len(remaining_employees)} records still pending. Continuing next cycle...")

                except Exception as exc:
                    err_msg = f"{type(exc).__name__}: {exc}"
                    logger.error(f"User {loginid} Attempt {attempt} Error: {err_msg}")

                    # Check if all records were actually completed before the error
                    if single_mode:
                        remaining_employees = get_single_sheet_pending_rows(EXCEL_FILE, loginid=loginid, month=month, year=year)
                    else:
                        remaining_employees = get_approved_payroll_values(payroll_file, loginid=None, include_done=False)

                    if not remaining_employees:
                        logger.info(f"All records in Excel are marked DONE for {loginid}!")
                        if single_mode:
                            mark_single_sheet_group_response(
                                EXCEL_FILE,
                                group_row_indices,
                                f"SUCCESS - Final Submitted | {month}-{year}",
                                "UPDATED",
                            )
                        else:
                            write_response(
                                row,
                                f"SUCCESS - Final Submitted | {month}-{year}",
                            )
                            write_update_status(
                                row,
                                "UPDATED",
                            )
                        user_success = True
                        success_count += 1
                        break

                    if attempt < MAX_RESUME_ATTEMPTS:
                        logger.info(
                            f"Portal disconnected/stopped: {err_msg}. "
                            f"{len(remaining_employees)} blank/pending records remain. "
                            f"Auto-logging in and resuming entry in 5 seconds..."
                        )
                        await asyncio.sleep(5)
                    else:
                        if single_mode:
                            mark_single_sheet_group_response(
                                EXCEL_FILE,
                                group_row_indices,
                                f"FAILED - {err_msg}",
                                f"NOT UPDATED - {err_msg}",
                            )
                        else:
                            write_response(
                                row,
                                f"FAILED - {err_msg}",
                            )
                            write_update_status(
                                row,
                                f"NOT UPDATED - {err_msg}",
                            )
                        logger.error(
                            f"User {loginid} failed after {MAX_RESUME_ATTEMPTS} attempts."
                        )
                        failed_count += 1

                # =================================

                # CLOSE CONTEXT

                # =================================



                finally:



                    if context:



                        try:

                            await context.close()

                        except Exception:

                            pass



        # =====================================

        # CLOSE BROWSER

        # =====================================



        await browser.close()



    # =========================================

    # SUMMARY

    # =========================================



    logger.info(

        "=================================================="

    )



    logger.info(

        "ESIC Portal Download Summary"

    )



    logger.info(

        f"SUCCESS : {success_count}"

    )



    logger.info(

        f"FAILED  : {failed_count}"

    )



    logger.info(

        "=================================================="

    )



    # -----------------------------------------

    # Downstream pipeline condition

    # -----------------------------------------



    if success_count > 0:



        if failed_count > 0:



            logger.warning(

                "Proceeding to downstream "

                "Merge/Compare pipeline with "

                f"{success_count} downloaded file(s)."

            )



        return True



    logger.error(

        "0 files downloaded. "

        "Cannot proceed to downstream processing."

    )



    return False





# ============================================================

# ENTRY POINT FOR N8N / WORKER

# ============================================================



def run(

    headless=False,

    auto_mode=True,

):

    """

    Entry point used by n8n worker.py.

    """



    return asyncio.run(

        run_portal_download(

            headless=headless,

            auto_mode=auto_mode,

        )

    )





# ============================================================

# DIRECT EXECUTION

# ============================================================



if __name__ == "__main__":



    headless_flag = (

        "--headless"

        in sys.argv

    )



    asyncio.run(

        run_portal_download(

            headless=headless_flag,

            auto_mode=True,

        )

    )
