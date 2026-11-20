import os
import sys
import pandas as pd
from playwright.sync_api import sync_playwright
from storage_state_cmhelpline import save_cmhelpline_storage_state

# Ensure root directory is in sys.path
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

try:
    from cmhelpline_code.func_download_excel import download_file
except ImportError:
    from func_download_excel import download_file

try:
    from cmhelpline_code.storage_state_cmhelpline import save_cmhelpline_storage_state
except ImportError:
    from storage_state_cmhelpline import save_cmhelpline_storage_state


def is_session_valid(state_file):
    """
    Checks if state_file exists and can successfully load the authenticated CM Helpline dashboard.
    Returns True if session is valid and active, False otherwise.
    """
    if not os.path.exists(state_file):
        return False

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(storage_state=state_file)
            page = context.new_page()
            try:
                page.goto('https://cmhelpline.mp.gov.in/MIS/dashboard.aspx', timeout=20000)
                page.wait_for_load_state('networkidle', timeout=10000)

                # Check if redirected to login page
                if "Default.aspx" in page.url or page.locator('//input[@id="txtUser"]').count() > 0:
                    return False

                # Check if ASP.NET server session error occurred
                if "Error" in page.title() or "Object reference" in page.title():
                    return False

                # Check if authenticated WIPLink element is attached
                wip = page.wait_for_selector("//a[@id='WIPLink']", state="attached", timeout=6000)
                return wip is not None

            except Exception:
                return False
            finally:
                browser.close()
    except Exception:
        return False


def ensure_valid_session(state_file=None, headless=True, captcha_value=None, captcha_callback=None, progress_callback=None):
    """
    Ensures a valid session exists. If state_file is missing, session is expired,
    or the dashboard fails to load, automatically calls save_cmhelpline_storage_state()
    to log in and renew the session.

    Returns:
        bool: True if a valid session is ready, False if login failed.
    """
    if state_file is None:
        state_file = os.path.join(project_root, 'cmhelpline_code', 'state_CM.json')

    print("\n" + "=" * 60)
    print("CHECKING CM HELPLINE SESSION")
    print("=" * 60)

    if progress_callback:
        progress_callback("checking_session", "Checking CM Helpline session validity...", 10)

    # Case 1: State file does not exist
    if not os.path.exists(state_file):
        msg = f"⚠️ State file '{state_file}' not found. Automatically calling save_cmhelpline_storage_state()..."
        print(msg)
        if progress_callback:
            progress_callback("authenticating", msg, 15)
        return save_cmhelpline_storage_state(
            state_file=state_file,
            headless=headless,
            captcha_value=captcha_value,
            captcha_callback=captcha_callback
        )

    # Case 2: State file exists - test if dashboard loads
    print("Testing saved session against dashboard...")
    if is_session_valid(state_file):
        print("✓ Saved session is valid and active.")
        if progress_callback:
            progress_callback("session_valid", "Session is active and valid.", 20)
        return True

    # Case 3: Session is expired or dashboard failed to load
    msg = "⚠️ Session is expired or dashboard not loading. Automatically calling save_cmhelpline_storage_state()..."
    print(msg)
    if progress_callback:
        progress_callback("authenticating", msg, 15)
    return save_cmhelpline_storage_state(
        state_file=state_file,
        headless=headless,
        captcha_value=captcha_value,
        captcha_callback=captcha_callback
    )


def download_all_cmhelpline_excel(
    headless=True,
    state_file=None,
    output_dir=None,
    progress_callback=None,
    captcha_value=None,
    captcha_callback=None
):
    """
    Downloads L1, L2, L3, and L4 complaint Excel files from CM Helpline portal.
    If login is unsuccessful (state file missing, session expired, or dashboard not loading),
    automatically calls save_cmhelpline_storage_state() to log in before continuing.

    Parameters:
        headless (bool): Run browser in headless mode. Default is True.
        state_file (str): Path to state_CM.json session file.
        output_dir (str): Directory where downloaded Excel files will be saved.
        progress_callback (callable): Optional callback for progress updates (stage, message, percent).

    Returns:
        bool: True if all files downloaded successfully, False otherwise.
    """
    if output_dir is None:
        output_dir = os.path.join(project_root, 'cmhelpline')
    os.makedirs(output_dir, exist_ok=True)

    if state_file is None:
        state_file = os.path.join(project_root, 'cmhelpline_code', 'state_CM.json')

    # Step 1: Ensure valid session (auto-triggers login if missing or expired)
    session_ok = ensure_valid_session(
        state_file=state_file,
        headless=headless,
        captcha_value=captcha_value,
        captcha_callback=captcha_callback,
        progress_callback=progress_callback
    )
    if not session_ok:
        print("\n✗ Unable to establish a valid session. Aborting download.")
        if progress_callback:
            progress_callback("error", "Unable to establish a valid session. Login failed.", 0)
        return False

    print("\n" + "=" * 60)
    print("STARTING COMPLAINT EXCEL DOWNLOADS")
    print("=" * 60)
    if progress_callback:
        progress_callback("downloading", "Starting complaint Excel downloads...", 25)

    # Step 2: Download complaints with validated session
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        try:
            print("\n1. Creating context with validated session...")
            context = browser.new_context(storage_state=state_file)
            page = context.new_page()

            # Navigate to dashboard
            print("\n2. Opening dashboard...")
            page.goto('https://cmhelpline.mp.gov.in/MIS/dashboard.aspx', timeout=30000)
            page.wait_for_load_state('networkidle', timeout=15000)

            # Download L1 complaints
            print("\n3. Downloading L1 complaints...")
            if progress_callback:
                progress_callback("downloading_l1", "Downloading L1 complaints...", 35)
            L1 = page.wait_for_selector("//a[@id='WIPLink']", state="attached", timeout=10000)
            L1.click(force=True)

            all_complain = page.wait_for_selector("//a[contains(@href,'complaintlist.aspx?type=1')]", timeout=5000)
            all_complain.click()

            l1_file_path = os.path.join(output_dir, 'cmhelpline_1.xls')
            download_file(
                page=page,
                locator='input#ctl00_ContentPlaceHolder1_btnExelExport',
                save_path=l1_file_path
            )
            if progress_callback:
                progress_callback("downloaded_l1", "L1 complaints downloaded successfully.", 50)

            # Safely navigate back to dashboard
            page.goto('https://cmhelpline.mp.gov.in/MIS/dashboard.aspx', timeout=30000)
            page.wait_for_load_state('networkidle', timeout=15000)

            # Click on 'उच्च लेवल पर लंबित शिकायतें' to view L2, L3, L4 table
            box = page.wait_for_selector("//div[contains(@class, 'cm-box-green') and .//span[text()='उच्च लेवल पर लंबित शिकायतें']]", timeout=10000)
            box.click()
            page.wait_for_load_state('networkidle', timeout=15000)

            levels = [
                ('L2', 2, os.path.join(output_dir, 'cmhelpline_L2.xls')),
                ('L3', 3, os.path.join(output_dir, 'cmhelpline_L3.xls')),
                ('L4', 4, os.path.join(output_dir, 'cmhelpline_L4.xls'))
            ]

            for name, col_idx, save_path in levels:
                print(f"\n--- Downloading {name} ---")
                pct = {"L2": 65, "L3": 75, "L4": 85}.get(name, 70)
                if progress_callback:
                    progress_callback(f"downloading_{name.lower()}", f"Downloading {name} complaints...", pct)

                if "TimeLimitExceededcomplaints.aspx" not in page.url or "TimeLimitExceededcomplaints_AttriWise.aspx" in page.url:
                    page.goto('https://cmhelpline.mp.gov.in/MIS/TimeLimitExceededcomplaints.aspx', timeout=30000)
                    page.wait_for_load_state('networkidle', timeout=15000)

                td_link = page.wait_for_selector(
                    f"//table[contains(@class, 'cm-margin-top-60') and contains(@class, 'table-hover')]//tr[1]//td[{col_idx}]//a",
                    timeout=10000
                )
                count_text = td_link.inner_text().strip()
                try:
                    count = int(count_text)
                except ValueError:
                    count = 0

                if count == 0:
                    print(f"   ✓ Pending complaints count for {name} is 0. Skipping.")
                    if progress_callback:
                        progress_callback(f"skipped_{name.lower()}", f"{name} count is 0. Skipped.", pct + 5)
                    continue

                td_link.click()
                page.wait_for_load_state('networkidle', timeout=15000)
                page.wait_for_selector(f"//a[@href='TimeLimitExceededcomplaintsByOfficerLevel.aspx?OfficrLevel2={name}']", timeout=15000).click()
                page.wait_for_load_state('networkidle', timeout=15000)

                download_file(
                    page=page,
                    locator='input#ctl00_ContentPlaceHolder1_btnExelExport',
                    save_path=save_path
                )
                if progress_callback:
                    progress_callback(f"downloaded_{name.lower()}", f"{name} complaints downloaded successfully.", pct + 8)

            print("\n✓ All complaint files processed successfully!")
            if progress_callback:
                progress_callback("completed", "All complaint files downloaded successfully!", 95)
            return True

        except Exception as e:
            print(f"\n✗ Error during download: {e}")
            return False

        finally:
            try:
                browser.close()
            except Exception:
                pass


if __name__ == "__main__":
    download_all_cmhelpline_excel(headless=False)
