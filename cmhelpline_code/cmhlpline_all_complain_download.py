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


def ensure_valid_session(state_file):
    """
    Ensures a valid session exists. If state_file is missing, session is expired,
    or the dashboard fails to load, automatically calls save_cmhelpline_storage_state()
    to log in and renew the session.

    Returns:
        bool: True if a valid session is ready, False if login failed.
    """
    print("\n" + "=" * 60)
    print("CHECKING CM HELPLINE SESSION")
    print("=" * 60)

    # Case 1: State file does not exist
    if not os.path.exists(state_file):
        print(f"⚠️ State file '{state_file}' not found.")
        print("Automatically calling save_cmhelpline_storage_state() to create session...")
        return save_cmhelpline_storage_state(state_file=state_file, headless=False)

    # Case 2: State file exists - test if dashboard loads
    print("Testing saved session against dashboard...")
    if is_session_valid(state_file):
        print("✓ Saved session is valid and active.")
        return True

    # Case 3: Session is expired or dashboard failed to load
    print("⚠️ Session is expired or dashboard not loading.")
    print("Automatically calling save_cmhelpline_storage_state() to re-authenticate...")
    return save_cmhelpline_storage_state(state_file=state_file, headless=False)


def download_all_cmhelpline_excel(headless=False, state_file=None, output_dir=None):
    """
    Downloads L1, L2, L3, and L4 complaint Excel files from CM Helpline portal.
    If login is unsuccessful (state file missing, session expired, or dashboard not loading),
    automatically calls save_cmhelpline_storage_state() to log in before continuing.

    Parameters:
        headless (bool): Run browser in headless mode or visible GUI. Default is False.
        state_file (str): Path to state_CM.json session file.
        output_dir (str): Directory where downloaded Excel files will be saved.

    Returns:
        bool: True if all files downloaded successfully, False otherwise.
    """
    if output_dir is None:
        output_dir = os.path.join(project_root, 'cmhelpline')
    os.makedirs(output_dir, exist_ok=True)

    if state_file is None:
        state_file = os.path.join(project_root, 'cmhelpline_code', 'state_CM.json')

    # Step 1: Ensure valid session (auto-triggers login if missing or expired)
    session_ok = ensure_valid_session(state_file)
    if not session_ok:
        print("\n✗ Unable to establish a valid session. Aborting download.")
        return False

    print("\n" + "=" * 60)
    print("STARTING COMPLAINT EXCEL DOWNLOADS")
    print("=" * 60)

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

            print("\n✓ All complaint files processed successfully!")
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
