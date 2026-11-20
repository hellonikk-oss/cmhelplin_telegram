import os
import sys
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()

# Ensure root directory is in sys.path
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if project_root not in sys.path:
    sys.path.insert(0, project_root)


def save_cmhelpline_storage_state(
    state_file=None,
    username=None,
    password=None,
    headless=True,
    screenshot_on_error=True,
    captcha_value=None,
    captcha_callback=None
):
    """
    Automates the login to CM Helpline portal and saves the browser storage state (session/cookies) to a JSON file.

    Parameters:
        state_file (str, optional): Target file path to save the session JSON.
                                    Defaults to './cmhelpline_code/state_CM.json'.
        username (str, optional): Portal username. If None, reads from CMHELIPLINE_USENAME or CMHELPLINE_USERNAME env var.
        password (str, optional): Portal password. If None, reads from CMHELPLINE_PASSWORD env var.
        headless (bool, optional): Run browser in headless mode. Defaults to True.
        screenshot_on_error (bool, optional): Save screenshot if an error occurs. Defaults to True.

    Returns:
        bool: True if the session was successfully saved, False otherwise.
    """
    if state_file is None:
        state_file = os.path.join(project_root, 'cmhelpline_code', 'state_CM.json')

    # Remove existing state file if present
    if os.path.exists(state_file):
        try:
            os.remove(state_file)
        except OSError:
            pass

    username = username or os.getenv('CMHELIPLINE_USENAME') or os.getenv('CMHELPLINE_USERNAME')
    password = password or os.getenv('CMHELPLINE_PASSWORD')

    if not username or not password:
        print("Error: Username or Password not provided and not found in environment variables (.env).", flush=True)
        return False

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=headless,
            args=['--disable-blink-features=AutomationControlled']
        )
        context = browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        )
        page = context.new_page()

        try:
            print("Opening login page...", flush=True)
            page.goto('https://cmhelpline.mp.gov.in/MIS/Default.aspx', wait_until='domcontentloaded', timeout=30000)
            try:
                page.wait_for_load_state('networkidle', timeout=10000)
            except Exception:
                pass
            print("Login page opened.", flush=True)

            # Enter username
            user_field = page.wait_for_selector('//input[@id="txtUser"]', timeout=15000)
            user_field.fill(username)
            print("Username entered.", flush=True)

            # Enter password
            pass_field = page.wait_for_selector('//input[@id="txtPass"]', timeout=5000)
            pass_field.fill(password)
            print("Password entered.", flush=True)

            # Capture CAPTCHA image
            captcha_img = page.locator('//img[contains(@src, "CaptchaImage")]')
            captcha_img.wait_for(state="visible", timeout=10000)

            captcha_local_path = os.path.join(os.path.dirname(os.path.abspath(state_file)), 'captcha.png')
            captcha_img.screenshot(path=captcha_local_path)

            brain_dir = os.environ.get('AGENT_BRAIN_DIR', r'C:\Users\acer\.gemini\antigravity-cli\brain\fc702f9c-9a82-4642-b6f0-974b7e02b3dd')
            if os.path.isdir(brain_dir):
                try:
                    captcha_img.screenshot(path=os.path.join(brain_dir, 'captcha.png'))
                except Exception:
                    pass

            print(f"CAPTCHA_READY:{captcha_local_path}", flush=True)
            if captcha_value:
                captcha_input = str(captcha_value).strip()
                print(f"Using provided CAPTCHA: {captcha_input}", flush=True)
            elif callable(captcha_callback):
                print("Requesting CAPTCHA via callback...", flush=True)
                captcha_input = str(captcha_callback(captcha_local_path)).strip()
                print(f"Callback returned CAPTCHA: {captcha_input}", flush=True)
            else:
                print("\nPlease view the CAPTCHA image and enter the displayed characters:", flush=True)
                try:
                    captcha_input = input("Enter CAPTCHA: ").strip()
                except Exception:
                    captcha_input = ""
            captcha_field = page.wait_for_selector('//input[@id="txtCaptcha"]', timeout=5000)
            captcha_field.fill(captcha_input)
            print("CAPTCHA entered.", flush=True)

            # Click submit button
            button = page.wait_for_selector("//input[@id='btnsubmit']", timeout=5000)
            button.click()
            print("Login button clicked.", flush=True)

            # Wait for successful navigation / dashboard load
            try:
                page.wait_for_load_state('networkidle', timeout=25000)
            except Exception:
                pass

            if "Default.aspx" in page.url or page.locator('//input[@id="txtUser"]').count() > 0:
                print("✗ Login failed: Still on login page. Please verify CAPTCHA or credentials.", flush=True)
                return False

            print("✓ Login successful.", flush=True)

            # Ensure target directory exists
            os.makedirs(os.path.dirname(os.path.abspath(state_file)), exist_ok=True)

            # Save session state
            context.storage_state(path=state_file)
            print(f"✓ Session state saved successfully to: {state_file}", flush=True)
            return True

        except Exception as e:
            print(f"✗ An error occurred: {e}", flush=True)
            if screenshot_on_error:
                error_screenshot_path = os.path.join(
                    os.path.dirname(os.path.abspath(state_file)),
                    'error_screenshot.png'
                )
                try:
                    page.screenshot(path=error_screenshot_path)
                    print(f"Screenshot saved as {error_screenshot_path} for debugging.", flush=True)
                except Exception:
                    pass
            return False

        finally:
            browser.close()


# Alias for convenience
create_cmhelpline_session = save_cmhelpline_storage_state


if __name__ == "__main__":
    save_cmhelpline_storage_state()
