import re
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

current_dir = Path(__file__).resolve().parent  # get the folder that contains the current file
logined_in_file = current_dir / "aha_auth.json"  # location of aha_auth.json
AHA_url = "https://atlas.heart.org/location"

def sign_in_visible(page) -> bool:  # page = current browser tab, PlayWright object
    # On the current page, find a visible Sign In button or link
    for role in ("link", "button"): # all elements that is either a link or button, can be hidden
        loc = page.get_by_role(role, name=re.compile(r"sign\s*in", re.I)) # visible name looks like ‘sign in’
        if loc.count() > 0: # if matching elements is > 0
            try:
                if loc.first.is_visible(): 
                    return True
            except Exception:
                pass
    return False

def main():
    print(f"save state to: {logined_in_file}")

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=False) # launch Microsoft Edge, visible window
        context = browser.new_context()
        page = context.new_page()

        # open AHA website
        page.goto(AHA_url, wait_until="domcontentloaded") # wait until the page’s basic HTML is loaded
        page.wait_for_timeout(1200) # pause for 1.2 sec
        page.screenshot(path=str(current_dir / "setup_1_atlas.png"), full_page=True) # Takes a screenshot & 
        # keep te file path in plain text str(), screenshot the entire page 

        #  don't touch this part...... prevent ssoVerifier missing
        clicked = False  # boolan flag, check if the script actually click on the Sign-in button or link
        for role in ("link", "button"): 
            loc = page.get_by_role(role, name=re.compile(r"sign\s*in", re.I))
            if loc.count() > 0:
                loc.first.click()
                clicked = True
                break

        if not clicked:
            page.screenshot(path=str(current_dir / "setup_signin_not_found.png"), full_page=True)
            browser.close()
            raise RuntimeError("Cannot find Sign In on AHA. Screenshot: setup_signin_not_found.png")

        # wait until redirect to the login page
        try:
            page.wait_for_url(re.compile(r"ahasso\.heart\.org/.*login", re.I), timeout=20000)
        except PlaywrightTimeoutError:
            page.screenshot(path=str(current_dir / "setup_not_redirected_to_login.png"), full_page=True)
            browser.close()
            raise RuntimeError("Did not redirect to login page. Screenshot: setup_not_redirected_to_login.png")

        page.screenshot(path=str(current_dir / "setup_2_login_page.png"), full_page=True)
        print("sign in manually in the browser window ( NOT close the browser).")
        print("The script should save aha_auth.json.")

        # wait. when you’re back on Atlas and “Sign In” is no longer visible, treat it as a successful login
        deadline = time.time() + 180  # now + 3 minute maximum waiting window.
        while time.time() < deadline:  # running timestamp < 3 min
            url = page.url.lower() # tab’s current url in lowercase

            on_atlas = "atlas.heart.org" in url # Ture or False. if the current url contains atlas.heart.org
            on_sso = "ahasso.heart.org" in url # True or False. If I'm on the sign in page or not
            still_sign_in = sign_in_visible(page) # whether ‘Sign In’ is visible right now in the current page

            if on_atlas and (not on_sso) and (not still_sign_in): # on an AHA page, not on the SSO login domain, the Sign In button or link is gone
                # browser context (contains session data), exports the current context’s session data into a json file
                context.storage_state(path=str(logined_in_file))
                page.screenshot(path=str(current_dir / "setup_success.png"), full_page=True)
                print(f" Saved login state: {logined_in_file}")
                print(" Screenshot: setup_success.png")
                browser.close()
                return

            time.sleep(1) # Stop the script for 1 second before checking again.

        # timeout：save screentshot for debugging
        page.screenshot(path=str(current_dir / "setup_timeout.png"), full_page=True)
        browser.close()
        raise RuntimeError("Login not detected within 3 minutes. Screenshot: setup_timeout.png")

if __name__ == "__main__":
    main()
