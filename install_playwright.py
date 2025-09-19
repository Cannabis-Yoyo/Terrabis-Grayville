# install_playwright.py
from playwright.sync_api import sync_playwright

def install_playwright_browsers():
    with sync_playwright() as p:
        p.chromium.install()  # Install Chromium
        p.firefox.install()   # Install Firefox (if required)
        p.webkit.install()    # Install Webkit (if required)

if __name__ == "__main__":
    install_playwright_browsers()
