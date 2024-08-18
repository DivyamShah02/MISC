import os
import pdb
import time
import subprocess
import pyautogui as pg
import pyperclip as pyc
from library.GetLogger import apply_logs_to_all_methods, log

@apply_logs_to_all_methods(log)
class ChromeHandler:
    def __init__(self, logger, config) -> None:
        self.logger = logger
        self.config = config
        self.process = None
        self.kill_all_chrome()

    def kill_all_chrome(self) -> None:
        """Terminate all Chrome processes."""
        try:
            self.logger.info("Terminating all Chrome processes.")
            os.system("taskkill /F /IM chrome.exe")
            self.logger.info("All Chrome processes terminated.")
        except Exception as e:
            self.logger.error("Failed to terminate Chrome processes.", exc_info=True)

    def start_chrome(self) -> bool:
        """Start Chrome browser."""
        try:
            self.logger.info("Starting Chrome.")
            self.process = subprocess.Popen([self.config.chrome_config.chrome_path])
            time.sleep(5)  # Wait for Chrome to fully start
            self.logger.info("Chrome started successfully.")
            return True
        except subprocess.SubprocessError as e:
            self.logger.error("Error while starting Chrome.", exc_info=True)
            return False

    def select_profile(self, profile_index: int) -> bool:
        """Select Chrome user profile."""
        try:
            self.logger.info("Selecting profile.")
            pg.press('tab')
            for _ in range(profile_index - 1):
                pg.press('tab', presses=3, interval=0.3)
                time.sleep(0.3)
            pg.press('enter')
            time.sleep(1)
            self.logger.info("Profile selected successfully.")
            return True
        except pg.PyAutoGUIException as e:
            self.logger.error("Error while selecting profile.", exc_info=True)
            return False

    def maximise_chrome(self) -> bool:
        """Maximize Chrome window."""
        try:
            self.logger.info("Maximizing Chrome window.")
            time.sleep(3)  # Wait for Chrome window to be ready
            pg.hotkey('ctrl', 'l')
            time.sleep(0.2)
            pg.hotkey('alt', 'space')
            time.sleep(0.5)
            pg.press('x')
            self.logger.info("Chrome window maximized successfully.")
            return True
        except pg.PyAutoGUIException as e:
            self.logger.error("Error while maximizing Chrome window.", exc_info=True)
            return False

    def load_url(self, url: str) -> bool:
        """Load a specific URL in Chrome."""
        try:
            pg.hotkey('ctrl', 'l')
            self.logger.info(f"Loading URL: {url}")
            pg.write(url)
            time.sleep(0.3)
            pg.press('enter')
            self.logger.info(f"URL {url} loaded successfully.")
            return True
        except pg.PyAutoGUIException as e:
            self.logger.error(f"Error while loading URL: {url}", exc_info=True)
            return False

    def template_fun(self) -> bool:
        """Placeholder function for custom logic."""
        try:
            self.logger.info("Executing template function.")
            # Custom logic goes here
            self.logger.info("Template function executed successfully.")
            return True
        except Exception as e:
            self.logger.error("Error while executing template function.", exc_info=True)
            return False
