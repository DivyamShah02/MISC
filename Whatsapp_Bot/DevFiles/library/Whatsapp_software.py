import os
import time
import psutil
import win32gui
import requests
import subprocess
import pyautogui as pg
import pyperclip as pyc
from pywinauto.application import Application
from pywinauto.findwindows import ElementNotFoundError


class WhatsAppHandler:
    def __init__(self, config, logger) -> None:
        self.config = config
        self.logger = logger
        self.app = None

    def _is_internet_on(self) -> bool:
        """Check if the internet is connected."""
        try:
            response = requests.get('http://www.google.com', timeout=5)
            return response.status_code == 200
        except requests.ConnectionError:
            self.logger.error("Internet connection is not available.")
            return False

    def _kill_whatsapp(self) -> bool:
        """Kills any running instance of WhatsApp."""
        try:
            for process in psutil.process_iter(['pid', 'name']):
                if process.info['name'] == 'WhatsApp.exe':
                    process.kill()
                    self.logger.info(f"Killed WhatsApp process with PID {process.info['pid']}")
            return True
        except Exception as e:
            self.logger.error("Error killing WhatsApp process", exc_info=True)
            return False

    def _start_whatsapp(self, retries=3, delay=5) -> bool:
        """Starts WhatsApp application."""
        for attempt in range(retries):
            if not self._is_internet_on():
                self.logger.error("Cannot start WhatsApp. Internet is not connected.")
                return False

            if self._kill_whatsapp():
                try:
                    self.logger.info("Starting WhatsApp")
                    subprocess.Popen(["cmd", "/C", "start whatsapp://send"], shell=True)
                    time.sleep(5)  # Give it some time to start
                    if self._is_whatsapp_running():
                        self.logger.info("WhatsApp started successfully.")
                        return True
                except Exception as e:
                    self.logger.error("Error starting WhatsApp", exc_info=True)
            time.sleep(delay)
        return False

    def _is_whatsapp_running(self) -> bool:
        """Checks if WhatsApp is running."""
        for process in psutil.process_iter(['name']):
            if process.info['name'] == 'WhatsApp.exe':
                return True
        return False

    def _connect_whatsapp(self, retries=3, delay=5) -> bool:
        """Connects to the WhatsApp application."""
        for attempt in range(retries):
            if self._is_whatsapp_running():
                try:
                    self.logger.info("Connecting to WhatsApp")
                    self.app = Application(backend='uia').connect(title='WhatsApp', timeout=100)
                    self.logger.info("Connected to WhatsApp")
                    # if self._is_whatsapp_foreground() and self._is_whatsapp_on_screen():
                    if self._is_whatsapp_foreground():
                        self.logger.info("WhatsApp is on screen and in focus.")
                        return True
                    else:
                        self.logger.error("WhatsApp is not on screen or not in focus.")
                except ElementNotFoundError:
                    self.logger.error("WhatsApp window not found", exc_info=True)
                except Exception as e:
                    self.logger.error("Error connecting to WhatsApp", exc_info=True)
            else:
                self.logger.error("WhatsApp is not running.")
            time.sleep(delay)
        return False

    def _get_whatsapp_hwnd(self):
        """Returns the window handle (HWND) of WhatsApp."""
        try:
            if self.app:
                return self.app.window(title='WhatsApp').handle
            return None
        except Exception as e:
            self.logger.error("Error getting WhatsApp window handle", exc_info=True)
            return None

    def _is_whatsapp_foreground(self) -> bool:
        """Check if WhatsApp is the foreground (active) window."""
        try:
            whatsapp_hwnd = self._get_whatsapp_hwnd()
            if whatsapp_hwnd:
                foreground_hwnd = win32gui.GetForegroundWindow()
                return foreground_hwnd == whatsapp_hwnd
            return False
        except Exception as e:
            self.logger.error("Error checking if WhatsApp is the foreground window", exc_info=True)
            return False

    def _is_whatsapp_on_screen(self) -> bool:
        """Check if WhatsApp is on the visible part of the screen and not minimized."""
        try:
            if self.app:
                whatsapp_window = self.app.window(title='WhatsApp')
                if whatsapp_window:
                    rect = whatsapp_window.rectangle()  # Gets the window rectangle (left, top, right, bottom)
                    if rect.width() > 0 and rect.height() > 0:  # Check if window is not minimized
                        screen_rect = win32gui.GetWindowRect(win32gui.GetDesktopWindow())
                        return (rect.left >= screen_rect[0] and
                                rect.top >= screen_rect[1] and
                                rect.right <= screen_rect[2] and
                                rect.bottom <= screen_rect[3])
            return False
        except Exception as e:
            self.logger.error("Error checking if WhatsApp is on screen", exc_info=True)
            return False

    def _perform_click(self, title, auto_id=None, control_type="Button", action='click'):
        """Helper method to perform a click action."""
        try:
            if auto_id:
                control = self.app.WhatsApp.child_window(auto_id=auto_id, control_type=control_type).wrapper_object()
            else:
                control = self.app.WhatsApp.child_window(title=title, control_type=control_type).wrapper_object()
            if action == 'click':
                control.click_input()
            elif action == 'select':
                control.select()
            time.sleep(0.3)
        except ElementNotFoundError:
            self.logger.error(f"Control with title '{title}' or auto_id '{auto_id}' not found")
            raise ElementNotFoundError
        except Exception as e:
            self.logger.error("Error performing click action", exc_info=True)
            raise e

    def send_message(self, number, name) -> bool:
        """Sends a message to a specific number."""
        try:
            self.logger.info(f"Sending message to {name} : {number}")
            pg.hotkey('ctrl', 'n')
            time.sleep(0.2)
            self._perform_click("Phone number", auto_id="PhoneNumberDialButton")
            pg.write(number)
            time.sleep(0.3)
            try:
                self._perform_click("Chat")
            except ElementNotFoundError:
                self.logger.info('Pressing ESC')
                for i in range(2):
                    pg.press('esc')
                    time.sleep(0.2)
                time.sleep(1)
                return False, 'Number not on WhatsApp' 

            message_input = self.app.WhatsApp.child_window(auto_id="InputBarTextBox", control_type="Edit").wrapper_object()
            message_input.click_input()
            time.sleep(0.3)

            with open(self.config.whatsapp_config.message_txt, 'r', encoding='utf-8') as msg_file:
                message = msg_file.read().replace('<name>', name)
            pyc.copy(message)
            time.sleep(0.3)
            pg.hotkey('ctrl', 'v')
            pg.press('enter')
            time.sleep(1)
            self.logger.info(f"Message sent to {name} : {number}")
            return True, 'Message Successfully Sent'

        except Exception as e:
            self.logger.error(f"Error sending message to {name} : {number}", exc_info=True)
            return False, 'Error in sending message'

    def send_message_with_attachment(self, number, name, attachment_path) -> bool:
        """Sends a message with an attachment to a specific number."""
        try:
            self.logger.info(f"Sending message with attachment to {name} : {number}")
            pg.hotkey('ctrl', 'n')
            time.sleep(0.2)
            self._perform_click("Phone number", auto_id="PhoneNumberDialButton")
            pg.write(number)
            time.sleep(0.3)
            try:
                self._perform_click("Chat")
            except ElementNotFoundError:
                self.logger.info('Pressing ESC')
                for i in range(2):
                    pg.press('esc')
                    time.sleep(0.2)
                
                time.sleep(1)
                return False, 'Number not on WhatsApp' 
            self._perform_click("Add attachment", auto_id="AttachButton")
            self._perform_click("Photos & videos", control_type="MenuItem", action='select')
            time.sleep(0.7)

            pg.write(attachment_path)
            pg.press('enter')
            time.sleep(1.5)

            with open(self.config.whatsapp_config.message_txt, 'r', encoding='utf-8') as msg_file:
                message = msg_file.read().replace('<name>', name)
            pyc.copy(message)
            time.sleep(0.3)
            pg.hotkey('ctrl', 'v')
            time.sleep(0.3)
            pg.press('enter')
            time.sleep(1)
            self.logger.info(f"Message with attachment sent to {name} : {number}")
            return True, 'Message Successfully Sent'
        except Exception as e:
            self.logger.error(f"Error sending message with attachment to {name} : {number}", exc_info=True)
            return False, 'Error in sending message'

    def extract_all_contact_groups(self) -> list:
        """Extracts all contacts from groups."""
        contact_details = []
        numbers = []

        try:
            self.logger.info("Extracting all contacts in groups")
            self._perform_click("Filter chats by", auto_id="FilterButton")
            self._perform_click("Groups", control_type="MenuItem")

            for _ in range(85):
                pg.press('tab')
                time.sleep(0.1)

            for _ in range(0):
                pg.press('down')

            for ind in range(85):
                pg.press('enter')
                self._perform_click(auto_id="SubtitleBlock", control_type="Text")
                self._perform_click("Members", auto_id="ParticipantsButton")

                try:
                    member_lst = self.app.WhatsApp.child_window(auto_id="MembersList", control_type="List").wrapper_object()
                    for _ in range(100):
                        member_lst.scroll('up', amount='page')
                except ElementNotFoundError:
                    self.logger.error("Members list not found", exc_info=True)
                    continue

                time.sleep(2)
                member_lst = self.app.WhatsApp.child_window(auto_id="MembersList", control_type="List").wrapper_object()
                time.sleep(1)
                numbers_remaining = True
                last_member_lst = []
                errors = 0

                while numbers_remaining:
                    try:
                        member_lst = self.app.WhatsApp.child_window(auto_id="MembersList", control_type="List").wrapper_object()
                        time.sleep(1)

                        if member_lst.items()[-2].texts() not in last_member_lst:
                            last_member_lst.append(member_lst.items()[-2].texts())
                            for member in member_lst.items():
                                member_data = member.texts()
                                for i in range(1, 3):
                                    try:
                                        number = str(member_data[i]).replace(' ', '').replace('+', '').replace('(', '').replace(')', '').replace('-', '')
                                        if number:
                                            if number not in numbers:
                                                numbers.append(number)
                                                contact_details.append({'Name': str(member_data[0]), 'Number': number})
                                                self.logger.info({'Name': str(member_data[0]), 'Number': number})
                                    except (IndexError, ValueError):
                                        continue

                            for _ in range(5):
                                member_lst.scroll('down', amount='page')
                                time.sleep(0.1)
                        else:
                            numbers_remaining = False
                    except Exception as e:
                        if errors >= 5:
                            numbers_remaining = False
                        else:
                            errors += 1
                            self.logger.error(f"Error occurred in extracting members of group {ind}", exc_info=True)

                message_input_exists = self.app.WhatsApp.child_window(auto_id="InputBarTextBox", control_type="Edit").exists()
                if message_input_exists:
                    pg.press('esc')
                    time.sleep(0.2)
                pg.press('esc')
                time.sleep(0.2)
                pg.press('down')
                time.sleep(1)

            self.logger.info(f"Extracted all contacts in groups: {len(contact_details)}")
            return contact_details
        except Exception as e:
            self.logger.error("Error extracting all contacts in groups", exc_info=True)
            return contact_details

    def fun_template(self) -> bool:
        """Template function for future use."""
        try:
            self.logger.info("Template function called")
            return True
        except Exception as e:
            self.logger.error("Error in template function", exc_info=True)
            return False
