import os
import time
import subprocess
import pyautogui as pg
import pyperclip as pyc
from library.GetLogger import apply_logs_to_all_methods, log

# @apply_logs_to_all_methods(log)
class ChromeHandler:

    def __init__(self, logger, config) -> None:
        self.logger = logger
        self.config = config
        self.kill_all_chrome()

    def kill_all_chrome(self):
        os.system("taskkill /F /IM chrome.exe")

    def start_chrome(self) -> bool:
        try:
            self.logger.info(f"# Starting chrome")
            self.process = subprocess.Popen([self.config.chrome_config.chrome_path])
            self.logger.info(f"     # Chrome started")
            time.sleep(5)
            return True
        
        except Exception as e:
            self.logger.info(f'     # Error while starting chrome')
            self.logger.error(e, exc_info=True)
            return False

    def select_profile(self, profile_index:int) -> bool:
        try:
            self.logger.info(f"# Selecting profile")
            pg.press('tab')
            if profile_index != 1:
                for i in range(profile_index):
                    pg.press('tab')
                    pg.press('tab')
                    pg.press('tab')
                    time.sleep(0.3)

            time.sleep(2)
            pg.press('enter')
            self.logger.info(f"     # Profile selected")
            return True

        except Exception as e:
            self.logger.info(f'     # Error while selecting profile')
            self.logger.error(e, exc_info=True)
            return False

    def maximise_chrome(self) -> bool:
        try:
            self.logger.info(f"# Maximising chrome")
            time.sleep(3)
            pg.hotkey('alt', 'space')
            time.sleep(0.5)
            pg.press('x')

            self.logger.info(f"     # Chrome maximised")
            return True

        except Exception as e:
            self.logger.info(f'     # Error while maximising chrome')
            self.logger.error(e, exc_info=True)
            return False

    def load_whatsapp(self) -> bool:
        try:
            pg.hotkey('ctrl', 'l')
            self.logger.info(f"# Starting whatsapp")
            pg.write(self.config.whatsapp_config.whatsapp_url)
            time.sleep(1)
            pg.press('enter')

            self.logger.info(f"     # Whatsapp started")
            return True

        except Exception as e:
            self.logger.info(f'     # Error while starting whatsapp')
            self.logger.error(e, exc_info=True)
            return False

    def send_message(self, name, number) -> bool:
        try:
            self.logger.info(f"     # Sending message to {name} : {number}")
            time.sleep(3)
            pg.hotkey('ctrl', 't')
            time.sleep(1)
            pg.hotkey('ctrl', 'tab')
            time.sleep(1)
            pg.hotkey('ctrl', 'w')
            time.sleep(1)
            
            message = self.create_message(name=name)
            pg.write(str(self.config.whatsapp_config.whatsapp_msg).replace("'", "").replace('<number>', number))
            time.sleep(0.5)
            pg.press('enter')
            time.sleep(10)
            
            pyc.copy(message)
            time.sleep(0.3)
            
            pg.hotkey('ctrl', 'v')
                        
            time.sleep(1)
            pg.press('enter')
            time.sleep(3)

            self.logger.info(f"         # Sent message to {name} : {number}")
            
            return True

        except Exception as e:
            self.logger.info(f'     # Error while sending message to {name} : {number}')
            self.logger.error(e, exc_info=True)
            return False                

    def create_message(self, name:str) -> str:
        try:
            self.logger.info(f"     # Generating message for {name}")
            with open(self.config.whatsapp_config.message_txt, 'r') as msg_file:
                message = msg_file.read()

            message = message.replace('<name>', name)
            self.logger.info(f"         # Generated message for {name}")
            
            return message
        
        except Exception as e:
            self.logger.info(f'         # Error while generating message for {name}')
            self.logger.error(e, exc_info=True)
            return False

    def template_fun(self) -> bool:
        try:
            self.logger.info(f"# ")

            self.logger.info(f"     # ")
            return True

        except Exception as e:
            self.logger.info(f'     # Error while ')
            self.logger.error(e, exc_info=True)
            return False
