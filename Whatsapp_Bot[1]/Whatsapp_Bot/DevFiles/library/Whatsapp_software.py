import os
import pdb
import time
import psutil
import subprocess
import pyautogui as pg
import pyperclip as pyc
from pywinauto.application import Application

class WhatsAppHandler:

    def __init__(self, config, logger) -> None:
        self.config = config
        self.logger = logger
        # self.app.WhatsApp.print_control_identifiers()
        # phone_number_btn = self.app.WhatsApp.child_window(title="Phone number", auto_id="PhoneNumberDialButton", control_type="Button").wrapper_object()
        # phone_number_btn.click()
        # time.sleep(0.3)

    def kill_whatsapp(self):
        try:
            # Iterate over all running processes
            for process in psutil.process_iter(['pid', 'name']):
                # Check if the process name matches "WhatsApp.exe"
                if process.info['name'] == 'WhatsApp.exe':
                    # Kill the process
                    process.kill()
                    self.logger.info(f"Killed WhatsApp process with PID {process.info['pid']}")
        
        except Exception as e:
            self.logger.error(f'# Error in killing whatsapp')
            self.logger.error(e)

    def start_whatsapp(self) -> bool:
        try:
            self.logger.info(f"# Starting whatsapp")
            self.kill_whatsapp()
            
            subprocess.Popen(["cmd", "/C", "start whatsapp://send"], shell=True)
            
            self.logger.info(f"# Whatsapp started")
            return True
        
        except Exception as e:
            self.logger.error(f"# Error in starting whatsapp")
            self.logger.error(e, exc_info=True)
            return False

    def connect_whatsapp(self) -> bool:
        try:
            self.logger.info(f"# Connecting to WhatsApp")
            self.app = Application(backend='uia').connect(title='WhatsApp', timeout=100)

            self.logger.info(f"# Connected to WhatsApp")
            return True
        
        except Exception as e:
            self.logger.error(f"# Error in connecting to WhatsApp")
            self.logger.error(e, exc_info=True)
            return False

    def send_message(self, number, name) -> bool:
        try:
            self.logger.info(f"#  Sending message to {name} : {number}")
            
            pg.hotkey('ctrl', 'n')
            time.sleep(0.2)

            phone_number_btn = self.app.WhatsApp.child_window(title="Phone number", auto_id="PhoneNumberDialButton", control_type="Button").wrapper_object()
            phone_number_btn.click()
            time.sleep(0.3)

            pg.write(number)
            time.sleep(0.3)

            chat_btn = self.app.WhatsApp.child_window(title="Chat", control_type="Button").wrapper_object()
            chat_btn.click()
            time.sleep(0.3)

            message_input = self.app.WhatsApp.child_window(title="Type a message", auto_id="PlaceholderTextContentPresenter", control_type="Text").wrapper_object()
            message_input.click_input()
            time.sleep(0.3)

            with open(self.config.whatsapp_config.message_txt, 'r', encoding='utf-8') as msg_file:
                message = msg_file.read()
            message = message.replace('<name>', name)
            pyc.copy(message)
            time.sleep(0.3)
            pg.hotkey('ctrl', 'v')            

            pg.press('enter')
            time.sleep(1)

            self.logger.info(f"# Message sent to {name} : {number}")
            return True

        except Exception as e:
            self.logger.error(f"# Error in sending message to {name} : {number}")
            self.logger.error(e, exc_info=True)
            return False

    def extract_all_contact_groups(self) -> list:
        try:
            self.logger.info(f"# Extracting all contact in groups")
            contact_details = []
            numbers = []

            filter_btn = self.app.WhatsApp.child_window(title="Filter chats by", auto_id="FilterButton", control_type="Button").wrapper_object()
            filter_btn.click_input()
            time.sleep(0.3)
            
            groups_btn = self.app.WhatsApp.child_window(title="Groups", control_type="MenuItem").wrapper_object()
            groups_btn.select()
            time.sleep(0.3)
            
            for i in range(1):
                pg.press('tab')
                time.sleep(0.1)
            
            for i in range(0):
                pg.press('down')

                
            for ind in range(85):
                member_lst = None
                pg.press('enter')
                print(i)

                group_name = self.app.WhatsApp.child_window(auto_id="SubtitleBlock", control_type="Text").wrapper_object()
                group_name.click_input()
                time.sleep(0.3)
                
                member_btn = self.app.WhatsApp.child_window(title="Members", auto_id="ParticipantsButton", control_type="ListItem").wrapper_object()
                member_btn.click_input()
                time.sleep(0.3)
                
                try:
                    member_lst = self.app.WhatsApp.child_window(auto_id="MembersList", control_type="List").wrapper_object()                    
                    for i in range(100):
                        member_lst.scroll('up', amount='page')

                except:
                    pass
                
                time.sleep(2)

                
                member_lst = self.app.WhatsApp.child_window(auto_id="MembersList", control_type="List").wrapper_object()

                time.sleep(1)
                print(f'------------------- Len of number : {len(numbers)}')
                print(f'------------------- Len of members {ind} group : {len(member_lst.items())}')

                numbers_remaining = True
                last_member_lst = []
                errors = 0
                
                while numbers_remaining:
                    try:
                        print('again while loop')
                        member_lst = self.app.WhatsApp.child_window(auto_id="MembersList", control_type="List").wrapper_object()
                        time.sleep(1)
                        print(f'------------------- Len of number : {len(numbers)}')
                        print(f'------------------- Len of members {ind} group : {len(member_lst.items())}')
                        
                        print(f'{member_lst.items()[-2].texts()} ----------------------------------------')
                        if member_lst.items()[-2].texts() not in last_member_lst:
                            last_member_lst.append(member_lst.items()[-2].texts())                    
                            for ind, member in enumerate(member_lst.items()):
                                try:
                                    member_data = member.texts()
                                    try:
                                        if float(str(member_data[1]).replace(' ','').replace('+', '').replace('(', '').replace(')', '').replace('-', '')):
                                                if str(member_data[1]) not in numbers:
                                                    numbers.append(str(member_data[1]))
                                                    contact_details.append({'Name':'', 'Number':str(member_data[1])})
                                                    print({'Name':'', 'Number':str(member_data[1])})
                                                    self.logger.info({'Name':'', 'Number':str(member_data[1])})
                                    except:
                                        pass    

                                    try:
                                        if float(str(member_data[2]).replace(' ','').replace('+', '').replace('(', '').replace(')', '').replace('-', '')):
                                            if str(member_data[2]) not in numbers:
                                                numbers.append(str(member_data[2]))
                                                contact_details.append({'Name':str(member_data[1]), 'Number':str(member_data[2])})                                
                                                print({'Name':str(member_data[1]), 'Number':str(member_data[2])})
                                                self.logger.info({'Name':str(member_data[1]), 'Number':str(member_data[2])})
                                    except:
                                        pass
                                
                                except Exception as e:
                                    self.logger.error(f'error occured in for loop')
                                    self.logger.error(e)

                            try:
                                for i in range(5):
                                    member_lst.scroll('down', amount='page')
                                    time.sleep(0.1)
                            
                            except:
                                pass
                        else:
                            print('else called of main if')
                            numbers_remaining = False
                    except Exception as e:
                        if errors >= 5:
                            numbers_remaining = False
                        else:
                            errors += 1
                            self.logger.error(f'error occured in while loop of {ind} group')
                            self.logger.error(e, exc_info=True)
                # message_input = self.app.WhatsApp.child_window(title="Type a message", auto_id="PlaceholderTextContentPresenter", control_type="Text").exists()
                message_input = self.app.WhatsApp.child_window(auto_id="InputBarTextBox", control_type="Edit").exists()
                # print(message_input)

                if message_input:
                    pg.press('esc')
                    time.sleep(0.2)
                                        
                pg.press('esc')
                time.sleep(0.2)
                pg.press('down')
                time.sleep(1)


            # member_lst.items()[36].texts()

            # self.app.WhatsApp.print_control_identifiers()
            
            self.logger.info(f"# Extracted all contact in groups: {len(contact_details)}")
            return contact_details

        except Exception as e:
            self.logger.error(f"# Error in extracting all contact in groups")
            self.logger.error(e, exc_info=True)
            return contact_details
   
    def fun_template(self) -> bool:
        try:
            self.logger.info(f"# ")

            self.logger.info(f"# ")
            return True
        
        except Exception as e:
            self.logger.error(f"# Error in")
            self.logger.error(e, exc_info=True)
            return False
