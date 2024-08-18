import os
import shutil
import datetime
import pandas as pd
from library.Config import Config
from library.GetLogger import GetLogger
from library.chrome_handler import ChromeHandler
from library.Messenger import show_success_message, show_danger_message

class ExcelHandler:

    def __init__(self, logger, config, chrome_handler) -> None:
        self.logger = logger
        self.config = config
        self.chrome_handler = chrome_handler
        
    def get_excels(self) -> list:
        try:
            self.logger.info(f"# Getting all excels in unprocessed folder")
            excel_files = [file for file in os.listdir(self.config.paths.unprocessed_path) if '.xlsx' in file.lower()]
            self.logger.info(f"     # Got {len(excel_files)} excels in unprocessed folder")
            return excel_files

        except Exception as e:
            self.logger.info(f'     # Error while getting all excels in unprocessed folder')
            self.logger.error(e, exc_info=True)
            return None
    
    def gen_df(self, excel_path) -> pd.DataFrame:
        try:
            self.logger.info(f"         # Generating df of excel")
            excel_path = os.path.join(self.config.paths.unprocessed_path, excel_path)
            if os.path.exists(excel_path):
                df = pd.read_excel(excel_path)
                self.logger.info(f"         # Generated df of excel")
                return df
            
            else:
                self.logger.info(f"         # Path '{excel_path}' doesnot exist")            
                return None

        except Exception as e:
            self.logger.info(f'     # Error while generating df of excel')
            self.logger.error(e, exc_info=True)
            return None

    def handle_excels(self, excel_files:list) -> bool:
        try:
            self.logger.info(f"# Processing {len(excel_files)} excel files")

            for excel in excel_files:
                self.logger.info(f'     # Processing {excel}')
                df = self.gen_df(excel_path=excel)
                self.logger.info(f'         # Total {len(df)} rows')
                
                error_df = self.handle_df(df=df)
                if error_df is not None:
                    error_df.to_excel(os.path.join(self.config.paths.error_path, f"Error_{datetime.datetime.now().strftime('%m-%d-%Y_%H-%M-%S')}.xlsx"), index=False)
                
                shutil.move(os.path.join(self.config.paths.unprocessed_path, excel), os.path.join(self.config.paths.processed_path, excel))


            self.logger.info(f"     # Processed {len(excel_files)} excel files")
            return True

        except Exception as e:
            self.logger.info(f'     # Error while processing {len(excel_files)} excel files')
            self.logger.error(e, exc_info=True)
            return False

    def handle_df(self, df:pd.DataFrame):
        try:
            self.logger.info(f"     # Processing df")
            error_df_lst = []
            for ind, row in df.iterrows():
                try:
                    if len(str(row['Mobile'])) == 10:
                        self.logger.info(f"     # Sending message to {str(row['Name'])}")                        
                        message_sent = self.chrome_handler.send_message(name=str(row['Name']), number=str(row['Mobile']))
                    
                    else:
                        self.logger.error(f'        # Mobile number invalid length of number : {len(str(row["Mobile"]))}')
                        message_sent = False
                
                except Exception as e:
                    self.logger.info(f'         # Error while sending message to {str(row["Name"])}')
                    self.logger.error(e, exc_info=True)
                    message_sent = False
                
                if not message_sent:
                    error_df_lst.append({'Name':str(row['Name']), 'Mobile':str(row['Mobile'])})

            if len(error_df_lst) > 0:
                error_df = pd.DataFrame(error_df_lst)
                return error_df

            else:
                return None
            
        except Exception as e:
            self.logger.info(f'         # Error while processing df')
            self.logger.error(e, exc_info=True)
            return None

    def template_fun(self) -> bool:
        try:
            self.logger.info(f"# ")

            self.logger.info(f"     # ")
            return True

        except Exception as e:
            self.logger.info(f'     # Error while ')
            self.logger.error(e, exc_info=True)
            return False


if __name__ == "__main__":
    cwd_path = os.getcwd()
    config_path = cwd_path.replace('DevFiles', 'BotConfig\\config.ini')
    config = Config(filename=config_path)

    logs_dir: str = config.paths.logs_path
    logging = GetLogger(log_file_dir=logs_dir, log_file_name=f"message_sender.log", file_handler=True)
    logger = logging.logger

    chrome_handler = ChromeHandler(logger=logger, config=config)
    chrome_started = chrome_handler.start_chrome()

    if chrome_started:
        profile_selected = chrome_handler.select_profile(profile_index=int(config.chrome_config.profile_index))

        if profile_selected:
            maximising_chrome = chrome_handler.maximise_chrome()
            
            if maximising_chrome:
                whatsapp_loaded = chrome_handler.load_whatsapp()
                
                if whatsapp_loaded:
                    excel_handler = ExcelHandler(logger=logger, config=config, chrome_handler=chrome_handler)
                    
                    excel_files = excel_handler.get_excels()
                    if len(excel_files) > 0:
                        excel_handler.handle_excels(excel_files=excel_files)
                        
                        chrome_handler.kill_all_chrome()
                        
                        # show_success_message()
