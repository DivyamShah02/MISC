from library.Config import Config
from library.GetLogger import GetLogger
from library.Whatsapp_software import WhatsAppHandler
import os
import pandas as pd
import time


if __name__ == "__main__":
    cwd_path = os.getcwd()
    config_path = cwd_path.replace('DevFiles', 'BotConfig\\config.ini')
    config = Config(filename=config_path)

    logs_dir: str = config.paths.logs_path
    logging = GetLogger(log_file_dir=logs_dir, log_file_name=f"message_sender.log", file_handler=True)
    logger = logging.logger

    whatsapp_handler = WhatsAppHandler(config=config, logger=logger)
    
    whatsapp_started = whatsapp_handler._start_whatsapp()
    
    if whatsapp_started:
        whatsapp_connected = whatsapp_handler._connect_whatsapp()
        
        if whatsapp_connected:
            time.sleep(5)
            excel_files = os.listdir(config.paths.unprocessed_path)
            for excel_file in excel_files:
                excel_df = pd.read_excel(os.path.join(config.paths.unprocessed_path, excel_file))
                for ind, row in excel_df.iterrows():
                    message_send = whatsapp_handler.send_message_with_attachment(name=str(row['Name']), number=str(row['Number']), attachment_path=r"{}".format(config.paths.attachment_photo))
                    print(f'Message sent to : {ind+1}')
                    logger.info(f'Message sent to : {ind+1}')
                    time.sleep(0.5)
                
            time.sleep(5)

            whatsapp_handler._kill_whatsapp()