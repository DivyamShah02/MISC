from library.Config import Config
from library.GetLogger import GetLogger
from library.Whatsapp_software import WhatsAppHandler
import os
import pandas as pd


if __name__ == "__main__":
    cwd_path = os.getcwd()
    config_path = cwd_path.replace('DevFiles', 'BotConfig\\config.ini')
    config = Config(filename=config_path)

    logs_dir: str = config.paths.logs_path
    logging = GetLogger(log_file_dir=logs_dir, log_file_name=f"message_sender.log", file_handler=True)
    logger = logging.logger

    whatsapp_handler = WhatsAppHandler(config=config, logger=logger)
    
    whatsapp_started = whatsapp_handler.start_whatsapp()
    
    if whatsapp_started:
        whatsapp_connected = whatsapp_handler.connect_whatsapp()
        
        if whatsapp_connected:
            groups = whatsapp_handler.extract_all_contact_groups()
            df = pd.DataFrame(groups)
            df.to_excel('test.xlsx', index=False)
