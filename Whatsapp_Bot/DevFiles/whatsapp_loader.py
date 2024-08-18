import os
from library.Config import Config
from library.GetLogger import GetLogger
from library.chrome_handler import ChromeHandler
from library.Messenger import show_success_message, show_danger_message


if __name__ == "__main__":
    cwd_path = os.getcwd()
    config_path = cwd_path.replace('DevFiles', 'BotConfig\\config.ini')
    config = Config(filename=config_path)

    logs_dir: str = config.paths.logs_path
    logging = GetLogger(log_file_dir=logs_dir, log_file_name=f"whatsapp_loader.log", file_handler=True)
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
                    show_success_message()
                
                else:
                    show_danger_message()
