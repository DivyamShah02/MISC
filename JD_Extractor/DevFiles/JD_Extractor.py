import os
import re
import sys
import time
import shutil
import pandas as pd
import pyautogui as pg
import pyperclip as pyc
from library.Config import Config
from library.GetLogger import GetLogger
from library.ChromeHandler import ChromeHandler

def main():
    """Main function to automate Chrome tasks."""
    try:
        cwd_path = os.getcwd()
        config_path = cwd_path.replace('DevFiles', 'BotConfig\\config.ini')
        config = Config(filename=config_path)

        logs_dir = config.paths.logs_path
        logger = GetLogger(log_file_dir=logs_dir, log_file_name="chrome_automater.log", file_handler=True).logger

        chrome_handler = ChromeHandler(logger=logger, config=config)

        if not chrome_handler.start_chrome():
            logger.error("Failed to start Chrome.")
            sys.exit(1)

        if not chrome_handler.select_profile(profile_index=int(config.chrome_config.profile_index)):
            logger.error("Failed to select profile.")
            chrome_handler.kill_all_chrome()
            sys.exit(1)

        if not chrome_handler.maximise_chrome():
            logger.error("Failed to maximize Chrome.")
            chrome_handler.kill_all_chrome()
            sys.exit(1)

        excel_files = [file for file in os.listdir(config.paths.unprocessed_path) if file.lower().endswith('.xlsx')]

        for excel in excel_files:
            process_excel_file(excel, config, chrome_handler, logger)

        chrome_handler.kill_all_chrome()
        logger.info("Chrome automation completed successfully.")

    except Exception as e:
        logger.error("An unexpected error occurred.", exc_info=True)
        sys.exit(1)

def process_excel_file(excel, config, chrome_handler, logger):
    """Process each Excel file for data extraction."""
    try:
        df = pd.read_excel(os.path.join(config.paths.unprocessed_path, excel))
        number_lst = []
        error_lst = []

        for ind, row in df.iterrows():
            pg.hotkey('ctrl', 't')
            time.sleep(0.3)

            if not chrome_handler.load_url(url=f"view-source:{row['Link']}"):
                logger.error("Failed to load URL for row %d.", ind)
                chrome_handler.kill_all_chrome()
                sys.exit(1)

            time.sleep(5)  # Waiting for the page to load

            pg.hotkey('ctrl', 'a')
            time.sleep(0.3)
            pg.hotkey('ctrl', 'c')
            time.sleep(0.3)

            web_data = pyc.paste()
            number_extracted, number = extract_number(web_data)

            entry = {
                'Name': row['Name'],
                'Address': row['Address'],
                'Number': number,
            }

            if number_extracted:
                number_lst.append(entry)
            else:
                error_lst.append(entry)

            pg.hotkey('ctrl', 'w')

        save_results(number_lst, error_lst, excel, config)
        archive_file(excel, config)

    except Exception as e:
        logger.error("Error processing file %s: %s", excel, e, exc_info=True)
        raise

def extract_number(web_data):
    """Extract number from the web data."""
    try:
        start_index = web_data.find('"msg_num":"')
        if start_index == -1:
            return False, "msg_num not found"

        start_index += len('"msg_num":"')
        end_index = web_data.find('}', start_index)
        json_part = web_data[start_index:end_index+1]

        wup_start_index = json_part.find('"wup":["') + len('"wup":["')
        wup_end_index = json_part.find('"]', wup_start_index)
        number = json_part[wup_start_index:wup_end_index]

        pattern = r'\d+'
        match = re.search(pattern, number)

        if match:
            return True, match.group(0)
        else:
            return False, "No number found"

    except Exception as e:
        raise RuntimeError(f"Error extracting number: {e}")

def save_results(number_lst, error_lst, excel, config):
    """Save results to Excel files and handle naming conflicts."""
    try:
        if number_lst:
            number_df = pd.DataFrame(number_lst)
            processed_file_name = generate_file_name(folder=config.paths.processed_path, file_name=excel)
            number_df.to_excel(processed_file_name, index=False)

        if error_lst:
            error_df = pd.DataFrame(error_lst)
            error_file_name = generate_file_name(folder=config.paths.error_path, file_name=excel)
            error_df.to_excel(error_file_name, index=False)
    except Exception as e:
        raise RuntimeError(f"Error saving results: {e}")

def archive_file(excel, config):
    """Move processed file to archive."""
    try:
        archive_file_name = generate_file_name(folder=config.paths.archive_path, file_name=excel)
        shutil.move(os.path.join(config.paths.unprocessed_path, excel), archive_file_name)
    except Exception as e:
        raise RuntimeError(f"Error archiving file {excel}: {e}")

def generate_file_name(folder, file_name):
    """Generate unique file name to avoid conflicts."""
    process_file_path = os.path.join(folder, file_name)
    counter = 0
    while os.path.exists(process_file_path):
        base_name, ext = os.path.splitext(file_name)
        process_file_path = os.path.join(folder, f"{base_name}_{counter}{ext}")
        counter += 1
    return process_file_path

if __name__ == "__main__":
    main()
