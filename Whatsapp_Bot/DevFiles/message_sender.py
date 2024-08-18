from library.Config import Config
from library.GetLogger import GetLogger
from library.Whatsapp_software import WhatsAppHandler
from library.summary_window import show_summary  # Import the summary utility
import os
import pandas as pd
import time
import datetime
import shutil


def move_file_to_processed(unprocessed_path, processed_path, file_name):
    """Move file to processed folder, handling naming conflicts."""
    process_file_path = os.path.join(processed_path, file_name)
    counter = 0
    while os.path.exists(process_file_path):
        base_name, ext = os.path.splitext(file_name)
        process_file_path = os.path.join(processed_path, f"{base_name}_{counter}{ext}")
        counter += 1
    shutil.move(os.path.join(unprocessed_path, file_name), process_file_path)
    return process_file_path


def process_excel_file(excel_file, whatsapp_handler, config, logger):
    """Process an individual Excel file to send messages."""
    error_df_lst = []
    excel_df = pd.read_excel(os.path.join(config.paths.unprocessed_path, excel_file))
    for ind, row in excel_df.iterrows():
        if len(str(row['Number'])) == 10:
            if str(config.whatsapp_config.message_with_attachemnt).lower() == 'true':
                message_sent, sending_status = whatsapp_handler.send_message_with_attachment(
                    name=str(row['Name']),
                    number=str(row['Number']),
                    attachment_path=config.whatsapp_config.attachment_photo
                )
            else:
                message_sent, sending_status = whatsapp_handler.send_message(
                    name=str(row['Name']),
                    number=str(row['Number'])
                )

            logger.info(f'Message sent to: {ind + 1}')
            time.sleep(0.5)
        else:
            logger.error(f'# Mobile number invalid length of number: {len(str(row["Number"]))} : {str(row["Number"])}')
            message_sent = False
            sending_status = f'Mobile number invalid length of number: {len(str(row["Number"]))}'

        if not message_sent:
            error_df_lst.append({'Name': str(row['Name']), 'Number': str(row['Number']), 'Remarks':sending_status})

    return error_df_lst


def main():
    cwd_path = os.getcwd()
    config_path = cwd_path.replace('DevFiles', 'BotConfig\\config.ini')
    config = Config(filename=config_path)

    logs_dir = config.paths.logs_path
    logging = GetLogger(log_file_dir=logs_dir, log_file_name="message_sender.log", file_handler=True)
    logger = logging.logger

    whatsapp_handler = WhatsAppHandler(config=config, logger=logger)
    
    if not whatsapp_handler._start_whatsapp():
        logger.error("Failed to start WhatsApp.")
        return

    if not whatsapp_handler._connect_whatsapp():
        logger.error("Failed to connect to WhatsApp.")
        return

    time.sleep(5)  # Give WhatsApp some time to connect

    # if not whatsapp_handler._is_whatsapp_foreground() or not whatsapp_handler._is_whatsapp_on_screen():
    #     logger.error("WhatsApp is not in focus or not on the screen. Exiting.")
    #     whatsapp_handler._kill_whatsapp()
    #     return

    excel_files = os.listdir(config.paths.unprocessed_path)
    all_errors = []
    total_messages_sent = 0

    for excel_file in excel_files:
        error_df_lst = process_excel_file(excel_file, whatsapp_handler, config, logger)
        all_errors.extend(error_df_lst)
        total_messages_sent += len(pd.read_excel(os.path.join(config.paths.unprocessed_path, excel_file)))
        move_file_to_processed(config.paths.unprocessed_path, config.paths.processed_path, excel_file)

    if len(all_errors) > 0:
        error_df = pd.DataFrame(all_errors)
        error_file_name = f"Error_{datetime.datetime.now().strftime('%m-%d-%Y_%H-%M-%S')}.xlsx"
        error_file_path = os.path.join(config.paths.error_path, error_file_name)
        error_df.to_excel(error_file_path, index=False)
        logger.info(f"Errors recorded in file: {error_file_path}")

    time.sleep(5)
    whatsapp_handler._kill_whatsapp()
    
    # Display summary
    if len(all_errors) > 0:
        show_summary(total_messages_sent-len(all_errors), len(all_errors), error_file_path=error_file_path)
    
    else:
        show_summary(total_messages_sent-len(all_errors), len(all_errors))


if __name__ == "__main__":
    main()
