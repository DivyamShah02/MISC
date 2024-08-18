import os
cwd_path = os.getcwd()
config_path = os.path.join(cwd_path, 'BotConfig', 'config.ini')

with open(config_path, 'r') as config_file:
    config_data = config_file.read()

config_data = config_data.replace('<BASE_DIR>', cwd_path)

with open(config_path, 'w') as config_file:
    config_file.write(config_data)
