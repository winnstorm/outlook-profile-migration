import winreg
import shutil
import os

# Define the path for the backup file
backup_path = r'C:\OutlookBackup'

# Define the registry paths for the Outlook profiles
outlook_2016_path = r'Software\Microsoft\Office\16.0\Outlook\Profiles'
outlook_2021_path = r'Software\Microsoft\Office\17.0\Outlook\Profiles'

# Create the backup directory if it doesn't exist
if not os.path.exists(backup_path):
    os.makedirs(backup_path)

# Backup the Outlook profiles
for path in [outlook_2016_path, outlook_2021_path]:
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path)
        profile_count = winreg.QueryInfoKey(key)[0]
        for i in range(profile_count):
            profile_name = winreg.EnumKey(key, i)
            profile_path = os.path.join(path, profile_name)
            backup_file = os.path.join(backup_path, f'{profile_name}.pst')
            if os.path.exists(backup_file):
                os.remove(backup_file)
            shutil.copytree(profile_path, backup_file)
            print(f'Profile "{profile_name}" backed up to "{backup_file}"')
    except FileNotFoundError:
        pass

# Restore the Outlook profiles to Outlook 2021
outlook_2021_profile_path = os.path.join(outlook_2021_path, profile_name)
for backup_file in os.listdir(backup_path):
    backup_file_path = os.path.join(backup_path, backup_file)
    profile_name = os.path.splitext(backup_file)[0]
    if not os.path.exists(outlook_2021_profile_path):
        os.makedirs(outlook_2021_profile_path)
    shutil.copytree(backup_file_path, os.path.join(outlook_2021_profile_path, backup_file))
    print(f'Profile "{profile_name}" restored to "{outlook_2021_profile_path}"')

print('Outlook profiles backup and restore complete.')
