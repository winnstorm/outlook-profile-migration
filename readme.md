# Outlook Profile Backup and Restore

This Python script allows you to backup your current Outlook profiles and settings and then restore them to Outlook 2021.

## Getting Started

### Prerequisites

* Windows Operating System
* Python 3.x

### Installing

1. Download the Python script to your local machine.
2. Open a command prompt or terminal and navigate to the directory where the script is saved.
3. Run the script using the command `python outlook_profile_backup_restore.py`.

### Usage

When you run the script, it will backup your current Outlook profiles and settings and save them to a directory of your choice. 

To restore the profiles to Outlook 2021:

1. Open Outlook 2021.
2. Click the "File" tab in the top left corner.
3. Click "Open & Export" on the left-hand side.
4. Click "Import/Export".
5. Select "Import from another program or file" and click "Next".
6. Select "Outlook Data File (.pst)" and click "Next".
7. Browse to the directory where you saved the backup files and select the file you want to restore.
8. Select the option to "Import items into the same folder in" and choose the new profile for Outlook 2021.
9. Click "Finish" to restore the backup.

## Customization

You can modify the following variables in the script to customize it for your specific needs:

* `backup_path`: The path where you want to store the backup files.
* `outlook_2016_path`: The registry path for the Outlook 2016 profiles.
* `outlook_2021_path`: The registry path for the Outlook 2021 profiles.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

* This script is based on code by [ChatGPT](https://github.com/ChatGPT).
* The [winreg](https://docs.python.org/3/library/winreg.html) and [shutil](https://docs.python.org/3/library/shutil.html) Python modules were used to backup and restore the Outlook profiles.
