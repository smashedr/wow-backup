# WoW Backup and Restore

A simple script to backup and restore your World of Warcraft settings, addons, and addon settings.

In short it creates a `.zip` archive of your `Interface` and `WTF` folder with a version and date stamp.

Only works on Windows 10.

## Download

Latest Release: [Click Here to Download](https://github.com/smashedr/wow-backup/releases/latest/download/WoW-Backup.exe).

Or head over to the [releases page](https://github.com/smashedr/wow-backup/releases).

## How it Works

#### Backup

- Attempts to automatically detect your WoW install directory, otherwise lets you choose.
- Lets you pick which version of WoW to backup (`_retail_`, `_classic_`, `_ptr_`, etc...)
- Lets you chose which directory to store the backup archive.
- Displays details of backup and requiers you to press `enter` to proceed.
- Create a `.zip` archive with your `WTF` and `Interface` folder.

#### Restore

- Lets you pick the backup archive to restore from.
- Attempts to automatically detect your WoW install directory, otherwise lets you choose.
- Attempts to automatically detect the WoW version from the backup filename, otherwise aborts.
- Displays details of restore and requires you to type `confirm` before proceeding.
- Deletes your `WTF` and `Interface` folder and replaces them from the backup archive as displayed in previous step.
