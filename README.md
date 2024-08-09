[Calibre Update Backup.ps1](https://github.com/DonGrobione/Calibre-Update-Backup-Script/blob/main/Calibre%20Update%20Backup.ps1) is my current script, which I use myself. Therefore you will need to change a few things, see comments for details.

It will do much the same as the old bat file (see section below), with a few extra steps.
Downloading the calibre update, creating a backup with changing path depending on the hostname, stopping OneDrive, appliying the update, starting OneDrive. Finally it will delete the update file and will only keep the latest three backups.

I implemented the hostname validation, as I am running this script on several machines and am to lazy to maintain specific scripts. Sometimes the calibre update failes due to some OneDrive file lock issue, therefore I will stop OnDrive temporarly for the Calibre Update.

---------------------------------------

All files in [Legacy](https://github.com/DonGrobione/Calibre-Update-Backup-Script/tree/main/Legacy) are no longer maintained.


[Calibre Update Backup.bat](https://github.com/DonGrobione/Calibre-Update-Backup-Script/blob/main/Legacy/Calibre%20Update%20Backup.bat) 

A simple script that will download the latest Calibre Portable update from https://calibre-ebook.com/download_portable. Then it will use 7zip to create a backup file of the Calibre Portable folder. If you have your library located in a sub folder, it will be backed up too.

I set it up, that the archive volume will be split every 4 GB, as I use OneDrive and sometimes there are issues with big files and/or some corporate policies. If this does not apply to you, just remove that parameter.