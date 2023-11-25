[Calibre Update Backup.bat](https://github.com/DonGrobione/Calibre-Update-Backup-Script/blob/main/Calibre%20Update%20Backup.bat) is depricated and will no longer updated.

A simple script that will download the latest Calibre Portable update from https://calibre-ebook.com/download_portable. Then it will use 7zip to create a backup file of the Calibre Portable folder. If you have your library located in a sub folder, it will be backed up too.

I set it up, that the archive volume will be split every 4 GB, as I use OneDrive and sometimes there are issues with big files and/or some corporate policies. If this does not apply to you, just remove that parameter.


[Calibre Update Backup.ps1](https://github.com/DonGrobione/Calibre-Update-Backup-Script/blob/main/Calibre%20Update%20Backup.ps1) is my current script. It will do the same as the old bat file, downloading the calibre update and creating a update. But it will change the path to the backup folder depending on oyu hostname. I am running this on several hosts and was to lazy to maintain several scripts.
Finally it will check your backup path and will only keep the latest three backups.