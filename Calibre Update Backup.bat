@echo off
cls

rem Calibe Portable Update Backup Script v1
rem This batch script will download the latest calibre portable installer, backup the current installation with 7zip, perform the update and finally delete the installer file.
rem latest version can be found here https://github.com/DonGrobione/Calibre-Update-Backup-Script/blob/717c4d9f7c0da1aa42557f3bb5328eb927f14449/Calibre%20Update%20Backup.bat

rem Var Definition, change as needed
rem CalibreFolder - full path to Calibe Portable location
set CalibreFolder="%OneDrive%\PortableApps\Calibre Portable"
rem CalibreInstaller - full path to update exe
set CalibreInstaller="%TEMP%\calibre-portable-installer.exe"
rem CalibreBackup - full path to backup, files will be named in YYYY-MM-DD
set CalibreBackup="%OneDrive%\_Backup\Calibre\CalibrePortableBackup_%DATE:~-4%-%DATE:~-7,2%-%DATE:~-10,2%.7z"

rem Check set var
echo CalibreFolder: %CalibreFolder%
echo CalibreBackup: %CalibreBackup%
echo CalibreInstaller: %CalibreInstaller%
echo

rem Download latest Calibre update exe
echo Download Upload file
echo Downloading %CalibreInstaller%
bitsadmin.exe /transfer "Calibre" https://calibre-ebook.com/dist/portable "%CalibreInstaller%"

rem Backup Current Calibre with 7zip
echo Create Backup
echo
"c:\Program Files\7-Zip\7z.exe" a -mx9 -v4g -bsp2 %CalibreBackup% %CalibreFolder%
rem a - create archive
rem mx9 - maximum compression
rem v4g - volume / file split after 4 GB
rem bsp - verboste activity stream

rem Perform Update
echo Updating Calibe Portable
echo
%CalibreInstaller% %CalibreFolder%

rem Delete Update
echo Delete Update file %CalibreInstaller%
echo
del "%CalibreInstaller%"
