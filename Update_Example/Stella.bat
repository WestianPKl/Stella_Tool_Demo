@echo off

echo Update in progress.

echo Downloading...
copy %~p0\dummy.txt
timeout /t 2
echo Update completed.

start dummy.txt

exit


