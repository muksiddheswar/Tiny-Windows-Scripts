@ECHO OFF

ECHO Execting file cleanup...
ECHO Recursively deleting all *.exe file in the current directory.

powershell -Command "& {ls *.txt -Recurse | foreach{echo $_}}"
powershell -Command "& {ls *.txt -Recurse | foreach{rm $_}}"


PAUSE

