@Echo Off
SET scriptpath=%~dp0
echo Handle Unlock (GUI for SysInternals Handle.exe) Uninstaller
echo Unreg %scriptpath%

call :IsAdmin

REG DELETE "HKCR\*\shell\HandleUnlocker" /F
REG DELETE "HKCR\Folder\shell\HandleUnlocker" /F
Exit

:IsAdmin
REG QUERY "HKU\S-1-5-19\Environment"
If Not %ERRORLEVEL% EQU 0 (
 Cls & Echo You must have administrator rights to continue ... 
 Pause & Exit
)
Cls
goto:eof