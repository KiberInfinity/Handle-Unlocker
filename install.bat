@Echo Off
SET scriptpath=%~dp0
echo Handle Unlock (GUI for SysInternals Handle.exe) Installer
echo Install from %scriptpath%

call :IsAdmin

REG DELETE "HKCR\*\shell\HandleUnlocker" /F
REG DELETE "HKCR\Folder\shell\HandleUnlocker" /F

REG ADD "HKCR\*\shell\HandleUnlocker" /ve /t REG_SZ /d "Unlock handle" /f
REG ADD "HKCR\*\shell\HandleUnlocker" /v "Icon" /t REG_SZ /d "\"%scriptpath%icon.ico\"" /f
REG ADD "HKCR\*\shell\HandleUnlocker\command" /ve /t REG_SZ /d "cmd /C PowerShell.exe -windowstyle hidden -ExecutionPolicy unrestricted -File \"%scriptpath%UnlockHandle.ps1\" -target \"%%1\"" /f
REG ADD "HKCR\Folder\shell\HandleUnlocker" /ve /t REG_SZ /d "Unlock handle" /f
REG ADD "HKCR\Folder\shell\HandleUnlocker" /v "Icon" /t REG_SZ /d "\"%scriptpath%icon.ico\"" /f
REG ADD "HKCR\Folder\shell\HandleUnlocker\command" /ve /t REG_SZ /d "cmd /C PowerShell.exe -windowstyle hidden -ExecutionPolicy unrestricted -File \"%scriptpath%UnlockHandle.ps1\" -target \"%%1\"" /f
Exit

:IsAdmin
REG QUERY "HKU\S-1-5-19\Environment"
If Not %ERRORLEVEL% EQU 0 (
 Cls & Echo You must have administrator rights to continue ... 
 Pause & Exit
)
Cls
goto:eof