@echo off
:: **********************************************
:: * SCRIPT IS BELOW - NO MODIFICATION REQUIRED *
:: **********************************************
:: Install or Uninstall trigger
:PROMPT
echo/
echo Read-Only and New Instance context menu (right-click) for Excel (Office16 version)
echo Do you want to install(i) or uninstall(u) the context menu for Excel (i/u)?
echo/
SET /p confirmation=
IF /I "%confirmation%" NEQ "I" GOTO UNINSTALL
:: Install
:: XLSX
:: Right-click open as read-only
reg add "HKCU\Software\Classes\Excel.Sheet.12\shell\open as read-only" /ve /t REG_SZ /d "Open as Read-Only"
reg add "HKCU\Software\Classes\Excel.Sheet.12\shell\open as read-only" /v "Icon" /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\""
reg add "HKCU\Software\Classes\Excel.Sheet.12\shell\open as read-only\command" /ve /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\" /r \"%1\""
:: Right-click open as new instance
reg add "HKCU\Software\Classes\Excel.Sheet.12\shell\open as new instance" /ve /t REG_SZ /d "Open as New Instance"
reg add "HKCU\Software\Classes\Excel.Sheet.12\shell\open as new instance" /v "Icon" /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\""
reg add "HKCU\Software\Classes\Excel.Sheet.12\shell\open as new instance\command" /ve /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\" /x \"%1\""
:: XLSM
:: Right-click open as read-only
reg add "HKCU\Software\Classes\Excel.SheetMacroEnabled.12\shell\open as read-only" /ve /t REG_SZ /d "Open as Read-Only"
reg add "HKCU\Software\Classes\Excel.SheetMacroEnabled.12\shell\open as read-only" /v "Icon" /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\""
reg add "HKCU\Software\Classes\Excel.SheetMacroEnabled.12\shell\open as read-only\command" /ve /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\" /r \"%1\""
:: Right-click open as new instance
reg add "HKCU\Software\Classes\Excel.SheetMacroEnabled.12\shell\open as new instance" /ve /t REG_SZ /d "Open as New Instance"
reg add "HKCU\Software\Classes\Excel.SheetMacroEnabled.12\shell\open as new instance" /v "Icon" /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\""
reg add "HKCU\Software\Classes\Excel.SheetMacroEnabled.12\shell\open as new instance\command" /ve /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\" /x \"%1\""
:: XLSB
:: Right-click open as read-only
reg add "HKCU\Software\Classes\Excel.SheetBinaryMacroEnabled.12\shell\open as read-only" /ve /t REG_SZ /d "Open as Read-Only"
reg add "HKCU\Software\Classes\Excel.SheetBinaryMacroEnabled.12\shell\open as read-only" /v "Icon" /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\""
reg add "HKCU\Software\Classes\Excel.SheetBinaryMacroEnabled.12\shell\open as read-only\command" /ve /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\" /r \"%1\""
:: Right-click open as new instance
reg add "HKCU\Software\Classes\Excel.SheetBinaryMacroEnabled.12\shell\open as new instance" /ve /t REG_SZ /d "Open as New Instance"
reg add "HKCU\Software\Classes\Excel.SheetBinaryMacroEnabled.12\shell\open as new instance" /v "Icon" /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\""
reg add "HKCU\Software\Classes\Excel.SheetBinaryMacroEnabled.12\shell\open as new instance\command" /ve /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\" /x \"%1\""
:: XLS
:: Right-click open as read-only
reg add "HKCU\Software\Classes\Excel.Sheet.8\shell\open as read-only" /ve /t REG_SZ /d "Open as Read-Only"
reg add "HKCU\Software\Classes\Excel.Sheet.8\shell\open as read-only" /v "Icon" /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\""
reg add "HKCU\Software\Classes\Excel.Sheet.8\shell\open as read-only\command" /ve /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\" /r \"%1\""
:: Right-click open as new instance
reg add "HKCU\Software\Classes\Excel.Sheet.8\shell\open as new instance" /ve /t REG_SZ /d "Open as New Instance"
reg add "HKCU\Software\Classes\Excel.Sheet.8\shell\open as new instance" /v "Icon" /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\""
reg add "HKCU\Software\Classes\Excel.Sheet.8\shell\open as new instance\command" /ve /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\" /x \"%1\""
:: CSV
:: Right-click open as read-only
reg add "HKCU\Software\Classes\Excel.CSV\shell\open as read-only" /ve /t REG_SZ /d "Open as Read-Only"
reg add "HKCU\Software\Classes\Excel.CSV\shell\open as read-only" /v "Icon" /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\""
reg add "HKCU\Software\Classes\Excel.CSV\shell\open as read-only\command" /ve /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\" /r \"%1\""
:: Right-click open as new instance
reg add "HKCU\Software\Classes\Excel.CSV\shell\open as new instance" /ve /t REG_SZ /d "Open as New Instance"
reg add "HKCU\Software\Classes\Excel.CSV\shell\open as new instance" /v "Icon" /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\""
reg add "HKCU\Software\Classes\Excel.CSV\shell\open as new instance\command" /ve /t REG_SZ /d "\"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE\" /x \"%1\""


echo ************************************************************************
echo Install complete.
echo ************************************************************************

:: Unistall
:UNINSTALL
IF /I "%confirmation%" NEQ "U" GOTO END

reg delete "HKCU\Software\Classes\Excel.Sheet.12" /f
reg delete "HKCU\Software\Classes\Excel.SheetMacroEnabled.12" /f
reg delete "HKCU\Software\Classes\Excel.SheetBinaryMacroEnabled.12" /f
reg delete "HKCU\Software\Classes\Excel.Sheet.8" /f
reg delete "HKCU\Software\Classes\Excel.CSV" /f

echo ************************************************************************
echo Uninstall complete.
echo ************************************************************************
:END
echo Continue to exit.
pause