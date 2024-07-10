@ECHO OFF
@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@REM
@REM Powershell Wraper Script
@REM
@SET SHELL=powershell.exe
@SET BASEPARM=-ExecutionPolicy Bypass -Command
@SET SCRIPTPATH=\\vdinas03\FolderRedirect\NSS\fbs_matsumoto_kazuhi\home\fbsmatsu\develop\AutoLogin
@SET MainScript=!SCRIPTPATH!\PROTO03.PS1
@SET ExBASE=DAT
@SET ExBook=ServerLoginList.xlsx
@SET ExSht1=DefineLoginSheet
@SET ExTbl1=DefineLogin

@REM +--------------------------------------------------------------------------------+
@REM +------+ Program Gimmick    PowerShell   ExcelBook         ExcelSheet ExcelTable
@REM +--------------------------------------------------------------------------------+
@SET LAUNCHER=!SHELL! !BASEPARM! !MainScript! !ExBASE!\!ExBook! !ExSht1! !ExTbl1!
@REM +--------------------------------------------------------------------------------+

@REM @ECHO !LAUNCHER!

!LAUNCHER!

@ENDLOCAL
@ECHO ON
