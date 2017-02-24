@echo OFF

reg Query "HKLM\Hardware\Description\System\CentralProcessor\0" | find /i "x86" > NUL && set OS=32BIT || set OS=64BIT


if %OS%==64BIT goto ESCAPE

:Net
start /wait prerequisite\NDP461-KB3102436-x86-x64-AllOS-ENU.EXE || goto ERROR

:office
start /wait prerequisite\vstor_redist.exe || goto ERROR

:Copy
XCOPY UnityTranslatorExcelAddins "%programfiles%/Schneider Electric" /s /i /y || goto ERROR

:registry
rem  set __COMPAT_LAYER=RunAsInvoker  
REGEDIT.EXE  /S  "excel32.reg"


echo installation successfull...
PAUSE
exit 


:ERROR
echo installation errors...
PAUSE
exit 

:ESCAPE
echo 64 bit system detected, please use x64 installer instead ...
PAUSE
exit 


