@echo Copy Libs to SysWOW64 Folder...
echo path: %~dp0
copy %~dp0\TABCTL32.OCX %windir%\SysWOW64\
copy %~dp0\MSCOMCTL.OCX %windir%\SysWOW64\
copy %~dp0\TeeChart5.OCX %windir%\SysWOW64\
copy %~dp0\msflxgrd.OCX %windir%\SysWOW64\
copy %~dp0\comdlg32.OCX %windir%\SysWOW64\
@echo Register Libs...
regsvr32 %windir%\SysWOW64\MSCOMCTL.OCX /s
regsvr32 %windir%\SysWOW64\TABCTL32.OCX /s
regsvr32 %windir%\SysWOW64\TeeChart5.OCX /s
regsvr32 %windir%\SysWOW64\msflxgrd.OCX /s
regsvr32 %windir%\SysWOW64\comdlg32.OCX /s
@echo Libs registerd...
@pause