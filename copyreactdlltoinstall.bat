set _targetdir=%INSTALLROOT%\x64\debug\devclient\FILES\PFILES\MSOFFICE\Office16
del %_targetdir%\cpprest140d_2_8.dll
mklink  %_targetdir%\cpprest140d_2_8.dll %TARGETROOT%\x64\debug\netui\x-none\cpprest140d_2_8.dll
del %_targetdir%\react-native-win32.dll
mklink  %_targetdir%\react-native-win32.dll %TARGETROOT%\x64\debug\netui\x-none\react-native-win32.dll


