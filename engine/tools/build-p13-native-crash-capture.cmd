@echo off
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat"
if errorlevel 1 exit /b %errorlevel%
cl.exe /nologo /EHsc /std:c++17 /Zi /Fe:engine\target\p13-native-crash-capture.exe /Fo:engine\target\p13-native-crash-capture.obj engine\tools\p13-native-crash-capture.cpp dbghelp.lib
exit /b %errorlevel%
