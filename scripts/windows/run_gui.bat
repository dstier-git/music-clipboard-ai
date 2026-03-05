@echo off
setlocal
for %%I in ("%~dp0..\..") do set "REPO_ROOT=%%~fI"
set "PYTHONPATH=%REPO_ROOT%\src;%REPO_ROOT%;%PYTHONPATH%"
pythonw -m music_clipboard.gui.app %*
exit /b %ERRORLEVEL%
