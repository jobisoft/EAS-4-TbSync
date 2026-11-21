@echo off
:: Install the outlook_web_helper_host native messaging host for Firefox on Windows.
:: Copies the runtime files into %APPDATA%, so the source folder can be deleted afterwards.

setlocal

set "INSTALL_DIR=%APPDATA%\Mozilla\NativeMessagingHosts\outlook_web_helper_host_helper"
set "PY_SRC=%~dp0outlook_web_helper_host.py"
set "MANIFEST_SRC=%~dp0outlook_web_helper_host.json"
set "PY_DEST=%INSTALL_DIR%\outlook_web_helper_host.py"
set "WRAPPER=%INSTALL_DIR%\outlook_web_helper_host_runner.bat"
set "MANIFEST_DEST=%INSTALL_DIR%\outlook_web_helper_host.json"
set "REG_KEY=HKCU\SOFTWARE\Mozilla\NativeMessagingHosts\outlook_web_helper_host"

set "PYTHON_PRESENT=0"
python -c "import sys; sys.exit(0)" >nul 2>&1
if not errorlevel 1 set "PYTHON_PRESENT=1"

echo.
echo This will install the helper app for the Firefox add-on
echo "Outlook Web Helper".
echo.
echo The following will happen:
if "%PYTHON_PRESENT%"=="1" echo   - Python 3: already installed, no action needed
if "%PYTHON_PRESENT%"=="0" echo   - Python 3: not detected, please install it from the Microsoft Store
echo   - Copy the helper app into:
echo       %INSTALL_DIR%
echo   - Register the helper app under:
echo       %REG_KEY%
echo.
echo The helper app listens on 127.0.0.1 only. Requests need the secret token,
echo see the options page of the add-on.
echo.
choice /c yn /n /m "Proceed with installation? [y/n] "
if errorlevel 2 (
  echo.
  echo Installation cancelled.
  pause
  endlocal
  exit /b 1
)
echo.

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
copy /y "%PY_SRC%" "%PY_DEST%" >nul

:: Firefox starts executables, so wrap the Python script in a .bat file
(
  echo @echo off
  echo python "%PY_DEST%" %%*
) > "%WRAPPER%"

:: Write the manifest with the path to the wrapper (backslashes must be doubled in JSON)
powershell -NoProfile -Command "$p = '%WRAPPER%'.Replace('\','\\'); (Get-Content -Raw '%MANIFEST_SRC%') -replace '/path/to/native-messaging-app/outlook_web_helper_host.py', $p | Set-Content -Encoding ASCII '%MANIFEST_DEST%'"

reg add "%REG_KEY%" /ve /t REG_SZ /d "%MANIFEST_DEST%" /f >nul

echo Installed app to: %INSTALL_DIR%
echo Registered under: %REG_KEY%
echo.
echo Restart Firefox to apply the changes.
echo The downloaded files can now be safely removed.
echo.
pause
endlocal
