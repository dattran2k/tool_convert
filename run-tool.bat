@echo off
echo ================================================
echo    Android Development Excel Tools
echo ================================================
echo.
echo Available tools:
echo 1. Excel Tree Flattener
echo 2. Spec to Requirements Converter
echo 3. Open both tools
echo 4. Exit
echo.
set /p choice="Please choose (1-4): "

if "%choice%"=="1" (
    echo Opening Excel Tree Flattener...
    start "" "%~dp0excel-tree-flattener.html"
) else if "%choice%"=="2" (
    echo Opening Spec to Requirements Converter...
    start "" "%~dp0spec-to-requirements-converter.html"
) else if "%choice%"=="3" (
    echo Opening both tools...
    start "" "%~dp0excel-tree-flattener.html"
    timeout /t 2 /nobreak > nul
    start "" "%~dp0spec-to-requirements-converter.html"
) else if "%choice%"=="4" (
    echo Goodbye!
    exit
) else (
    echo Invalid choice. Please run the script again.
)

echo.
echo Tools opened in default browser.
echo Press any key to exit...
pause > nul