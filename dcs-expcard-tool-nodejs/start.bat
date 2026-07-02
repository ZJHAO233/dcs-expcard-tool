@echo off
title ExpCard Converter
color 0B

cls
echo.
echo    ===================================================
echo    EEEEEEEE XXXXXXXX PPPPPPPP  CCCCCCCC  AAA  RRRRRR  DDDDDDD
echo    EEE       XXX       PP    PP CC        AAAA  RR   RR DD    DD
echo    EEE       XXX       PP    PP CC       AA  AA RR  RR  DD    DD
echo    EEEEEEEE  XXXXXXXX  PPPPPPPP CC      AAAAAAAA RRRRRR  DD    DD
echo    EEE       XXX       PP       CC      AA    AA RR  RR  DD    DD
echo    EEE       XXX       PP       CC      AA    AA RR   RR DD    DD
echo    EEEEEEEE  XXXXXXXX  PP        CCCCCCCC AA    AA RR    RR DDDDDDD
echo    ===================================================
echo                      [ Excel to Markdown ]
echo.
echo.
echo    [1] Start Server
echo    [2] Edit Config
echo    [3] Help
echo    [0] Exit
echo.
echo    ---------------------------------------------------
echo.

set /p choice="    Select: "

if "%choice%"=="1" goto start
if "%choice%"=="2" goto config
if "%choice%"=="3" goto help
if "%choice%"=="0" goto exit

echo.
echo    [!] Invalid option
timeout /t 2 >nul
goto menu

:start
cls
echo.
echo    ===================================================
echo                      Starting Server...
echo    ===================================================
echo.
echo    Address : http://localhost:3210
echo    Stop    : Ctrl+C
echo.
echo    ---------------------------------------------------
echo.

if exist "expcard-converter.exe" (
    expcard-converter.exe
) else (
    node -v >nul 2>&1
    if errorlevel 1 (
        echo    [!] Node.js not found!
        echo        Download: https://nodejs.org/
        pause
        exit /b 1
    )
    node server.js
)
goto end

:config
cls
echo.
echo    ===================================================
echo                       Config File
echo    ===================================================
echo.
echo    File : config.js
echo.
echo    Edit:
echo      - LOGIC_OPERATORS
echo      - SPECIAL_SEPARATORS
echo      - SKIP_HEADERS
echo      - SECTION_HEADERS
echo.
echo    Note: Restart server after edit.
echo.
echo    ---------------------------------------------------
echo.
if exist "config.js" (
    echo    Opening config.js...
    notepad config.js
) else (
    echo    [!] config.js not found!
)
pause
goto menu

:help
cls
echo.
echo    ===================================================
echo                         Help
echo    ===================================================
echo.
echo    1. Double click to start
echo    2. Open http://localhost:3210
echo    3. Upload .xlsx file
echo    4. Select sheet and range
echo    5. Click Convert
echo    6. Preview and export
echo.
echo    Support:
echo      - Sub-number mode (x.y format)
echo      - Pure number mode (column nesting)
echo.
echo    ---------------------------------------------------
echo.
pause
goto menu

:exit
exit /b 0

:end
pause
