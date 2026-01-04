@echo off
chcp 1251 >nul 2>&1
setlocal enabledelayedexpansion

set "WINDOWS_PRO_KEY=W269N-WFGWX-YVC9B-4J6C9-T83GX"
set "WINDOWS_HOME_KEY=TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"
set "WINDOWS_ENTERPRISE_KEY=NPPR9-FWDCX-D2C8J-H872K-2YT43"
set "OFFICE_KEY=YVY2H-8F9FJ-YGQFQ-9M9YX-9GVY3"  :: Office 2021 ProPlus (универсальный KMS-ключ)

title Активация Windows и Office
color 07

:main_menu
cls
echo ===============================================
echo    Активация Windows и Microsoft Office
echo ===============================================
echo.
echo Выберите действие:
echo.
echo   1. Активировать Windows 10/11
echo   2. Активировать Microsoft Office
echo   3. Выход
echo.
set /p choice="Введите номер (1, 2 или 3): "

if "%choice%"=="1" goto activate_windows
if "%choice%"=="2" goto activate_office
if "%choice%"=="3" exit /b 0
goto main_menu

:: ==============================
:activate_windows
cls
echo.
echo Вы выбрали: Активация Windows 10/11
echo.

:: Проверка прав администратора
net session >nul 2>&1
if not %errorlevel% equ 0 (
    echo Ошибка: Для активации Windows нужны права администратора!
    echo Запустите скрипт "От имени администратора".
    echo.
    pause >nul
    goto main_menu
)

echo Шаг 1: Установка ключа Windows Pro...
cscript //nologo %windir%\system32\slmgr.vbs /ipk %WINDOWS_PRO_KEY%
if not %errorlevel% equ 0 (
    echo Ошибка: Не удалось установить ключ Windows.
    pause >nul
    goto main_menu
)

echo Шаг 2: Настройка KMS-сервера...
cscript //nologo %windir%\system32\slmgr.vbs /skms kms8.msguides.com >nul 2>&1
if not %errorlevel% equ 0 (
    echo Предупреждение: Не удалось подключиться к kms8.msguides.com. Пробуем резервный...
    cscript //nologo %windir%\system32\slmgr.vbs /skms kms.core.bz >nul 2>&1
)

echo Шаг 3: Активация Windows...
cscript //nologo %windir%\system32\slmgr.vbs /ato

echo.
echo Шаг 4: Проверка статуса активации:
cscript //nologo %windir%\system32\slmgr.vbs /xpr

echo.
echo Активация Windows завершена.
pause >nul
goto main_menu

:: ==============================
:activate_office
cls
echo.
echo Вы выбрали: Активация Microsoft Office
echo.

net session >nul 2>&1
if not %errorlevel% equ 0 (
    echo Ошибка: Для активации Office нужны права администратора!
    echo Запустите скрипт "От имени администратора".
    echo.
    pause >nul
    goto main_menu
)

set "office_path="
for /f "tokens=*" %%i in ('dir /b /s "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" 2^>nul') do set "office_path=%%i"
if not defined office_path (
    for /f "tokens=*" %%i in ('dir /b /s "%ProgramFiles%\Microsoft Office\root\Office16\ospp.vbs" 2^>nul') do set "office_path=%%i"
)
if not defined office_path (
    for /f "tokens=*" %%i in ('dir /b /s "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" 2^>nul') do set "office_path=%%i"
)
if not defined office_path (
    for /f "tokens=*" %%i in ('dir /b /s "%ProgramFiles(x86)%\Microsoft Office\root\Office16\ospp.vbs" 2^>nul') do set "office_path=%%i"
)

if not defined office_path (
    echo Ошибка: Не найден Microsoft Office (2016/2019/2021/365).
    echo Убедитесь, что Office установлен.
    echo.
    pause >nul
    goto main_menu
)

echo Найден Office: "!office_path!"
echo.

echo Шаг 1: Установка KMS-ключа для Office...
cscript //nologo "!office_path!" /inpkey:%OFFICE_KEY%
if not %errorlevel% equ 0 (
    echo Ошибка: Не удалось установить ключ Office.
    pause >nul
    goto main_menu
)

echo Шаг 2: Настройка KMS-сервера...
cscript //nologo "!office_path!" /sethst:kms8.msguides.com >nul 2>&1
if not %errorlevel% equ 0 (
    echo Предупреждение: Не удалось подключиться к kms8.msguides.com. Пробуем резервный...
    cscript //nologo "!office_path!" /sethst:kms.core.bz >nul 2>&1
)

echo Шаг 3: Активация Office...
cscript //nologo "!office_path!" /act

echo.
echo Шаг 4: Проверка статуса активации Office...
cscript //nologo "!office_path!" /dstatus

echo.
echo Активация Office завершена.
pause >nul
goto main_menu