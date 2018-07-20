@echo off
SetLocal EnableExtensions EnableDelayedExpansion

set MAX_Processes_Count=30
set MAX_Timout=80
set timer=0

chcp 866 >NUL
cd /d "%~dp0"

title RegSystemLibraries by Dragokas ver.1.1

echo Скрипт восстановления регистрации системных библиотек
echo.

call :CheckPrivilegies
color 1A

:settings

:: Проверка версии ОС
echo Версия ОС:
ver

:: Проверка разрядности ОС
Set "xOS=AMD64"& If "%PROCESSOR_ARCHITECTURE%"=="x86" If Not Defined PROCESSOR_ARCHITEW6432 Set "xOS=x86"
echo [ %xOS% ]

echo.
echo Это может занять до 10 минут и более.
echo Процессор может быть нагружен до 100 %%.
echo Для ускорения операции временно отключите антивирусное ПО.
echo.

echo -------------------------------------------------
echo Текущие настройки:
echo.
echo - Одновременно работает процессов: %MAX_Processes_Count%
echo - Таймаут завершения процесса:     %MAX_Timout% сек.
echo -------------------------------------------------
echo.

echo.
echo Нажмите любую клавишу чтобы принять настройки и начать регистрацию,
echo или N чтобы изменить настройки.
echo.

set ch=
set /p ch=
if /i "%ch%" neq "N" if /i "%ch%" neq "т" goto begin
echo.
echo Введите ограничение кол-ва одновременно работающих процессов
set /p MAX_Processes_Count=" (для слабых машин указывайте меньшее число): "
echo.
echo Введите таймаут для завершения процесса
set /p MAX_Timout=" (для слабых машин указывайте большее число): "
cls
goto settings

:begin
cd /d "%~dp0"

echo Пожалуйста, подождите ...
echo.

taskkill /F /T /IM regsvr32.exe >NUL 2>NUL

:: для прогрессбара
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%A in (1) do rem"') do set "DEL=%%a"

:: Регистрация библиотек, необходимых для корректной работы скрипта

echo Начата регистрация основных библиотек и WMI ...

start "" /min "%SystemRoot%\System32\regsvr32.exe" %SystemRoot%\System32\advapi32.dll /s
start "" /min "%SystemRoot%\System32\regsvr32.exe" %SystemRoot%\System32\kernel32.dll /s
start "" /min "%SystemRoot%\System32\regsvr32.exe" %SystemRoot%\System32\msvcrt.dll /s
start "" /min "%SystemRoot%\System32\regsvr32.exe" %SystemRoot%\System32\user32.dll /s
start "" /min "%SystemRoot%\System32\regsvr32.exe" %SystemRoot%\System32\ntdll.dll /s
start "" /min "%SystemRoot%\System32\regsvr32.exe" %SystemRoot%\System32\ole32.dll /s
start "" /min "%SystemRoot%\System32\regsvr32.exe" %SystemRoot%\System32\version.dll /s
start "" /min "%SystemRoot%\System32\regsvr32.exe" %SystemRoot%\System32\mpr.dll /s
start "" /min "%SystemRoot%\System32\regsvr32.exe" %SystemRoot%\System32\oleaut32.dll /s
start "" /min "%SystemRoot%\System32\regsvr32.exe" %SystemRoot%\System32\secur32.dll /s
start "" /min "%SystemRoot%\System32\regsvr32.exe" %SystemRoot%\System32\ws2_32.dll /s
start "" /min "%SystemRoot%\System32\regsvr32.exe" %SystemRoot%\System32\framedynos.dll /s
start "" /min "%SystemRoot%\System32\regsvr32.exe" %SystemRoot%\System32\netapi32.dll /s
start "" /min "%SystemRoot%\System32\regsvr32.exe" %SystemRoot%\System32\dbghelp.dll /s
start "" /min "%SystemRoot%\System32\regsvr32.exe" %SystemRoot%\System32\shlwapi.dll /s

echo .dll - %SystemRoot%\System32\wbem\
echo.

For /F %%F in ('dir /B /A-D "%SystemRoot%\System32\wbem\*.dll"') do (
  start "" /min "%SystemRoot%\System32\regsvr32.exe" "%SystemRoot%\System32\wbem\%%F" /s
)

call :Wait
call :Wait
call :Wait
call :Wait
call :Wait

echo OK.

call :RegLib "%SystemRoot%\System32\*.dll"
call :RegLib "%SystemRoot%\System32\*.ocx"
call :RegLib "%SystemRoot%\System32\*.tlb"

if %xOS%==AMD64 (
call :RegLib "%SystemRoot%\SysWOW64\wbem\*.dll"
call :RegLib "%SystemRoot%\SysWOW64\*.dll"
call :RegLib "%SystemRoot%\SysWOW64\*.ocx"
call :RegLib "%SystemRoot%\SysWOW64\*.tlb"
)

echo.
pause
exit /B




:CheckPrivilegies
  :: проверка наличия прав администратора
  >NUL 2>NUL reg add HKLM\Software\Test /v Test /F && (
  >NUL 2>NUL reg delete HKLM\Software\Test /v Test /F) || (
  echo Запустите скрипт правой кнопкой мыши "От имени администратора"
  echo.
  pause
  exit
  )
Exit /B


:RegLib [_in_Полный путь к библиотеке. Допускается маска.]
  echo Подсчет библиотек %~x1 - %~dp1
  
  SetLocal
  set PercentView=0
  set libs=0
  :: записываю все имена файлов в массив
  For /F "delims=" %%F in ('dir /B /A-D-L "%~1"') do set /a libs+=1& set lib.!libs!=%~dp1%%F
  echo [ %libs% ]

  if %libs%==0 goto :RegLib_exit

  set LastProcess=%MAX_Processes_Count%
  :: если кол-во библиотек меньше числа ограничения
  if %MAX_Processes_Count% GTR %libs% set LastProcess=%libs%

  set /p "x=Прогресс...   0%%"<NUL

  :: Запуск первой партии процессов регистрации
  call :RunProcesses 1 %LastProcess%

  :RegLib_loop

  :: Следим за кол-вом
  call :WaitForReady "regsvr32.exe" TotalP

  :: Можно запустить процессов
  set /a AvailP=%MAX_Processes_Count% - %TotalP%

  :: Номер первого процесса, от которого сейчас можно запускать
  set /a FirstProcess=%LastProcess% + 1

  :: Номер последнего процесса, до которого сейчас можно запускать
  set /a LastProcess+=%AvailP%

  :: Если # последнего запускаемого больше общего кол-ва библиотек - усекаем
  if %LastProcess% GTR %libs% set LastProcess=%libs%

  ::echo TotalP=%TotalP%, FirstProcess=%FirstProcess%, LastProcess=%LastProcess%

  :: Конечный # >= Начального # - запускаем процессы по цепочке
  if %LastProcess% GEQ %FirstProcess% call :RunProcesses %FirstProcess% %LastProcess%

  :: Если обрабатваемые процессы закончились, а необходимых для обработки также нет - выходим
  if %TotalP%==0 if %FirstProcess% GTR %LastProcess% goto RegLib_exit

  goto RegLib_loop

  :RegLib_exit

  EndLocal &:: очистка массива
  echo %DEL%%DEL%%DEL%%DEL%100 %%
Exit /B


:RunProcesses [# первого процесса] [# последнего процесса]
  :: Функция запускает несколько процессов в фоне.
  :: Требует глобальных переемнных:
  :: PercentView - должен быть инициализирован числом 0.
  :: libs - общее кол-во библиотек. Не должно равняться 0.
  For /L %%C in (%~1 1 %~2) do start "Stream %%C" /min "%SystemRoot%\System32\regsvr32.exe" !lib.%%C! /s
  set /a NextPercent=%PercentView% / 10 * 10 + 10
  set /a CurPercent=%~1 * 100 / %libs%
  ::echo CurPercent=%CurPercent%, NextPercent=%NextPercent%
  if %CurPercent% GEQ %NextPercent% (set PercentView=%NextPercent%& set /p "x=%DEL%%DEL%%DEL%%NextPercent%%%"<NUL)
  ::echo %~1 - %~2
Exit /B


:WaitForReady [_in_Имя процесса] [_out_Буфер для хранения кол-ва запущенных процессов]
  :: Ожидает завершения указанного числа процессов, после чего возвращает управление
  :: Если процесс не был завершен за указанное число секунд, принудительно завершает его
  ::
  :: Дополнительно требует глобальных переменных:
  :: CurPID - PID текущего процесса
  :: MAX_Timout - максимальный таймаут для процесса, который завис
  :: call :Wait
  Set /A PIDs=0
  :: Перечисляю все процессы. Проверяю, указан ли для каждого из PID свой timer.
  :: Если не указан, заношу в массив TimeOut.PID = таймер.
  :: Если указан, сверяю разницу между ним и таймером. Если превышает допустимый - убиваю процесс и высвобождаю элемент массива.
  For /F "tokens=1,2 delims=," %%A in ('tasklist /FI "imagename EQ %~1" /FO CSV /NH 2^>NUL ^| find /i "%~1"') do (
    rem echo %%A - %%B
    if not Defined TimeOut.%%~B (
      set TimeOut.%%~B=%timer%
      set /A PIDs+=1
    ) else (
      set /a diff=%timer% - !TimeOut.%%~B!
      if !diff! GEQ %MAX_Timout% (taskkill /F /PID %%~B >NUL 2>NUL& set TimeOut.%%~B=) else (set /A PIDs+=1)
    )
  )
  set %~2=%PIDs%
  call :Wait
Exit /B

:Wait { void }
  ping -n 2 127.1 >NUL
  set /a timer+=1&:: Псевдотаймер
Exit /B