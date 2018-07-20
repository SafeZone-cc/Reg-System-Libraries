@echo off
SetLocal EnableExtensions EnableDelayedExpansion

set MAX_Processes_Count=30
set MAX_Timout=80
set timer=0

chcp 866 >NUL
cd /d "%~dp0"

title RegSystemLibraries by Dragokas ver.1.1

echo ��ਯ� ����⠭������� ॣ����樨 ��⥬��� ������⥪
echo.

call :CheckPrivilegies
color 1A

:settings

:: �஢�ઠ ���ᨨ ��
echo ����� ��:
ver

:: �஢�ઠ ࠧ�來��� ��
Set "xOS=AMD64"& If "%PROCESSOR_ARCHITECTURE%"=="x86" If Not Defined PROCESSOR_ARCHITEW6432 Set "xOS=x86"
echo [ %xOS% ]

echo.
echo �� ����� ������ �� 10 ����� � �����.
echo ������ ����� ���� ����㦥� �� 100 %%.
echo ��� �᪮७�� ����樨 �६���� �⪫��� ��⨢���᭮� ��.
echo.

echo -------------------------------------------------
echo ����騥 ����ன��:
echo.
echo - �����६���� ࠡ�⠥� ����ᮢ: %MAX_Processes_Count%
echo - ������� �����襭�� �����:     %MAX_Timout% ᥪ.
echo -------------------------------------------------
echo.

echo.
echo ������ ���� ������� �⮡� �ਭ��� ����ன�� � ����� ॣ������,
echo ��� N �⮡� �������� ����ன��.
echo.

set ch=
set /p ch=
if /i "%ch%" neq "N" if /i "%ch%" neq "�" goto begin
echo.
echo ������ ��࠭�祭�� ���-�� �����६���� ࠡ����� ����ᮢ
set /p MAX_Processes_Count=" (��� ᫠��� ��設 㪠�뢠�� ����襥 �᫮): "
echo.
echo ������ ⠩���� ��� �����襭�� �����
set /p MAX_Timout=" (��� ᫠��� ��設 㪠�뢠�� ����襥 �᫮): "
cls
goto settings

:begin
cd /d "%~dp0"

echo ��������, �������� ...
echo.

taskkill /F /T /IM regsvr32.exe >NUL 2>NUL

:: ��� �ண��ᡠ�
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%A in (1) do rem"') do set "DEL=%%a"

:: ��������� ������⥪, ����室���� ��� ���४⭮� ࠡ��� �ਯ�

echo ���� ॣ������ �᭮���� ������⥪ � WMI ...

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
  :: �஢�ઠ ������ �ࠢ �����������
  >NUL 2>NUL reg add HKLM\Software\Test /v Test /F && (
  >NUL 2>NUL reg delete HKLM\Software\Test /v Test /F) || (
  echo ������� �ਯ� �ࠢ�� ������� ��� "�� ����� �����������"
  echo.
  pause
  exit
  )
Exit /B


:RegLib [_in_����� ���� � ������⥪�. ����᪠���� ��᪠.]
  echo ������ ������⥪ %~x1 - %~dp1
  
  SetLocal
  set PercentView=0
  set libs=0
  :: �����뢠� �� ����� 䠩��� � ���ᨢ
  For /F "delims=" %%F in ('dir /B /A-D-L "%~1"') do set /a libs+=1& set lib.!libs!=%~dp1%%F
  echo [ %libs% ]

  if %libs%==0 goto :RegLib_exit

  set LastProcess=%MAX_Processes_Count%
  :: �᫨ ���-�� ������⥪ ����� �᫠ ��࠭�祭��
  if %MAX_Processes_Count% GTR %libs% set LastProcess=%libs%

  set /p "x=�ண���...   0%%"<NUL

  :: ����� ��ࢮ� ���⨨ ����ᮢ ॣ����樨
  call :RunProcesses 1 %LastProcess%

  :RegLib_loop

  :: ������ �� ���-���
  call :WaitForReady "regsvr32.exe" TotalP

  :: ����� �������� ����ᮢ
  set /a AvailP=%MAX_Processes_Count% - %TotalP%

  :: ����� ��ࢮ�� �����, �� ���ண� ᥩ�� ����� ����᪠��
  set /a FirstProcess=%LastProcess% + 1

  :: ����� ��᫥����� �����, �� ���ண� ᥩ�� ����� ����᪠��
  set /a LastProcess+=%AvailP%

  :: �᫨ # ��᫥����� ����᪠����� ����� ��饣� ���-�� ������⥪ - �ᥪ���
  if %LastProcess% GTR %libs% set LastProcess=%libs%

  ::echo TotalP=%TotalP%, FirstProcess=%FirstProcess%, LastProcess=%LastProcess%

  :: ������ # >= ��砫쭮�� # - ����᪠�� ������ �� 楯�窥
  if %LastProcess% GEQ %FirstProcess% call :RunProcesses %FirstProcess% %LastProcess%

  :: �᫨ ��ࠡ�⢠��� ������ �����稫���, � ����室���� ��� ��ࠡ�⪨ ⠪�� ��� - ��室��
  if %TotalP%==0 if %FirstProcess% GTR %LastProcess% goto RegLib_exit

  goto RegLib_loop

  :RegLib_exit

  EndLocal &:: ���⪠ ���ᨢ�
  echo %DEL%%DEL%%DEL%%DEL%100 %%
Exit /B


:RunProcesses [# ��ࢮ�� �����] [# ��᫥����� �����]
  :: �㭪�� ����᪠�� ��᪮�쪮 ����ᮢ � 䮭�.
  :: �ॡ�� ��������� ��॥�����:
  :: PercentView - ������ ���� ���樠����஢�� �᫮� 0.
  :: libs - ��饥 ���-�� ������⥪. �� ������ ࠢ������ 0.
  For /L %%C in (%~1 1 %~2) do start "Stream %%C" /min "%SystemRoot%\System32\regsvr32.exe" !lib.%%C! /s
  set /a NextPercent=%PercentView% / 10 * 10 + 10
  set /a CurPercent=%~1 * 100 / %libs%
  ::echo CurPercent=%CurPercent%, NextPercent=%NextPercent%
  if %CurPercent% GEQ %NextPercent% (set PercentView=%NextPercent%& set /p "x=%DEL%%DEL%%DEL%%NextPercent%%%"<NUL)
  ::echo %~1 - %~2
Exit /B


:WaitForReady [_in_��� �����] [_out_���� ��� �࠭���� ���-�� ����饭��� ����ᮢ]
  :: ������� �����襭�� 㪠������� �᫠ ����ᮢ, ��᫥ 祣� �����頥� �ࠢ�����
  :: �᫨ ����� �� �� �����襭 �� 㪠������ �᫮ ᥪ㭤, �ਭ㤨⥫쭮 �����蠥� ���
  ::
  :: �������⥫쭮 �ॡ�� ��������� ��६�����:
  :: CurPID - PID ⥪�饣� �����
  :: MAX_Timout - ���ᨬ���� ⠩���� ��� �����, ����� �����
  :: call :Wait
  Set /A PIDs=0
  :: ������� �� ������. �஢����, 㪠��� �� ��� ������� �� PID ᢮� timer.
  :: �᫨ �� 㪠���, ������ � ���ᨢ TimeOut.PID = ⠩���.
  :: �᫨ 㪠���, ᢥ��� ࠧ���� ����� ��� � ⠩��஬. �᫨ �ॢ�蠥� �����⨬� - 㡨��� ����� � ��᢮������ ����� ���ᨢ�.
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
  set /a timer+=1&:: �ᥢ��⠩���
Exit /B