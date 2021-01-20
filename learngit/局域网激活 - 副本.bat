@echo off
TITLE Microsfot Windows Office Active V1.0
cls
color a
REM 获得管理员权限
openfiles >nul 2>nul ||(
  echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
  echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
  "%temp%\getadmin.vbs" >nul 2>&1
  goto:eof
)
del /f /q "%temp%\getadmin.vbs" >nul 2>nul
pushd "%~dp0"

REM 设置KMS服务器地址,请自行更换。
set kmserver=kms.03k.org
set kmserver=kms.03k.org
set kmserver=kms.mrxn.net

echo =========================================================================================================
echo Author:ITS
echo Create:2019/6/14
echo UpdataDate:2019/6/17 
echo Version:V1.0
echo note:增加人机对话界面，用户可以自行选择
echo Function:
echo        本版本支持的服务器操作系统：Windows Server2008/2008R2/2012/2012R2/2016/2019,Standard,Enterprise
echo        本版本支持的桌面操作系统：7/8/8.1/Win10/LTSC/LSTB,Professional,Enterprise
echo        本版本支持的office：2010/2013/2016/2019 professional
echo        只支持VL版本，不支持Retail版本
echo ===========================================================================================================

rem KMS服务器可用性检查
:ping
set /a cont0+=1
@echo.
echo  正在检查KMS服务器是否可达......
echo  正在进行第%cont0%次尝试......
ping /n 1 /w 2000 %kmserver% >nul && (goto ok) || (goto tem1)

:tem1
if %cont0% GTR 2 goto error
goto ping

:ok
@echo.
@echo.服务器连接正常！请选择需要激活操作
@echo.

:select
echo.1. 激活 Windows ALL
echo.2. 激活 office2019(包括visio与project)
echo.3. 激活 office2016(包括visio与project)
echo.4. 激活 office2013(包括visio与project)
echo.5. 激活 office2010(包括visio与project)
echo.6. Office2019 Retail Convert VOLUME,并激活
echo 0. 退出
echo;
set /p flag="请输入(数字):"
if %flag%==1 goto windows
if %flag%==2 goto active_office_2019
if %flag%==3 goto active_office_2016
if %flag%==4 goto active_office_2013
if %flag%==5 goto active_office_2010
if %flag%==6 goto Office_Retail2VOL
if %flag%==0 exit

:windows
rem 获取操作系统版本....
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v "EditionSubstring" /t REG_SZ /D "" /f
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v "RegisteredOrganization" /t REG_SZ /D "" /f
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v "RegisteredOwner" /t REG_SZ /D "" /f
setlocal enabledelayedexpansion
for /f "delims=" %%i in ('wmic os get Caption /value') do set "%%~i" >nul 2>nul

@echo 检测到系统版本:%Caption%
if "%Caption%" == "Microsoft Windows 10 Enterprise LTSC 2019" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk M7XTQ-FN8P6-TTKYV-9D4CC-J462D & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 10 企业版 LTSC" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk M7XTQ-FN8P6-TTKYV-9D4CC-J462D & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 10 Enterprise LTSB 2016" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 10 企业版 LTSB" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 10 Professional" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 10 Enterprise" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43 & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 10 专业版" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 10 企业版" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43 & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 10 企业版 G" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk YYVX9-NTFWV-6MDM3-9PT4T-4M68B & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 8 Professional" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk NG4HW-VH26C-733KW-K6F98-J8CK4 & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 8 Enterprise" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk 32JNW-9KQ84-P47T8-D8GGY-CWCK7 & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 8 专业版" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk NG4HW-VH26C-733KW-K6F98-J8CK4 & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 8 企业版" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk 32JNW-9KQ84-P47T8-D8GGY-CWCK7 & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 8.1 Professional" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk GCRJD-8NW9H-F2CDX-CCM8D-9D6T9 & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 8.1 Enterprise" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk MHF9N-XY6XB-WVXMC-BTDCT-MKKG7 & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 8.1 专业版" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk GCRJD-8NW9H-F2CDX-CCM8D-9D6T9 & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 8.1 企业版" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk MHF9N-XY6XB-WVXMC-BTDCT-MKKG7 & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 7 Professional" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4 & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 7 Enterprise" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk 33PXH-7Y6KF-2VJC9-XBBR8-HVTHH & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 7 专业版" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4 & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows 7 企业版" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk 33PXH-7Y6KF-2VJC9-XBBR8-HVTHH & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server 2012 Datacenter" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk 48HP8-DN98B-MYWDG-T2DCC-8W83P & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server 2012 Standard" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk XC9B7-NBPP2-83J2H-RHMBY-92BT4 & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server 2012 R2 Datacenter" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9 & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server 2012 R2 Standard" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk D2N9P-3P6X9-2R39C-7RTCD-MDVJX & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server 2019 Datacenter" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk WMDGN-G9PQG-XVVXX-R3X43-63DFG & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server 2019 Standard" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk N69G4-B89J2-4G8F4-WWYCC-J464C & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server 2016 Datacenter" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk CB7KF-BWN84-R7R2Y-793K2-8XDDG & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server 2016 Standard" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server Datacenter" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk 6NMRW-2C8FM-D24W7-TQWMY-CWH2D & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server Standard" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk N2KJX-J94YW-TQVFB-DG9YT-724CC & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server 2008 Standard" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk YC6KT-GKW9T-YTKYR-T4X34-R7VHC & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server 2008 Enterprise" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk YQGMW-MPWTJ-34KDK-48M3W-X4Q6V & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server 2008 Datacenter" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk 7M67G-PC374-GR742-YH8V4-TCBY3 & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server 2008 R2 Standard" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk YC6KT-GKW9T-YTKYR-T4X34-R7VHC & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server 2008 R2 Enterprise" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk 489J6-VHDMP-X63PK-3K798-CPX3Y & goto active_windows & set w=1) else (set w=0)
if "%Caption%" == "Microsoft Windows Server 2008 R2 Datacenter" (cscript //Nologo %windir%\system32\slmgr.vbs /ipk 74YFP-3QFB3-KQT8W-PMXWJ-7M648 & goto active_windows & set w=1) else (set w=0)
if %w% EQU 0 (echo "没找到匹配的系统版本")
pause.
exit

:active_windows
cscript //Nologo %windir%\system32\slmgr.vbs /skms %kmserver%
cscript //Nologo %windir%\system32\slmgr.vbs /skhc
cscript //Nologo %windir%\system32\slmgr.vbs /ato  && echo windows激活成功，请查看详情  ||echo windows激活失败
cscript //Nologo %windir%\system32\slmgr.vbs /dlv 
pause.
exit

:active_office_2019
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
cscript //nologo ospp.vbs /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
cscript //nologo ospp.vbs /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB
cscript //nologo ospp.vbs /sethst:%kmserver%
cscript //nologo ospp.vbs /act && echo office激活成功 ||echo Office激活失败你可能的安装的不是VOLUME版本,请退出程序选择第6项
@echo office2019激活可能需要重启计算机才生效，请按任意键重启计算机
pause
shutdown -r -t 0
exit

:active_office_2016
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
cscript //nologo ospp.vbs /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT
cscript //nologo ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK
cscript //nologo ospp.vbs /sethst:%kmserver%
cscript //nologo ospp.vbs /act && echo office激活成功 ||echo Office激活失败你可能的安装的不是VOLUME版本
pause
exit

:active_office_2013
if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office15"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office15"
cscript //nologo ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT
cscript //nologo ospp.vbs /inpkey:FN8TT-7WMH6-2D4X9-M337T-2342K
cscript //nologo ospp.vbs /inpkey:C2FG9-N6J68-H8BTJ-BW3QX-RM3B3
cscript //nologo ospp.vbs /sethst:%kmserver%
cscript //nologo ospp.vbs /act && echo office激活成功 ||echo Office激活失败你可能的安装的不是VOLUME版本
pause
exit

:active_office_2010
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office14"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office14"
cscript //nologo ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
cscript //nologo ospp.vbs /inpkey:YGX6F-PGV49-PGW3J-9BTGG-VHKC6
cscript //nologo ospp.vbs /inpkey:7MCW8-VRQVK-G677T-PDJCM-Q8TCP
cscript //nologo ospp.vbs /sethst:%kmserver%
cscript //nologo ospp.vbs /act && echo office激活成功 ||echo Office激活失败你可能的安装的不是VOLUME版本
pause
exit

REM 将零售版 Office 转换为 VOL 版  
:Office_Retail2VOL
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
cls
echo 正在重置Office2019零售激活...
cscript ospp.vbs /rearm
echo 正在安装 KMS 许可证...
for /f %%u in ('dir /b ..\root\Licenses16\ProPlusVL_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%u" >nul
for /f %%v in ('dir /b ..\root\Licenses16\ProPlus2019VL_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%v" >nul
for /f %%w in ('dir /b ..\root\Licenses16\ProjectProVL_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%w" >nul
for /f %%x in ('dir /b ..\root\Licenses16\ProjectPro2019VL_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%y in ('dir /b ..\root\Licenses16\VisioProVL_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%y" >nul
for /f %%z in ('dir /b ..\root\Licenses16\VisioPro2019VL_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%z" >nul
echo 正在安装 MAK 许可证...
for /f %%u in ('dir /b ..\root\Licenses16\ProPlusVL_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%u" >nul
for /f %%v in ('dir /b ..\root\Licenses16\ProPlus2019VL_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%v" >nul
for /f %%w in ('dir /b ..\root\Licenses16\ProjectProVL_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%w" >nul
for /f %%x in ('dir /b ..\root\Licenses16\ProjectPro2019VL_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%y in ('dir /b ..\root\Licenses16\VisioProVL_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%y" >nul
for /f %%z in ('dir /b ..\root\Licenses16\VisioPro2019VL_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%z" >nul
echo 转换完成，开始激活
goto active_office_2019

:error
@echo.
@echo.=========================================================
echo 无法连接KMS服务器，请检查网络状态或者检查KMS服务状态
echo 或者更换kmServer服务器地址
@echo.=========================================================
echo 是否重试连接？输入[y]重试，输入[n]以退出：
set /p choi="请选择："
set cont0=0
if %choi% ==y goto ping
pause
@echo.
exit