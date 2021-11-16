<!-- : Begin batch script
@echo off
color 07
setlocal EnableDelayedExpansion
:: Fix for special characters in file name by @abbodi1406
setlocal DisableDelayedExpansion
set "work=%~dp0"
if [%work:~-1%]==[\] set "work=%work:~0,-1%"
setlocal EnableDelayedExpansion
pushd "!work!"
%systemroot%\system32\fsutil.exe dirty query %systemdrive% >nul 2>&1 || (
    %systemroot%\System32\cscript.exe //nologo "%~nx0?.wsf" //job:ELAV /File:"!work!\%~nx0"
    exit /b
)
set "_path=!work!"
set "uiver=1.0"
set "batf=!work!\%~nx0"
title EzHWID %uiver% by Exe Csrss
mode con cols=83 lines=32
set _args=%*
if not defined _args goto :noargs
set _args=%_args:"=%
set "hwid="
set "k38="
set "silent="
set "pkey="
set "force="
for %%A in (%_args%) do (
if /i "%%A"=="/dl" (
    set dl=1
) else if /i "%%A"=="/k38" (
    set k38=1
) else if /i "%%A"=="/silent" (
    set silent=1
) else if /i "%%A"=="/force" (
    set force=1
) else if /i "%%A"=="/pkey" (
    set pkey=1
) else if /i "%%A"=="-dl" (
    set dl=1
) else if /i "%%A"=="-k38" (
    set k38=1
) else if /i "%%A"=="-silent" (
    set silent=1
) else if /i "%%A"=="-force" (
    set force=1
) else if /i "%%A"=="-pkey" (
    set pkey=1
) else (
    if defined pkey (
        set "_customkey=%%A"
    )
)
)
:noargs
for %%# in (powershell.exe) do if [%%~$PATH:#]==[] if not exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" (
    call :colortxt 0c ERROR: /n
    echo PowerShell is not installed in your system.
    echo Install PowerShell and run this script again.
    echo Would you like to install PowerShell?
    echo Choose an option:
    echo [Y]es
    echo [N]o
    choice /c yn /n /m "Enter your choice: "
    set _errtemp=!errorlevel!
    if /i [!_errtemp!]==[1] (
        start "" https://www.catalog.update.microsoft.com/Search.aspx?q=KB968930
    ) else (
        echo Press any key to exit...
        pause >nul
    )
    exit /b
)
for %%# in (powershell.exe) do if not [%%~$PATH:#]==[] (
   set "_psc="%%~$PATH:#" -nop -c"
   set "_ps=%%~$PATH:#"
)
if not defined _psc (
   set "_psc="%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -nop -c"
   set "_ps=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
)
:: Set buffer height
%_psc% "&{$H=get-host;$W=$H.ui.rawui;$B=$W.buffersize;$B.height=250;$W.buffersize=$B;}"
:: Declare some variables for convenience
set "red=0c"
set "green=0a"
set "_csc=%systemroot%\System32\cscript.exe //nologo"
set "_slm=%_csc% %systemroot%\System32\slmgr.vbs"
set "_wmi=%_csc% //job:WMI"
set "_wmim=%_csc% //job:WmiMethod"
set "_wact=%_csc% //job:ATO"
set "_ipk=%_csc% //job:IPK"
set "_refrsh=%_csc% //job:refresh"
set "_reapp=%_csc% //job:reapp"
set "_xpr=%_csc% //job:XPR"
set "_nul=1>nul 2>nul"
set "_eline=call :colortxt %red% ERROR: /n"
set "_pak=echo Press any key to exit... & pause %_nul% & goto exit"
set "_slp=SoftwareLicensingProduct"
set "_sls=SoftwareLicensingService"
set "_wApp=55c92734-d682-4d71-983e-d6ec3f16059f"
set "_tmp=%systemroot%\Temp"
if defined silent set "_pak=goto exit"
if not exist "!work!\bin" (
    %_eline%
    echo 'bin' folder is missing.
    echo Most likely you ran this script directly from your archiving software instead of extracting.
    %_pak%
)
if defined dl (
    call :dl 
    goto exit
)
if defined k38 (
    call :k38 
    goto exit
)
:mainmenu
cls
if not defined dl if not defined k38 (
echo:
echo:
echo:
echo EzHWID %uiver% by Exe Csrss
echo          _______________________________________________________________
echo         ^| Activation:                                                   ^|
echo         ^|                      [D] Digital License                      ^|
echo         ^|                                                               ^|
echo         ^|                           [K] KMS38                           ^|
echo         ^|_______________________________________________________________^|
echo         ^| Help and support:                                             ^|
echo         ^|                          [R] Read-me                          ^|
echo         ^|                                                               ^|
echo         ^|                       [H] Discord Server                      ^|
echo         ^|                                                               ^|
echo         ^|                        [G] Github repo                        ^|
echo         ^|_______________________________________________________________^|
echo         ^| Configuration:                                                ^|
if defined force (
call :colortxt 07 "        |"
call :colortxt 0e "                    [F] Forceful mode: [On]                    "
call :colortxt 07 "|" /n
) else (
echo         ^|                    [F] Forceful mode: [Off]                   ^|
)
echo         ^|_______________________________________________________________^|
echo         ^| Miscellaneous:                                                ^|
echo         ^|                 [X] Display Activation Status                 ^|
echo         ^|                                                               ^|
echo         ^|                       [S] Switch Edition                      ^|
echo         ^|                                                               ^|
echo         ^|                    [C] Create $OEM$ Folder                    ^|
echo         ^|_______________________________________________________________^|
echo:
set "_errtemp="
choice /c dkrhgfxsce /n /m "Enter your choice, or press E to exit: "
set _errtmp=!errorlevel!
if /i [!_errtmp!]==[1] call :dl
if /i [!_errtmp!]==[2] call :k38
if /i [!_errtmp!]==[3] start "%systemroot%\notepad.exe" "!work!\bin\readme.txt"
if /i [!_errtmp!]==[4] start "" https://discord.gg/M2Pu6cTee2
if /i [!_errtmp!]==[5] start "" https://www.github.com/ExeCsrss/EzHWID
if /i [!_errtmp!]==[6] call :toggleforce
if /i [!_errtmp!]==[7] call :dispstat2usr
if /i [!_errtmp!]==[8] call :switch
if /i [!_errtmp!]==[9] call :createoem
if /i [!_errtmp!]==[10] goto exit
) else exit /b
goto mainmenu
:toggleforce
if defined force (
set "force="
) else (
set "force=1"
)
exit /b

:dl
cls
set "_missingfiles="
for %%# in (8664_gatherosstate.exe 8664_slc.dll buildkeys.ini keys.ini sku.ini) do (if not exist "!work!\bin\%%#" (if defined _missingfiles (set "_missingfiles=!_missingfiles!, %%#") else (set "_missingfiles=%%#")))
if defined _missingfiles (
    %_eline%
    echo The following required files are missing from 'bin' folder:
    echo %_missingfiles%
    %_pak%
)
call :dispstat
if /i [%LicenseStatus%]==[1] if /i [%GracePeriodRemaining%]==[0] if not defined force (
    call :colortxt %green% "Windows is already permanently activated." /n
    echo Running this script is unnecessary.
    if defined silent goto exit
    set "_errtemp="
    echo Choose an option:
    echo [C]ontinue
    echo [E]xit
    choice /c ce /n /m "Enter your choice: "
    set _errtemp=!errorlevel!
    if /i [!_errtemp!]==[2] (
        goto mainmenu
    )
)
if [%_eval%]==[1] (
    %_eline%
    echo Your OS is an evaluation copy.
    echo Activating evaluation OSes is not supported.
    echo Please reinstall a non-evaluation OS.
    echo Would you like to download genuine installation media for non-evaluation Windows?
    if defined silent goto exit
    set "_errtemp="
    echo Choose an option:
    echo [Y]es
    echo [N]o
    choice /c yn /n /m "Enter your choice: "
    set _errtemp=!errorlevel!
    if /i [!_errtemp!]==[1] (
        start "" https://pastebin.com/raw/jduBSazJ
    )
    %_pak%
)
call :setkey
if not defined key (
    %_eline%
    echo Your OS isn't supported.
    echo Please use alternative activation exploits.
    echo Would you like to see a list of all operating systems and their possible activation exploits?
    if defined silent goto exit
    set "_errtemp="
    echo Choose an option:
    echo [Y]es
    echo [N]o
    choice /c yn /n /m "Enter your choice: "
    set _errtemp=!errorlevel!
    if /i [!_errtemp!]==[1] (
        start "" https://forums.mydigitallife.net/threads/windows-10-activation-in-vm.80712/#post-1656042
    )
    %_pak%
)
call :colortxt 07 "Checking internet connection...["
ping -n 1 1.1.1.1 %_nul% && set intcon=1
if not defined intcon ping -n 1 1.1.1.1 %_nul% && set intcon=1
if not defined intcon ping -n 1 1.1.1.1 %_nul% && set intcon=1
if not defined intcon ping -n 1 1.1.1.1 %_nul% && set intcon=1
if not defined intcon ping -n 1 1.1.1.1 %_nul% && set intcon=1
if defined intcon (
    call :colortxt 0a "Connected"
) else (
    call :colortxt 0c "Not connected"
)
echo ]
if not defined intcon (
    echo Digital license activation requires internet connection.
    echo Please connect to the internet and try again.
    %_pak%
)
call :colortxt 07 "Checking service ClipSVC...["
call :checkservice "ClipSVC" "cl"
echo !cl_run! ^| !cl_stat! start]
if [!cl_stat!]==[Disabled] (
    echo Configuring ClipSVC to Demand start...
    sc config ClipSVC start= Demand %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
        set "cl_change=1"
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
if [!cl_run!]==[Stopped] (
    echo Starting service ClipSVC...
    %_wmim% "!batf!?.wsf" "Win32_Service.Name='ClipSVC'" "StartService" %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
call :colortxt 07 "Checking service wlidsvc...["
call :checkservice "wlidsvc" "wl"
echo !wl_run! ^| !wl_stat! start]
if [!wl_stat!]==[Disabled] (
    echo Configuring wlidsvc to Demand start...
    sc config wlidsvc start= Demand %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
        set "wl_change=1"
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
if [!wl_run!]==[Stopped] (
    echo Starting service wlidsvc...
    %_wmim% "!batf!?.wsf" "Win32_Service.Name='wlidsvc'" "StartService" %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
call :colortxt 07 "Checking service sppsvc...["
call :checkservice "sppsvc" "sp"
echo !sp_run! ^| !sp_stat! start]
if not [!sp_stat!]==[Auto] if not [!sp_stat!]==[Delayed-Auto] (
    echo Configuring sppsvc to Delayed-Auto start...
    sc config sppsvc start= Delayed-Auto %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
        set "sp_change=1"
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
if [!sp_run!]==[Stopped] (
    echo Starting service sppsvc...
    %_wmim% "!batf!?.wsf" "Win32_Service.Name='sppsvc'" "StartService" %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
call :colortxt 07 "Checking service wuauserv...["
call :checkservice "wuauserv" "wu"
echo !wu_run! ^| !wu_stat! start]
if not [!wu_stat!]==[Auto] (
    echo Configuring wuauserv to Auto start...
    sc config wuauserv start= Auto %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
        set "wu_change=1"
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
if [!wu_run!]==[Stopped] (
    echo Starting service wuauserv...
    %_wmim% "!batf!?.wsf" "Win32_Service.Name='wuauserv'" "StartService" %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
set "_keyeligible="
for %%# in (Retail OEM MAK) do (
echo %ProductKeyChannel% | findstr /i "%%#" %_nul% && set _keyeligible=1
)
if not defined force if not defined _keyeligible call :instkey
if defined force call :instkey
for /f "tokens=2 delims==" %%# in ('%_wmi% "!batf!?.wsf" "%_slp%" "PartialProductKey is not NULL" "ID"') do set "ID=%%#"
call :patchg
echo Generating ticket...
call "!_tmp!\gatherosstate.exe"
if not exist "!_tmp!\GenuineTicket.xml" call "!_tmp!\gatherosstate.exe"
if not exist "!_tmp!\GenuineTicket.xml" call "!_tmp!\gatherosstate.exe"
if not exist "!_tmp!\GenuineTicket.xml" call "!_tmp!\gatherosstate.exe"
if not exist "!_tmp!\GenuineTicket.xml" call "!_tmp!\gatherosstate.exe"
if exist "!_tmp!\GenuineTicket.xml" (
call :colortxt %green% "Successful" /n
) else (
call :colortxt %red% "Unsuccessful" /n
goto :cleanup
)
echo Copying ticket...
copy /y "!_tmp!\GenuineTicket.xml" "%systemdrive%\ProgramData\Microsoft\Windows\ClipSVC\GenuineTicket\GenuineTicket.xml" %_nul%
if exist "%systemdrive%\ProgramData\Microsoft\Windows\ClipSVC\GenuineTicket\GenuineTicket.xml" (
call :colortxt %green% "Successful" /n
) else (
call :colortxt %red% "Unsuccessful" /n
goto :cleanup
)
echo Applying ticket...
%_wmim% "!batf!?.wsf" "Win32_Service.Name='ClipSVC'" "StopService" %_nul%
%_wmim% "!batf!?.wsf" "Win32_Service.Name='ClipSVC'" "StartService" %_nul%
timeout /t 5 >nul
if not exist "%systemdrive%\ProgramData\Microsoft\Windows\ClipSVC\GenuineTicket\GenuineTicket.xml" (
call :colortxt %green% "Successful" /n
) else (
call :colortxt %red% "Unsuccessful" /n
goto :cleanup
)
echo Activating...
set "ID="
call :chkvalues %_nul%
%_wmim% "!batf!?.wsf" "%_slp%.ID='%ID%'" "Activate" %_nul%
%_wmi% "!batf!?.wsf" "%_slp%" "LicenseStatus='1' and GracePeriodRemaining='0' and PartialProductKey is not NULL" "Name" | findstr /i "Windows" %_nul% && set success=1 || set "success="
if defined success (
    call :colortxt %green% "Product activation successful." /n
) else (
    call :retry
)
:cleanup
echo Cleaning up files...
del "!_tmp!\gatherosstate.exe" %_nul%
del "!_tmp!\slc.dll" %_nul%
del "!_tmp!\GenuineTicket.xml" %_nul%
call :checkservice "ClipSVC" "cln"
if [!cl_run!]==[Stopped] if [!cln_run!]==[Running] (
    echo Re-stopping ClipSVC...
    %_wmim% "!batf!?.wsf" "Win32_Service.Name='ClipSVC'" "StopService" %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
if defined cl_change if not [!cln_stat!]==[!cl_stat!] (
    echo Configuring ClipSVC start back to !cl_stat!...
    sc config ClipSVC start= !cl_stat! %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
call :checkservice "wlidsvc" "wln"
if [!wl_run!]==[Stopped] if [!wln_run!]==[Running] (
    echo Re-stopping wlidsvc...
    %_wmim% "!batf!?.wsf" "Win32_Service.Name='wlidsvc'" "StopService" %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
if defined wl_change if not [!wln_stat!]==[!wl_stat!] (
    echo Configuring wlidsvc start back to !wl_stat!...
    sc config wlidsvc start= !wl_stat! %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
call :checkservice "sppsvc" "spn"
if [!sp_run!]==[Stopped] if [!spn_run!]==[Running] (
    echo Re-stopping sppsvc...
    %_wmim% "!batf!?.wsf" "Win32_Service.Name='sppsvc'" "StopService" %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
if defined sp_change if not [!spn_stat!]==[!sp_stat!] (
    echo Configuring sppsvc start back to !sp_stat!...
    sc config sppsvc start= !sp_stat! %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
call :checkservice "wuauserv" "wun"
if [!wu_run!]==[Stopped] if [!wun_run!]==[Running] (
    echo Re-stopping wuauserv...
    %_wmim% "!batf!?.wsf" "Win32_Service.Name='wuauserv'" "StopService" %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
if defined wu_change if not [!wun_stat!]==[!wu_stat!] (
    echo Configuring wuauserv start back to !wu_stat!...
    sc config wuauserv start= !wu_stat! %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
%_pak%
exit /b

:k38
cls
set "_missingfiles="
for %%# in (8664_gatherosstate.exe 8664_slc.dll kmsbuildkeys.ini kmskeys.ini sku.ini ClipUp.exe) do (if not exist "!work!\bin\%%#" (if defined _missingfiles (set "_missingfiles=!_missingfiles!, %%#") else (set "_missingfiles=%%#")))
if defined _missingfiles (
    %_eline%
    echo The following required files are missing from 'bin' folder:
    echo %_missingfiles%
    %_pak%
)
call :dispstat
if /i [%LicenseStatus%]==[1] if /i [%GracePeriodRemaining%]==[0] if not defined force (
    call :colortxt %green% "Windows is already permanently activated." /n
    echo Running this script is unnecessary.
    if defined silent goto exit
    set "_errtemp="
    echo Choose an option:
    echo [C]ontinue
    echo [E]xit
    choice /c ce /n /m "Enter your choice: "
    set _errtemp=!errorlevel!
    if /i [!_errtemp!]==[2] (
        goto mainmenu
    )
)
if %osbuild% LSS 14393 (
    %_eline%
    echo You must be running minimum Windows 10 Anniversary Update {build 14393} or Server 2016 {build 14393} to use this script.
    %_pak%
)
if [%_eval%]==[1] (
    %_eline%
    echo Your OS is an evaluation copy.
    echo Activating evaluation OSes is not supported.
    echo Please reinstall a non-evaluation OS.
    echo Would you like to download genuine installation media for non-evaluation Windows?
    if defined silent goto exit
    set "_errtemp="
    echo Choose an option:
    echo [Y]es
    echo [N]o
    choice /c yn /n /m "Enter your choice: "
    set _errtemp=!errorlevel!
    if /i [!_errtemp!]==[1] (
        start "" https://pastebin.com/raw/jduBSazJ
    )
    %_pak%
)
call :setkmskey
if not defined key (
    %_eline%
    echo Your OS isn't supported.
    echo Please use alternative activation exploits.
    echo Would you like to see a list of all operating systems and their possible activation exploits?
    if defined silent goto exit
    set "_errtemp="
    echo Choose an option:
    echo [Y]es
    echo [N]o
    choice /c yn /n /m "Enter your choice: "
    set _errtemp=!errorlevel!
    if /i [!_errtemp!]==[1] (
        start "" https://forums.mydigitallife.net/threads/windows-10-activation-in-vm.80712/#post-1656042
    )
    %_pak%
)
call :colortxt 07 "Checking service ClipSVC...["
call :checkservice "ClipSVC" "cl"
echo !cl_run! ^| !cl_stat! start]
if [!cl_stat!]==[Disabled] (
    echo Configuring ClipSVC to Demand start...
    sc config ClipSVC start= Demand %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
        set "cl_change=1"
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
if [!cl_run!]==[Stopped] (
    echo Starting service ClipSVC...
    %_wmim% "!batf!?.wsf" "Win32_Service.Name='ClipSVC'" "StartService" %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
)
call :colortxt 07 "Checking service sppsvc...["
call :checkservice "sppsvc" "sp"
echo !sp_run! ^| !sp_stat! start]
if not [!sp_stat!]==[Auto] if not [!sp_stat!]==[Delayed-Auto] (
    echo Configuring sppsvc to Delayed-Auto start...
    sc config sppsvc start= Delayed-Auto %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
        set "sp_change=1"
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
if [!sp_run!]==[Stopped] (
    echo Starting service sppsvc...
    %_wmim% "!batf!?.wsf" "Win32_Service.Name='sppsvc'" "StartService" %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
set "_keyeligible="
echo %ProductKeyChannel% | findstr /i "GVLK" %_nul% && set _keyeligible=1
if not defined force if not defined _keyeligible call :instkey
if defined force call :instkey
for /f "tokens=2 delims==" %%# in ('%_wmi% "!batf!?.wsf" "%_slp%" "PartialProductKey is not NULL" "ID"') do set "ID=%%#"
%_wmim% "!batf!?.wsf" "%_slp%.ID='%ID%'" "ClearKeyManagementServiceMachine" %_nul%
%_wmim% "!batf!?.wsf" "%_slp%.ID='%ID%'" "ClearKeyManagementServicePort" %_nul%
echo Setting KMS host to [127.0.0.2]...
%_csc% //job:skmsm "!batf!?.wsf" "%ID%" "127.0.0.2" %_nul%
if [%errorlevel%]==[1] (
call :colortxt %red% "Unsuccessful" /n
) else (
call :colortxt %green% "Successful" /n
)
%_csc% //job:skmsp "!batf!?.wsf" "%ID%" "1688" %_nul%
call :patchg
if not exist "%systemroot%\System32\ClipUp.exe" (
    echo Copying ClipUp.exe...
    copy "!work!\bin\ClipUp.exe" "%systemroot%\System32\ClipUp.exe"
    set "cl_copy=1"
)
echo Generating ticket...
call "!_tmp!\gatherosstate.exe"
if not exist "!_tmp!\GenuineTicket.xml" call "!_tmp!\gatherosstate.exe"
if not exist "!_tmp!\GenuineTicket.xml" call "!_tmp!\gatherosstate.exe"
if not exist "!_tmp!\GenuineTicket.xml" call "!_tmp!\gatherosstate.exe"
if not exist "!_tmp!\GenuineTicket.xml" call "!_tmp!\gatherosstate.exe"
if exist "!_tmp!\GenuineTicket.xml" (
call :colortxt %green% "Successful" /n
) else (
call :colortxt %red% "Unsuccessful" /n
goto :kcleanup
)
echo Copying ticket...
copy /y "!_tmp!\GenuineTicket.xml" "%systemdrive%\ProgramData\Microsoft\Windows\ClipSVC\GenuineTicket\GenuineTicket.xml" %_nul%
if exist "%systemdrive%\ProgramData\Microsoft\Windows\ClipSVC\GenuineTicket\GenuineTicket.xml" (
call :colortxt %green% "Successful" /n
) else (
call :colortxt %red% "Unsuccessful" /n
goto :kcleanup
)
echo Applying ticket...
%_wmim% "!batf!?.wsf" "Win32_Service.Name='ClipSVC'" "StopService" %_nul%
%_wmim% "!batf!?.wsf" "Win32_Service.Name='ClipSVC'" "StartService" %_nul%
timeout /t 5 >nul
if not exist "%systemdrive%\ProgramData\Microsoft\Windows\ClipSVC\GenuineTicket\GenuineTicket.xml" (
call :colortxt %green% "Successful" /n
) else (
call :colortxt %red% "Unsuccessful" /n
goto :kcleanup
)
echo Unlocking KMS activation limit...
%_reapp% "!batf!?.wsf" "%_wApp%"
%_wmim% "!batf!?.wsf" "%_slp%.ID='%ID%'" "ReArmsku"
if [%errorlevel%]==[0] (
    call :colortxt 0a "Successful" /n
) else (
    call :colortxt 0c "Unsuccessful" /n
)
%_refrsh% "!batf!?.wsf"
echo Activating...
%_wmim% "!batf!?.wsf" "%_slp%.ID='%ID%'" "Activate" %_nul%
%_refrsh% "!batf!?.wsf" %_nul%
set "GracePeriodRemaining="
call :chkvalues %_nul%
set "gprtemp=%GracePeriodRemaining%"
set "success="
if %gprtemp% GTR 259200 (
    set "success=1"
)
if defined success (
    call :colortxt %green% "Product activation successful." /n
) else (
    call :catch
)
:kcleanup
echo Cleaning up files...
del "!_tmp!\gatherosstate.exe" %_nul%
del "!_tmp!\slc.dll" %_nul%
del "!_tmp!\GenuineTicket.xml" %_nul%
if defined cl_copy del "%systemroot%\System32\ClipUp.exe" %_nul%
call :checkservice "ClipSVC" "cln"
if [!cl_run!]==[Stopped] if [!cln_run!]==[Running] (
    echo Re-stopping ClipSVC...
    %_wmim% "!batf!?.wsf" "Win32_Service.Name='ClipSVC'" "StopService" %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
if defined cl_change if not [!cln_stat!]==[!cl_stat!] (
    echo Configuring ClipSVC start back to !cl_stat!...
    sc config ClipSVC start= !cl_stat! %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
call :checkservice "sppsvc" "spn"
if [!sp_run!]==[Stopped] if [!spn_run!]==[Running] (
    echo Re-stopping sppsvc...
    %_wmim% "!batf!?.wsf" "Win32_Service.Name='sppsvc'" "StopService" %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
if defined sp_change if not [!spn_stat!]==[!sp_stat!] (
    echo Configuring sppsvc start back to !sp_stat!...
    sc config sppsvc start= !sp_stat! %_nul%
    if [%errorlevel%]==[0] (
        call :colortxt 0a "Successful" /n
    ) else (
        call :colortxt 0c "Unsuccessful" /n
    )
)
%_pak%
exit /b

:switch
cls
call :chkvalues
cls
if %osbuild% LSS 17134 (
    %_eline%
    echo You must be running minimum Windows 10 April 2018 Update {build 17134} to use this option.
    %_pak%
)
if [%ostype%]==[Server] (
    %_eline%
    echo You must use Client OS to run this option.
    %_pak%
)
for %%# in (
    Core
    CoreN
    CoreSingleLanguage
    CoreCountrySpecific
) do if [%osedition%]==[%%#] set "_coreed=1"
if defined _coreed (
    echo Note: This option *CANNOT* change core editions ^(Home, Home N, etc.^) to non-core editions ^(Pro, Enterprise, Education, etc.^)
    echo To switch to these editions, follow the guide in the read-me.
    for %%# in (
        Core
        CoreN
        CoreSingleLanguage
        CoreCountrySpecific
    ) do (
        dism /english /online /get-targeteditions 2>nul | find /i "%%#" >nul && set _elig=1
    )
    if not defined _elig (
        %_eline%
        echo This script does not support switching editions on this particular SKU.
        %_pak%
    )
    echo You can switch %osedition% SKU to one of the following SKUs:
    echo:
    for /f "tokens=4" %%a in ('dism /online /english /Get-TargetEditions ^| findstr /i /c:"Target Edition : "') do (
        findstr /i /r /c:"^^%%a~.*" "!work!\bin\keys.ini" %_nul% && (
            echo %%a | find /i "Core" %_nul% && (
                set "%%a=1"
                echo %%a
            )
        )
    )
    echo:
) else (
  for %%# in (
      Education
      EducationN
      Enterprise
      EnterpriseN
      Professional
      ProfessionalN
      ProfessionalEducation
      ProfessionalEducationN
      ProfessionalWorkstation
      ProfessionalWorkstationN
      ServerRdsh
      IoTEnterprise
      CloudEdition
      CloudEditionN
    ) do (
        dism /english /online /get-targeteditions 2>nul | find /i "%%#" >nul && set _elig=1
    )
    if not defined _elig (
        %_eline%
        echo This script does not support switching editions on this particular SKU.
        %_pak%
    )
    echo You can switch %osedition% SKU to one of the following SKUs:
    echo:
    for /f "tokens=4" %%a in ('dism /online /english /Get-TargetEditions ^| findstr /i /c:"Target Edition : "') do (
        findstr /i /r /c:"^^%%a~.*" "!work!\bin\keys.ini" %_nul% && (
            set "%%a=1"
            echo %%a
        )
    )
    echo:
)
:ask
set /p "_edchc=Enter your choice, or type 'exit' to exit: "
if /i [%_edchc%]==[exit] %_pak%
if not defined %_edchc% (
    echo Your choice is not valid. Please try again.
    goto :ask
)
set "swkey="
for /f "usebackq eol=; delims=" %%# in ("!work!\bin\keys.ini") do for /f "tokens=1-3 delims=~" %%A in ("%%#") do if /i [%_edchc%]==[%%A] (
    set "swkey=%%B"
    set "swkeytype=%%C"
)
if not defined swkey (
    echo Your choice is not valid. Please try again.
    goto :ask
)
echo Installing %_edchc% %swkeytype% key [%swkey%]...
%_ipk% "!batf!?.wsf" "%swkey%" %_nul%
if not [%errorlevel%]==[0] (
call :colortxt %red% "Unsuccessful" /n
) else (
call :colortxt %green% "Successful" /n
%_refrsh% "!batf!?.wsf" %_nul%
echo Please reboot your system for the edition to be properly changed.
)
%_pak%
exit /b
:createoem
cls
echo Please choose the type of $OEM$ folder you want:
echo [D] Digital license activation
echo [K] KMS38 activation
echo [M] Digital license, fallback to KMS38 activation
set "_errtemp="
choice /c dkme /n /m "Enter your choice, or press E to exit: "
set _errtmp=!errorlevel!
if /i [!_errtmp!]==[1] set "_supp=h"
if /i [!_errtmp!]==[2] set "_supp=k"
if /i [!_errtmp!]==[3] set "_supp=kh"
if /i [!_errtmp!]==[4] %_pak%
echo:
if exist "!work!\$OEM$" (
    %_eline%
    echo $OEM$ folder already exists in the current directory.
    %_pak%
)
md "!work!\$OEM$\$$\Setup\Scripts\bin" %_nul%
copy "!batf!" "!work!\$OEM$\$$\Setup\Scripts" %_nul%
copy "!work!\bin\setupcomplete_%_supp%.cmd" "!work!\$OEM$\$$\Setup\Scripts\setupcomplete.cmd" %_nul%
if not exist "!work!\$OEM$" (
    %_eline%
    echo Failed creating $OEM$ folder.
) else (
    echo $OEM$ was successfully created in the current directory.
    echo You can place the folder in the "sources" directory of Windows installation media/ISO to pre-activate.
)
%_pak%
exit /b

:setkey
set "key="
for /f "usebackq eol=; delims=" %%# in ("!work!\bin\keys.ini") do for /f "tokens=1-3 delims=~" %%A in ("%%#") do if /i [%osedition%]==[%%A] (
    set "key=%%B"
    set "keytype=%%C"
)
for /f "usebackq eol=; delims=" %%# in ("!work!\bin\buildkeys.ini") do for /f "tokens=1-4 delims=~" %%A in ("%%#") do if /i [%osedition%]==[%%A] if %osbuild% GEQ %%B (
    set "key=%%C"
    set "keytype=%%D"
)
if defined _customkey (
    set "key=%_customkey%"
    set "keytype=Custom"
)
exit /b

:setkmskey
set "key="
for /f "usebackq eol=; delims=" %%# in ("!work!\bin\kmskeys.ini") do for /f "tokens=1-2 delims=~" %%A in ("%%#") do if /i [%osedition%]==[%%A] (
    set "key=%%B"
    set "keytype=Volume:GVLK"
)
for /f "usebackq eol=; delims=" %%# in ("!work!\bin\kmsbuildkeys.ini") do for /f "tokens=1-3 delims=~" %%A in ("%%#") do if /i [%osedition%]==[%%A] if %osbuild% GEQ %%B (
    set "key=%%C"
    set "keytype=Volume:GVLK"
)
if defined _customkey (
    set "key=%_customkey%"
    set "keytype=Custom"
)
exit /b

:checkservice
set "2=%~2"
set "tempstat="
for /f "tokens=2 delims==" %%a in ('%_wmi% "!batf!?.wsf" "Win32_Service" "Name='%~1'" "State"') do set "!2!_run=%%a"
for /f "tokens=1-4 delims=: " %%a in ('sc qc %~1') do (if /i [%%a]==[start_type] set "tempstat=%%c %%d")
if /i "%tempstat%"=="auto_start (delayed)" set !2!_stat=Delayed-Auto
if /i "%tempstat%"=="auto_start "          set !2!_stat=Auto
if /i "%tempstat%"=="demand_start "        set !2!_stat=Demand
if /i "%tempstat%"=="disabled "            set !2!_stat=Disabled
exit /b

:instkey
echo Installing %keytype% key [%key%]...
%_ipk% "!batf!?.wsf" "%key%"
if not [%errorlevel%]==[0] (
call :colortxt %red% "Unsuccessful" /n
) else (
call :colortxt %green% "Successful" /n
%_refrsh% "!batf!?.wsf" %_nul%
)
exit /b

:patchg
echo Patching gatherosstate...
call :patch
if [!src!]==[arm] (
call :colortxt %green% "Successful" /n
) else (
if exist "!_tmp!\gatherosstatemodified.exe" (
call :colortxt %green% "Successful" /n
) else (
call :colortxt %red% "Unsuccessful" /n
)
)
if exist "!_tmp!\gatherosstatemodified.exe" (
    del "!_tmp!\gatherosstate.exe" %_nul%
    ren "!_tmp!\gatherosstatemodified.exe" "gatherosstate.exe" %_nul%
)
exit /b

:patch
set "src=8664"
echo %PROCESSOR_ARCHITECTURE% | find /i "ARM" %_nul% && set src=arm
copy /y "!work!\bin\%src%_slc.dll" "!_tmp!\slc.dll" %_nul%
copy /y "!work!\bin\%src%_gatherosstate.exe" "!_tmp!\gatherosstate.exe" %_nul%
pushd "!_tmp!"
rundll32 "!_tmp!\slc.dll",PatchGatherosstate
popd
exit /b

:retry
%_reapp% "!batf!?.wsf" "%_wApp%" %_nul%
%_wmim% "!batf!?.wsf" "%_slp%.ID='%ID%'" "ReArmsku" %_nul%
%_refrsh% "!batf!?.wsf" %_nul%
%_wmim% "!batf!?.wsf" "%_slp%.ID='%ID%'" "Activate" %_nul%
%_wmi% "!batf!?.wsf" "%_slp%" "LicenseStatus='1' and GracePeriodRemaining='0' and PartialProductKey is not NULL" "Name" | findstr /i "Windows" %_nul% && set success=1 || set "success="
if defined success (
    call :colortxt %green% "Product activation successful." /n
) else (
    call :catch
)
exit /b

:catch
%_wmim% "!batf!?.wsf" "%_slp%.ID='%ID%'" "Activate" %_nul%
set tempcode=%errorlevel%
cmd /c exit /b %tempcode%
if %tempcode% NEQ 0 set "ErrCode=0x%=ExitCode%"
call :colortxt %red% "Product activation failed." /n
if defined ErrCode (
    echo Error: %ErrCode%
    echo On a computer running Microsoft Windows non-core edition, run 'slui.exe 0x2a %ErrCode%' to display the error text.
)
exit /b

:dispstat2usr
call :dispstat
%_pak%
exit /b

:dispstat
cls
call :chkvalues
cls
echo Operating System: [%fullosname%]
echo OS Version: [%osver%]
echo OS Type: [%ostype%]
echo Architecture: [%PROCESSOR_ARCHITECTURE%]
echo OS Edition: [%osedition%]
echo OS Build: [%osbuildstr%]
if defined ProductKeyChannel echo Product Key Channel: [%ProductKeyChannel%]
call :colortxt 07 "License Status: ["
call :colortxt %_clr% "%_licstat%"
echo ]
if not [%LicenseStatusReason%]==[0] (
echo License Status Reason: [%_licreas%]
)
if defined _licstatmsg echo %_licstatmsg:}= %.
if not [%GracePeriodRemaining%]==[0] (
if not [%LicenseStatus%]==[1] (
echo Grace Period Remaining: [%GracePeriodRemaining% minute^(s^)/%_gprdays% day^(s^)]
) else (
echo Activation Expiration: [%GracePeriodRemaining% minute^(s^)/%_gprdays% day^(s^)]
)
)
if not [%GracePeriodRemaining%]==[0] (
if not [%LicenseStatus%]==[1] (
echo Grace Period Expires on: [%xpr%]
) else (
echo Activation Expires on: [%xpr%]
)
)
echo SKU Value: [%OperatingSystemSKU%]
echo SKU ID: [%ID%]
echo Product PFN: [%pfn%]
echo Partial Product Key: [%_ppk%]
if [%_eval%]==[1] (
echo Evaluation: [Yes]
) else (
    echo Evaluation: [No]
)
echo Motherboard: [%_mobo%]
echo:
exit /b
:chkvalues
echo Please wait...
echo Running WMI queries...
:: Detect OS version
for /f "tokens=4,5 delims=[]. " %%G in ('ver') do set osver=%%G.%%H
:: Detect OS build
for /f "tokens=6 delims=[]. " %%L in ('ver') do set osbuild=%%L
if %osbuild% LSS 10240 (
    %_eline%
    echo You must be running minimum Windows 10 RTM {build 10240} or Server 2016 {build 14393} to use this script.
    %_pak%
)
:: Detect OS edition
for /f "skip=2 tokens=2*" %%G in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID 2^>nul') do set "oseditionreg=%%H"
:: Detect OS name
for /f "skip=2 tokens=2*" %%G in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName 2^>nul') do set "fullosname=%%H"
:: Detect full build string
for /f "skip=2 tokens=2*" %%G in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v BuildLabEx 2^>nul') do set "osbuildstr=%%H"
for /f "skip=2 tokens=2*" %%G in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v OSProductPfn 2^>nul') do set "pfn=%%H"
:: Detect OS type by @abbodi1406
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
set ostype=Server
)
if not defined ostype set ostype=Client
:: Check if OS is evaluation by @abbodi1406
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" (
set _eval=1
)
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalEdition~*.mum" (
set _eval=1
)
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalCorEdition~*.mum" (
set _eval=1
)
for %%# in (ID,LicenseStatus,PartialProductKey,_ppk,LicenseStatusReason,GracePeriodRemaining,Description,ProductKeyChannel) do set "%%#="
for /f "tokens=2 delims==" %%# in ('%_wmi% "!batf!?.wsf" "%_slp%" "ApplicationID='%_wApp%' and PartialProductKey is not null" "ID"') do set "ID=%%#"
for /f "tokens=2 delims==" %%# in ('%_wmi% "!batf!?.wsf" "%_sls%" "Version"') do set "slsver=%%#"
if not defined ID (
    set "LicenseStatus=0"
    set "PartialProductKey=N/A"
    set "_ppk=N/A"
    set "LicenseStatusReason=3221549076"
    set "GracePeriodRemaining=0"
    set "Description=N/A"
    set "ProductKeyChannel=N/A"
)
for %%# in (LicenseStatus ProductKeyChannel LicenseStatusReason GracePeriodRemaining PartialProductKey Description) do (
    if not defined %%# (
        for /f "tokens=* delims=" %%# in ('%_wmi% "!batf!?.wsf" "%_slp%" "ID='%ID%'" "%%#"') do set "%%#"
    )
)
for %%# in (OperatingSystemSKU Caption) do (
    for /f "tokens=* delims=" %%# in ('%_wmi% "!batf!?.wsf" "Win32_OperatingSystem" "%%#"') do set "%%#"
)
if defined Caption set "fullosname=%Caption%"
if %osbuild% LSS 7600 (
    set "fullosname=!fullosname:Microsoftr=Microsoft!"
    set "fullosname=!fullosname:VistaT=Vista!"
    set "fullosname=!fullosname:Serverr=Server!"
)
set "space= "
if [!fullosname:~-1!]==[!space!] set "fullosname=!fullosname:~0,-1!"
for /f "usebackq eol=; delims=" %%# in ("!work!\bin\sku.ini") do for /f "tokens=1-2 delims=:" %%A in ("%%#") do if /i [%OperatingSystemSKU%]==[%%A] set "osedition=%%B"
if %osbuild% LSS 7600 if /i [%OperatingSystemSKU%]==[33] set "osedition=ServerSBSPrime"
if [%LicenseStatus%]==[0] (
set "_licstat=Unlicensed"
set "_clr=0c"
)
if [%LicenseStatus%]==[1] (
set "_licstat=Licensed"
set "_clr=0a"
)
if [%LicenseStatus%]==[2] (
set "_licstat=Initial grace period"
set "_clr=0e"
)
if [%LicenseStatus%]==[3] (
set "_licstat=Additional grace period"
set "_clr=0e"
)
if [%LicenseStatus%]==[4] (
set "_licstat=Non-genuine grace period"
set "_clr=0e"
)
if [%LicenseStatus%]==[5] (
set "_licstat=Notification"
set "_clr=0c"
)
if [%LicenseStatus%]==[6] (
set "_licstat=Extended grace period"
set "_clr=0e"
)
if not defined _licstat (
    set "_licstat=Unavailable"
    set "_clr=08"
)
set /a _gprdays=(GracePeriodRemaining+1439)/1440
cmd /c exit /b %LicenseStatusReason%
set _licreas=0x%=ExitCode%
for %%# in (
0x4004F00C:The}Software}Licensing}Service}reported}that}the}application}is}running}within}the}valid}grace}period
0x4004F00D:The}Software}Licensing}Service}reported}that}the}application}is}running}within}the}valid}out}of}tolerance}grace}period
0x4004F040:The}Software}Licensing}Service}reported}that}the}product}was}activated}but}the}owner}should}verify}the}Product}Use}Rights
0x4004F401:The}Software}Licensing}Service}reported}that}the}application}has}a}store}license
0xC004F001:The}Software}Licensing}Service}reported}an}internal}error
0xC004F007:The}Software}Licensing}Service}reported}that}the}license}could}not}be}found
0xC004F008:The}Software}Licensing}Service}reported}that}the}license}could}not}be}found
0xC004F009:The}Software}Licensing}Service}reported}that}the}grace}period}has}expired
0xC004F014:The}Software}Licensing}Service}reported}that}the}product}key}is}not}available
0xC004F034:The}Software}Licensing}Service}reported}that}the}license}could}not}be}found}or}was}invalid
0xC004F035:The}Software}Licensing}Service}reported}that}the}computer}could}not}be}activated}with}a}Volume}license}product}key
0xC004F057:The}Software}Licensing}Service}reported}that}the}computer}BIOS}is}missing}a}required}license
0xC004F058:The}Software}Licensing}Service}reported}that}the}computer}BIOS}is}missing}a}required}license
0xC004F059:The}Software}Licensing}Service}reported}that}a}license}in}the}computer}BIOS}is}invalid
0xC004F200:The}Software}Licensing}Service}reported}that}current}state}is}not}genuine
) do for /f "delims=: tokens=1-2" %%A in ("%%#") do if /i [%_licreas%]==[%%A] set "_licstatmsg=%%B"
if not defined _ppk set _ppk=*****-*****-*****-*****-%PartialProductKey%
for /f "tokens=2 delims=," %%A in ("%Description%") do set "_channel=%%A"
for /f "tokens=1 delims= " %%A in ("%_channel%") do set "_channel=%%A"
if [%LicenseStatus%]==[0] set "_channel="
:: Detect motherboard
for %%# in (manufacturer product) do (
    for /f "tokens=* delims=" %%# in ('%_wmi% "!batf!?.wsf" "Win32_BaseBoard"  "%%#"') do set "%%#"
)
set "_mobo=%manufacturer% %product%"
if not [%GracePeriodRemaining%]==[0] for /f "tokens=* delims=" %%# in ('%_xpr% "!batf!?.wsf" "%GracePeriodRemaining%"') do set "xpr=%%#"
exit /b
:colortxt
setlocal enableDelayedExpansion
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:colorPrint Color  Str  [/n]
setlocal
set "s=%~2"
call :colorPrintVar %1 s %3
exit /b

:colorPrintVar  Color  StrVar  [/n]
if not defined DEL call :initColorPrint
setlocal enableDelayedExpansion
pushd .
':
cd \
set "s=!%~2!"
:: The single blank line within the following IN() clause is critical - DO NOT REMOVE
for %%n in (^"^

^") do (
  set "s=!s:\=%%~n\%%~n!"
  set "s=!s:/=%%~n/%%~n!"
  set "s=!s::=%%~n:%%~n!"
)
for /f delims^=^ eol^= %%s in ("!s!") do (
  if "!" equ "" setlocal disableDelayedExpansion
  if %%s==\ (
    findstr /a:%~1 "." "\'" nul
    <nul set /p "=%DEL%%DEL%%DEL%"
  ) else if %%s==/ (
    findstr /a:%~1 "." "/.\'" nul
    <nul set /p "=%DEL%%DEL%%DEL%%DEL%%DEL%"
  ) else (
    >colorPrint.txt (echo %%s\..\')
    findstr /a:%~1 /f:colorPrint.txt "."
    <nul set /p "=%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%"
  )
)
if /i "%~3"=="/n" echo(
popd
exit /b


:initColorPrint
for /f %%A in ('"prompt $H&for %%B in (1) do rem"') do set "DEL=%%A %%A"
<nul >"%temp%\'" set /p "=."
subst ': "%temp%" >nul
exit /b


:cleanupColorPrint
2>nul del "%temp%\'"
2>nul del "%temp%\colorPrint.txt"
>nul subst ': /d
exit /b

:exit
exit /b
----- Begin wsf script --->
<package>
   <job id="ELAV">
       <script language="VBScript">
           Set strArg=WScript.Arguments.Named

           If Not strArg.Exists("File") Then
               Wscript.Echo "Switch /File:<File> is missing."
               WScript.Quit 1
           End If

           Set strRdlproc = CreateObject("WScript.Shell").Exec("rundll32 kernel32,Sleep")
           With GetObject("winmgmts:\\.\root\CIMV2:Win32_Process.Handle='" & strRdlproc.ProcessId & "'")
               With GetObject("winmgmts:\\.\root\CIMV2:Win32_Process.Handle='" & .ParentProcessId & "'")
                   If InStr (.CommandLine, WScript.ScriptName) <> 0 Then
                       strLine = Mid(.CommandLine, InStr(.CommandLine , "/File:") + Len(strArg("File")) + 8)
                   End If
               End With
               .Terminate
           End With

          CreateObject("Shell.Application").ShellExecute "cmd", "/c " & chr(34) & chr(34) & strArg("File") & chr(34) & " " & strLine & chr(34), "", "runas", 1
       </script>
   </job>
   <job id="WMI">
      <script language="VBScript">
         If WScript.Arguments.Count = 3 Then
            wExc = "Select " & WScript.Arguments.Item(2) & " from " & WScript.Arguments.Item(0) & " where " & WScript.Arguments.Item(1)
         Else
            wExc = "Select " & WScript.Arguments.Item(1) & " from " & WScript.Arguments.Item(0)
         End If
         Set objCol = GetObject("winmgmts:\\.\root\CIMV2").ExecQuery(wExc,,48)
         For Each objItm in objCol
            For each Prop in objItm.Properties_
               WScript.Echo Prop.Name & "=" & Prop.Value
            Next
         Next
      </script>
   </job>
   <job id="WmiMethod">
      <script language="VBScript">
         On Error resume next
         wPath = WScript.Arguments.Item(0)
         wMethod = WScript.Arguments.Item(1)
         Set objCol = GetObject("winmgmts:\\.\root\CIMV2:" & wPath)
         objCol.ExecMethod_(wMethod)
         WScript.Quit Err.Number
      </script>
   </job>
   <job id="ATO">
      <script language="VBScript">
         On Error resume next
         Set objCol = GetObject("WinMgmts:").ExecQuery("Select * From SoftwareLicensingProduct Where ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' And PartialProductKey is not NULL")
         For Each objItm in objCol
            objItm.ExecMethod_("Activate")
         Next
         WScript.Quit Err.Number
      </script>
   </job>
   <job id="IPK">
      <script language="VBScript">
         On error resume next
         pkey = Wscript.Arguments.Item(0)
         Set objWMIService = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT Version FROM SoftwareLicensingService")
         For each colService in objWMIService
            Exit For
         Next
         set objService = colService
         objService.InstallProductKey(pkey)
         WScript.Quit Err.Number
      </script>
   </job>
   <job id="skmsm">
      <script language="VBScript">
         On error resume next
         kms = Wscript.Arguments.Item(1)
         Set objWMIService = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT * FROM SoftwareLicensingProduct where ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey is not NULL")
         For each colService in objWMIService
            Exit For
         Next
         set objService = colService
         objService.SetKeyManagementServiceMachine(kms)
         WScript.Quit Err.Number
      </script>
   </job>
   <job id="skmsp">
      <script language="VBScript">
         On error resume next
         kms = Wscript.Arguments.Item(1)
         Set objWMIService = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT * FROM SoftwareLicensingProduct where ID='" & Wscript.Arguments.Item(0) & "'")
         For each colService in objWMIService
            Exit For
         Next
         set objService = colService
         objService.SetKeyManagementServicePort(kms)
         WScript.Quit Err.Number
      </script>
   </job>
   <job id="reapp">
      <script language="VBScript">
         On error resume next
         appid = Wscript.Arguments.Item(0)
         Set objWMIService = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT Version FROM SoftwareLicensingService")
         For each colService in objWMIService
            Exit For
         Next
         set objService = colService
         objService.ReArmApp(appid)
         WScript.Quit Err.Number
      </script>
   </job>
   <job id="refresh">
      <script language="VBScript">
         On error resume next
         strComputer = "."
         Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
         Set objCol = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("Select Version from SoftwareLicensingService")
         For Each objItm in objCol
            slsver = objItm.Version
         Next
         qry = "SoftwareLicensingService.Version='" & slsver & "'"
         Set objShare = objWMIService.Get(qry)
         Set objOutParams = objWMIService.ExecMethod(qry, "RefreshLicenseStatus")
         WScript.Quit Err.Number          
      </script>
   </job>
   <job id="XPR">
      <script language="VBScript">
         Wscript.Echo DateAdd("n", WScript.Arguments.Item(0), Now)
      </script>
   </job>
</package>