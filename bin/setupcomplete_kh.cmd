<!-- : Begin batch script
@ECHO OFF
cd /d "%~dp0"
for /f "delims=" %%# in ('dir /b Ez*.cmd') do set "script=%%#"
call "%script%" /dl /silent
for /f "tokens=2 delims==" %%# in ('C:\Windows\System32\cscript.exe //nologo //job:WMI "%~nx0?.wsf" "SoftwareLicensingProduct" "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey is not null" "LicenseStatus"') do set "licstat=%%#"
if not [%licstat%]==[1] (
call "%script%" /k38 /silent
)
attrib -R -A -S -H *.*
RMDIR /S /Q "%WINDIR%\Setup\Scripts"
exit /b
----- Begin wsf script --->
<package>
   <job id="WMI">
      <script language="VBScript">
         wExc = "Select " & WScript.Arguments.Item(2) & " from " & WScript.Arguments.Item(0) & " where " & WScript.Arguments.Item(1)
         Set objCol = GetObject("winmgmts:\\.\root\CIMV2").ExecQuery(wExc,,48)
         For Each objItm in objCol
            For each Prop in objItm.Properties_
               WScript.Echo Prop.Name & "=" & Prop.Value
            Next
         Next
      </script>
   </job>
</package>