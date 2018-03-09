Echo This will start the firmware update automation tool
ping localhost -n 15
C:
CD "C:\CopTrax Support\Tools\FirmwareAutomation"
schtasks /Delete /TN Automation /F
schtasks /Create /XML autorun.xml /TN Automation
Del /S /F user.flg
ping localhost -n 5
Start FirmwareUpdateTool.exe
exit