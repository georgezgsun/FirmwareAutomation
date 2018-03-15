#RequireAdmin

#pragma compile(FileVersion, 1.2.20.12)
#pragma compile(FileDescription, Firmware Update Automation Tool)
#pragma compile(ProductName, AutomationTest)
#pragma compile(ProductVersion, 1.1)
#pragma compile(CompanyName, 'Stalker')
#pragma compile(Icon, automation.ico)

;Firmware update
;1. set autostart: copy all neccessary files to C; read target version number
;2. Run Validation tool
;3. Trying to connect to firmware. in case failed, run PIC programmer
;4. After connect, check current firmware version first;
;5. In case the firmware version is not the target version, go to Bootload
;6. In case the firmware version is correct, do cleanup and restart the Coptrax
;7. In PIC programmer, try to connect the bootload;
;8. In case connected, load the hex file and program it;
;9. In case not connected, reboot;
;10. After 5 times reboot, do the cleanup;

#include <File.au3>
#include <Misc.au3>
#include <MsgBoxConstants.au3>

_Singleton('Firmware Update Automation Tool')

Global $userName = "Unknown"
Global $firmwareVersion = ""
Global $libraryVersion = ""
Global $targetVersion = "1.3.0.2"
Global $updatingEnd = False
Global $pCopTrax = "IncaXPCApp.exe"

If WinExists("", "Open CopTrax") Then
	ControlClick("", "Open CopTrax", "[NAME:panelCopTrax]")
	Sleep(1000)
EndIf

Global $workDir = @ScriptDir & "\"
Local $destDir = "C:\CopTrax Support\Tools\FirmwareAutomation\"
Local $rst = False
Local $i = 0
If ($workDir <> $destDir) Or Not FileExists($workDir & "user.flg") Then
	While Not $rst And ($i < 3)
		$rst = DirCopy($workDir, $destDir, 1)	; copies the directory $sourceDir and all sub-directories and files to $destDir in overwrite mode
		$i += 1
		Sleep(500)
	WEnd

	If Not $rst Then
		MsgBox($MB_OK, "Firmware update automation tool", "Cannot autostart the firmware update tool.", 10)
		Exit
	Else
		$workDir = $destDir
		FileChangeDir($destDir)
		FileDelete("user.flg")
		FileCopy("cleanup.bat", "C:\CopTrax Support\Tools", 1)
		Run(@comSpec & " /c schtasks /Delete /TN Automation /F")
		Sleep(500)
		Run(@comSpec & " /c schtasks /Create /XML autorun.xml /TN Automation", $workDir)
;		Run($workDir & "SetupFirmwareUpdate.bat")
		MsgBox($MB_OK, "Firmware update automation tool", "The firmware update tool has been setup.", 2)
;		Exit
	EndIf
EndIf

Local $versionFile = FileOpen($workDir & "version.cfg")
$targetVersion = FileReadLine($versionFile)
$userName = FileReadLine($versionFile)
FileClose($versionFile)

If Not StringRegExp($targetVersion, "([0-9]+\.[0-9]+\.[0-9]+\.?[0-9a-zA-Z]*)") Then
	MsgBox($MB_OK, "Target firmware version has a wrong format. Please assign a correct firmware version before updating the firmware.", 10)
	Exit
EndIf

Global $firmwareFile = $workDir & "TriggerBox 2.0 App " & $targetVersion & ".hex"
If Not FileExists($firmwareFile) Then
	MsgBox($MB_OK, "Cannot find " & $firmwareFile & ". Please provide a valid firmware hex file before updating.", 10)
	$updatingEnd = True
	Exit
EndIf

HotKeySet("{Esc}", "HotKeyPressed") ; Esc to stop testing
HotKeySet("q", "HotKeyPressed") ; Esc to stop testing
;OnAutoItExitRegister("OnAutoItExit")	; Register OnAutoItExit to be called when the script is closed.
AutoItSetOption ("WinTitleMatchMode", 2)	; match any substring in the title
AutoItSetOption("SendKeyDelay", 100)

Global $filename = "C:\CopTrax Support\Tools\FirmwareUpdater.log"
Global $logFile = FileOpen($filename, $FO_APPEND)
Local $flagFile = FileOpen($workDir & "user.flg", $FO_APPEND)
FileWrite($flagFile, "x")
FileClose($flagFile)
RegWrite("HKEY_CURRENT_USER\Control Panel\Desktop", "AutoEndTasks", "REG_SZ", 1)	; Shut down without user's response

EndCopTrax()

Local $startTimes = FileGetSize($workDir & "user.flg")
LogWrite("Reboot times is " & $startTimes & ". CopTrax App is terminated to continue the firmware update.")

If $startTimes > 5 Then
	CleanUp()
EndIf

If $startTimes = 1 Then
	LogWrite("Try to update the firmware to version " & $targetVersion)
EndIf

RunValidationTool()
RunValidationTool()
Cleanup()
Exit

Func Cleanup()
	FileClose($logFile)
	Local $CopTraxAppDir = @ProgramFilesDir & "\IncaX\CopTrax\"
	Local $currentDir = "C:\CopTrax Support\Tools\"
	Run($CopTraxAppDir & $pCopTrax, $CopTraxAppDir)
	Run(@comSpec & " /c schtasks /Delete /TN Automation /F")
	FileMove($filename, $currentDir & $userName & ".log", 1)
	ProcessClose("CopTraxBoxII.exe")
	ProcessClose("PIC32UBL.exe")
	Run($currentDir & "Cleanup.bat")
	Exit
EndFunc

Func LogWrite($s, $show = True)
	_FileWriteLog($logFile,$s)
	If Not $show Then Return
	MsgBox($MB_OK, "Firmware update automation tool", $s, 2)
EndFunc

Func EndCopTrax()
	If ProcessExists($pCopTrax) And ProcessClose($pCopTrax) Then
		Return
	EndIf

	Sleep(1000)
	If ProcessExists($pCopTrax) Then
		LogWrite("Unable to end " & $pCopTrax)
		Exit
	EndIf
EndFunc

Func RunFirmwareTool()
	LogWrite("Run " & $workDir & "PIC32UBL.exe at " &$workDir)
	Run($workDir & "PIC32UBL.exe", $workDir)

	Local $hWnd = GetHandleWindowWait("PIC32")
	If $hWnd = 0 Then
		LogWrite("Unable to run PIC32UBL.exe.")
		Shutdown(1+4+16)
		Exit
	EndIf

	ControlClick($hwnd, "", "[CLASS:Button; INSTANCE:13]")	; click on enable USB connection checkbox
	Sleep(1000)
	ControlClick($hwnd, "", "[CLASS:Button; INSTANCE:7]")	; click on connect button
	Sleep(1000)

	Local $txt = WinGetText($hWnd)
	If StringInStr($txt, "reset device") Or WinExists("Error") Then
		ProcessClose("PIC32UBL.exe")
		LogWrite("Unable to connect to firmware.")
		Shutdown(1+4+16)
		Exit
	EndIf

	ControlClick($hWnd, "", "[CLASS:Button; INSTANCE:1]")	; click on Load Hex File
	Sleep(500)
	Send(" " & $firmwareFile & "{ENTER}")
;		ControlSend("Open", "", "[CLASS:Edit; INSTANCE:1]", $firmwareFile)
	If Not WaitFor($hWnd, "loaded successfully") Then Exit

	ControlClick($hWnd, "", "[CLASS:Button; INSTANCE:4]")	; click on Erase
	If Not WaitFor($hWnd, "Flash Erased") Then Exit

	ControlClick($hWnd, "", "[CLASS:Button; INSTANCE:2]")	; click on Program
	If Not WaitFor($hWnd, "Programming completed") Then Exit

	ControlClick($hWnd, "", "[CLASS:Button; INSTANCE:3]")	; click on Verification
	If Not WaitFor($hWnd, "Verification successfull") Then Exit

	LogWrite("New firmware has been programmed successfully. Reboot now to check the final result.")
	FileClose($logFile)

	Sleep(2000)	; add more delay before click on Run Application button
	ControlClick($hWnd, "", "[CLASS:Button; INSTANCE:5]")	; click on Run Application button
	Sleep(200)
	Shutdown(1+4+16)
EndFunc

Func WaitFor($hwnd, $txt)
	Local $done = False
	Local $i = 0
	Do
		If StringInStr(WinGetText($hWnd), $txt) Then
			LogWrite($txt & " in " & $i & "s.")
			$done = True
		EndIf
		$i += 1
		Sleep(1000)
	Until $done Or $i > 45

	If Not $done Then
		LogWrite("Programming failed. Get window text as " & WinGetText($hWnd))
		ProcessClose("PIC32UBL.exe")
		Sleep(200)
		Return False
	EndIf

	Return $done
EndFunc

Func RunValidationTool()
	LogWrite("Run " & $workDir & "CopTraxBoxII.exe at " &$workDir)
	Run($workDir & "CopTraxBoxII.exe", $workDir)

	Local $hWnd = GetHandleWindowWait("Trigger")
	If $hWnd = 0 Then
		LogWrite("Unable to trigger Validation Tool.")
		Exit
	EndIf

	ControlClick($hWnd, "", "[NAME:libConnect]")
	Sleep(1000)
	If WinExists("CopTraxII", "OK") Then
		ControlClick("CopTraxII", "OK", "OK")
		RunFirmwareTool()
		Shutdown(1+2+16)
	EndIf

	Local $title = WinGetTitle($hwnd) ; CopTraxII -  Library Version:  1.0.1.5, Firmware Version:  2.1.1
	Local $splittedTitle = StringRegExp($title, "([0-9]+\.[0-9]+\.[0-9]+\.?[0-9a-zA-Z]*)", $STR_REGEXPARRAYGLOBALMATCH)
	If IsArray($splittedTitle) And UBound($splittedTitle) = 2 Then
		$libraryVersion = $splittedTitle[0]
		$firmwareVersion = $splittedTitle[1]
	Else
		LogWrite("Cannot connect to firmware. Try again. " & $title)
		WinClose($hWnd)
		Return
	EndIf
	$splittedTitle = StringRegExp(WinGetText($hwnd), "(?:Product: )([A-Za-z]{2}[0-9]{6})", $STR_REGEXPARRAYMATCH )
	If IsArray($splittedTitle) Then
		$userName = $splittedTitle[0]
		Local $versionFile = FileOpen($workDir & "version.cfg", 2)
		FileWriteLine($versionFile, $targetVersion)
		FileWriteLine($versionFile, $userName)
		FileWriteLine($versionFile, $firmwareVersion)
		FileClose($versionFile)
	EndIf
	LogWrite("Reading from validation tool, the serial number of the box is " & $userName & ", the firmware version is " & $firmwareVersion & ", the library version is " & $libraryVersion)

	;ControlClick($hWnd, "", "[NAME:radioButton_HBOff]")	; set the heartbeat to off, preventing unnecessary reboot

	If _VersionCompare($firmwareVersion, $targetVersion) = 0 Then
		LogWrite("The firmware has been uptodated to " & $firmwareVersion & ".")
		Sleep(2000)
		WinClose($hWnd)
		Cleanup()
		Exit
	EndIf

	MouseWheel("down")	; move screen down

	ControlClick($hWnd, "", "Advance")	; set the heartbeat to off, preventing unnecessary reboot

	Local $hAdv = GetHandleWindowWait("Advance")
	If $hAdv = 0 Then
		LogWrite("Unable to open Advance Settings window.")
		$updatingEnd = True
		Exit
	EndIf

	LogWrite("Click on Enter Bootload button.")
	ControlClick($hAdv, "", "Enter Bootload")	; Click on Enter Bootload button, and this may introduce a reboot
;	Sleep(30000)	; wait long enough for the Bootload to take effect
;	LogWrite("Unable to click on Enter Bootload button.")
;	$updatingEnd = True
	Exit
EndFunc

Func HotKeyPressed()
	Switch @HotKeyPressed ; The last hotkey pressed.
		Case "{Esc}", "q" ; KeyStroke is the {ESC} hotkey. to stop testing and quit
			LogWrite("Firmware update stopped by operator.")
			FileClose($logFile)
			Exit
	EndSwitch
EndFunc

Func GetHandleWindowWait($title, $seconds = 10)
	Local $hWnd = 0
	Local $i = 0
	If $seconds < 1 Then $seconds = 1
	If $seconds > 1000 Then $seconds = 1000
	While ($hWnd = 0) And ($i < $seconds)
		WinActivate($title)
		$hWnd = WinWaitActive($title, "", 1)
		$i += 1
	WEnd
	Return $hWnd
EndFunc
;[CLASS:DirectUIHWND; INSTANCE:0]
;[CLASS:IPTip_Main_Window}