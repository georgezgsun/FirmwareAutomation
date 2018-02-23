#RequireAdmin

#pragma compile(FileVersion, 1.2.20.3)
#pragma compile(FileDescription, Firmware Update Automation Tool)
#pragma compile(ProductName, AutomationTest)
#pragma compile(ProductVersion, 2.11)
#pragma compile(CompanyName, 'Stalker')
#pragma compile(Icon, automation.ico)

;Firmware update
;1. set autostart, write automation.bat;
;2. Open Validation tool, click on advance
;3. Click on Enter Bootload
;4. Box reboot; force reboot? set autostart;
;5. Run automation, read status, kill the autostart
;6. Kill Coptrax process
;7. start program tool
;8. click connection, two clicks
;9. load hex file
;10. click to program and write and verify the hex
;11. click to run app

#include <File.au3>
#include <Misc.au3>
#include <MsgBoxConstants.au3>

_Singleton('Firmware Update Automation Tool')

Global $userName = "Unknown"
Global $firmwareVersion = ""
Global $libraryVersion = ""
Global  $targetVersion = "1.3.0.2"
Global $updatingEnd = False

Global $workDir = "C:\FirmwareAutomation\"
Global $filename = "C:\CopTrax Support\Tools\FirmwareUpdater.log"
Global $logFile = FileOpen($filename, $FO_APPEND)
Local $flagFile = FileOpen($workDir & "user.flg", $FO_APPEND)
FileWrite($flagFile, "x")
FileClose($flagFile)

Local $versionFile = FileOpen($workDir & "version.cfg")
$targetVersion = FileReadLine($versionFile)
$userName = FileReadLine($versionFile)
FileClose($versionFile)

If Not StringRegExp($targetVersion, "([0-9]+\.[0-9]+\.[0-9]+\.?[0-9a-zA-Z]*)") Then
	LogWrite("Target firmware version has a wrong format. Please assign a correct firmware version before updating the firmware.")
	$updatingEnd = True
	Exit
EndIf

Global $firmwareFile = $workDir & "TriggerBox 2.0 App " & $targetVersion & ".hex"
If Not FileExists($firmwareFile) Then
	LogWrite("Cannot find " & $firmwareFile & ". Please provide a valid firmware hex file before updating.")
	$updatingEnd = True
	Exit
EndIf

HotKeySet("{Esc}", "HotKeyPressed") ; Esc to stop testing
HotKeySet("q", "HotKeyPressed") ; Esc to stop testing
OnAutoItExitRegister("OnAutoItExit")	; Register OnAutoItExit to be called when the script is closed.
AutoItSetOption ("WinTitleMatchMode", 2)	; match any substring in the title
AutoItSetOption("SendKeyDelay", 100)

EndCopTrax()
Local $startTimes = FileGetSize($workDir & "user.flg")

Switch $startTimes
	Case 1
		LogWrite("Running application is terminated to continue the firmware updating.")
		LogWrite("Target firmware version read from version.cfg is " & $targetVersion)
		Run("schtasks /Delete /TN Automation /F", "", @SW_HIDE)
		Run("schtasks /Create /XML " & $workDir & "autorun.xml /TN Automation", "", @SW_HIDE)
		FileCopy($workDir & "Cleanup.bat", "C:\CopTrax Support\Tools\Cleanup.bat", 1)	; force copy the cleanup.bat to the working folder
		LogWrite("Firmware update automation tool begin the first part of firmware updating.")
		RunValidationTool()
		Sleep(1000)
		RunValidationTool()
		Exit
	Case 2
		LogWrite("Read start times is " & $startTimes)
		LogWrite("Firmware update automation tool begin the second part of firmware updating.")
		RunFirmwareTool()
		Sleep(1000)
		RunFirmwareTool()
		Exit
	Case 3, 4
		LogWrite("Read start times is " & $startTimes)
		LogWrite("Firmware update automation tool begin the last part of firmware updating.")
		RunValidationTool()
		Sleep(1000)
		RunValidationTool()
		Exit
	Case Else
		LogWrite("Read start times is " & $startTimes)
		LogWrite("Unable to update the firmware by automation tool. Please update the firmware manually. ")
		LogWrite("Save the log file for further investigation.")
		Cleanup()
EndSwitch

Exit

Func Cleanup()
	FileClose($logFile)
	Run("schtasks /Delete /TN Automation /F", "", @SW_HIDE)
	Local $currentDir = "C:\CopTrax Support\Tools\"
	FileMove($filename, $currentDir & $userName & ".log", 1)
	Run($currentDir & "Cleanup.bat")
	$updatingEnd = True
EndFunc

Func OnAutoItExit()
	FileClose($logFile)
    If Not $updatingEnd Then
		Shutdown(2+4)
	EndIf
 EndFunc   ;==>OnAutoItExit

Func LogWrite($s)
	_FileWriteLog($logFile,$s)
	MsgBox($MB_OK, "Firmware update automation tool", $s, 2)
EndFunc

Func EndCopTrax()
	Local $pCopTrax = "IncaXPCApp.exe"
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
		$updatingEnd = True
		Exit
	EndIf

	ControlClick($hwnd, "", "[CLASS:Button; INSTANCE:13]")	; click on enable USB connection checkbox
	Sleep(1000)
	ControlClick($hwnd, "", "[CLASS:Button; INSTANCE:7]")	; click on connect button
	Sleep(1000)

	Local $txt = WinGetText($hWnd)
	If StringInStr($txt, "reset device") Or WinExists("Error") Then
		WinClose("Error")
		LogWrite("Unable to connect to firmware. Try again.")

		WinClose($hWnd)
		Sleep(1000)

		$hWnd = GetHandleWindowWait("Exit")
		If $hWnd = 0 Then
			LogWrite("Unable to run PIC32UBL.exe.")
			$updatingEnd = True
			Exit
		EndIf

		ControlClick($hWnd, "", "&Yes")
		Return
	EndIf

	ControlClick($hWnd, "", "[CLASS:Button; INSTANCE:1]")	; click on Load Hex File
	Sleep(500)
	Send($firmwareFile & "{ENTER}")
	Sleep(1000)
	ControlClick($hWnd, "", "[CLASS:Button; INSTANCE:8]")	; click on Erase-Program-Verify button

	Local $done = False
	Local $i = 0
	Local $erased = False
	Do
		$txt = WinGetText($hWnd)
		If Not $erased And StringInStr($txt, "Flash Erased") Then
			LogWrite("Flash Erased in " & $i & "s.")
			$erased = True
		EndIf

		If StringInStr($txt, "Programming completed") Then
			LogWrite("Programming completed in " & $i & "s.")
		EndIf

		If StringInStr($txt, "Verification successfull") Then
			LogWrite("Verification successfull in " & $i & "s.")
			$done = True
		EndIf
		$i += 1
		Sleep(1000)
	Until $done Or $i > 45

	If Not $done Then
		LogWrite("Programming failed. Get window text as " & $txt)
		$updatingEnd = True
		Exit
	EndIf

	LogWrite("Wait for hard reset to complete the firmware update.")
	FileClose($logFile)

	MsgBox($MB_OK, "Firmware update automation tool", "Please press the hard reset button when the power LED turn solid red to complete the firmware update.", 10)
	ControlClick($hWnd, "", "[CLASS:Button; INSTANCE:5]")	; click on Run Application button
	Exit
EndFunc

Func RunValidationTool()
	LogWrite("Run " & $workDir & "CopTraxBoxII.exe at " &$workDir)
	Run($workDir & "CopTraxBoxII.exe", $workDir)

	Local $hWnd = GetHandleWindowWait("Trigger")
	If $hWnd = 0 Then
		LogWrite("Unable to trigger Validation Tool.")
		$updatingEnd = True
		Exit
	EndIf

	ControlClick($hWnd, "", "[NAME:libConnect]")
	Sleep(1000)
	If WinExists("CopTraxII", "OK") Then
		ControlClick("CopTraxII", "OK", "OK")
		LogWrite("A hard reset is required to complete the update.")
		MsgBox($MB_OK, "Firmware update automation tool", "It is required to press the hard reset button to complete the firmware update. " & @CRLF & "Do not click this OK button.")
		Exit
	EndIf

	$title = WinGetTitle($hwnd) ; CopTraxII -  Library Version:  1.0.1.5, Firmware Version:  2.1.1
	Local $splittedTitle = StringRegExp($title, "([0-9]+\.[0-9]+\.[0-9]+\.?[0-9a-zA-Z]*)", $STR_REGEXPARRAYGLOBALMATCH)
	If IsArray($splittedTitle) And $splittedTitle[0] = 2 Then
		$libraryVersion = $splittedTitle[0]
		$firmwareVersion = $splittedTitle[1]
	Else
		LogWrite("Cannot connect to firmware. Try again.")
		WinClose($hWnd)
		Return
	EndIf
	$splittedTitle = StringRegExp(WinGetText($hwnd), "(?:Product: )([A-Za-z]{2}[0-9]{6})", $STR_REGEXPARRAYMATCH )
	If IsArray($splittedTitle) Then
		$userName = $splittedTitle[0]
		Local $versionFile = FileOpen($workDir & "version.cfg", 2)
		FileWriteLine($versionFile, $targetVersion)
		FileWriteLine($versionFile, $userName)
		FileClose($versionFile)
	EndIf
	LogWrite("Reading from validation tool, the serial number of the box is " & $userName & ", the firmware version is " & $firmwareVersion & ", the library version is " & $libraryVersion)

	ControlClick($hWnd, "", "[NAME:radioButton_HBOff]")	; set the heartbeat to off, preventing unnecessary reboot

	If _VersionCompare($firmwareVersion, $targetVersion) = 0 Then
		LogWrite("The firmware has been uptodated to " & $firmwareVersion & ". Exit validation tool now.")
		Sleep(5000)
		WinClose($hWnd)
		Cleanup()
		Run("c:\Program Files (x86)\IncaX\CopTrax\IncaXPCApp.exe", "c:\Program Files (x86)\IncaX\CopTrax")
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
	Sleep(1000)
	LogWrite("Unable to click on Enter Bootload button.")
	$updatingEnd = True
	Exit
EndFunc

Func HotKeyPressed()
	Switch @HotKeyPressed ; The last hotkey pressed.
		Case "{Esc}", "q" ; KeyStroke is the {ESC} hotkey. to stop testing and quit
			$updatingEnd = True	;	Stop testing marker
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