; Centr_ScriptsGUI.au3
; version 1.1.04
; updated 9/3/2010 by Peter Stein
; adding mult buttons
; added command line parameters to read ini file
; added ability to cancel the phone note msgBox.
; The main script for getting to letter directory
; right now, all is in one script
; change Dim to Local

#include <GUIConstantsEx.au3>
#include <file.au3>
#include <array.au3>

;; CONSTANTS
Const $Version = "1.1.04"
Const $DataFileNameConst = @ScriptDir & "\auData.ini"  ; name of the ini data file
Const $sendKeyDelay = 3 ;; was 20 to begin with
Const $ProgramTitle = "Stein Utilities, Version " & $Version
;;Const $BUTTONS[2][5] = [["Get Letters", "New Flag", "New Flag to Recipient", "Phone Note to Recipient", "F/U Procedure Flag"], _
;;	["getLettersAction", "newFlagAction", "newFlagToRecipientAction", "phoneNoteToRecipientAction", "fuProcedureFlagAction"]]
Const $BUTTONS[2][4] = [["Get Letters", "New Flag", "New Flag to Recipient", "Phone Note to Recipient"], _
	["getLettersAction", "newFlagAction", "newFlagToRecipientAction", "phoneNoteToRecipientAction"]]
Const $BTN_DELTA = 15
Const $BTN_WIDTH = 150


;; CONSTANTS TO MATCH INI FILE SECTIONS AND KEYS
Const $ProviderSection = "provider"
Const $AssistantSection = "assistant"
Const $ApplicationSection = "application"
Const $OpenLettersSection = "open letters"
Const $NewFlagSection = "new flag"
Const $PhoneNoteSection = "phone note"

Const $ProviderLNameKey = "Last Name"
Const $ProviderFNameKey = "First Name"
Const $ProviderTitleKey = "Title"
Const $AssistantLNameKey = "Last Name"
Const $AssistantFNameKey = "First Name"
Const $AssistantTitleKey = "Title"
Const $ApplicationKey = "Application"
Const $OpenLetterKey = "Open Letter Script"
Const $NewFlagScript1Key = "New Flag Script1"
Const $NewFlagScript2Key = "New Flag Script2"
Const $PhoneNoteScript1Key = "Phone Note Script1"
Const $PhoneNoteScript2Key = "Phone Note Script2"
Const $PhoneNoteScript3Key = "Phone Note Script3"
Const $PhoneNoteScript4Key = "Phone Note Script4"

;; GLOBAL APPLICATION VARIABLES
Local $iniData[1][2] ;; holds data from ini file
Local $mainwindow ;; main window for application


;; below filled by the fillGlobalVars method
Local $app		; application name
Local $dataFileName = $DataFileNameConst
Local $providerLName	; doc's last name
Local $providerFName	; doc's first name
Local $providerTitle	; doc's title
Local $assistantLName	; rn's last name
Local $assistantFName	; rn's first name
Local $assistantTitle	; rn's title
Local $openLetterScript ; file name for open letter script
Local $newFlagScript1 ; file name for new flag script, first part
Local $newFlagScript2 ; file name for new flag script, second part
Local $phoneNoteScript1 ; file name for new phone note script, first part
Local $phoneNoteScript2 ; file name for new phone note script, second part
Local $phoneNoteScript3 ; file name for new phone note script, third part
Local $phoneNoteScript4 ; file name for new phone note script, fourth part

initializeNonGui()
initializeGui()

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; non-GUI section.  Fills variables and set us.

;; set up all non-gui
Func initializeNonGui()
	;; get the ini file section names
	If $CmdLine[0] > 0 Then ;; if cmd line parameters
		$dataFileName = $CmdLine[1]
	EndIf
	;; else default to the constant

	$sectionNames = IniReadSectionNames($dataFileName)

	If @error Then
		MsgBox(4096, "", "Error occurred, probably no INI file.")
	Else
		;; set up array to hold section names and section data, the latter in an array
		ReDim $iniData[$sectionNames[0]][2]

		For $i = 1 To $sectionNames[0]
			;; read in key value pairs in the current section. section data holds all key/val pairs in 2 d array
			$sectionData = IniReadSection($dataFileName, $sectionNames[$i])

			;; section name goes into ini data array [][0]
			$iniData[$i-1][0] = $sectionNames[$i]

			;; section data array goes into ini data array [][1]
			$iniData[$i-1][1] = $sectionData
		Next
	EndIf

	fillGlobalVars()

EndFunc

;; uses Strings from ini file sections and key
Func fillGlobalVars()
	;; main App
	$app = getSectionKeyValue($iniData, $ApplicationSection, $ApplicationKey)

	;; Provider section
	$providerLName = getSectionKeyValue($iniData, $ProviderSection, $ProviderLNameKey)
	$providerFName = getSectionKeyValue($iniData, $ProviderSection, $ProviderFNameKey)
	$providerTitle = getSectionKeyValue($iniData, $ProviderSection, $ProviderTitleKey)

	;; Assistant section
	$assistantLName = getSectionKeyValue($iniData, $AssistantSection, $AssistantLNameKey)
	$assistantFName = getSectionKeyValue($iniData, $AssistantSection, $AssistantFNameKey)
	$assistantTitle = getSectionKeyValue($iniData, $AssistantSection, $AssistantTitleKey)

	;; Get open Letters Script file name; then prepend the script directory path
	$openLetterScript = getSectionKeyValue($iniData, $OpenLettersSection, $OpenLetterKey)
	$openLetterScript = @ScriptDir & "\" & $openLetterScript

	;; Get new Flags Script file name; then prepend the script directory path
	$newFlagScript1 = getSectionKeyValue($iniData, $NewFlagSection, $NewFlagScript1Key)
	$newFlagScript2 = getSectionKeyValue($iniData, $NewFlagSection, $NewFlagScript2Key)
	$newFlagScript1 = @ScriptDir & "\" & $newFlagScript1
	$newFlagScript2 = @ScriptDir & "\" & $newFlagScript2

	;; Get new Phone Note Script file name; then prepend the script directory path
	$phoneNoteScript1 = getSectionKeyValue($iniData, $PhoneNoteSection, $PhoneNoteScript1Key)
	$phoneNoteScript2 = getSectionKeyValue($iniData, $PhoneNoteSection, $PhoneNoteScript2Key)
	$phoneNoteScript3 = getSectionKeyValue($iniData, $PhoneNoteSection, $PhoneNoteScript3Key)
	$phoneNoteScript4 = getSectionKeyValue($iniData, $PhoneNoteSection, $PhoneNoteScript4Key)
	$phoneNoteScript1 = @ScriptDir & "\" & $phoneNoteScript1
	$phoneNoteScript2 = @ScriptDir & "\" & $phoneNoteScript2
	$phoneNoteScript3 = @ScriptDir & "\" & $phoneNoteScript3
	$phoneNoteScript4 = @ScriptDir & "\" & $phoneNoteScript4

EndFunc

;; get value from key/value pair held by section
Func getSectionKeyValue($iniDataParam, $sectionNameParam, $keyString)

	;; get index section in ini data array
	$sectionIndex = _ArraySearch($iniDataParam, $sectionNameParam)

	;; get section key/value array held there. Note that [0] holds the section name
	$sectionDataLocal = $iniDataParam[$sectionIndex][1]

	;; get index of key/value row corresponding to key string
	$retValueIndex = _ArraySearch($sectionDataLocal, $keyString)

	;; get the value associated with the key. key is at [0], value at [1]
	$retValue = $sectionDataLocal[$retValueIndex][1]
	Return $retValue
EndFunc

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; GUI section.  initializes and runs GUI

Func initializeGui()
	Opt("WinTitleMatchMode", 2)  ; allows match with substring
	Opt("GUIOnEventMode", 1)  ; Set to OnEvent mode

	;; create main GUI and place at near bottom of screen
	$tempWidth = UBound($BUTTONS, 2) * ($BTN_DELTA + $BTN_WIDTH) + $BTN_DELTA
	$mainwindow = GUICreate($ProgramTitle, $tempWidth, 75, 20, @DesktopHeight - 140)

	GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClickedAction")

	$provFullName = $providerFName & " " & $providerLName & ", " & $providerTitle
	$xPos = $BTN_DELTA + 5
	$yPos = 10
	$providerLabel = GUICtrlCreateLabel("Provider: " & $provFullName, $xPos, $yPos)

	$xPos = 180
	$ctrlWidth = 250
	$assistantFullName = $assistantFName & " " & $assistantLName & ", " & $assistantTitle
	$assistantLabel = GUICtrlCreateLabel("Flag/Note Recipient: " & $assistantFullName, $xPos, $yPos, $ctrlWidth)


	;; Make buttons
	$yPos = 40
	$xPos = $BTN_DELTA
	$width = $BTN_WIDTH
	For $btnIndex = 0 to UBound($BUTTONS, 2) - 1
		$button = GUICtrlCreateButton($BUTTONS[0][$btnIndex], $xPos, $yPos, $width)
		GUICtrlSetOnEvent($button, $BUTTONS[1][$btnIndex])
		$xPos += $width + $BTN_DELTA
	Next

	;; show the GUI!
	GUISetState(@SW_SHOW)

	;; the OnEvent GUI loop
	While 1
	  Sleep(1000)  ; Idle around
	WEnd
EndFunc

;; get into letters section
Func getLettersAction()
	;;displayDataInIniData()
	myAppActivate($app)
	getCommandsAndRun($openLetterScript)
EndFunc

Func newFlagAction()
	myAppActivate($app)
	getCommandsAndRun($newFlagScript1)
EndFunc

Func newFlagToRecipientAction()
	myAppActivate($app)
	getCommandsAndRun($newFlagScript1)
	$text = $assistantLName & " " & $assistantTitle & ", " & $assistantFName
	Send($text)
	getCommandsAndRun($newFlagScript2)
EndFunc

Func phoneNoteToRecipientAction()
	GUISetState(@SW_DISABLE, $mainwindow) ;; disable main window

	$phoneNoteTopic = InputBox("Phone Note Topic", "Enter phone note topic below:", "Lab Results")
	If @error <> 0 Then
		GUISetState(@SW_ENABLE, $mainwindow) ;; enable main window
		Return
	EndIf

	myAppActivate($app)
	getCommandsAndRun($phoneNoteScript1)
	Send($phoneNoteTopic) ;; send topic into phone note topic text field
	getCommandsAndRun($phoneNoteScript2)
	$retValue = MsgBox(0x30 + 1, "Complete Phone Note", "To route phone note to Recipient, press OK when done entering text")

	If $retValue = 1 Then ;; if press OK, then send phone note to Recipient
		myAppActivate($app)
		getCommandsAndRun($phoneNoteScript3)
		$text = $assistantLName & " " & $assistantTitle & ", " & $assistantFName
		Send($text)
		getCommandsAndRun($phoneNoteScript4)
	EndIf
	GUISetState(@SW_ENABLE, $mainwindow) ;; enable main window
	myAppActivate($app)

EndFunc

Func phoneNoteFinishAction()
EndFunc

Func fuProcedureFlagAction()

EndFunc

;; exit when main window's close button pressed
Func CLOSEClickedAction()
	If @GUI_WinHandle = $mainwindow Then
		Exit
	EndIf
EndFunc

;; test that data is OK
Func displayDataInIniData()
	$sectionNames = IniReadSectionNames($dataFileName)

	For $i = 1 To $sectionNames[0]
		$foo = _ArraySearch($iniData, $sectionNames[$i])
		$fooArray = $iniData[$foo][1]
		_ArrayDisplay($fooArray, $iniData[$foo][0])
	Next
EndFunc

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; open text command file and run commands it contains.  Used by all buttons

Func getCommandsAndRun($fileName)
	$fileLines = getCommandArrayFromFile($fileName)
	;;!! _ArrayDisplay($fileLines) ;; show that commands have been read from file
	sendKeys($fileLines) ; send commands from text file
EndFunc   ;==>getCommandsAndRun

Func getCommandArrayFromFile($fileName)
	$count = _FileCountLines($fileName)
	Local $fileLines[1] ;; create an array with one empty element

	$file = FileOpen($fileName, 0)
	If @error Then
		MsgBox(4096, "Error", "Error opening file: " & $fileName)
		Exit
	EndIf

	$arrayIndex = 0
	For $i = 0 To $count - 1
		$line = FileReadLine($file)
		;;MsgBox(0, "foo", $line & "; StringLen($line): " & StringLen($line))
		If StringLen($line) > 0 Then
			If StringLeft($line, 1) <> ";" Then
				_ArrayAdd($fileLines, $line)
			EndIf
		EndIf

	Next
	_ArrayDelete($fileLines, 0) ;; get rid of first empty element
	FileClose($file)
	Return $fileLines
EndFunc   ;==>getCommandArrayFromFile

;  pass array of string commands.  If no array, then default performed
Func sendKeys($commandStringArray = False)

	;; if array passed to function
	If $commandStringArray <> False Then
		If IsArray($commandStringArray) Then
			For $commandString in $commandStringArray
				Send($commandString)
			Next
		EndIf
	EndIf
EndFunc   ;==>sendKeys

;; activate a window. exit program if fails
;; also set sendkey delay
Func myAppActivate($windowNameParam)
	;; activate my app
	;; Opt("WinTitleMatchMode", 2)  ; allows match with substring.  May already be true
	WinActivate($windowNameParam)
	$var1 = WinWaitActive($windowNameParam, "", 3)

	If $var1 = 0 Then
		MsgBox(4096, "Error", "Window could not be open: " & $windowNameParam)
		Exit
	EndIf

	Opt("SendKeyDelay", $sendKeyDelay) ; send key with a delay of $sendKeyDelay ms

EndFunc
