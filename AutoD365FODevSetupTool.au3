#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.1
 Author:         Andrej Graf

 Script Function:
	02.08.2023  - Start DEV
				- path to file
				- 4 Parameter wurden hinzugefügt
	03.08.2023  - Plan erstellt
				- Button Cancel
				- Checkbox on/off (manuell)
				- File Open, Error, wenn es kein File gibt
				- String wurde mit Array geändert
				- func GetStringFromXML - nimmt die Info von XML zwischen > und < und fügt das in Input in GUI hinzu (abhängig von linenumber)
				- func WriteStringToXML - ändert XML abhängig vom Input des Users (abhängig von linenumber)
				- die andere Parameter wurden hinzufügen
	04.08.2023 	- func GetStringFromXMLForCheckbox - nimmt die Info von XML zwischen > und < (true oder false) und macht Checkbox on/off
				- func newstringTF, ändert XML abhängig vom Checkbox (on/off in GUI)
				- clear code and comments
				- bug fix - - - 3 Stunden
				- variable wurden richtig formatiert
				- for loop musste suchen variable in XML, funtionierte nicht korrekt
				- while loop funktioniert, muss verbessert sein
	07.08.2023 - bugfix (\ - wurden entfernt in XML) - - - 3 Stunde
				- neue Function GetLineNumber
				- clear code and comments


#ce ----------------------------------------------------------------------------


Opt("MustDeclareVars", 1)

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Array.au3>
#include <File.au3>
#include <String.au3>




;~ Local $FilePath = @UserProfileDir & "\Documents\Visual Studio Dynamics 365\DynamicsDevConfig.xml" ; TODO allgemein
Local $FilePath = @UserProfileDir & "\OneDrive - Roedl Dynamics GmbH\Documents\Visual Studio Dynamics 365\DynamicsDevConfig.xml" ; ich habe OneDrive eingerichtet, so liegen jetzt meine Documents in diesem Folder
Local $FileOpen = FileOpen($FilePath)
If $FileOpen = -1 Then ; Prüfen, ob File existiert, wenn nicht, dann Error
	MsgBox(16, "Error", "Error, there is no file" & @CRLF & $FilePath)
	Exit
EndIf
FileClose($FileOpen)




;GUI
#Region ; Const für strings <>
Local $Form_D365_FO_DEV_Setup = GUICreate("Rödl Dynamics GmbH - D365 FO DEV Setup", 615, 692, 187, 140)

Const $AddProjectToExistingSolution_Const = "AddProjectToExistingSolution"
Local $CB_AddProjectToExistingSolution = GUICtrlCreateCheckbox($AddProjectToExistingSolution_Const, 16, 8, 209, 17)
GetStringFromXMLForCheckbox($AddProjectToExistingSolution_Const)

Const $AosWebsiteName_Const = "AosWebsiteName"
Local $Label_AosWebsiteName = GUICtrlCreateLabel($AosWebsiteName_Const, 16, 32, 95, 17)
Local $Input_AosWebsiteName = GUICtrlCreateInput(GetStringFromXML($AosWebsiteName_Const), 16, 48, 553, 21)

Const $ApplicationHostConfigFile_Const = "ApplicationHostConfigFile"
Local $Label_ApplicationHostConfigFile = GUICtrlCreateLabel($ApplicationHostConfigFile_Const, 16, 74, 130, 17)
Local $Input_ApplicationHostConfigFile = GUICtrlCreateInput(GetStringFromXML($ApplicationHostConfigFile_Const), 16, 90, 553, 21)

Const $BuildModulesInParallel_Const = "BuildModulesInParallel"
Local $CB_BuildModulesInParallel = GUICtrlCreateCheckbox($BuildModulesInParallel_Const, 16, 116, 209, 17)
GetStringFromXMLForCheckbox($BuildModulesInParallel_Const)

Const $BusinessDatabaseName_Const = "BusinessDatabaseName"
Local $Label_BusinessDatabaseName = GUICtrlCreateLabel($BusinessDatabaseName_Const, 16, 140, 132, 17)
Local $Input_BusinessDatabaseName = GUICtrlCreateInput(GetStringFromXML($BusinessDatabaseName_Const), 16, 156, 553, 21)

Const $DBSyncInBuild_Const = "DBSyncInBuild"
Local $CB_DBSyncInBuild = GUICtrlCreateCheckbox($DBSyncInBuild_Const, 16, 182, 209, 17)
GetStringFromXMLForCheckbox($DBSyncInBuild_Const)

Const $DatabaseServer_Const = "DatabaseServer"
Local $Label_DatabaseServer = GUICtrlCreateLabel($DatabaseServer_Const, 16, 206, 93, 17)
Local $Input_DatabaseServer = GUICtrlCreateInput(GetStringFromXML($DatabaseServer_Const), 16, 222, 553, 21)

Const $DefaultModelForNewProjects_Const = "DefaultModelForNewProjects"
Local $Label_DefaultModelForNewProjects = GUICtrlCreateLabel($DefaultModelForNewProjects_Const, 16, 248, 154, 17)
Local $Input_DefaultModelForNewProjects = GUICtrlCreateInput(GetStringFromXML($DefaultModelForNewProjects_Const), 16, 264, 553, 21)

Const $DefaultWebBrowser_Const = "DefaultWebBrowser"
Local $Label_DefaultWebBrowser = GUICtrlCreateLabel($DefaultWebBrowser_Const, 16, 290, 111, 17)
Local $Input_DefaultWebBrowser = GUICtrlCreateInput(GetStringFromXML($DefaultWebBrowser_Const), 16, 306, 553, 21)

Const $DisableBPCheck_Const = "DisableBPCheck"
Local $CB_DisableBPCheck = GUICtrlCreateCheckbox($DisableBPCheck_Const, 16, 332, 209, 17)
GetStringFromXMLForCheckbox($DisableBPCheck_Const)

Const $DisableFormStaticCompile_Const = "DisableFormStaticCompile"
Local $CB_DisableFormStaticCompile = GUICtrlCreateCheckbox($DisableFormStaticCompile_Const, 16, 356, 209, 17)
GetStringFromXMLForCheckbox($DisableFormStaticCompile_Const)

Const $EmitTraceEvents_Const = "EmitTraceEvents"
Local $CB_EmitTraceEvents = GUICtrlCreateCheckbox($EmitTraceEvents_Const, 16, 380, 209, 17)
GetStringFromXMLForCheckbox($EmitTraceEvents_Const)

Const $EnableNativeDebugging_Const = "EnableNativeDebugging"
Local $CB_EnableNativeDebugging = GUICtrlCreateCheckbox($EnableNativeDebugging_Const, 16, 404, 209, 17)
GetStringFromXMLForCheckbox($EnableNativeDebugging_Const)

Const $EnableOfflineAuthentication_Const = "EnableOfflineAuthentication"
Local $CB_EnableOfflineAuthentication = GUICtrlCreateCheckbox($EnableOfflineAuthentication_Const, 16, 428, 209, 17)
GetStringFromXMLForCheckbox($EnableOfflineAuthentication_Const)

Const $EnableSymbolLoadingForSolutionOnly_Const = "EnableSymbolLoadingForSolutionOnly"
Local $CB_EnableSymbolLoadingForSolutionOnly = GUICtrlCreateCheckbox($EnableSymbolLoadingForSolutionOnly_Const, 16, 452, 209, 17)
GetStringFromXMLForCheckbox($EnableSymbolLoadingForSolutionOnly_Const)

Const $FallbackToNativeSync_Const = "FallbackToNativeSync"
Local $CB_FallbackToNativeSync = GUICtrlCreateCheckbox($FallbackToNativeSync_Const, 16, 476, 209, 17)
GetStringFromXMLForCheckbox($FallbackToNativeSync_Const)

Const $OrganizeElementsInProject_Const = "OrganizeElementsInProject"
Local $CB_OrganizeElementsInProject = GUICtrlCreateCheckbox($OrganizeElementsInProject_Const, 16, 500, 209, 17)
GetStringFromXMLForCheckbox($OrganizeElementsInProject_Const)

Local $ButtonOk = GUICtrlCreateButton("OK", 368, 648, 100, 25)
Local $ButtonClose = GUICtrlCreateButton("Cancel", 487, 648, 100, 25)
GUISetState(@SW_SHOW)
#EndRegion




While 1
	Local $nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

		Case $ButtonOk
			WriteStringToXMLForCheckbox($AddProjectToExistingSolution_Const, $CB_AddProjectToExistingSolution)
			WriteStringToXML($AosWebsiteName_Const, $Input_AosWebsiteName)
			WriteStringToXML($ApplicationHostConfigFile_Const, $Input_ApplicationHostConfigFile)
			WriteStringToXMLForCheckbox($BuildModulesInParallel_Const, $CB_BuildModulesInParallel)
			WriteStringToXML($BusinessDatabaseName_Const, $Input_BusinessDatabaseName)
			WriteStringToXMLForCheckbox($DBSyncInBuild_Const, $CB_DBSyncInBuild)
			WriteStringToXML($DatabaseServer_Const, $Input_DatabaseServer)
			WriteStringToXML($DefaultModelForNewProjects_Const, $Input_DefaultModelForNewProjects)
			WriteStringToXML($DefaultWebBrowser_Const, $Input_DefaultWebBrowser)
			WriteStringToXMLForCheckbox($DisableBPCheck_Const, $CB_DisableBPCheck)
			WriteStringToXMLForCheckbox($DisableFormStaticCompile_Const, $CB_DisableFormStaticCompile)
			WriteStringToXMLForCheckbox($EmitTraceEvents_Const, $CB_EmitTraceEvents)
			WriteStringToXMLForCheckbox($EnableNativeDebugging_Const, $CB_EnableNativeDebugging)
			WriteStringToXMLForCheckbox($EnableOfflineAuthentication_Const, $CB_EnableOfflineAuthentication)
			WriteStringToXMLForCheckbox($EnableSymbolLoadingForSolutionOnly_Const, $CB_EnableSymbolLoadingForSolutionOnly)
			WriteStringToXMLForCheckbox($FallbackToNativeSync_Const, $CB_FallbackToNativeSync)
			WriteStringToXMLForCheckbox($OrganizeElementsInProject_Const, $CB_OrganizeElementsInProject)

		Case $ButtonClose
			Exit
	EndSwitch
WEnd




Func GetLineNumber($SearchText) ;  man muss den Text zwischen < und > hinzufügen, am Ende bekommt man die Nummer von Linien, wo das steht
	Local $FileStrings = FileReadToArray($FilePath)
	Local $i = 0
	Global $linenumber
	While $i < UBound($FileStrings)
		If StringInStr($FileStrings[$i], "<" & $SearchText & ">") Then
			$linenumber = $i + 1
			ExitLoop
		EndIf
		$i += 1
	WEnd
	ConsoleWrite($linenumber & @CRLF)
	Return $linenumber
EndFunc




Func GetStringFromXML($SearchText) ; nimmt info von XML, was steht zwischen > und <, und fügt in GUI hinzu
	GetLineNumber($SearchText)
	Local $FileStrings = FileReadToArray($FilePath)
	Local $result = StringRegExpReplace($FileStrings[$linenumber - 1], ".*>(.*?)<.*", "$1")
	ConsoleWrite($result & @CRLF)
	Return $result
EndFunc




Func GetStringFromXMLForCheckbox($SearchText) ; nimmt info von XML, macht Checkbox on/off, wenn in XML true/false
	GetLineNumber($SearchText)
	Local $FileStrings = FileReadToArray($FilePath)
	Local $result = StringRegExpReplace($FileStrings[$linenumber - 1], ".*>(.*?)<.*", "$1")
	ConsoleWrite($result & @CRLF)
	If $result = "true" Then
		GUICtrlSetState(-1, $GUI_CHECKED)
	ElseIf $result = "false" Then
		GUICtrlSetState(-1, Not $GUI_CHECKED)
	Else
		MsgBox(16, "Error", "Please, proof XML-File, the string number " & $linenumber & " should contain only true or false")
	EndIf
EndFunc




Func WriteStringToXML($SearchText, $Input) ; schreibt Input von User in XML
	GetLineNumber($SearchText)
	Local $FileStrings = FileReadToArray($FilePath)
	Local $InputSpecChar
	$InputSpecChar = GUICtrlRead($Input)
	$InputSpecChar = StringRegExpReplace($InputSpecChar, "[.*?^${}()|[\]\\]", "\\$0") ; ohne das wurden spec character entfernt
	$FileStrings[$linenumber - 1] = StringRegExpReplace($FileStrings[$linenumber - 1], "(?<=>).*?(?=<)" , $InputSpecChar)
 	_FileWriteFromArray($FilePath, $FileStrings) ;changes file
	ConsoleWrite($FileStrings[$linenumber - 1] & @CRLF)
EndFunc




Func WriteStringToXMLForCheckbox($SearchText, $Checkbox) ; schreibt true/false in XML, wenn Checkbox von User on/off gestellt wurde
	GetLineNumber($SearchText)
	Local $FileStrings = FileReadToArray($FilePath)
	Local $statusOfCheckbox = GUICtrlRead($Checkbox)
	If $statusOfCheckbox = 1 Then
		$FileStrings[$linenumber - 1] = StringRegExpReplace($FileStrings[$linenumber - 1], "(?<=>).*(?=<)" , "true")
		ConsoleWrite($FileStrings[$linenumber - 1] & @CRLF)
 		_FileWriteFromArray($FilePath, $FileStrings) ;changes file
	ElseIf $statusOfCheckbox = 4 Then
		$FileStrings[$linenumber - 1] = StringRegExpReplace($FileStrings[$linenumber - 1], "(?<=>).*(?=<)" , "false")
		ConsoleWrite($FileStrings[$linenumber - 1] & @CRLF)
 		_FileWriteFromArray($FilePath, $FileStrings) ;changes file
	Else
		MsgBox(16, "Error", "Unknown error")
	EndIf
EndFunc