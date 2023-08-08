#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.1
 Author:         Andrej Graf

 Script Function:
	02.08.2023  - Start DEV
				- path to file
				- 4 Parameter added
	03.08.2023  - To Do list
				- Button Cancel
				- Checkbox on/off
				- File Open, error, if there is no file
				- Change strings by arrays
				- function _GetStringFromXML - get info from XML between > and <  and set in input in GUI (depend on linenumber)
				- function _WriteStringToXML - write input from user in XML (depend on linenumber)
				- the rest of parameters added
	04.08.2023 	- function _GetStringFromXMLForCheckbox - get info from XML between > and <  (true or false) and check the box (on or off) (depend on linenumber)
				- function _WriteStringToXMLForCheckbox - write in XML true or false, if checkbox is on or off (depend on linenumber)
				- clear code and comments
				- format of variables
	07.08.2023  - bugfix (special characters (\) were deleted in XML)
				- function _GetLineNumber - search parameter between < and >, return linenumber
				- _GetLineNumber was added in all of 4 functions. Now functions aren´t depending on linenumber
				- constants for each input and checkbox with a name of parameter (between < and >)
	08.08.2023	- add if statement with error message in _GetLineNumber
				- clear code and comments

				- Button Ok safe file and now close
				- add icon
				- add _ for my functions
				- GUICtrlSetState in GetStringFromXMLForCheckbox is now with name of checkbox, earlier war with "-1" (added second parameter in this function)
				- func _WriteStringToXMLForCombo




#ce ----------------------------------------------------------------------------


Opt("MustDeclareVars", 1)
#AutoIt3Wrapper_icon=zahnrad.ico


#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Array.au3>
#include <File.au3>
#include <String.au3>




Local $FilePath = @MyDocumentsDir & "\Visual Studio Dynamics 365\DynamicsDevConfig.xml"
Local $FileOpen = FileOpen($FilePath)
If $FileOpen = -1 Then ; check, if the file exists. If not, then error
	MsgBox(16, "Error", "Error, there is no file" & @CRLF & $FilePath)
	Exit
EndIf
FileClose($FileOpen)




;GUI. For each input and checkbox there are the constants with the name of parameter. Constants are used in each function.
;Variables of checkboxes are named with "CB_" at the beginning. For input there are two variables: "Label_..." - the name of input, "Input_..." - box with input.
#Region
Local $Form_D365_FO_DEV_Setup = GUICreate("Rödl Dynamics GmbH - D365 FO DEV Setup", 615, 692, 187, 140)

Const $AddProjectToExistingSolution_Const = "AddProjectToExistingSolution"
Local $CB_AddProjectToExistingSolution = GUICtrlCreateCheckbox($AddProjectToExistingSolution_Const, 16, 8, 209, 17)
_GetStringFromXMLForCheckbox($AddProjectToExistingSolution_Const, $CB_AddProjectToExistingSolution)

Const $AosWebsiteName_Const = "AosWebsiteName"
Local $Label_AosWebsiteName = GUICtrlCreateLabel($AosWebsiteName_Const, 16, 32, 95, 17)
Local $Input_AosWebsiteName = GUICtrlCreateInput(_GetStringFromXML($AosWebsiteName_Const), 16, 48, 553, 21)

Const $ApplicationHostConfigFile_Const = "ApplicationHostConfigFile"
Local $Label_ApplicationHostConfigFile = GUICtrlCreateLabel($ApplicationHostConfigFile_Const, 16, 74, 130, 17)
Local $Input_ApplicationHostConfigFile = GUICtrlCreateInput(_GetStringFromXML($ApplicationHostConfigFile_Const), 16, 90, 553, 21)

Const $BuildModulesInParallel_Const = "BuildModulesInParallel"
Local $CB_BuildModulesInParallel = GUICtrlCreateCheckbox($BuildModulesInParallel_Const, 16, 116, 209, 17)
_GetStringFromXMLForCheckbox($BuildModulesInParallel_Const, $CB_BuildModulesInParallel)

Const $BusinessDatabaseName_Const = "BusinessDatabaseName"
Local $Label_BusinessDatabaseName = GUICtrlCreateLabel($BusinessDatabaseName_Const, 16, 140, 132, 17)
Local $Input_BusinessDatabaseName = GUICtrlCreateInput(_GetStringFromXML($BusinessDatabaseName_Const), 16, 156, 553, 21)

Const $DBSyncInBuild_Const = "DBSyncInBuild"
Local $CB_DBSyncInBuild = GUICtrlCreateCheckbox($DBSyncInBuild_Const, 16, 182, 209, 17)
_GetStringFromXMLForCheckbox($DBSyncInBuild_Const, $CB_DBSyncInBuild)

Const $DatabaseServer_Const = "DatabaseServer"
Local $Label_DatabaseServer = GUICtrlCreateLabel($DatabaseServer_Const, 16, 206, 93, 17)
Local $Input_DatabaseServer = GUICtrlCreateInput(_GetStringFromXML($DatabaseServer_Const), 16, 222, 553, 21)

Const $DefaultModelForNewProjects_Const = "DefaultModelForNewProjects"
Local $Label_DefaultModelForNewProjects = GUICtrlCreateLabel($DefaultModelForNewProjects_Const, 16, 248, 154, 17)
Local $Combo_DefaultModelForNewProjects = GUICtrlCreateCombo(_GetStringFromXML($DefaultModelForNewProjects_Const), 16, 264, 553, 21)
;~ Local $Path_DefaultModelForNewProjects = "K:\AosService\PackagesLocalDirectory"; path to folder
Local $Path_DefaultModelForNewProjects = @MyDocumentsDir & "\test\" ; path to folder
Local $Array_DefaultModelForNewProjects = _FileListToArray($Path_DefaultModelForNewProjects, Default, $FLTA_FOLDERS) ; the foldernames from folder are to combo
For $i = 1 To $Array_DefaultModelForNewProjects[0]
	GUICtrlSetData($Combo_DefaultModelForNewProjects, $Array_DefaultModelForNewProjects[$i])
Next

Const $DefaultWebBrowser_Const = "DefaultWebBrowser"
Local $Label_DefaultWebBrowser = GUICtrlCreateLabel($DefaultWebBrowser_Const, 16, 290, 111, 17)
Local $Input_DefaultWebBrowser = GUICtrlCreateInput(_GetStringFromXML($DefaultWebBrowser_Const), 16, 306, 553, 21)

Const $DisableBPCheck_Const = "DisableBPCheck"
Local $CB_DisableBPCheck = GUICtrlCreateCheckbox($DisableBPCheck_Const, 16, 332, 209, 17)
_GetStringFromXMLForCheckbox($DisableBPCheck_Const, $CB_DisableBPCheck)

Const $DisableFormStaticCompile_Const = "DisableFormStaticCompile"
Local $CB_DisableFormStaticCompile = GUICtrlCreateCheckbox($DisableFormStaticCompile_Const, 16, 356, 209, 17)
_GetStringFromXMLForCheckbox($DisableFormStaticCompile_Const, $CB_DisableFormStaticCompile)

Const $EmitTraceEvents_Const = "EmitTraceEvents"
Local $CB_EmitTraceEvents = GUICtrlCreateCheckbox($EmitTraceEvents_Const, 16, 380, 209, 17)
_GetStringFromXMLForCheckbox($EmitTraceEvents_Const, $CB_EmitTraceEvents)

Const $EnableNativeDebugging_Const = "EnableNativeDebugging"
Local $CB_EnableNativeDebugging = GUICtrlCreateCheckbox($EnableNativeDebugging_Const, 16, 404, 209, 17)
_GetStringFromXMLForCheckbox($EnableNativeDebugging_Const, $CB_EnableNativeDebugging)

Const $EnableOfflineAuthentication_Const = "EnableOfflineAuthentication"
Local $CB_EnableOfflineAuthentication = GUICtrlCreateCheckbox($EnableOfflineAuthentication_Const, 16, 428, 209, 17)
_GetStringFromXMLForCheckbox($EnableOfflineAuthentication_Const, $CB_EnableOfflineAuthentication)

Const $EnableSymbolLoadingForSolutionOnly_Const = "EnableSymbolLoadingForSolutionOnly"
Local $CB_EnableSymbolLoadingForSolutionOnly = GUICtrlCreateCheckbox($EnableSymbolLoadingForSolutionOnly_Const, 16, 452, 209, 17)
_GetStringFromXMLForCheckbox($EnableSymbolLoadingForSolutionOnly_Const, $CB_EnableSymbolLoadingForSolutionOnly)

Const $FallbackToNativeSync_Const = "FallbackToNativeSync"
Local $CB_FallbackToNativeSync = GUICtrlCreateCheckbox($FallbackToNativeSync_Const, 16, 476, 209, 17)
_GetStringFromXMLForCheckbox($FallbackToNativeSync_Const, $CB_FallbackToNativeSync)

Const $OrganizeElementsInProject_Const = "OrganizeElementsInProject"
Local $CB_OrganizeElementsInProject = GUICtrlCreateCheckbox($OrganizeElementsInProject_Const, 16, 500, 209, 17)
_GetStringFromXMLForCheckbox($OrganizeElementsInProject_Const, $CB_OrganizeElementsInProject)

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
			_WriteStringToXMLForCheckbox($AddProjectToExistingSolution_Const, $CB_AddProjectToExistingSolution)
			_WriteStringToXML($AosWebsiteName_Const, $Input_AosWebsiteName)
			_WriteStringToXML($ApplicationHostConfigFile_Const, $Input_ApplicationHostConfigFile)
			_WriteStringToXMLForCheckbox($BuildModulesInParallel_Const, $CB_BuildModulesInParallel)
			_WriteStringToXML($BusinessDatabaseName_Const, $Input_BusinessDatabaseName)
			_WriteStringToXMLForCheckbox($DBSyncInBuild_Const, $CB_DBSyncInBuild)
			_WriteStringToXML($DatabaseServer_Const, $Input_DatabaseServer)
			_WriteStringToXMLForCombo($DefaultModelForNewProjects_Const, $Combo_DefaultModelForNewProjects)
			_WriteStringToXML($DefaultWebBrowser_Const, $Input_DefaultWebBrowser)
			_WriteStringToXMLForCheckbox($DisableBPCheck_Const, $CB_DisableBPCheck)
			_WriteStringToXMLForCheckbox($DisableFormStaticCompile_Const, $CB_DisableFormStaticCompile)
			_WriteStringToXMLForCheckbox($EmitTraceEvents_Const, $CB_EmitTraceEvents)
			_WriteStringToXMLForCheckbox($EnableNativeDebugging_Const, $CB_EnableNativeDebugging)
			_WriteStringToXMLForCheckbox($EnableOfflineAuthentication_Const, $CB_EnableOfflineAuthentication)
			_WriteStringToXMLForCheckbox($EnableSymbolLoadingForSolutionOnly_Const, $CB_EnableSymbolLoadingForSolutionOnly)
			_WriteStringToXMLForCheckbox($FallbackToNativeSync_Const, $CB_FallbackToNativeSync)
			_WriteStringToXMLForCheckbox($OrganizeElementsInProject_Const, $CB_OrganizeElementsInProject)
			Exit

		Case $ButtonClose
			Exit
	EndSwitch
WEnd




; Function takes a name of parameter (between < and >), searches that in XML, returns the linenumber of parameter. If there is no parameter with the name of constant, then error and exit
Func _GetLineNumber($SearchText)
	Local $FileStrings = FileReadToArray($FilePath)
	Local $i = 0
	Local $linenumber
	While $i < UBound($FileStrings)
		If StringInStr($FileStrings[$i], "<" & $SearchText & ">") Then
			$linenumber = $i + 1
			ExitLoop
		EndIf
		$i += 1
	WEnd

	If $i = UBound($FileStrings) Then
		MsgBox(16, "Error", "The parameter """ & $SearchText & """ is not found")
		Exit
	EndIf
;~ 	ConsoleWrite($linenumber & @CRLF)
	Return $linenumber
EndFunc




; Function takes info from XML, what is between > and < and adds it in GUI. It depends on name of parameter (between < and >, by function _GetLineNumber)
Func _GetStringFromXML($SearchText)
	Local $linenumber = _GetLineNumber($SearchText)
	Local $FileStrings = FileReadToArray($FilePath)
	Local $result = StringRegExpReplace($FileStrings[$linenumber - 1], ".*>(.*?)<.*", "$1")
;~ 	ConsoleWrite($result & @CRLF)
	Return $result
EndFunc




; Function takes info from XML, what is between > and < (true or false) and check the box in GUI on or off. It depends on name of parameter (between < and >, by function _GetLineNumber)
; The function is only for checkboxes. If there is not true and not false between > and <, then error
Func _GetStringFromXMLForCheckbox($SearchText, $Checkbox)
	Local $linenumber = _GetLineNumber($SearchText)
	Local $FileStrings = FileReadToArray($FilePath)
	Local $result = StringRegExpReplace($FileStrings[$linenumber - 1], ".*>(.*?)<.*", "$1")
;~ 	ConsoleWrite($result & @CRLF)
	If $result = "true" Then
		GUICtrlSetState($Checkbox, $GUI_CHECKED)
	ElseIf $result = "false" Then
		GUICtrlSetState($Checkbox, Not $GUI_CHECKED)
	Else
		MsgBox(16, "Error", "Please, proof XML-File, the string number """ & $linenumber & """ should contain only true or false")
	EndIf
EndFunc




;Function takes the input from user or the info from input in GUI, if user didn´t change the inputbox, then escape the special characters, then adds the input in XML between > and < and saves the file
;It depends on name of parameter (between < and >, by function _GetLineNumber)
Func _WriteStringToXML($SearchText, $Input)
	Local $linenumber = _GetLineNumber($SearchText)
	Local $FileStrings = FileReadToArray($FilePath)
	Local $InputSpecChar = GUICtrlRead($Input)
	$InputSpecChar = StringRegExpReplace($InputSpecChar, "[.*?^${}()|[\]\\]", "\\$0")
	$FileStrings[$linenumber - 1] = StringRegExpReplace($FileStrings[$linenumber - 1], "(?<=>).*?(?=<)" , $InputSpecChar)
 	_FileWriteFromArray($FilePath, $FileStrings)
;~ 	ConsoleWrite($FileStrings[$linenumber - 1] & @CRLF)
EndFunc




;Function takes the poistion of checkbox (on/off). User could it change, or not change. If checkbox is on, than function writes "true" in XML between > and < and changes the file. If checkbox is off, than write "false".
;It depends on name of parameter (between < and >, by function _GetLineNumber)
Func _WriteStringToXMLForCheckbox($SearchText, $Checkbox)
	Local $linenumber = _GetLineNumber($SearchText)
	Local $FileStrings = FileReadToArray($FilePath)
	Local $statusOfCheckbox = GUICtrlRead($Checkbox)
	If $statusOfCheckbox = 1 Then
		$FileStrings[$linenumber - 1] = StringRegExpReplace($FileStrings[$linenumber - 1], "(?<=>).*(?=<)" , "true")
;~ 		ConsoleWrite($FileStrings[$linenumber - 1] & @CRLF)
 		_FileWriteFromArray($FilePath, $FileStrings)
	ElseIf $statusOfCheckbox = 4 Then
		$FileStrings[$linenumber - 1] = StringRegExpReplace($FileStrings[$linenumber - 1], "(?<=>).*(?=<)" , "false")
;~ 		ConsoleWrite($FileStrings[$linenumber - 1] & @CRLF)
 		_FileWriteFromArray($FilePath, $FileStrings)
	EndIf
EndFunc




; Function takes the input from user or selected element of combo, adds in XML between > and < and saves the file
;It depends on name of parameter (between < and >, by function _GetLineNumber)
Func _WriteStringToXMLForCombo($SearchText, $Combo)
	Local $linenumber = _GetLineNumber($SearchText)
	Local $FileStrings = FileReadToArray($FilePath)
	$FileStrings[$linenumber - 1] = StringRegExpReplace($FileStrings[$linenumber - 1], "(?<=>).*?(?=<)" , GUICtrlRead($Combo))
	_FileWriteFromArray($FilePath, $FileStrings)
EndFunc