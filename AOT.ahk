EnvGet, USERDOMAIN, USERDOMAIN

global cmtraceContext := USERDOMAIN . "\" . A_UserName	
global cmtraceWarning = 2
global cmtraceError = 3
global cmtraceVerbose = 4
global cmtraceDebug = 5
global cmtraceInformation = 6
global cmtraceThread := DllCall("GetCurrentProcessId")
global cmtraceFile := A_ScriptFullPath

global aotLogPath := A_Temp . "\AOT.LOG"
global aotLogRotationPath := A_Temp . "\AOT.LO_"
global aotIconPath := A_ScriptDir . "\AOT.ico"

; Programs that run in the background, but are not visible to the user.
global ExcludedClasses := Array("Desktop User Picture","Shell_TrayWnd","DV2ControlHost","Progman","Button","LyncConversationWindowClass")

msgboxButtonsOK := 0
msgboxButtonsYesNo := 4
msgboxIconError := 16
msgboxIconQuestion := 32
msgboxIconExclamation := 48
msgboxIconInformation := 64
msgboxBotton2Default := 256
msgboxSytleAlwaysOnTop := 262144

wmi := ComObjGet("winmgmts:")
queryEnum := wmi.ExecQuery("Select * From Win32_Process WHERE ProcessId=" . cmtraceThread)._NewEnum()
if queryEnum[process] {
	global cmtraceComponent := process.Name
}

GoSub, RotateLog
WriteCmtraceEntry("AlwaysOnTop initialized", cmtraceInformation)

Menu, Tray, Icon, %aotIconPath%
^SPACE:: 
{
	GoSub, RotateLog
	WriteCmtraceEntry("CTRL+SPACE detected", cmtraceInformation)
	
	caller_id := WinExist("A")
	WinGetTitle, caller_title, ahk_id %caller_id%
	WinGetClass, caller_class, ahk_id %caller_id%
	WriteCmtraceEntry("Caller: '" . caller_title . "'`r`nID: " . caller_id . "'`r`nClass: " . caller_class, cmtraceDebug)
	
	ObjectsToSet := Object()
	ObjectsToClear := Object()
	
	; Loop through all open windows
	WinGet WindowList, List
	Loop %WindowList%
	{
		this_Action := "Scan"
		this_id := WindowList%A_Index%
		WinGet, this_PID, PID, ahk_id %this_id%
		WinGetTitle, this_title, ahk_id %this_id%
		WinGetClass, this_class, ahk_id %this_id%
		WinGet, this_ExStyle, ExStyle, ahk_id %this_id%
		; 0x8 is WS_EX_TOPMOST.
		if (this_ExStyle & 0x8) {
			this_TopMost := true
		} else {
			this_TopMost := false
		}
		this_ExcludedClass := JokeHasKey(ExcludedClasses, this_class)
		this_CanSet := true
		this_CanClear := true
		if (StrLen(Trim(this_title)) = 0) {
			this_CanSet := false
			this_CanClear := false
		}
		if (this_ExcludedClass = true) {
			this_CanSet := false
			this_CanClear := false
		}
		this_Selected := (this_id = caller_id)
		WriteCmtraceEvent(this_title, this_PID, this_id, this_Class, this_Action, this_ExcludedClass, this_Selected, this_CanSet, this_CanClear, this_TopMost, cmtraceInformation)
		if (this_TopMost = false and this_Selected = true and this_CanSet = true) {
			ObjectsToSet.Push(this_id)
		}
		if (this_TopMost = true and this_CanClear = true) {
			ObjectsToClear.Push(this_id)
		}
	}
	WriteCmtraceEntry(ObjectsToSet.Length() . " available to Set.", cmtraceInformation)
	WriteCmtraceEntry(ObjectsToClear.Length() . " available to Clear.", cmtraceInformation)
	if (ObjectsToClear.Length() > 0) {
		for index, element in ObjectsToClear
		{
			this_id := ObjectsToClear[index]
			WinGet, this_PID, PID, ahk_id %this_id%
			WinGetClass, this_class, ahk_id %this_id%
			WinGetTitle, this_title, ahk_id %this_id%
			MsgBox, % (msgboxButtonsYesNo + msgboxIconQuestion + msgboxSytleAlwaysOnTop + msgboxBotton2Default), AlwaysOnTop, Would you like to clear AlwaysOnTop for '%this_title%'?
			IfMsgBox Yes
			{
				this_Action := "Clear"
				WriteCmtraceEntry("User chose to clear '" . this_title . "' WS_EX_TOPMOST", cmtraceInformation)
				Loop 5
				{
					Winset, Alwaysontop,, ahk_id %this_id%
					WinGet, this_ExStyle, ExStyle, ahk_id ahk_id %this_id%
				} Until NOT (this_ExStyle & 0x8)
				if (this_ExStyle & 0x8) {
					this_TopMost := true
				} else {
					this_TopMost := false
				}
				this_ExcludedClass := JokeHasKey(ExcludedClasses, this_class)
				this_CanSet := true
				this_CanClear := true
				if (StrLen(Trim(this_title)) = 0) {
					this_CanSet := false
					this_CanClear := false
				}
				if (this_ExcludedClass = true) {
					this_CanSet := false
					this_CanClear := false
				}
				this_Selected := (this_id = caller_id)
				WriteCmtraceEvent(this_title, this_PID, this_id, this_Class, this_Action, this_ExcludedClass, this_Selected, this_CanSet, this_CanClear, this_TopMost, cmtraceInformation)
				if (this_TopMost = false) {
					WriteCmtraceEntry("WS_EX_TOPMOST cleared on '" . this_title . "'", cmtraceInformation)
					MsgBox, % (msgboxButtonsOK + msgboxIconInformation + msgboxSytleAlwaysOnTop), AlwaysOnTop, AlwaysOnTop successfully cleared.
				} else {
					WriteCmtraceEntry("WS_EX_TOPMOST could not be cleared on '" . this_title . "'", cmtraceError)
					MsgBox, % (msgboxButtonsOK + msgboxIconError + msgboxSytleAlwaysOnTop), AlwaysOnTop, Unable to clear AlwaysOnTop.
				}
			} else {
				WriteCmtraceEntry("User chose not to clear '" . this_title . "' WS_EX_TOPMOST", cmtraceInformation)
				MsgBox, % (msgboxButtonsOK + msgboxIconInformation + msgboxSytleAlwaysOnTop), AlwaysOnTop, You chose not to clear '%this_title%' to AlwaysOnTop.
			}
		}
	} else if (ObjectsToSet.Length() = 1) {
		this_id := ObjectsToSet[1]
		WinGet, this_PID, PID, ahk_id %this_id%
		WinGetClass, this_class, ahk_id %this_id%
		WinGetTitle, this_title, ahk_id %this_id%
		MsgBox, % (msgboxButtonsYesNo + msgboxIconQuestion + msgboxSytleAlwaysOnTop + msgboxBotton2Default), AlwaysOnTop, Would you like to set AlwaysOnTop for '%this_title%'?
		IfMsgBox Yes
		{
			this_Action := "Set"
			WriteCmtraceEntry("User chose to set '" . this_title . "' WS_EX_TOPMOST", cmtraceInformation)
			Loop 5
			{
				Winset, Alwaysontop,, ahk_id %this_id%
				WinGet, this_ExStyle, ExStyle, ahk_id %this_id%
			} Until (this_ExStyle & 0x8)
			if (this_ExStyle & 0x8) {
				this_TopMost := true
			} else {
				this_TopMost := false
			}
			this_ExcludedClass := JokeHasKey(ExcludedClasses, this_class)
			this_CanSet := true
			this_CanClear := true
			if (StrLen(Trim(this_title)) = 0) {
				this_CanSet := false
				this_CanClear := false
			}
			if (this_ExcludedClass = true) {
				this_CanSet := false
				this_CanClear := false
			}
			this_Selected := (this_id = caller_id)
			WriteCmtraceEvent(this_title, this_PID, this_id, this_Class, this_Action, this_ExcludedClass, this_Selected, this_CanSet, this_CanClear, this_TopMost, cmtraceInformation)
			if (this_TopMost = true) {
				WriteCmtraceEntry("WS_EX_TOPMOST set on '" . this_title . "'", cmtraceInformation)
				MsgBox, % (msgboxButtonsOK + msgboxIconInformation + msgboxSytleAlwaysOnTop), AlwaysOnTop, AlwaysOnTop successfully set.
			} else {
				this_Context := GetProcessOwner(this_PID)
				if (this_Context <> cmtraceContext) {
					WriteCmtraceEntry("Alternate credentials detected.`r`n`r`n'" . this_title . "' is running as (" . this_Context . ") and cannot be set to AlwaysOnTop for " . cmtraceContext . ".", cmtraceWarning)
					MsgBox, % (msgboxButtonsOK + msgboxIconError + msgboxSytleAlwaysOnTop), AlwaysOnTop, % "Alternate credentials detected.`r`n`r`n'" . this_title . "' is running as (" . this_Context . ") and cannot be set to AlwaysOnTop for " . cmtraceContext . "."
				} else {
					WriteCmtraceEntry("An unknown error occurred attempting to set WS_EX_TOPMOST on '" . this_title . "'", cmtraceError)
					MsgBox, % (msgboxButtonsOK + msgboxIconError + msgboxSytleAlwaysOnTop), AlwaysOnTop, % "An unknown error occurred attempting to set WS_EX_TOPMOST on '" . this_title . "'"
				}
			}
		} else {
			WriteCmtraceEntry("User chose not to set '" . this_title . "' WS_EX_TOPMOST", cmtraceInformation)
			MsgBox, % (msgboxButtonsOK + msgboxIconInformation + msgboxSytleAlwaysOnTop), AlwaysOnTop, You chose not to set '%this_title%' to AlwaysOnTop.
		}
	} else {
		WriteCmtraceEntry("No items process.", cmtraceInformation)
	}
return
}

RotateLog:
	; Log rotation
	logRotationFileSize := 20
	logRotationUnits := "M"
	FileGetSize, aotLogFileSize, %aotLogPath%, %logRotationUnits%
	if (aotLogFileSize >= logRotationFileSize) {
		FileMove, %aotLogPath%, %aotLogRotationPath% , 1
		WriteCmtraceEntry("Log file has been rotated. (logfile > " . logRotationFileSize . logRotationUnits . ")", cmtraceInformation)
	}
	return

WriteCmtraceEvent(sTitle, iPID, iID, sClass, sAction, bExcludedClass, bSelected, bCanSet, bCanClear, bTopMost, cmtraceEntryType) {
	WriteCmtraceEntry(""
		. "Title: " . sTitle
		. "`r`nPID: " . iPID
		. "`r`nID: " . iId
		. "`r`nContext: " . sContext
		. "`r`nClass: " . sClass
		. "`r`nAction: " . sAction
		. "`r`nExcluded Class: " . bExcludedClass
		. "`r`nSelected: " . bSelected
		. "`r`nCan Set: " . bCanSet
		. "`r`nCan Clear: " . bCanClear
		. "`r`nWS_EX_TOPMOST: " . bTopMost
		, cmtraceEntryType)
}

GetProcessOwner(ProcessId) {
	wmi := ComObjGet("winmgmts:")
	queryEnum := wmi.ExecQuery("Select * from Win32_Process WHERE ProcessId = " . ProcessId)._NewEnum()
	if queryEnum[process] {
		process.GetOwner(user, domain)
		VarSetCapacity(var1, 24, 0), user := ComObject(0x400C, &var1)  ; Requires v1.1.17+
		VarSetCapacity(var2, 24, 0), domain := ComObject(0x400C, &var2)  ; Requires v1.1.17+
		process.GetOwner(user, domain)
		return domain[] . "\" . user[]
	}
}

JokeHasKey(Obj, Key) {
	Exists := false
	for index, element in Obj {
		if (element = Key) {
			Exists := true
		}
	}
	return Exists
}

WriteCmtraceEntry(cmtraceMessage, cmtraceType) {
	dtmNow := A_Now
	FormatTime, cmtraceTime, dtmNow, HH:mm:ss.000+000
	FormatTime, cmtraceDate, dtmNow, MM-dd-yyyy	
	cmtraceEntry := "<![LOG[" . cmtraceMessage . "]LOG]!><time=""" . cmtraceTime . """ date=""" . cmtraceDate . """ component=""" . cmtraceComponent . """ context=""" . cmtraceContext . """ type=""" . cmtraceType . """ thread=""" . cmtraceThread . """ file=""" . cmtraceFile . """>`r`n"
	FileAppend, %cmtraceEntry%, %aotLogPath%
}

return
