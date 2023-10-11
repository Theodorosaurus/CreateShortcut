Option Explicit

Dim x, z, t, y, q, p
Function PrintNstartup()
	x = MsgBox("Will the Restaurant/Shop be using a Printer?                    ", 4 +vbSystemModal, "--kiosk-printing")
		If x = 6 Then
			Shortcut.Arguments = "--app=https://web.eu.restaurant-partners.com/ --kiosk-printing --start-maximized --profile-directory=Default"
			Shortcut.Save
			z = MsgBox("Done!                    ", 64 +vbSystemModal, "Add --kiosk-printing")
		Else
			Shortcut.Arguments = "--app=https://web.eu.restaurant-partners.com/ --start-maximized --profile-directory=Default"
			Shortcut.Save
			t = MsgBox("Done!                    ", 64 +vbSystemModal, "No --kiosk-printing")
		End If
	y = MsgBox("Do you wish GoWeb to launch on Startup, automatically?                    ", 4 +vbSystemModal, "GoWeb on Startup")
		If y = 6 Then
			obj.CopyFile strCurDir & "\efood.lnk", strStUpDir & "\efood.lnk", True
			q = MsgBox("Done!                    ", 64 +vbSystemModal, "GoWeb on Startup")
		Else
			p = MsgBox("Done!                    ", 64 +vbSystemModal, "No Startup")
		End If
End Function

Dim fso, obj
Set fso = CreateObject("WScript.Shell")
Set obj = CreateObject("Scripting.FileSystemObject")

Dim strStUpDir, strCurDir
strStUpDir = fso.ExpandEnvironmentStrings("%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\")
strCurDir = obj.GetAbsolutePathName("")

Dim Shortcut
Set Shortcut = fso.CreateShortcut(strCurDir & "\efood.lnk") 

Dim strComputer, objReg, strKeyPathLM, arrEntryNamesLM, arrEntryNamesLMi, strValueLM, strKeyPathCU, arrEntryNamesCU, strValueCU
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CURRENT_USER  = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

strKeyPathLM = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\brave.exe"
objReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPathLM, arrEntryNamesLM

If IsArray(arrEntryNamesLM) Then
	objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPathLM, arrEntryNamesLMi, strValueLM
	Shortcut.TargetPath = strValueLM
	Shortcut.WorkingDirectory = LEFT(strValueLM, (LEN(strValueLM)-9))
	Shortcut.IconLocation = "https://web.restaurant-partners.com/icons/favicon.c250a8bf2b775a3db19e33c4bf393cf5.ico"
	Shortcut.Save
	PrintNStartup()
Else
	strKeyPathCU = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\brave.exe"
	objReg.GetStringValue HKEY_CURRENT_USER, strKeyPathCU, arrEntryNamesCU, strValueCU
	Shortcut.TargetPath = strValueCU
	Shortcut.Arguments = "--app=https://web.eu.restaurant-partners.com/ --kiosk-printing --start-maximized --profile-directory=Default"
	Shortcut.WorkingDirectory = LEFT(strValueCU, (LEN(strValueCU)-9))
	Shortcut.IconLocation = "https://web.restaurant-partners.com/icons/favicon.c250a8bf2b775a3db19e33c4bf393cf5.ico"
	Shortcut.Save
	PrintNstartup()
End If