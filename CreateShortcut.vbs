Option Explicit

Dim x, z, t, y, q, p

Function PrintNstartup() 
	x = MsgBox("Will the User be using a Printer?                    ", 4 +vbSystemModal, "--kiosk-printing") 'asking the user whether they wish to use auto-printing w/ their web app, or not.
		If x = 6 Then
			Shortcut.Arguments = "--app=https://example.com --kiosk-printing --start-maximized --profile-directory=Default" 'args w/ auto-print enabled.
			Shortcut.Save
			z = MsgBox("Done!                    ", 64 +vbSystemModal, "Add --kiosk-printing")
		Else
			Shortcut.Arguments = "--app=https://example.com --start-maximized --profile-directory=Default"
			Shortcut.Save
			t = MsgBox("Done!                    ", 64 +vbSystemModal, "No --kiosk-printing") 'args w/o auto-print enabled.
		End If
	y = MsgBox("Do you wish App to launch on Startup, automatically?                    ", 4 +vbSystemModal, "App on Startup") 'asking the user whether they wish to use make the App launch on PC startup.
		If y = 6 Then
			obj.CopyFile strCurDir & "\App.lnk", strStUpDir & "\App.lnk", True ' this will not only copy the shortcut icon on Startup, but overwrite any existing ones w/ the same name.
			q = MsgBox("Done!                    ", 64 +vbSystemModal, "App on Startup")
		Else
			p = MsgBox("Done!                    ", 64 +vbSystemModal, "No Startup")
		End If
End Function

Dim fso, obj
Set fso = CreateObject("WScript.Shell")
Set obj = CreateObject("Scripting.FileSystemObject")

Dim strStUpDir, strCurDir
strStUpDir = fso.ExpandEnvironmentStrings("%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\") 'creating a var for the standard location of the Startup folder
strCurDir = obj.GetAbsolutePathName("") 'using this obj instead of the Special Folders one, because the "Desktop" won't work with Win11 users using OneDrive for their setup, i.e. C:\Users\Username\OneDrive\MyComputer\desktop files

Dim Shortcut
Set Shortcut = fso.CreateShortcut(strCurDir & "\App.lnk") 'creating the shortuct on the same path where our script is saved & launched, so we must launch it from the desktop location

Dim strComputer, objReg, strKeyPathLM, arrEntryNamesLM, arrEntryNamesLMi, strValueLM, strKeyPathCU, arrEntryNamesCU, strValueCU
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CURRENT_USER  = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv") 'pretty standard template for getting Registry's objects

strKeyPathLM = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MyInternetBrowser.exe"
objReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPathLM, arrEntryNamesLM 'checking Registry's Enumarated Values for the installation path of the browser the user will be using w/ the web based App

If IsArray(arrEntryNamesLM) Then 'if our search finds an Array in the Local Machine location, then the browser was installed w/ Admin rights in Program Files 
	objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPathLM, arrEntryNamesLMi, strValueLM
	Shortcut.TargetPath = strValueLM
	Shortcut.WorkingDirectory = LEFT(strValueLM, (LEN(strValueLM)-21)) 'truncating the full path of the browser by 21 characters to get its full directory path excluding the exe file
	Shortcut.IconLocation = "https://www.example.com/icons/favicon.app.ico"
	Shortcut.Save
	PrintNStartup() 'calling my function from above
Else
	strKeyPathCU = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MyInternetBrowser.exe" 'if our search doesn't find an Array in the Local Machine location, then the browser was installed w/o Admin rights in the local user's profile paths
	objReg.GetStringValue HKEY_CURRENT_USER, strKeyPathCU, arrEntryNamesCU, strValueCU
	Shortcut.TargetPath = strValueCU
	Shortcut.WorkingDirectory = LEFT(strValueCU, (LEN(strValueCU)-9))
	Shortcut.IconLocation = "https://www.example.com/icons/favicon.app.ico"
	Shortcut.Save
	PrintNstartup()
End If
