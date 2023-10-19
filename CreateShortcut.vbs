Option Explicit

Dim x, z, t, y, q, p

Function PrintNstartup()
	x = MsgBox("Will you be using a Printer with this browser App?                    ", 4 +vbSystemModal, "--kiosk-printing")
		If x = 6 Then
			Shortcut.Arguments = "--app=https://www.example.com/ --kiosk-printing --start-maximized --profile-directory=Default"
			Shortcut.Save
			z = MsgBox("Done!                    ", 64 +vbSystemModal, "Add --kiosk-printing")
		Else
			Shortcut.Arguments = "--app=https://www.example.com/ --start-maximized --profile-directory=Default"
			Shortcut.Save
			t = MsgBox("Done!                    ", 64 +vbSystemModal, "No --kiosk-printing")
		End If
	y = MsgBox("Do you wish your browser App to launch on Startup, automatically?                    ", 4 +vbSystemModal, "BrowserApp on Startup")
		If y = 6 Then
			obj.CopyFile strCurDir & "\BrowserApp.lnk", strStUpDir & "\BrowserApp.lnk", True
			q = MsgBox("Done!                    ", 64 +vbSystemModal, "Browser App on Startup")
		Else
			p = MsgBox("Done!                    ", 64 +vbSystemModal, "No Startup")
		End If
End Function

Function IconFolder()
	If obj.FolderExists(strMyDocs & "\MyIcon") Then
		obj.DeleteFolder strMyDocs & "\MyIcon", True
		obj.CreateFolder strMyDocs & "\MyIcon"
	Else
		obj.CreateFolder strMyDocs & "\MyIcon"
	End If
End Function

Dim fso, obj
Set fso = CreateObject("WScript.Shell")
Set obj = CreateObject("Scripting.FileSystemObject")

Dim strStUpDir, strMyDocs, strCurDir
strStUpDir = fso.SpecialFolders("Startup")
strMyDocs = fso.SpecialFolders("MyDocuments")
strCurDir = obj.GetAbsolutePathName("")

efoodFolder()

Dim xHttp, bStrm
Set xHttp = createobject("Microsoft.XMLHTTP")
Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", "https://www.example.com/icons/favicon.ico", False
xHttp.Send

with bStrm
    .type = 1
    .open
    .write xHttp.responseBody
    .savetofile strMyDocs & "\MyIcon\favicon.ico", 2
end with

Dim Shortcut
Set Shortcut = fso.CreateShortcut(strCurDir & "\BrowserApp.lnk") 
Shortcut.IconLocation = strMyDocs & "\MyIcon\favicon.ico"
Shortcut.Save

Dim strComputer, objReg, strKeyPathLM, arrEntryNamesLM, arrEntryNamesLMi, strValueLM, strKeyPathCU, arrEntryNamesCU, strValueCU
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CURRENT_USER  = &H80000001
Const REG_SZ = 1
strComputer = "."
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

strKeyPathLM = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\mybrowser.exe"
objReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPathLM, arrEntryNamesLM

If IsArray(arrEntryNamesLM) Then
	objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPathLM, arrEntryNamesLMi, strValueLM
	Shortcut.TargetPath = strValueLM
	Shortcut.WorkingDirectory = LEFT(strValueLM, (LEN(strValueLM)-13))
	Shortcut.Save
	PrintNStartup()
Else
	strKeyPathCU = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\mybrowser.exe"
	objReg.GetStringValue HKEY_CURRENT_USER, strKeyPathCU, arrEntryNamesCU, strValueCU
	Shortcut.TargetPath = strValueCU
	Shortcut.WorkingDirectory = LEFT(strValueCU, (LEN(strValueCU)-13))
	Shortcut.Save
	PrintNstartup()
End If
