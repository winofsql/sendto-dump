if WScript.Arguments.Count = 1 then
	strMessage = "���邩����s���ĉ�����" & vbCrLf & vbCrLf
	strMessage = strMessage & "�� �����N��̍Ō�̐����̓R�}���h�v�����v�g�̍s���ł��@�@�@" & vbCrLf
	strMessage = strMessage & "�� �v���p�e�B���E�C���h�E���ő剻������@������܂��@�@�@" & vbCrLf
	Call MsgBox(strMessage,0,"lightbox")
	Wscript.Quit
end if

Set WshShell = CreateObject( "WScript.Shell" )   
Set Fso = CreateObject( "Scripting.FileSystemObject" )

strCurPath = WScript.ScriptFullName
Set obj = Fso.GetFile( strCurPath )
Set obj = obj.ParentFolder
strCurPath = obj.Path

strCommand = "cmd /k mode CON lines="&WScript.Arguments(0)&" & cscript.exe """ & _
strCurPath & "\dump_c.vbs"" """ & WScript.Arguments(1) & """ | more & pause"
Call WshShell.Run( strCommand )
