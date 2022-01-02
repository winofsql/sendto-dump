if WScript.Arguments.Count = 1 then
	strMessage = "送るから実行して下さい" & vbCrLf & vbCrLf
	strMessage = strMessage & "※ リンク先の最後の数字はコマンドプロンプトの行数です　　　" & vbCrLf
	strMessage = strMessage & "※ プロパティよりウインドウを最大化する方法もあります　　　" & vbCrLf
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
