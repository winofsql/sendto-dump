' ****************************************************
' �t�@�C�����P�U�i���Ń_���v���܂�
' ****************************************************
Dim Fs,Stream
Dim InFile
Dim Kana
Dim KjFlg
Kana = Array( _
"�","�","�","�","�","�","�","�","�","�","�","�","�","�","�", _
"�","�","�","�","�","�","�","�","�","�","�","�","�","�","�","�", _
"�","�","�","�","�","�","�","�","�","�","�","�","�","�","�","�", _
"�","�","�","�","�","�","�","�","�","�","�","�","�","�","�","�" )

Set Fs = CreateObject( "Scripting.FileSystemObject" )
Set Stream = CreateObject("ADODB.Stream")

InFile = WScript.Arguments(0)

Dim LineBuffer,DispBuffer,CWork,nCnt,strBuff,i,j

if not Fs.FileExists( InFile ) then
	Wscript.Echo "�t�@�C�������݂��܂���"
	Wscript.Quit
end if

' ------------------------------------------------------
' Stream �̃I�[�v��
Stream.Open
 
' ------------------------------------------------------
' Stream �^�C�v�̎w��
Stream.Type = 1		' StreamTypeEnum �� adTypeBinary
 
' ------------------------------------------------------
' �����t�@�C���̓��e�� Stream �ɓǂݍ���
Stream.LoadFromFile InFile
 
' ------------------------------------------------------
' �o�C�i���^�� Stream �I�u�W�F�N�g�����ǂݎ���ĉ��H
Bcnt = 0
nCnt = 0
KjFlg = ""

Do while not Stream.EOS

	if ( nCnt MOD 16 ) = 0 then
		Wscript.Echo "          0  1  2  3  4  5  6  7" _
		& "  8  9  A  B  C  D  E  F"
		Wscript.Echo "--------------------------------" _
		& "------------------------------------------"
	end if

	' 16 �o�C�g�̓Ǎ�
	LineBuffer = Stream.Read(16)

	strBuff = ""
	For i = 1 to LenB( LineBuffer )
		CWork = MidB(LineBuffer,i,1)
		Cwork = AscB(Cwork)
		Cwork = Hex(Cwork)
		Cwork = Ucase(Cwork)
		Cwork = Right( "0" & Cwork, 2 )
		DispBuffer = DispBuffer & Cwork & " "
		strBuff = strBuff & CharConv( Cwork )
	Next

	Wscript.Echo _
		Right( _
			"00000000" & Ucase(Hex( nCnt * 16 )), 8 _
		) & " " & _
		Left(DispBuffer & String(49," "), 49 ) & strBuff
	DispBuffer = ""

	nCnt = nCnt + 1
 
Loop
 
' ------------------------------------------------------
' Stream �����
Stream.Close

Set Stream = Nothing
Stream = Empty
Set Fs = Nothing
Fs = Empty

' ****************************************************
' ���f�[�^�̃e�L�X�g
' ****************************************************
function CharConv( HexCode )

	Dim nCode

	nCode = Cint( "&H" & HexCode )

	if KjFlg = "" then
		if &H81 <= nCode and nCode <= &H84 or _
			&H88 <= nCode and nCode <= &H9f or _
			&HE0 <= nCode and nCode <= &HEA then
			KjFlg = HexCode
			CharConv = ""
			Exit Function
		end if
	else
		if HexCode <> "00" then
			KjFlg = KjFlg & HexCode
			CharConv = Chr( Cint( "&H" & KjFlg ) )
		else
			CharConv = ".."
		end if
		KjFlg = ""
		Exit Function
	end if

	if 0 <= nCode and nCode <= &H1F then
		CharConv = "."
	end if
	if &H20 <= nCode and nCode <= &H7E then
		CharConv = Chr(nCode)
	end if
	if &H7F <= nCode and nCode <= &HA0 then
		CharConv = "."
	end if
	if &HA1 <= nCode and nCode <= &HDF then
		CharConv = Kana(nCode-&HA1)
	end if
	if &HE0 <= nCode and nCode <= &HFF then
		CharConv = "."
	end if

end function
