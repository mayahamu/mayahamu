Attribute VB_Name = "ProcessData"

Sub L7�p�Ƃ���a����b�Ɉړ�����o�b�`�t�@�C�����쐬()
'-----------------------------------------------------------------------------'
' L7�p�ɊY�����̃f�[�^�S����a����b�t�H���_�Ɉړ�����o�b�`�t�@�C�����쐬����
' �����Excel VBA Version 1.00
' 
' Copyright 2015 Masanori Tanaka (Genonsha Co.,Ltd.)
'-----------------------------------------------------------------------------'
'
' ���s����ƁAL7�p�ɊY�����̔ԑg�S����a�t�H���_����b�t�H���_�ֈړ�����
' �o�b�`�t�@�C���udataMove_L7_a2b.cmd�v���f�X�N�g�b�v�ɕۑ����Ă���܂��B
'

' �A�N�e�B�u�V�[�g�̖��O���擾���āA�l�X�ȏ�ʂŊ��p�ł���悤�ɂ��Ă���
Dim activeWSName As String
Dim activeWSNameMid As String

activeWSName = Replace(ActiveSheet.Name, ".","")
activeWSNameMid = Replace(Mid(ActiveSheet.Name, 3), ".","") ' ��u2015.07�v���u1507�v��


' �e�L�X�g�t�@�C�����f�X�N�g�b�v�ɍ쐬����
' ���̐�uClose #1�v����܂ŁA�f�[�^���uPrint #1,�i���e�j�v�Ƃ��Ĉ�s���������߂�
	Dim exportPath As String
	exportPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\dataMove_L7_a2b.cmd"
	Open exportPath For Output As #1

	Print #1, "@echo off" & vbCrLf;
	Print #1, "title ""L7�p�ɑO�����̃f�[�^�S����a����b�t�H���_�Ɉړ�����o�b�`�t�@�C��"""
	Print #1, "SETLOCAL enabledelayedexpansion"


' �A�N�e�B�u�V�[�g�̃t�@�C�������X�g�i�e��j����s���ǂ�Ń��[�v����
' �ŏI�s�������I�Ɏ擾����̂ŁA�ԑg�����ς���Ă��󂫃`�����l���������Ă����v
Dim searchStrings As Range
Dim currentProgramName As Range
Dim processFiles As Long

processFiles = 0

For Each searchStrings In ActiveSheet.Range("F5:F" & Cells(Rows.Count, 6).End(xlUp).Row)

If searchStrings <> "nha061620001ma" Then ' ���k�W�F�b�g�ȊO������

	'���o�[�W�����̃t�@�C�����Ȃ�
	If searchStrings Like "######" OR searchStrings Like "####[a-h]#" Then
	' ����Ȋ����̕�����ɂ���
	' move e:\ANA201507\L7_a\150705*.mp3 e:\ANA201507\L7_b\ >> Result_L7_move_a2b.txt
	Print #1, "move e:\ANA20" & activeWSNameMid & "\L7_a\" & searchStrings & "*.mp3 " & "e:\ANA20" & activeWSNameMid & "\L7_b\ >> Result_L7_move_a2b.txt"

	'�V�o�[�W�����̃t�@�C�����Ȃ�
	ElseIf searchStrings Like "nha*" Then
	' ����Ȋ����̕�����ɂ���
	' move e:\ANA201603\L7_a\nha0316001*.mp3 e:\ANA201603\L7_b\ >> Result_L7_move_a2b.txt
	Print #1, "move e:\ANA20" & activeWSNameMid & "\L7_a\" & left(searchStrings, 10) & "*.mp3 " & "e:\ANA20" & activeWSNameMid & "\L7_b\ >> Result_L7_move_a2b.txt"
	End If

	' �Ȑ����J�E���g���Ă݂�
	processFiles = processFiles + CLng(searchStrings.offset(0,11).Value)

End If
Next searchStrings


' �㏈��
	Print #1, "echo �t�@�C���̈ړ����������܂����B"
	Print #1, "echo �\�z�����t�@�C������ " & processFiles & "�t�@�C���ł��B"
	Print #1, "pause"
	Print #1, "ENDLOCAL"
	Print #1, "exit /b"

	Close #1


msgbox "�udataMove_L7_a2b.cmd�v�ł�������"

End Sub






Sub L7�p��b����final�ɃR�s�[����o�b�`�t�@�C�����쐬()
'-----------------------------------------------------------------------------'
' L7�p�̃f�[�^��b����finalSelect�t�H���_�ɃR�s�[����o�b�`�t�@�C�����쐬����
' �����Excel VBA Version 1.00
' 
' Copyright 2015 Masanori Tanaka (Genonsha Co.,Ltd.)
'-----------------------------------------------------------------------------'
'
' ���s����ƁAL7�p�̔ԑg��b�t�H���_����finalSelect�t�H���_�փR�s�[����
' �o�b�`�t�@�C���udataCopy_L7_b2Final.cmd�v���f�X�N�g�b�v�ɕۑ����Ă���܂��B
'

' �A�N�e�B�u�V�[�g�̖��O���擾���āA�l�X�ȏ�ʂŊ��p�ł���悤�ɂ��Ă���
Dim activeWSName As String
Dim activeWSNameMid As String

activeWSName = Replace(ActiveSheet.Name, ".","")
activeWSNameMid = Replace(Mid(ActiveSheet.Name, 3), ".","") ' ��u2015.07�v���u1507�v��


' �e�L�X�g�t�@�C�����f�X�N�g�b�v�ɍ쐬����
' ���̐�uClose #1�v����܂ŁA�f�[�^���uPrint #1,�i���e�j�v�Ƃ��Ĉ�s���������߂�
	Dim exportPath As String
	exportPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\dataCopy_L7_b2Final.cmd"
	Open exportPath For Output As #1

	Print #1, "@echo off" & vbCrLf;
	Print #1, "title ""L7�p�̃f�[�^��b����finalSelect�t�H���_�ɃR�s�[����o�b�`�t�@�C��"""
	Print #1, "SETLOCAL enabledelayedexpansion"


' �A�N�e�B�u�V�[�g�̃t�@�C�������X�g�i�e��j����s���ǂ�Ń��[�v����
' �ŏI�s�������I�Ɏ擾����̂ŁA�ԑg�����ς���Ă��󂫃`�����l���������Ă����v
Dim searchStrings As Range
Dim currentProgramName As Range
Dim processFiles As Long

processFiles = 0

For Each searchStrings In ActiveSheet.Range("F5:F" & Cells(Rows.Count, 6).End(xlUp).Row)

If searchStrings.Offset(0,-2).Value = "��" Then ' �c��Ɂ�������΁i��L7�Ή��`�����l����������j

	'���o�[�W�����̃t�@�C�����Ȃ�
	If searchStrings Like "######" OR searchStrings Like "####[a-h]#" Then
	' ����Ȋ����̕�����ɂ���
	' copy e:\ANA201507\L7_b\150705*.mp3 e:\ANA201507\L7_b\ >> Result_L7_move_a2b.txt
	Print #1, "copy e:\ANA20" & activeWSNameMid & "\L7_b\" & searchStrings & "*.mp3 " & "e:\ANA20" & activeWSNameMid & "\L7_finalSelect\ >> Result_L7_copy_b2Final.txt"

	'�V�o�[�W�����̃t�@�C�����Ȃ�
	ElseIf searchStrings Like "nha*" Then
	' ����Ȋ����̕�����ɂ���
	' copy e:\ANA201603\L7_a\nha0316001*.mp3 e:\ANA201603\L7_b\ >> Result_L7_move_a2b.txt
	Print #1, "copy e:\ANA20" & activeWSNameMid & "\L7_b\" & left(searchStrings, 10) & "*.mp3 " & "e:\ANA20" & activeWSNameMid & "\L7_finalSelect\ >> Result_L7_copy_b2Final.txt"
	End If

	' �Ȑ����J�E���g���Ă݂�
	processFiles = processFiles + CLng(searchStrings.offset(0,11).Value)

End If
Next searchStrings


' �㏈��
	Print #1, "echo �t�@�C���̃R�s�[���������܂����B"
	Print #1, "echo �\�z�����t�@�C������ " & processFiles & "�t�@�C���ł��B"
	Print #1, "pause"
	Print #1, "ENDLOCAL"
	Print #1, "exit /b"

	Close #1


msgbox "�udataCopy_L7_b2Final.cmd�v�ł�������"

End Sub





Sub Bluebox�Ώ۔ԑg�̃f�[�^����������o�b�`�t�@�C�����쐬()
'-----------------------------------------------------------------------------'
' Bluebox�Ώ۔ԑg�̃f�[�^���e�t�H���_����W�߂Ă���o�b�`�t�@�C�����쐬����
' �����Excel VBA Version 1.03
' 
' Copyright 2015, 2018-2019 Masanori Tanaka (Genonsha Co.,Ltd.)
'-----------------------------------------------------------------------------'
'
' Bluebox�Ώ۔ԑg�̃f�[�^���e�t�H���_���瓖����WiFi�t�H���_�փR�s�[����
' �o�b�`�t�@�C���udataCollection_WiFi.cmd�v���f�X�N�g�b�v�ɕۑ����Ă���܂��B
' �ȑO��Wi-Fi�p���������ǁA���݂�Bluebox�p�Ƃ��Ċ��p���܂��B


' �A�N�e�B�u�V�[�g�̖��O���擾���āA�l�X�ȏ�ʂŊ��p�ł���悤�ɂ��Ă���
Dim activeWSName As String
Dim activeWSNameMid As String

activeWSName = Replace(ActiveSheet.Name, ".","")
activeWSNameMid = Replace(Mid(ActiveSheet.Name, 3), ".","") ' ��u2015.07�v���u1507�v��


' �e�L�X�g�t�@�C�����f�X�N�g�b�v�ɍ쐬����
' ���̐�uClose #1�v����܂ŁA�f�[�^���uPrint #1,�i���e�j�v�Ƃ��Ĉ�s���������߂�
	Dim exportPath As String
	exportPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\dataCollection_Bluebox.cmd"
	Open exportPath For Output As #1

	Print #1, "@echo off" & vbCrLf;
	Print #1, "title ""Bluebox�Ώ۔ԑg�̃f�[�^���e�t�H���_����W�߂Ă���o�b�`�t�@�C��"""
	Print #1, "SETLOCAL enabledelayedexpansion"


' �A�N�e�B�u�V�[�g�̃t�@�C�������X�g�i�e��j����s���ǂ�Ń��[�v����
' �ŏI�s�������I�Ɏ擾����̂ŁA�ԑg�����ς���Ă��󂫃`�����l���������Ă����v
Dim searchStrings As Range
Dim processFiles As Long
Dim totalPrograms As Long


processFiles = 0
totalPrograms = 0

For Each searchStrings In ActiveSheet.Range("F5:F" & Cells(Rows.Count, 6).End(xlUp).Row)

If searchStrings.Offset(0, -3).Value <> "" Then ' �b��ɉ��������I�Ȃ��̂������Ă���΁i��Wi-Fi�Ή��`�����l����������j

	'���o�[�W�����̃t�@�C�����Ȃ�
	If searchStrings Like "######" Then
		' ����Ȋ����̕�����ɂ���
		' copy e:\ANA201603\mp3\160305*.mp3 e:\ANA201603\Bluebox\ >> Result_WiFi_copy.txt
		Print #1, "copy e:\ANA20" & left(searchStrings, 4) & "\mp3\" & searchStrings & "*.mp3 " & "e:\ANA20" & activeWSNameMid & "\Bluebox\ >> Result_BB_copy.txt"

	'�V�o�[�W�����̃t�@�C�����Ȃ�
	ElseIf searchStrings Like "nha*" Then
		' ����Ȋ����̕�����ɂ���
		' copy e:\ANA201603\mp3\nha0316001*.mp3 e:\ANA201603\Bluebox\ >> Result_WiFi_copy.txt
		Print #1, "copy e:\ANA20" & mid(searchStrings, 6, 2) & mid(searchStrings, 4, 2) & "\mp3\" & left(searchStrings, 10) & "*.mp3 " & "e:\ANA20" & activeWSNameMid & "\Bluebox\ >> Result_WiFi_copy.txt"
	End If

	' �Ȑ����J�E���g���Ă݂�
	processFiles = processFiles + CLng(searchStrings.offset(0,11).Value)

	' �ԑg�����J�E���g���Ă݂�
	totalPrograms = totalPrograms + 1

End If
Next searchStrings


' �㏈��
	Print #1, "echo �t�@�C���̃R�s�[���������܂����B"
	Print #1, "echo ���v " & totalPrograms & "�ԑg�B�\�z�����t�@�C������ " & processFiles & "�t�@�C���ł��B"
	Print #1, "echo �󕨌�𑫂�����" & processFiles +1 & "�t�@�C���ˁB"
	Print #1, "pause"
	Print #1, "ENDLOCAL"
	Print #1, "exit /b"

	Close #1


msgbox "�udataCollection_Bluebox.cmd�v�ł�������" & vbCr & "Bluebox�p�� " & totalPrograms & " �ԑg�ł�"

End Sub




Sub �����p�t�@�C�����t�H���_�ɐU�蕪����PS1���쐬()
'-----------------------------------------------------------------------------'
' ��������MP3���A�ԑg���̂����t�H���_�ɐU�蕪����PS1�t�@�C�����쐬����
' �����Excel VBA Version 2.10
' 
' Copyright 2016-2019 Masanori Tanaka (Genonsha Co.,Ltd.)
'-----------------------------------------------------------------------------'
'
' PowerShell�p�t�@�C���udataCollection_mp3.ps1�v���f�X�N�g�b�v�ɕۑ����Ă���܂��B
' ����𓖌�����MP3�t�@�C���������Ă���t�H���_�ɓ���Ď��s����Ƌ����̌������I

' 2.10 - [2019/04/16] �U�蕪�����t�@�C�������ɖ߂��@�\��t����

Dim stream As New ADODB.Stream
stream.Type = adTypeText ' ��������2�i1�̓o�C�i���j
stream.Charset = "UTF-8" ' �����R�[�h
stream.LineSeparator = 10 ' ���s�R�[�h


' �A�N�e�B�u�V�[�g�̖��O���擾���āA�l�X�ȏ�ʂŊ��p�ł���悤�ɂ��Ă���
Dim activeWSName As String
Dim activeWSNameMid As String

activeWSName = Replace(ActiveSheet.Name, ".","")
activeWSNameMid = Replace(Mid(ActiveSheet.Name, 3), ".","") ' ��u2015.07�v���u1507�v��


' �t�@�C���i�ɂȂ��ԁj�ɓ��e����������ł���
stream.Open

	stream.WriteText "# " & Replace(ActiveSheet.Range("A2").value," ","") & "�̔ԑg�݂̂��g�p���܂�", 1
	stream.WriteText "# ���̃t�@�C����" & Replace(ActiveSheet.Range("A2").value," ","") & "��MP3�t�@�C���������Ă���t�H���_��", 1
	stream.WriteText "# ����Ď��s����ƁA�ԑg���̂����t�H���_���쐬���ēK�X�t�@�C����U�蕪���܂��i���ɖ߂��@�\�t���j", 1
	stream.WriteText "", 1
	stream.WriteText "If (Get-ChildItem -Filter *.mp3) {", 1
	stream.WriteText "", 1
	stream.WriteText "[array] $allProgramNames = @()", 1
	stream.WriteText "", 1

' �A�N�e�B�u�V�[�g�̃t�@�C�������X�g�i�e��j����s���ǂ�Ń��[�v����
' �ŏI�s�������I�Ɏ擾����̂ŁA�ԑg�����ς���Ă��󂫃`�����l���������Ă����v
Dim searchStrings As Range
Dim processFiles As Long
Dim totalPrograms As Long
Dim programString As String
Dim programName As String

totalPrograms = 0

For Each searchStrings In ActiveSheet.Range("F5:F" & Cells(Rows.Count, 6).End(xlUp).Row)

If searchStrings Like "nha" & right(activeWSNameMid, 2) & left(activeWSNameMid, 2) & "*" Then ' �t�@�C�������unha�v�̂��Ɠ����i�Ⴆ��0316�j��������

	' �t�H���_���Ƃ��Ďg���Ȃ��������������ꍇ�ɑS�p�ɂ��Ă����i���n�I�I�j
	ProgramName = searchStrings.Offset(0, 1).Value
	ProgramName = Replace(ProgramName,"/","�^")
	ProgramName = Replace(ProgramName,"\","��")
	ProgramName = Replace(ProgramName,":","�F")
	ProgramName = Replace(ProgramName,"*","��")
	ProgramName = Replace(ProgramName,"?","�H")
	ProgramName = Replace(ProgramName,"""","''")
	ProgramName = Replace(ProgramName,"<","��")
	ProgramName = Replace(ProgramName,">","��")
	ProgramName = Replace(ProgramName,"|","�b")
	programName = Replace(programName, Chr$(&H8167), "�e") ' �S�p�́g�h������ɔ��p�ɂȂ�G���[�ɂȂ�̂ŁA�������̂��Ɓe�f��
	programName = Replace(ProgramName, Chr$(&H8168), "�f")

	'�O�̂��߃Z�����̉��s��]���ȃX�y�[�X������
	ProgramName = Replace(ProgramName, vbLf, "")
	ProgramName = Trim(ProgramName)


	programString =  """" & FORMAT(searchStrings.Offset(0, -5).Value,"000_") & ProgramName & """"

	' �t�H���_�����u001_�ԑg���v�݂����ɂ��Ĕz��ɒǉ�
	stream.WriteText "$allProgramNames += " & programString, 1

	' �ԑg�����J�E���g���Ă݂�
	totalPrograms = totalPrograms + 1

End If
Next searchStrings


	' �t�H���_������ăt�@�C�����ړ������Ă���
	stream.WriteText "", 1
	stream.WriteText "for ($i=0; $i -lt $allProgramNames.Length; $i++)", 1
	stream.WriteText "{", 1
	stream.WriteText vbTab & "New-Item $allProgramNames[$i] -Force -ItemType Directory", 1
	stream.WriteText vbTab & "$moji = $allProgramNames[$i].Substring(0, 3) + ""*.mp3""", 1
	stream.WriteText vbTab & "Get-ChildItem ./nha" & right(activeWSNameMid, 2) & left(activeWSNameMid, 2) & "$moji" & " | Move-Item -Destination $allProgramNames[$i]", 1
	stream.WriteText "}", 1
	stream.WriteText "", 1

	'�g�O�����ۂ������ŃA���h�D���镔��
	stream.WriteText "}else{", 1
	stream.WriteText "# �ċA�I��MP3��T���āA�����������̂��t���p�X�ɂ��āA���̃X�N���v�g������K�w�Ɏ����Ă���", 1
	stream.WriteText "Get-ChildItem -Filter *.mp3 -Recurse | Select-Object FullName | ForEach-Object{Move-Item -literalPath $_.Fullname .\}", 1
	stream.WriteText "", 1
	stream.WriteText "# �t�@�C�����̓����uNH_yymm-W-A-�v�Ŏn�܂��Ă���΁A�������������", 1
	stream.WriteText "Get-ChildItem *.mp3 | Rename-Item -NewName { $_.Name -replace 'NH_\d{4}-W-A-','' }", 1
	stream.WriteText "", 1
	stream.WriteText "# �Ō�Ƀt�H���_���폜�i�O�̂��߁A�t�H���_���Ƀt�@�C�����c���Ă�����~�߂�j", 1
	stream.WriteText "Get-ChildItem -Directory | Remove-Item -Recurse", 1
	stream.WriteText "Read-Host ""������ɂ� Enter �L�[�������Ă�������...""", 1
	stream.WriteText "}", 1
	stream.WriteText "", 1

' �w�肵���t�@�C���ɕۑ����ĕ���
stream.SaveToFile "C:\Users\tanaka\Desktop\dataCollection" & activeWSName & ".ps1", adSaveCreateOverWrite
stream.Close

msgbox "�udataCollection.ps1�v�ł�������" & vbCr & "�����̐V�K�`�����l���� " & totalPrograms & " �ԑg�ł�"

End Sub


