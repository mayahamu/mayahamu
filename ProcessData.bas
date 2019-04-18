Attribute VB_Name = "ProcessData"

Sub L7用としてaからbに移動するバッチファイルを作成()
'-----------------------------------------------------------------------------'
' L7用に該当月のデータ全部をaからbフォルダに移動するバッチファイルを作成する
' そんなExcel VBA Version 1.00
' 
' Copyright 2015 Masanori Tanaka (Genonsha Co.,Ltd.)
'-----------------------------------------------------------------------------'
'
' 実行すると、L7用に該当月の番組全部をaフォルダからbフォルダへ移動する
' バッチファイル「dataMove_L7_a2b.cmd」をデスクトップに保存してくれます。
'

' アクティブシートの名前を取得して、様々な場面で活用できるようにしておく
Dim activeWSName As String
Dim activeWSNameMid As String

activeWSName = Replace(ActiveSheet.Name, ".","")
activeWSNameMid = Replace(Mid(ActiveSheet.Name, 3), ".","") ' 例「2015.07」を「1507」に


' テキストファイルをデスクトップに作成する
' この先「Close #1」するまで、データを「Print #1,（内容）」として一行ずつ書き込める
	Dim exportPath As String
	exportPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\dataMove_L7_a2b.cmd"
	Open exportPath For Output As #1

	Print #1, "@echo off" & vbCrLf;
	Print #1, "title ""L7用に前月分のデータ全部をaからbフォルダに移動するバッチファイル"""
	Print #1, "SETLOCAL enabledelayedexpansion"


' アクティブシートのファイル名リスト（Ｆ列）を一行ずつ読んでループを回す
' 最終行を自動的に取得するので、番組数が変わっても空きチャンネルがあっても大丈夫
Dim searchStrings As Range
Dim currentProgramName As Range
Dim processFiles As Long

processFiles = 0

For Each searchStrings In ActiveSheet.Range("F5:F" & Cells(Rows.Count, 6).End(xlUp).Row)

If searchStrings <> "nha061620001ma" Then ' 東北ジェット以外を処理

	'旧バージョンのファイル名なら
	If searchStrings Like "######" OR searchStrings Like "####[a-h]#" Then
	' こんな感じの文字列にする
	' move e:\ANA201507\L7_a\150705*.mp3 e:\ANA201507\L7_b\ >> Result_L7_move_a2b.txt
	Print #1, "move e:\ANA20" & activeWSNameMid & "\L7_a\" & searchStrings & "*.mp3 " & "e:\ANA20" & activeWSNameMid & "\L7_b\ >> Result_L7_move_a2b.txt"

	'新バージョンのファイル名なら
	ElseIf searchStrings Like "nha*" Then
	' こんな感じの文字列にする
	' move e:\ANA201603\L7_a\nha0316001*.mp3 e:\ANA201603\L7_b\ >> Result_L7_move_a2b.txt
	Print #1, "move e:\ANA20" & activeWSNameMid & "\L7_a\" & left(searchStrings, 10) & "*.mp3 " & "e:\ANA20" & activeWSNameMid & "\L7_b\ >> Result_L7_move_a2b.txt"
	End If

	' 曲数をカウントしてみる
	processFiles = processFiles + CLng(searchStrings.offset(0,11).Value)

End If
Next searchStrings


' 後処理
	Print #1, "echo ファイルの移動が完了しました。"
	Print #1, "echo 予想されるファイル数は " & processFiles & "ファイルです。"
	Print #1, "pause"
	Print #1, "ENDLOCAL"
	Print #1, "exit /b"

	Close #1


msgbox "「dataMove_L7_a2b.cmd」できあがり"

End Sub






Sub L7用にbからfinalにコピーするバッチファイルを作成()
'-----------------------------------------------------------------------------'
' L7用のデータをbからfinalSelectフォルダにコピーするバッチファイルを作成する
' そんなExcel VBA Version 1.00
' 
' Copyright 2015 Masanori Tanaka (Genonsha Co.,Ltd.)
'-----------------------------------------------------------------------------'
'
' 実行すると、L7用の番組をbフォルダからfinalSelectフォルダへコピーする
' バッチファイル「dataCopy_L7_b2Final.cmd」をデスクトップに保存してくれます。
'

' アクティブシートの名前を取得して、様々な場面で活用できるようにしておく
Dim activeWSName As String
Dim activeWSNameMid As String

activeWSName = Replace(ActiveSheet.Name, ".","")
activeWSNameMid = Replace(Mid(ActiveSheet.Name, 3), ".","") ' 例「2015.07」を「1507」に


' テキストファイルをデスクトップに作成する
' この先「Close #1」するまで、データを「Print #1,（内容）」として一行ずつ書き込める
	Dim exportPath As String
	exportPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\dataCopy_L7_b2Final.cmd"
	Open exportPath For Output As #1

	Print #1, "@echo off" & vbCrLf;
	Print #1, "title ""L7用のデータをbからfinalSelectフォルダにコピーするバッチファイル"""
	Print #1, "SETLOCAL enabledelayedexpansion"


' アクティブシートのファイル名リスト（Ｆ列）を一行ずつ読んでループを回す
' 最終行を自動的に取得するので、番組数が変わっても空きチャンネルがあっても大丈夫
Dim searchStrings As Range
Dim currentProgramName As Range
Dim processFiles As Long

processFiles = 0

For Each searchStrings In ActiveSheet.Range("F5:F" & Cells(Rows.Count, 6).End(xlUp).Row)

If searchStrings.Offset(0,-2).Value = "○" Then ' Ｄ列に○があれば（＝L7対応チャンネルだったら）

	'旧バージョンのファイル名なら
	If searchStrings Like "######" OR searchStrings Like "####[a-h]#" Then
	' こんな感じの文字列にする
	' copy e:\ANA201507\L7_b\150705*.mp3 e:\ANA201507\L7_b\ >> Result_L7_move_a2b.txt
	Print #1, "copy e:\ANA20" & activeWSNameMid & "\L7_b\" & searchStrings & "*.mp3 " & "e:\ANA20" & activeWSNameMid & "\L7_finalSelect\ >> Result_L7_copy_b2Final.txt"

	'新バージョンのファイル名なら
	ElseIf searchStrings Like "nha*" Then
	' こんな感じの文字列にする
	' copy e:\ANA201603\L7_a\nha0316001*.mp3 e:\ANA201603\L7_b\ >> Result_L7_move_a2b.txt
	Print #1, "copy e:\ANA20" & activeWSNameMid & "\L7_b\" & left(searchStrings, 10) & "*.mp3 " & "e:\ANA20" & activeWSNameMid & "\L7_finalSelect\ >> Result_L7_copy_b2Final.txt"
	End If

	' 曲数をカウントしてみる
	processFiles = processFiles + CLng(searchStrings.offset(0,11).Value)

End If
Next searchStrings


' 後処理
	Print #1, "echo ファイルのコピーが完了しました。"
	Print #1, "echo 予想されるファイル数は " & processFiles & "ファイルです。"
	Print #1, "pause"
	Print #1, "ENDLOCAL"
	Print #1, "exit /b"

	Close #1


msgbox "「dataCopy_L7_b2Final.cmd」できあがり"

End Sub





Sub Bluebox対象番組のデータを引っ張るバッチファイルを作成()
'-----------------------------------------------------------------------------'
' Bluebox対象番組のデータを各フォルダから集めてくるバッチファイルを作成する
' そんなExcel VBA Version 1.03
' 
' Copyright 2015, 2018-2019 Masanori Tanaka (Genonsha Co.,Ltd.)
'-----------------------------------------------------------------------------'
'
' Bluebox対象番組のデータを各フォルダから当月のWiFiフォルダへコピーする
' バッチファイル「dataCollection_WiFi.cmd」をデスクトップに保存してくれます。
' 以前はWi-Fi用だったけど、現在はBluebox用として活用します。


' アクティブシートの名前を取得して、様々な場面で活用できるようにしておく
Dim activeWSName As String
Dim activeWSNameMid As String

activeWSName = Replace(ActiveSheet.Name, ".","")
activeWSNameMid = Replace(Mid(ActiveSheet.Name, 3), ".","") ' 例「2015.07」を「1507」に


' テキストファイルをデスクトップに作成する
' この先「Close #1」するまで、データを「Print #1,（内容）」として一行ずつ書き込める
	Dim exportPath As String
	exportPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\dataCollection_Bluebox.cmd"
	Open exportPath For Output As #1

	Print #1, "@echo off" & vbCrLf;
	Print #1, "title ""Bluebox対象番組のデータを各フォルダから集めてくるバッチファイル"""
	Print #1, "SETLOCAL enabledelayedexpansion"


' アクティブシートのファイル名リスト（Ｆ列）を一行ずつ読んでループを回す
' 最終行を自動的に取得するので、番組数が変わっても空きチャンネルがあっても大丈夫
Dim searchStrings As Range
Dim processFiles As Long
Dim totalPrograms As Long


processFiles = 0
totalPrograms = 0

For Each searchStrings In ActiveSheet.Range("F5:F" & Cells(Rows.Count, 6).End(xlUp).Row)

If searchStrings.Offset(0, -3).Value <> "" Then ' Ｃ列に何か文字的なものが入っていれば（＝Wi-Fi対応チャンネルだったら）

	'旧バージョンのファイル名なら
	If searchStrings Like "######" Then
		' こんな感じの文字列にする
		' copy e:\ANA201603\mp3\160305*.mp3 e:\ANA201603\Bluebox\ >> Result_WiFi_copy.txt
		Print #1, "copy e:\ANA20" & left(searchStrings, 4) & "\mp3\" & searchStrings & "*.mp3 " & "e:\ANA20" & activeWSNameMid & "\Bluebox\ >> Result_BB_copy.txt"

	'新バージョンのファイル名なら
	ElseIf searchStrings Like "nha*" Then
		' こんな感じの文字列にする
		' copy e:\ANA201603\mp3\nha0316001*.mp3 e:\ANA201603\Bluebox\ >> Result_WiFi_copy.txt
		Print #1, "copy e:\ANA20" & mid(searchStrings, 6, 2) & mid(searchStrings, 4, 2) & "\mp3\" & left(searchStrings, 10) & "*.mp3 " & "e:\ANA20" & activeWSNameMid & "\Bluebox\ >> Result_WiFi_copy.txt"
	End If

	' 曲数をカウントしてみる
	processFiles = processFiles + CLng(searchStrings.offset(0,11).Value)

	' 番組数をカウントしてみる
	totalPrograms = totalPrograms + 1

End If
Next searchStrings


' 後処理
	Print #1, "echo ファイルのコピーが完了しました。"
	Print #1, "echo 合計 " & totalPrograms & "番組。予想されるファイル数は " & processFiles & "ファイルです。"
	Print #1, "echo 空物語を足したら" & processFiles +1 & "ファイルね。"
	Print #1, "pause"
	Print #1, "ENDLOCAL"
	Print #1, "exit /b"

	Close #1


msgbox "「dataCollection_Bluebox.cmd」できあがり" & vbCr & "Bluebox用は " & totalPrograms & " 番組です"

End Sub




Sub 試聴用ファイルをフォルダに振り分けるPS1を作成()
'-----------------------------------------------------------------------------'
' 当月分のMP3を、番組名のついたフォルダに振り分けるPS1ファイルを作成する
' そんなExcel VBA Version 2.10
' 
' Copyright 2016-2019 Masanori Tanaka (Genonsha Co.,Ltd.)
'-----------------------------------------------------------------------------'
'
' PowerShell用ファイル「dataCollection_mp3.ps1」をデスクトップに保存してくれます。
' これを当月分のMP3ファイルが入っているフォルダに入れて実行すると驚きの結末が！

' 2.10 - [2019/04/16] 振り分けたファイルを元に戻す機能を付ける

Dim stream As New ADODB.Stream
stream.Type = adTypeText ' もしくは2（1はバイナリ）
stream.Charset = "UTF-8" ' 文字コード
stream.LineSeparator = 10 ' 改行コード


' アクティブシートの名前を取得して、様々な場面で活用できるようにしておく
Dim activeWSName As String
Dim activeWSNameMid As String

activeWSName = Replace(ActiveSheet.Name, ".","")
activeWSNameMid = Replace(Mid(ActiveSheet.Name, 3), ".","") ' 例「2015.07」を「1507」に


' ファイル（になる空間）に内容を書き込んでいく
stream.Open

	stream.WriteText "# " & Replace(ActiveSheet.Range("A2").value," ","") & "の番組のみを使用します", 1
	stream.WriteText "# このファイルを" & Replace(ActiveSheet.Range("A2").value," ","") & "のMP3ファイルが入っているフォルダに", 1
	stream.WriteText "# 入れて実行すると、番組名のついたフォルダを作成して適宜ファイルを振り分けます（元に戻す機能付き）", 1
	stream.WriteText "", 1
	stream.WriteText "If (Get-ChildItem -Filter *.mp3) {", 1
	stream.WriteText "", 1
	stream.WriteText "[array] $allProgramNames = @()", 1
	stream.WriteText "", 1

' アクティブシートのファイル名リスト（Ｆ列）を一行ずつ読んでループを回す
' 最終行を自動的に取得するので、番組数が変わっても空きチャンネルがあっても大丈夫
Dim searchStrings As Range
Dim processFiles As Long
Dim totalPrograms As Long
Dim programString As String
Dim programName As String

totalPrograms = 0

For Each searchStrings In ActiveSheet.Range("F5:F" & Cells(Rows.Count, 6).End(xlUp).Row)

If searchStrings Like "nha" & right(activeWSNameMid, 2) & left(activeWSNameMid, 2) & "*" Then ' ファイル名が「nha」のあと当月（例えば0316）だったら

	' フォルダ名として使えない文字があった場合に全角にしておく（原始的！）
	ProgramName = searchStrings.Offset(0, 1).Value
	ProgramName = Replace(ProgramName,"/","／")
	ProgramName = Replace(ProgramName,"\","￥")
	ProgramName = Replace(ProgramName,":","：")
	ProgramName = Replace(ProgramName,"*","＊")
	ProgramName = Replace(ProgramName,"?","？")
	ProgramName = Replace(ProgramName,"""","''")
	ProgramName = Replace(ProgramName,"<","＜")
	ProgramName = Replace(ProgramName,">","＞")
	ProgramName = Replace(ProgramName,"|","｜")
	programName = Replace(programName, Chr$(&H8167), "‘") ' 全角の“”が勝手に半角になりエラーになるので、いっそのこと‘’に
	programName = Replace(ProgramName, Chr$(&H8168), "’")

	'念のためセル内の改行や余分なスペースも除去
	ProgramName = Replace(ProgramName, vbLf, "")
	ProgramName = Trim(ProgramName)


	programString =  """" & FORMAT(searchStrings.Offset(0, -5).Value,"000_") & ProgramName & """"

	' フォルダ名を「001_番組名」みたいにして配列に追加
	stream.WriteText "$allProgramNames += " & programString, 1

	' 番組数をカウントしてみる
	totalPrograms = totalPrograms + 1

End If
Next searchStrings


	' フォルダを作ってファイルを移動させていく
	stream.WriteText "", 1
	stream.WriteText "for ($i=0; $i -lt $allProgramNames.Length; $i++)", 1
	stream.WriteText "{", 1
	stream.WriteText vbTab & "New-Item $allProgramNames[$i] -Force -ItemType Directory", 1
	stream.WriteText vbTab & "$moji = $allProgramNames[$i].Substring(0, 3) + ""*.mp3""", 1
	stream.WriteText vbTab & "Get-ChildItem ./nha" & right(activeWSNameMid, 2) & left(activeWSNameMid, 2) & "$moji" & " | Move-Item -Destination $allProgramNames[$i]", 1
	stream.WriteText "}", 1
	stream.WriteText "", 1

	'トグルっぽい感じでアンドゥする部分
	stream.WriteText "}else{", 1
	stream.WriteText "# 再帰的にMP3を探して、見つかったものをフルパスにして、このスクリプトがある階層に持ってくる", 1
	stream.WriteText "Get-ChildItem -Filter *.mp3 -Recurse | Select-Object FullName | ForEach-Object{Move-Item -literalPath $_.Fullname .\}", 1
	stream.WriteText "", 1
	stream.WriteText "# ファイル名の頭が「NH_yymm-W-A-」で始まっていれば、それを消去する", 1
	stream.WriteText "Get-ChildItem *.mp3 | Rename-Item -NewName { $_.Name -replace 'NH_\d{4}-W-A-','' }", 1
	stream.WriteText "", 1
	stream.WriteText "# 最後にフォルダを削除（念のため、フォルダ内にファイルが残っていたら止める）", 1
	stream.WriteText "Get-ChildItem -Directory | Remove-Item -Recurse", 1
	stream.WriteText "Read-Host ""続けるには Enter キーを押してください...""", 1
	stream.WriteText "}", 1
	stream.WriteText "", 1

' 指定したファイルに保存して閉じる
stream.SaveToFile "C:\Users\tanaka\Desktop\dataCollection" & activeWSName & ".ps1", adSaveCreateOverWrite
stream.Close

msgbox "「dataCollection.ps1」できあがり" & vbCr & "今月の新規チャンネルは " & totalPrograms & " 番組です"

End Sub


