'//////////////////////////////////////////////////////////////////////////////////////
' -------------------------------------------------------------------------------------
' 項目数が不定の可変長ファイルを所定の項目数に揃えるため、末尾に区切り文字を追加する
' -------------------------------------------------------------------------------------
' [引数]：なし
' -------------------------------------------------------------------------------------
'//////////////////////////////////////////////////////////////////////////////////////
Option Explicit

'//////////////////////////////////////////////////////////////////////////////////////
' 定数宣言部
'//////////////////////////////////////////////////////////////////////////////////////
' 実行環境に合わせて編集する ここから -------------------------------------------------

' 変換元ファイルが格納されたフォルダのパス (末尾の\は不要)
Private Const IN_FILE_DIRECTORY_PATH = "C:XXXXXXXXXXXXX\Input"
' 変換後ファイルを格納するフォルダのパス (末尾の\は不要)
Private Const OUT_FILE_DIRECTORY_PATH = "C:XXXXXXXXXXXXX\Output"
' 区切り文字 (および末尾に補完する区切り記号) (カンマの場合は"C"を指定する)
Private Const DELIMITER_STRING = "T"
' 最大項目数
Private Const MAX_ITEM_COUNT = 33

' 実行環境に合わせて編集する ここまで -------------------------------------------------
'//////////////////////////////////////////////////////////////////////////////////////
' 変数宣言部
'//////////////////////////////////////////////////////////////////////////////////////
Dim inputTextFile
Dim outputTextFile
Dim currentTextLine
Dim splitLine
Dim delimiter
Dim fso
Dim objInDirecotry
Dim objInFile
Dim appendCount
Dim i

'//////////////////////////////////////////////////////////////////////////////////////
' メインロジック
'//////////////////////////////////////////////////////////////////////////////////////
Wscript.Echo "入力フォルダ： " & IN_FILE_DIRECTORY_PATH
Wscript.Echo "出力フォルダ： " & OUT_FILE_DIRECTORY_PATH

If DELIMITER_STRING = "C" Then
    delimiter = ","
ElseIf DELIMITER_STRING = "T" Then
    delimiter = vbTab
Else
    Wscript.Echo "第3引数不正。区切り文字は、既定値のタブを使用します"
	delimiter = vbTab
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set objInDirecotry = fso.GetFolder(IN_FILE_DIRECTORY_PATH)

For Each objInFile in objInDirecotry.Files
    Wscript.Echo Now() & " " & objInFile.Path & " の変換を開始しました"
    Set inputTextFile = fso.OpenTextFile(objInFile.Path)
	' 出力ファイルを作成
	Set outputTextFile = fso.CreateTextFile(OUT_FILE_DIRECTORY_PATH & "\" & objInFile.Name)
	
	' 入力ファイルを1行ずつ読み取る
	Do Until inputTextFile.AtEndOfStream
		currentTextLine = inputTextFile.ReadLine
	    ' 読み取った1行を、パラメータで指定されたデリミタを使って配列に展開する
		splitLine = Split(currentTextLine, delimiter)
		
		' 区切った結果、最大項目数より多い = 想定外の形式
		If UBound(splitLine) + 1 > MAX_ITEM_COUNT Then
		    Wscript.Echo Now() & " " & "想定される最大項目数を超えるレコードがあります。ファイル名:" & objInFile.Name
			' 現在の入出力ファイルをクローズする
			If (inputTextFile Is Nothing) = False Then
			    inputTextFile.Close
		    End If
			If (outputTextFile Is Nothing) = False Then
			    outputTextFile.Close
		    End If
			WScript.Quit 99
		' 最大項目数と一致 = 何もしないでそのまま書き込む
		ElseIf UBound(splitLine) + 1 = MAX_ITEM_COUNT Then
		    outputTextFile.WriteLine (currentTextLine)
		' 最大項目数より少ない = 区切り文字を補完する
		Else
			outputTextFile.WriteLine (AppendDelimiter(splitLine, delimiter))
		End If
	Loop
	' 現在の入出力ファイルをクローズする
    If (inputTextFile Is Nothing) = False Then
        inputTextFile.Close
    End If
    If (outputTextFile Is Nothing) = False Then
        outputTextFile.Close
    End If
	Wscript.Echo Now() & " " & objInFile.Path & " の変換が終了しました"
Next

Wscript.Echo Now() & " 全体処理完了！"

Set objInDirecotry = Nothing
Set objInFile = Nothing
Set fso = Nothing
Set inputTextFile = Nothing
Set outputTextFile = Nothing

Function AppendDelimiter(splitLine, delimiter)
    Dim appndedLine
    Dim appendCount
    Dim i
    
    ' 最大項目数から、いくつの区切り文字を末尾に追加するかを計算
    appendCount = MAX_ITEM_COUNT - (UBound(splitLine) + 1)
    
    ' 配列で渡された行の内容を文字列に戻す
    appndedLine = Join(splitLine, delimiter)
    
    ' 末尾に区切り文字を追加して整形
    For i = 1 To appendCount
        appndedLine = appndedLine & delimiter
    Next
    
    AppendDelimiter = appndedLine
End Function