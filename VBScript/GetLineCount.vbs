Option Explicit

'////////////////////////////////////////////////////////////////////////////////
' 変数定義部
'////////////////////////////////////////////////////////////////////////////////
Dim fso
Dim targetDirectory
Dim filter
Dim objDirectory
Dim objFile
Dim sumLineCount

'////////////////////////////////////////////////////////////////////////////////
' メインロジック
'////////////////////////////////////////////////////////////////////////////////
Set fso = CreateObject("Scripting.FileSystemObject")

'　パラメータでフォルダが指定されていれば適用し、そうでない場合はスクリプトと同階層のファイルを対象とする
If WScript.Arguments.Length > 0 Then
	targetDirectory = WScript.Arguments(0)
Else
	targetDirectory fso.getParentFolderName(WScript.ScriptFullName)
End If

'　拡張子を決定する。パラメータで指定がない場合は、.txt を対象とする
If WScript.Arguments.Length > 1 Then
	filter = UCase(WScript.Arguments(1))
Else
	filter = "TXT"
End If

' フォルダパスからファイルのリストを取得
Set objDirectory = fso.GetFolder(targetDirectory)
For Each objFile In objDirectory.Files
	If UCase(fso.GetExtensionName(objFile.Path)) = filter Then
		With fso.OpenTextFile(objFile.Path, 8)
			sumLineCount = sumLineCount + .Line - 1
			.Close
		End With
	End If
Next

Wscript.Echo "拡張子[" & filter & "]：合計行数[" & sumLineCount &"]"

'////////////////////////////////////////////////////////////////////////////////
' 後始末
'////////////////////////////////////////////////////////////////////////////////
Set fso = Nothing
Set objDirectory = Nothing
Set objFile = Nothing