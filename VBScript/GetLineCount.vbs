Option Explicit

'////////////////////////////////////////////////////////////////////////////////
' �ϐ���`��
'////////////////////////////////////////////////////////////////////////////////
Dim fso
Dim targetDirectory
Dim filter
Dim objDirectory
Dim objFile
Dim sumLineCount

'////////////////////////////////////////////////////////////////////////////////
' ���C�����W�b�N
'////////////////////////////////////////////////////////////////////////////////
Set fso = CreateObject("Scripting.FileSystemObject")

'�@�p�����[�^�Ńt�H���_���w�肳��Ă���ΓK�p���A�����łȂ��ꍇ�̓X�N���v�g�Ɠ��K�w�̃t�@�C����ΏۂƂ���
If WScript.Arguments.Length > 0 Then
	targetDirectory = WScript.Arguments(0)
Else
	targetDirectory fso.getParentFolderName(WScript.ScriptFullName)
End If

'�@�g���q�����肷��B�p�����[�^�Ŏw�肪�Ȃ��ꍇ�́A.txt ��ΏۂƂ���
If WScript.Arguments.Length > 1 Then
	filter = UCase(WScript.Arguments(1))
Else
	filter = "TXT"
End If

' �t�H���_�p�X����t�@�C���̃��X�g���擾
Set objDirectory = fso.GetFolder(targetDirectory)
For Each objFile In objDirectory.Files
	If UCase(fso.GetExtensionName(objFile.Path)) = filter Then
		With fso.OpenTextFile(objFile.Path, 8)
			sumLineCount = sumLineCount + .Line - 1
			.Close
		End With
	End If
Next

Wscript.Echo "�g���q[" & filter & "]�F���v�s��[" & sumLineCount &"]"

'////////////////////////////////////////////////////////////////////////////////
' ��n��
'////////////////////////////////////////////////////////////////////////////////
Set fso = Nothing
Set objDirectory = Nothing
Set objFile = Nothing