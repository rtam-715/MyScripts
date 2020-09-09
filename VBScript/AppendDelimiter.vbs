'//////////////////////////////////////////////////////////////////////////////////////
' -------------------------------------------------------------------------------------
' ���ڐ����s��̉ϒ��t�@�C��������̍��ڐ��ɑ����邽�߁A�����ɋ�؂蕶����ǉ�����
' -------------------------------------------------------------------------------------
' [����]�F�Ȃ�
' -------------------------------------------------------------------------------------
'//////////////////////////////////////////////////////////////////////////////////////
Option Explicit

'//////////////////////////////////////////////////////////////////////////////////////
' �萔�錾��
'//////////////////////////////////////////////////////////////////////////////////////
' ���s���ɍ��킹�ĕҏW���� �������� -------------------------------------------------

' �ϊ����t�@�C�����i�[���ꂽ�t�H���_�̃p�X (������\�͕s�v)
Private Const IN_FILE_DIRECTORY_PATH = "C:XXXXXXXXXXXXX\Input"
' �ϊ���t�@�C�����i�[����t�H���_�̃p�X (������\�͕s�v)
Private Const OUT_FILE_DIRECTORY_PATH = "C:XXXXXXXXXXXXX\Output"
' ��؂蕶�� (����і����ɕ⊮�����؂�L��) (�J���}�̏ꍇ��"C"���w�肷��)
Private Const DELIMITER_STRING = "T"
' �ő區�ڐ�
Private Const MAX_ITEM_COUNT = 33

' ���s���ɍ��킹�ĕҏW���� �����܂� -------------------------------------------------
'//////////////////////////////////////////////////////////////////////////////////////
' �ϐ��錾��
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
' ���C�����W�b�N
'//////////////////////////////////////////////////////////////////////////////////////
Wscript.Echo "���̓t�H���_�F " & IN_FILE_DIRECTORY_PATH
Wscript.Echo "�o�̓t�H���_�F " & OUT_FILE_DIRECTORY_PATH

If DELIMITER_STRING = "C" Then
    delimiter = ","
ElseIf DELIMITER_STRING = "T" Then
    delimiter = vbTab
Else
    Wscript.Echo "��3�����s���B��؂蕶���́A����l�̃^�u���g�p���܂�"
	delimiter = vbTab
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set objInDirecotry = fso.GetFolder(IN_FILE_DIRECTORY_PATH)

For Each objInFile in objInDirecotry.Files
    Wscript.Echo Now() & " " & objInFile.Path & " �̕ϊ����J�n���܂���"
    Set inputTextFile = fso.OpenTextFile(objInFile.Path)
	' �o�̓t�@�C�����쐬
	Set outputTextFile = fso.CreateTextFile(OUT_FILE_DIRECTORY_PATH & "\" & objInFile.Name)
	
	' ���̓t�@�C����1�s���ǂݎ��
	Do Until inputTextFile.AtEndOfStream
		currentTextLine = inputTextFile.ReadLine
	    ' �ǂݎ����1�s���A�p�����[�^�Ŏw�肳�ꂽ�f���~�^���g���Ĕz��ɓW�J����
		splitLine = Split(currentTextLine, delimiter)
		
		' ��؂������ʁA�ő區�ڐ���葽�� = �z��O�̌`��
		If UBound(splitLine) + 1 > MAX_ITEM_COUNT Then
		    Wscript.Echo Now() & " " & "�z�肳���ő區�ڐ��𒴂��郌�R�[�h������܂��B�t�@�C����:" & objInFile.Name
			' ���݂̓��o�̓t�@�C�����N���[�Y����
			If (inputTextFile Is Nothing) = False Then
			    inputTextFile.Close
		    End If
			If (outputTextFile Is Nothing) = False Then
			    outputTextFile.Close
		    End If
			WScript.Quit 99
		' �ő區�ڐ��ƈ�v = �������Ȃ��ł��̂܂܏�������
		ElseIf UBound(splitLine) + 1 = MAX_ITEM_COUNT Then
		    outputTextFile.WriteLine (currentTextLine)
		' �ő區�ڐ���菭�Ȃ� = ��؂蕶����⊮����
		Else
			outputTextFile.WriteLine (AppendDelimiter(splitLine, delimiter))
		End If
	Loop
	' ���݂̓��o�̓t�@�C�����N���[�Y����
    If (inputTextFile Is Nothing) = False Then
        inputTextFile.Close
    End If
    If (outputTextFile Is Nothing) = False Then
        outputTextFile.Close
    End If
	Wscript.Echo Now() & " " & objInFile.Path & " �̕ϊ����I�����܂���"
Next

Wscript.Echo Now() & " �S�̏��������I"

Set objInDirecotry = Nothing
Set objInFile = Nothing
Set fso = Nothing
Set inputTextFile = Nothing
Set outputTextFile = Nothing

Function AppendDelimiter(splitLine, delimiter)
    Dim appndedLine
    Dim appendCount
    Dim i
    
    ' �ő區�ڐ�����A�����̋�؂蕶���𖖔��ɒǉ����邩���v�Z
    appendCount = MAX_ITEM_COUNT - (UBound(splitLine) + 1)
    
    ' �z��œn���ꂽ�s�̓��e�𕶎���ɖ߂�
    appndedLine = Join(splitLine, delimiter)
    
    ' �����ɋ�؂蕶����ǉ����Đ��`
    For i = 1 To appendCount
        appndedLine = appndedLine & delimiter
    Next
    
    AppendDelimiter = appndedLine
End Function