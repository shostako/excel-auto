Attribute VB_Name = "m�]�L_�V�[�g���`���@��"
Option Explicit

Sub �]�L_�V�[�g���`���@��()
    ' �������ݒ�
    Application.ScreenUpdating = False
    Application.StatusBar = "���`���@�ʓ]�L�������J�n..."
    
    On Error GoTo ErrorHandler
    
    ' �e�[�u���擾
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("���`���@��")
    
    Dim srcTable As ListObject, tgtTable As ListObject
    Set srcTable = ws.ListObjects("_���`���@��a")
    Set tgtTable = ws.ListObjects("_���`���@��b")
    
    ' �f�[�^�͈̓`�F�b�N
    If srcTable.DataBodyRange Is Nothing Then
        Application.StatusBar = "�\�[�X�e�[�u���Ƀf�[�^�Ȃ�"
        GoTo Cleanup
    End If
    
    ' �K�v�ȗ�C���f�b�N�X�擾�i�\�[�X�j
    Dim srcCols As Object
    Set srcCols = CreateObject("Scripting.Dictionary")
    srcCols("���t") = srcTable.ListColumns("���t").Index
    srcCols("�@�B") = srcTable.ListColumns("�@�B").Index
    srcCols("����") = srcTable.ListColumns("����").Index
    srcCols("�s��") = srcTable.ListColumns("�s��").Index
    srcCols("�ғ�����") = srcTable.ListColumns("�ғ�����").Index
    
    ' �]�L��̗�C���f�b�N�X�擾
    Dim tgtCols As Object
    Set tgtCols = CreateObject("Scripting.Dictionary")
    tgtCols("���t") = tgtTable.ListColumns("���t").Index
    
    ' 1���@�p��
    tgtCols("1���@������") = GetColumnIndexSafe(tgtTable, "1���@������")
    tgtCols("1���@���s�ǎ���") = GetColumnIndexSafe(tgtTable, "1���@���s�ǎ���")
    tgtCols("1���@���ғ�����") = GetColumnIndexSafe(tgtTable, "1���@���ғ�����")
    
    ' 2���@�p��
    tgtCols("2���@������") = GetColumnIndexSafe(tgtTable, "2���@������")
    tgtCols("2���@���s�ǎ���") = GetColumnIndexSafe(tgtTable, "2���@���s�ǎ���")
    tgtCols("2���@���ғ�����") = GetColumnIndexSafe(tgtTable, "2���@���ғ�����")
    
    ' 3���@�p��
    tgtCols("3���@������") = GetColumnIndexSafe(tgtTable, "3���@������")
    tgtCols("3���@���s�ǎ���") = GetColumnIndexSafe(tgtTable, "3���@���s�ǎ���")
    tgtCols("3���@���ғ�����") = GetColumnIndexSafe(tgtTable, "3���@���ғ�����")
    
    ' 4���@�p��
    tgtCols("4���@������") = GetColumnIndexSafe(tgtTable, "4���@������")
    tgtCols("4���@���s�ǎ���") = GetColumnIndexSafe(tgtTable, "4���@���s�ǎ���")
    tgtCols("4���@���ғ�����") = GetColumnIndexSafe(tgtTable, "4���@���ғ�����")
    
    ' 5���@�p��
    tgtCols("5���@������") = GetColumnIndexSafe(tgtTable, "5���@������")
    tgtCols("5���@���s�ǎ���") = GetColumnIndexSafe(tgtTable, "5���@���s�ǎ���")
    tgtCols("5���@���ғ�����") = GetColumnIndexSafe(tgtTable, "5���@���ғ�����")
    
    ' �f�[�^�]�L����
    Dim srcData As Range, tgtData As Range
    Set srcData = srcTable.DataBodyRange
    Set tgtData = tgtTable.DataBodyRange
    
    If tgtData Is Nothing Then
        Application.StatusBar = "�]�L��e�[�u������"
        GoTo Cleanup
    End If
    
    Dim i As Long, j As Long
    Dim srcDate As Date, machine As String
    Dim transferred As Long: transferred = 0
    Dim totalRows As Long: totalRows = srcData.Rows.Count
    
    ' �]�L��e�[�u������U�N���A
    Application.StatusBar = "�]�L����N���A��..."
    For j = 1 To tgtData.Rows.Count
        ' 1���@��̃N���A
        If tgtCols("1���@������") > 0 Then tgtData.Cells(j, tgtCols("1���@������")).ClearContents
        If tgtCols("1���@���s�ǎ���") > 0 Then tgtData.Cells(j, tgtCols("1���@���s�ǎ���")).ClearContents
        If tgtCols("1���@���ғ�����") > 0 Then tgtData.Cells(j, tgtCols("1���@���ғ�����")).ClearContents
        
        ' 2���@��̃N���A
        If tgtCols("2���@������") > 0 Then tgtData.Cells(j, tgtCols("2���@������")).ClearContents
        If tgtCols("2���@���s�ǎ���") > 0 Then tgtData.Cells(j, tgtCols("2���@���s�ǎ���")).ClearContents
        If tgtCols("2���@���ғ�����") > 0 Then tgtData.Cells(j, tgtCols("2���@���ғ�����")).ClearContents
        
        ' 3���@��̃N���A
        If tgtCols("3���@������") > 0 Then tgtData.Cells(j, tgtCols("3���@������")).ClearContents
        If tgtCols("3���@���s�ǎ���") > 0 Then tgtData.Cells(j, tgtCols("3���@���s�ǎ���")).ClearContents
        If tgtCols("3���@���ғ�����") > 0 Then tgtData.Cells(j, tgtCols("3���@���ғ�����")).ClearContents
        
        ' 4���@��̃N���A
        If tgtCols("4���@������") > 0 Then tgtData.Cells(j, tgtCols("4���@������")).ClearContents
        If tgtCols("4���@���s�ǎ���") > 0 Then tgtData.Cells(j, tgtCols("4���@���s�ǎ���")).ClearContents
        If tgtCols("4���@���ғ�����") > 0 Then tgtData.Cells(j, tgtCols("4���@���ғ�����")).ClearContents
        
        ' 5���@��̃N���A
        If tgtCols("5���@������") > 0 Then tgtData.Cells(j, tgtCols("5���@������")).ClearContents
        If tgtCols("5���@���s�ǎ���") > 0 Then tgtData.Cells(j, tgtCols("5���@���s�ǎ���")).ClearContents
        If tgtCols("5���@���ғ�����") > 0 Then tgtData.Cells(j, tgtCols("5���@���ғ�����")).ClearContents
    Next j
    
    ' �@�B�ʃf�[�^�̓]�L
    For i = 1 To totalRows
        ' �i���\���i10�s���ƂɍX�V���ď������x�D��j
        If i Mod 10 = 0 Or i = totalRows Then
            Application.StatusBar = "���`���@�ʓ]�L������... " & Format(i / totalRows, "0%") & _
                                  " (" & i & "/" & totalRows & "�s)"
            DoEvents ' ��ʍX�V
        End If
        
        ' �\�[�X�f�[�^�擾
        srcDate = srcData.Cells(i, srcCols("���t")).Value
        machine = Trim(srcData.Cells(i, srcCols("�@�B")).Value)
        
        ' �@�B���ΏۊO�Ȃ�X�L�b�v
        If machine <> "SS01" And machine <> "SS02" And machine <> "SS03" And _
           machine <> "SS04" And machine <> "SS05" Then
            GoTo NextRow
        End If
        
        ' �]�L��̓��t����
        For j = 1 To tgtData.Rows.Count
            If tgtData.Cells(j, tgtCols("���t")).Value = srcDate Then
                ' �@�B�ɉ����ē]�L
                Select Case machine
                    Case "SS01"
                        ' 1���@�̓]�L
                        If tgtCols("1���@������") > 0 Then
                            tgtData.Cells(j, tgtCols("1���@������")).Value = srcData.Cells(i, srcCols("����")).Value
                        End If
                        If tgtCols("1���@���s�ǎ���") > 0 Then
                            tgtData.Cells(j, tgtCols("1���@���s�ǎ���")).Value = srcData.Cells(i, srcCols("�s��")).Value
                        End If
                        If tgtCols("1���@���ғ�����") > 0 Then
                            tgtData.Cells(j, tgtCols("1���@���ғ�����")).Value = srcData.Cells(i, srcCols("�ғ�����")).Value
                        End If
                        transferred = transferred + 1
                        
                    Case "SS02"
                        ' 2���@�̓]�L
                        If tgtCols("2���@������") > 0 Then
                            tgtData.Cells(j, tgtCols("2���@������")).Value = srcData.Cells(i, srcCols("����")).Value
                        End If
                        If tgtCols("2���@���s�ǎ���") > 0 Then
                            tgtData.Cells(j, tgtCols("2���@���s�ǎ���")).Value = srcData.Cells(i, srcCols("�s��")).Value
                        End If
                        If tgtCols("2���@���ғ�����") > 0 Then
                            tgtData.Cells(j, tgtCols("2���@���ғ�����")).Value = srcData.Cells(i, srcCols("�ғ�����")).Value
                        End If
                        transferred = transferred + 1
                        
                    Case "SS03"
                        ' 3���@�̓]�L
                        If tgtCols("3���@������") > 0 Then
                            tgtData.Cells(j, tgtCols("3���@������")).Value = srcData.Cells(i, srcCols("����")).Value
                        End If
                        If tgtCols("3���@���s�ǎ���") > 0 Then
                            tgtData.Cells(j, tgtCols("3���@���s�ǎ���")).Value = srcData.Cells(i, srcCols("�s��")).Value
                        End If
                        If tgtCols("3���@���ғ�����") > 0 Then
                            tgtData.Cells(j, tgtCols("3���@���ғ�����")).Value = srcData.Cells(i, srcCols("�ғ�����")).Value
                        End If
                        transferred = transferred + 1
                        
                    Case "SS04"
                        ' 4���@�̓]�L
                        If tgtCols("4���@������") > 0 Then
                            tgtData.Cells(j, tgtCols("4���@������")).Value = srcData.Cells(i, srcCols("����")).Value
                        End If
                        If tgtCols("4���@���s�ǎ���") > 0 Then
                            tgtData.Cells(j, tgtCols("4���@���s�ǎ���")).Value = srcData.Cells(i, srcCols("�s��")).Value
                        End If
                        If tgtCols("4���@���ғ�����") > 0 Then
                            tgtData.Cells(j, tgtCols("4���@���ғ�����")).Value = srcData.Cells(i, srcCols("�ғ�����")).Value
                        End If
                        transferred = transferred + 1
                        
                    Case "SS05"
                        ' 5���@�̓]�L
                        If tgtCols("5���@������") > 0 Then
                            tgtData.Cells(j, tgtCols("5���@������")).Value = srcData.Cells(i, srcCols("����")).Value
                        End If
                        If tgtCols("5���@���s�ǎ���") > 0 Then
                            tgtData.Cells(j, tgtCols("5���@���s�ǎ���")).Value = srcData.Cells(i, srcCols("�s��")).Value
                        End If
                        If tgtCols("5���@���ғ�����") > 0 Then
                            tgtData.Cells(j, tgtCols("5���@���ғ�����")).Value = srcData.Cells(i, srcCols("�ғ�����")).Value
                        End If
                        transferred = transferred + 1
                End Select
                
                Exit For ' ���t�����������玟�̍s��
            End If
        Next j
        
NextRow:
    Next i
    
    ' �������̃X�e�[�^�X�o�[�\��
    Application.StatusBar = "���`���@�ʓ]�L����: " & transferred & "���̃f�[�^��]�L"
    
    ' 1�b�ҋ@���Ă���N���A
    Application.Wait Now + TimeValue("0:00:01")
    Application.StatusBar = False
    
Cleanup:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    ' �G���[���������b�Z�[�W�{�b�N�X
    MsgBox "���`���@�ʓ]�L�����ŃG���[����" & vbCrLf & vbCrLf & _
           "�G���[���e: " & Err.Description & vbCrLf & _
           "�G���[�ԍ�: " & Err.Number, vbCritical, "�]�L�G���["
    Application.StatusBar = False
    Resume Cleanup
End Sub

' ��C���f�b�N�X�����S�Ɏ擾����w���p�[�֐�
Private Function GetColumnIndexSafe(tbl As ListObject, colName As String) As Long
    On Error Resume Next
    GetColumnIndexSafe = tbl.ListColumns(colName).Index
    If Err.Number <> 0 Then
        GetColumnIndexSafe = 0
        Debug.Print "�x��: ��u" & colName & "�v��������܂���"
    End If
    On Error GoTo 0
End Function

