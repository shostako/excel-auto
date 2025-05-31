Attribute VB_Name = "m�]�L_�V�[�gTG�i�ԕ�"
Option Explicit

Sub �]�L_�V�[�gTG�i�ԕ�()
    ' �������ݒ�
    Application.ScreenUpdating = False
    Application.StatusBar = "TG�i�ԕʓ]�L�������J�n..."
    
    On Error GoTo ErrorHandler
    
    ' �e�[�u���擾
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("TG�i�ԕ�")
    
    Dim srcTable As ListObject, tgtTable As ListObject
    Set srcTable = ws.ListObjects("_TG�i�ԕ�a")
    Set tgtTable = ws.ListObjects("_TG�i�ԕ�b")
    
    ' �f�[�^�͈̓`�F�b�N
    If srcTable.DataBodyRange Is Nothing Then
        Application.StatusBar = "�\�[�X�e�[�u���Ƀf�[�^�Ȃ�"
        GoTo Cleanup
    End If
    
    ' �K�v�ȗ�C���f�b�N�X�擾
    Dim srcCols As Object
    Set srcCols = CreateObject("Scripting.Dictionary")
    srcCols("���t") = srcTable.ListColumns("���t").Index
    srcCols("�i��") = srcTable.ListColumns("�i��").Index
    srcCols("����") = srcTable.ListColumns("����").Index
    srcCols("�s��") = srcTable.ListColumns("�s��").Index
    srcCols("�ғ�����") = srcTable.ListColumns("�ғ�����").Index
    
    ' �]�L��̗�C���f�b�N�X�擾
    Dim tgtCols As Object
    Set tgtCols = CreateObject("Scripting.Dictionary")
    tgtCols("���t") = tgtTable.ListColumns("���t").Index
    
    ' RH�p��
    tgtCols("RH������") = GetColumnIndexSafe(tgtTable, "RH������")
    tgtCols("RH���s�ǎ���") = GetColumnIndexSafe(tgtTable, "RH���s�ǎ���")
    tgtCols("RH���ғ�����") = GetColumnIndexSafe(tgtTable, "RH���ғ�����")
    
    ' LH�p��
    tgtCols("LH������") = GetColumnIndexSafe(tgtTable, "LH������")
    tgtCols("LH���s�ǎ���") = GetColumnIndexSafe(tgtTable, "LH���s�ǎ���")
    tgtCols("LH���ғ�����") = GetColumnIndexSafe(tgtTable, "LH���ғ�����")
    
    ' ���v�p��
    tgtCols("���v������") = GetColumnIndexSafe(tgtTable, "���v������")
    tgtCols("���v���s�ǎ���") = GetColumnIndexSafe(tgtTable, "���v���s�ǎ���")
    tgtCols("���v���ғ�����") = GetColumnIndexSafe(tgtTable, "���v���ғ�����")
    
    ' �f�[�^�]�L����
    Dim srcData As Range, tgtData As Range
    Set srcData = srcTable.DataBodyRange
    Set tgtData = tgtTable.DataBodyRange
    
    If tgtData Is Nothing Then
        Application.StatusBar = "�]�L��e�[�u������"
        GoTo Cleanup
    End If
    
    ' ���t���Ƃ̍��v�l���i�[���鎫��
    Dim dailyTotals As Object
    Set dailyTotals = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, j As Long
    Dim srcDate As Date, hinban As String
    Dim transferred As Long: transferred = 0
    Dim totalRows As Long: totalRows = srcData.Rows.Count
    
    ' �܂��A���t���Ƃ̍��v���v�Z
    Application.StatusBar = "���v�l���v�Z��..."
    For i = 1 To totalRows
        srcDate = srcData.Cells(i, srcCols("���t")).Value
        hinban = Trim(srcData.Cells(i, srcCols("�i��")).Value)
        
        ' �Ώەi�Ԃ̂ݏ���
        If hinban = "53827-60050" Or hinban = "53828-60080" Then
            Dim dateKey As String
            dateKey = Format(srcDate, "yyyy-mm-dd")
            
            ' ���t�L�[�����݂��Ȃ��ꍇ�͏�����
            If Not dailyTotals.Exists(dateKey) Then
                dailyTotals(dateKey) = Array(0, 0, 0) ' ���сA�s�ǁA�ғ����Ԃ̏�
            End If
            
            ' ���v�l�����Z
            Dim totals As Variant
            totals = dailyTotals(dateKey)
            totals(0) = totals(0) + srcData.Cells(i, srcCols("����")).Value
            totals(1) = totals(1) + srcData.Cells(i, srcCols("�s��")).Value
            totals(2) = totals(2) + srcData.Cells(i, srcCols("�ғ�����")).Value
            dailyTotals(dateKey) = totals
        End If
    Next i
    
    ' �]�L��e�[�u������U�N���A�i�i�ԕʃf�[�^�j
    Application.StatusBar = "�]�L����N���A��..."
    For j = 1 To tgtData.Rows.Count
        ' RH��̃N���A
        If tgtCols("RH������") > 0 Then tgtData.Cells(j, tgtCols("RH������")).ClearContents
        If tgtCols("RH���s�ǎ���") > 0 Then tgtData.Cells(j, tgtCols("RH���s�ǎ���")).ClearContents
        If tgtCols("RH���ғ�����") > 0 Then tgtData.Cells(j, tgtCols("RH���ғ�����")).ClearContents
        
        ' LH��̃N���A
        If tgtCols("LH������") > 0 Then tgtData.Cells(j, tgtCols("LH������")).ClearContents
        If tgtCols("LH���s�ǎ���") > 0 Then tgtData.Cells(j, tgtCols("LH���s�ǎ���")).ClearContents
        If tgtCols("LH���ғ�����") > 0 Then tgtData.Cells(j, tgtCols("LH���ғ�����")).ClearContents
        
        ' ���v��̃N���A
        If tgtCols("���v������") > 0 Then tgtData.Cells(j, tgtCols("���v������")).ClearContents
        If tgtCols("���v���s�ǎ���") > 0 Then tgtData.Cells(j, tgtCols("���v���s�ǎ���")).ClearContents
        If tgtCols("���v���ғ�����") > 0 Then tgtData.Cells(j, tgtCols("���v���ғ�����")).ClearContents
    Next j
    
    ' �i�ԕʃf�[�^�̓]�L
    For i = 1 To totalRows
        ' �i���\���i10�s���ƂɍX�V���ď������x�D��j
        If i Mod 10 = 0 Or i = totalRows Then
            Application.StatusBar = "TG�i�ԕʓ]�L������... " & Format(i / totalRows, "0%") & _
                                  " (" & i & "/" & totalRows & "�s)"
            DoEvents ' ��ʍX�V
        End If
        
        ' �\�[�X�f�[�^�擾
        srcDate = srcData.Cells(i, srcCols("���t")).Value
        hinban = Trim(srcData.Cells(i, srcCols("�i��")).Value)
        
        ' �i�Ԃ��ΏۊO�Ȃ�X�L�b�v
        If hinban <> "53827-60050" And hinban <> "53828-60080" Then
            GoTo NextRow
        End If
        
        ' �]�L��̓��t����
        For j = 1 To tgtData.Rows.Count
            If tgtData.Cells(j, tgtCols("���t")).Value = srcDate Then
                ' �i�Ԃɉ����ē]�L
                If hinban = "53827-60050" Then
                    ' RH�i�Ԃ̓]�L
                    If tgtCols("RH������") > 0 Then
                        tgtData.Cells(j, tgtCols("RH������")).Value = srcData.Cells(i, srcCols("����")).Value
                    End If
                    If tgtCols("RH���s�ǎ���") > 0 Then
                        tgtData.Cells(j, tgtCols("RH���s�ǎ���")).Value = srcData.Cells(i, srcCols("�s��")).Value
                    End If
                    If tgtCols("RH���ғ�����") > 0 Then
                        tgtData.Cells(j, tgtCols("RH���ғ�����")).Value = srcData.Cells(i, srcCols("�ғ�����")).Value
                    End If
                    transferred = transferred + 1
                    
                ElseIf hinban = "53828-60080" Then
                    ' LH�i�Ԃ̓]�L
                    If tgtCols("LH������") > 0 Then
                        tgtData.Cells(j, tgtCols("LH������")).Value = srcData.Cells(i, srcCols("����")).Value
                    End If
                    If tgtCols("LH���s�ǎ���") > 0 Then
                        tgtData.Cells(j, tgtCols("LH���s�ǎ���")).Value = srcData.Cells(i, srcCols("�s��")).Value
                    End If
                    If tgtCols("LH���ғ�����") > 0 Then
                        tgtData.Cells(j, tgtCols("LH���ғ�����")).Value = srcData.Cells(i, srcCols("�ғ�����")).Value
                    End If
                    transferred = transferred + 1
                End If
                
                Exit For ' ���t�����������玟�̍s��
            End If
        Next j
        
NextRow:
    Next i
    
    ' ���v�l�̓]�L
    Application.StatusBar = "���v�l��]�L��..."
    Dim totalTransferred As Long: totalTransferred = 0
    For j = 1 To tgtData.Rows.Count
        Dim tgtDate As Date
        tgtDate = tgtData.Cells(j, tgtCols("���t")).Value
        dateKey = Format(tgtDate, "yyyy-mm-dd")
        
        If dailyTotals.Exists(dateKey) Then
            totals = dailyTotals(dateKey)
            
            ' ���v�l��]�L
            If tgtCols("���v������") > 0 Then
                tgtData.Cells(j, tgtCols("���v������")).Value = totals(0)
            End If
            If tgtCols("���v���s�ǎ���") > 0 Then
                tgtData.Cells(j, tgtCols("���v���s�ǎ���")).Value = totals(1)
            End If
            If tgtCols("���v���ғ�����") > 0 Then
                tgtData.Cells(j, tgtCols("���v���ғ�����")).Value = totals(2)
            End If
            totalTransferred = totalTransferred + 1
        End If
    Next j
    
    ' �������̃X�e�[�^�X�o�[�\��
    Application.StatusBar = "TG�i�ԕʓ]�L����: " & transferred & "���̕i�ԕʃf�[�^�A" & _
                           totalTransferred & "���̍��v�f�[�^��]�L"
    
    ' 1�b�ҋ@���Ă���N���A
    Application.Wait Now + TimeValue("0:00:01")
    Application.StatusBar = False
    
Cleanup:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    ' �G���[���������b�Z�[�W�{�b�N�X
    MsgBox "TG�i�ԕʓ]�L�����ŃG���[����" & vbCrLf & vbCrLf & _
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

