Attribute VB_Name = "m�]�L_�V�[�g���H�i�ԕ�"
Option Explicit

Sub �]�L_�V�[�g���H�i�ԕ�()
    ' �������ݒ�
    Application.ScreenUpdating = False
    Application.StatusBar = "���H�i�ԕ� �]�L�������J�n..."
    
    On Error GoTo ErrorHandler
    
    ' �e�[�u���擾
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("���H�i�ԕ�")
    
    Dim srcTable As ListObject, tgtTable As ListObject
    Set srcTable = ws.ListObjects("_���H�i�ԕ�a")
    Set tgtTable = ws.ListObjects("_���H�i�ԕ�b")
    
    ' �f�[�^�͈̓`�F�b�N
    If srcTable.DataBodyRange Is Nothing Then
        Application.StatusBar = "�\�[�X�e�[�u���Ƀf�[�^�Ȃ�"
        GoTo Cleanup
    End If
    
    ' �K�v�ȗ�C���f�b�N�X�擾
    Dim srcCols As Object
    Set srcCols = CreateObject("Scripting.Dictionary")
    srcCols("���t") = srcTable.ListColumns("���t").Index
    srcCols("�ʏ�") = srcTable.ListColumns("�ʏ�").Index
    srcCols("����") = srcTable.ListColumns("����").Index
    srcCols("�s��") = srcTable.ListColumns("�s��").Index
    srcCols("�ғ�����") = srcTable.ListColumns("�ғ�����").Index
    
    Dim tgtDateCol As Long
    tgtDateCol = tgtTable.ListColumns("���t").Index
    
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
    
    ' ���t�E�ʏ̂��Ƃ̍��v�l���i�[���鎫���i���ꂪ�C���̃|�C���g�j
    Dim dateNicknameTotals As Object
    Set dateNicknameTotals = CreateObject("Scripting.Dictionary")
    
    ' ���t���Ƃɏo�������ʏ̂��L�^���鎫���i�]�L��N���A�p�j
    Dim dailyNicknames As Object
    Set dailyNicknames = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, j As Long
    Dim srcDate As Date, nickname As String
    Dim transferred As Long: transferred = 0
    Dim totalRows As Long: totalRows = srcData.Rows.Count
    
    ' �܂��A���t�E�ʏ̂��Ƃ̍��v�l���v�Z
    Application.StatusBar = "�f�[�^���W�v��..."
    For i = 1 To totalRows
        srcDate = srcData.Cells(i, srcCols("���t")).Value
        nickname = Trim(srcData.Cells(i, srcCols("�ʏ�")).Value)
        
        If nickname <> "" Then
            Dim dateKey As String
            dateKey = Format(srcDate, "yyyy-mm-dd")
            
            ' ���t�E�ʏ̂��L�[�Ƃ��������L�[
            Dim compositeKey As String
            compositeKey = dateKey & "|" & nickname
            
            ' ���t�E�ʏ̂��Ƃ̏W�v
            If Not dateNicknameTotals.Exists(compositeKey) Then
                dateNicknameTotals(compositeKey) = Array(0, 0, 0) ' ���сA�s�ǁA�ғ����Ԃ̏�
            End If
            
            Dim nicknameTotals As Variant
            nicknameTotals = dateNicknameTotals(compositeKey)
            nicknameTotals(0) = nicknameTotals(0) + srcData.Cells(i, srcCols("����")).Value
            nicknameTotals(1) = nicknameTotals(1) + srcData.Cells(i, srcCols("�s��")).Value
            nicknameTotals(2) = nicknameTotals(2) + srcData.Cells(i, srcCols("�ғ�����")).Value
            dateNicknameTotals(compositeKey) = nicknameTotals
            
            ' ���t���Ƃ̍��v�l�������Ɍv�Z
            If Not dailyTotals.Exists(dateKey) Then
                dailyTotals(dateKey) = Array(0, 0, 0) ' ���сA�s�ǁA�ғ����Ԃ̏�
                Set dailyNicknames(dateKey) = CreateObject("Scripting.Dictionary")
            End If
            
            Dim totals As Variant
            totals = dailyTotals(dateKey)
            totals(0) = totals(0) + srcData.Cells(i, srcCols("����")).Value
            totals(1) = totals(1) + srcData.Cells(i, srcCols("�s��")).Value
            totals(2) = totals(2) + srcData.Cells(i, srcCols("�ғ�����")).Value
            dailyTotals(dateKey) = totals
            
            ' �ʏ̂��L�^
            dailyNicknames(dateKey)(nickname) = True
        End If
    Next i
    
    ' �]�L��e�[�u������U�N���A�i�Y������ʏ̂ƍ��v�̂݁j
    Application.StatusBar = "�]�L����N���A��..."
    For j = 1 To tgtData.Rows.Count
        Dim tgtDate As Date
        tgtDate = tgtData.Cells(j, tgtDateCol).Value
        dateKey = Format(tgtDate, "yyyy-mm-dd")
        
        ' �Y�����t�ɒʏ̂����݂���ꍇ�̂݃N���A
        If dailyNicknames.Exists(dateKey) Then
            Dim nick As Variant
            For Each nick In dailyNicknames(dateKey).Keys
                ' �ʏ̂̎��сA�s�ǁA�ғ����Ԃ��N���A
                ClearValue tgtTable, tgtData, j, CStr(nick) & "������"
                ClearValue tgtTable, tgtData, j, CStr(nick) & "���s�ǎ���"
                ClearValue tgtTable, tgtData, j, CStr(nick) & "���ғ�����"
            Next nick
            
            ' ���v����N���A
            ClearValue tgtTable, tgtData, j, "���v������"
            ClearValue tgtTable, tgtData, j, "���v���s�ǎ���"
            ClearValue tgtTable, tgtData, j, "���v���ғ�����"
        End If
    Next j
    
    ' �W�v���ꂽ�f�[�^��]�L
    Application.StatusBar = "�W�v�f�[�^��]�L��..."
    Dim key As Variant
    Dim processedKeys As Long: processedKeys = 0
    Dim totalKeys As Long: totalKeys = dateNicknameTotals.Count
    
    For Each key In dateNicknameTotals.Keys
        processedKeys = processedKeys + 1
        
        ' �i���\��
        If processedKeys Mod 10 = 0 Or processedKeys = totalKeys Then
            Application.StatusBar = "���H�i�ԕ� �]�L������... " & Format(processedKeys / totalKeys, "0%") & _
                                  " (" & processedKeys & "/" & totalKeys & "��)"
            DoEvents
        End If
        
        ' �L�[�𕪉����ē��t�ƒʏ̂��擾
        Dim keyParts() As String
        keyParts = Split(key, "|")
        dateKey = keyParts(0)
        nickname = keyParts(1)
        
        ' �]�L��̓��t����
        For j = 1 To tgtData.Rows.Count
            If Format(tgtData.Cells(j, tgtDateCol).Value, "yyyy-mm-dd") = dateKey Then
                ' �W�v���ꂽ�f�[�^��]�L
                nicknameTotals = dateNicknameTotals(key)
                TransferValue tgtTable, tgtData, j, nickname & "������", nicknameTotals(0)
                TransferValue tgtTable, tgtData, j, nickname & "���s�ǎ���", nicknameTotals(1)
                TransferValue tgtTable, tgtData, j, nickname & "���ғ�����", nicknameTotals(2)
                transferred = transferred + 1
                Exit For
            End If
        Next j
    Next key
    
    ' ���v�l�̓]�L
    Application.StatusBar = "���v�l��]�L��..."
    Dim totalTransferred As Long: totalTransferred = 0
    For j = 1 To tgtData.Rows.Count
        tgtDate = tgtData.Cells(j, tgtDateCol).Value
        dateKey = Format(tgtDate, "yyyy-mm-dd")
        
        If dailyTotals.Exists(dateKey) Then
            totals = dailyTotals(dateKey)
            
            ' ���v�l��]�L
            TransferValue tgtTable, tgtData, j, "���v������", totals(0)
            TransferValue tgtTable, tgtData, j, "���v���s�ǎ���", totals(1)
            TransferValue tgtTable, tgtData, j, "���v���ғ�����", totals(2)
            totalTransferred = totalTransferred + 1
        End If
    Next j
    
    ' �������̃X�e�[�^�X�o�[�\��
    Application.StatusBar = "���H�i�ԕʓ]�L����: " & transferred & "���̕i�ԕʃf�[�^�A" & _
                           totalTransferred & "���̍��v�f�[�^��]�L"
    
    ' 1�b�ҋ@���Ă���N���A
    Application.Wait Now + TimeValue("0:00:01")
    Application.StatusBar = False
    
Cleanup:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    ' �G���[���������b�Z�[�W�{�b�N�X
    MsgBox "���H�i�ԕ� �]�L�����ŃG���[����" & vbCrLf & vbCrLf & _
           "�G���[���e: " & Err.Description & vbCrLf & _
           "�G���[�ԍ�: " & Err.Number, vbCritical, "�]�L�G���["
    Application.StatusBar = False
    Resume Cleanup
End Sub

' �l�]�L�p�w���p�[�֐�
Private Sub TransferValue(tbl As ListObject, data As Range, _
                         row As Long, colName As String, val As Variant)
    On Error Resume Next
    Dim colIdx As Long
    colIdx = tbl.ListColumns(colName).Index
    If colIdx > 0 Then data.Cells(row, colIdx).Value = val
    On Error GoTo 0
End Sub

' �l�N���A�p�w���p�[�֐�
Private Sub ClearValue(tbl As ListObject, data As Range, _
                      row As Long, colName As String)
    On Error Resume Next
    Dim colIdx As Long
    colIdx = tbl.ListColumns(colName).Index
    If colIdx > 0 Then data.Cells(row, colIdx).ClearContents
    On Error GoTo 0
End Sub

