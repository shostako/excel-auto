Attribute VB_Name = "m�]�L_�V�[�g���[��FR��"
Option Explicit

' ���[��FR�ʓ]�L�}�N���i�������Łj
' �u_���[��FR��a�v�e�[�u������u_���[��FR��b�v�e�[�u���փf�[�^��]�L
Sub �]�L_�V�[�g���[��FR��()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim srcTable As ListObject
    Dim tgtTable As ListObject
    Dim srcData As Range
    Dim tgtData As Range
    Dim srcCols As Object
    Dim tgtCols As Object
    
    ' ��{�ݒ�
    Set wb = ThisWorkbook
    
    ' �������ݒ�i���ꂪ�d�v�I�j
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' �X�e�[�^�X�o�[�\��
    Application.StatusBar = "���[��FR�ʓ]�L�������J�n..."
    
    On Error GoTo ErrorHandler
    
    ' �V�[�g�擾�i�V�[�g���́u���[��FR�ʁv�Ƒz��j
    Set ws = wb.Worksheets("���[��FR��")
    
    ' �e�[�u���擾
    Set srcTable = ws.ListObjects("_���[��FR��a")
    Set tgtTable = ws.ListObjects("_���[��FR��b")
    
    ' �f�[�^�͈̓`�F�b�N
    If srcTable.DataBodyRange Is Nothing Then
        Application.StatusBar = "�\�[�X�e�[�u���Ƀf�[�^������܂���"
        GoTo Cleanup
    End If
    
    If tgtTable.DataBodyRange Is Nothing Then
        Application.StatusBar = "�]�L��e�[�u���Ƀf�[�^������܂���"
        GoTo Cleanup
    End If
    
    ' �\�[�X�e�[�u���̗�C���f�b�N�X�擾
    Set srcCols = CreateObject("Scripting.Dictionary")
    srcCols("���t") = srcTable.ListColumns("���t").Index
    srcCols("F/R") = srcTable.ListColumns("F/R").Index
    srcCols("����") = srcTable.ListColumns("����").Index
    srcCols("�s��") = srcTable.ListColumns("�s��").Index
    srcCols("�ғ�����") = srcTable.ListColumns("�ғ�����").Index
    
    ' �]�L��̗�C���f�b�N�X�擾
    Set tgtCols = CreateObject("Scripting.Dictionary")
    tgtCols("���t") = tgtTable.ListColumns("���t").Index
    
    ' ���[��F��
    tgtCols("���[��F������") = GetColumnIndexSafe(tgtTable, "���[��F������")
    tgtCols("���[��F���s�ǐ�") = GetColumnIndexSafe(tgtTable, "���[��F���s�ǐ�")
    tgtCols("���[��F���ғ�����") = GetColumnIndexSafe(tgtTable, "���[��F���ғ�����")
    
    ' ���[��R��
    tgtCols("���[��R������") = GetColumnIndexSafe(tgtTable, "���[��R������")
    tgtCols("���[��R���s�ǐ�") = GetColumnIndexSafe(tgtTable, "���[��R���s�ǐ�")
    tgtCols("���[��R���ғ�����") = GetColumnIndexSafe(tgtTable, "���[��R���ғ�����")
    
    ' �f�[�^�͈͎擾
    Set srcData = srcTable.DataBodyRange
    Set tgtData = tgtTable.DataBodyRange
    
    Dim i As Long, j As Long
    Dim srcDate As Date, frType As String
    Dim transferred As Long: transferred = 0
    Dim totalRows As Long: totalRows = srcData.Rows.Count
    
    ' �]�L��e�[�u�����ꊇ�N���A�i�������j
    Application.StatusBar = "�]�L����N���A��..."
    
    ' ���[��F��̈ꊇ�N���A
    If tgtCols("���[��F������") > 0 Then
        tgtData.Columns(tgtCols("���[��F������")).ClearContents
    End If
    If tgtCols("���[��F���s�ǐ�") > 0 Then
        tgtData.Columns(tgtCols("���[��F���s�ǐ�")).ClearContents
    End If
    If tgtCols("���[��F���ғ�����") > 0 Then
        tgtData.Columns(tgtCols("���[��F���ғ�����")).ClearContents
    End If
    
    ' ���[��R��̈ꊇ�N���A
    If tgtCols("���[��R������") > 0 Then
        tgtData.Columns(tgtCols("���[��R������")).ClearContents
    End If
    If tgtCols("���[��R���s�ǐ�") > 0 Then
        tgtData.Columns(tgtCols("���[��R���s�ǐ�")).ClearContents
    End If
    If tgtCols("���[��R���ғ�����") > 0 Then
        tgtData.Columns(tgtCols("���[��R���ғ�����")).ClearContents
    End If
    
    ' �������̊́F�]�L��̓��t�ƍs�ԍ��̑Ή���Dictionary�Ɋi�[
    Application.StatusBar = "�C���f�b�N�X�쐬��..."
    Dim dateIndex As Object
    Set dateIndex = CreateObject("Scripting.Dictionary")
    
    For j = 1 To tgtData.Rows.Count
        Dim tgtDate As Date
        tgtDate = tgtData.Cells(j, tgtCols("���t")).Value
        ' ���t���L�[�ɂ��čs�ԍ����i�[
        dateIndex(CLng(tgtDate)) = j
    Next j
    
    ' �f�[�^�̓]�L�i�������Łj
    Application.StatusBar = "�f�[�^�]�L��..."
    For i = 1 To totalRows
        ' �i���\���i100�s���ƂɍX�V - DoEvents�����炷�j
        If i Mod 100 = 0 Or i = totalRows Then
            Application.StatusBar = "���[��FR�ʓ]�L������... " & Format(i / totalRows, "0%") & _
                                  " (" & i & "/" & totalRows & "�s)"
            ' DoEvents�͍ŏ�����
            If i Mod 500 = 0 Then DoEvents
        End If
        
        ' �\�[�X�f�[�^�擾
        srcDate = srcData.Cells(i, srcCols("���t")).Value
        frType = Trim(srcData.Cells(i, srcCols("F/R")).Value)
        
        ' ���t�ɑΉ�����]�L��̍s�ԍ����擾�i���������j
        If dateIndex.Exists(CLng(srcDate)) Then
            j = dateIndex(CLng(srcDate))
            
            ' F/R�^�C�v�ɉ����ē]�L
            If frType = "F" Then
                ' ���[��F��ւ̓]�L
                If tgtCols("���[��F������") > 0 Then
                    tgtData.Cells(j, tgtCols("���[��F������")).Value = srcData.Cells(i, srcCols("����")).Value
                End If
                If tgtCols("���[��F���s�ǐ�") > 0 Then
                    tgtData.Cells(j, tgtCols("���[��F���s�ǐ�")).Value = srcData.Cells(i, srcCols("�s��")).Value
                End If
                If tgtCols("���[��F���ғ�����") > 0 Then
                    tgtData.Cells(j, tgtCols("���[��F���ғ�����")).Value = srcData.Cells(i, srcCols("�ғ�����")).Value
                End If
                transferred = transferred + 1
                
            ElseIf frType = "R" Then
                ' ���[��R��ւ̓]�L
                If tgtCols("���[��R������") > 0 Then
                    tgtData.Cells(j, tgtCols("���[��R������")).Value = srcData.Cells(i, srcCols("����")).Value
                End If
                If tgtCols("���[��R���s�ǐ�") > 0 Then
                    tgtData.Cells(j, tgtCols("���[��R���s�ǐ�")).Value = srcData.Cells(i, srcCols("�s��")).Value
                End If
                If tgtCols("���[��R���ғ�����") > 0 Then
                    tgtData.Cells(j, tgtCols("���[��R���ғ�����")).Value = srcData.Cells(i, srcCols("�ғ�����")).Value
                End If
                transferred = transferred + 1
            End If
        End If
    Next i
    
    ' �����_�ȉ�2���̏����ݒ�i��S�̂Ɉꊇ�K�p�j
    Application.StatusBar = "�����ݒ蒆..."
    
    ' ���[��F���ғ����ԁi��S�́j
    If tgtCols("���[��F���ғ�����") > 0 Then
        tgtTable.ListColumns("���[��F���ғ�����").DataBodyRange.NumberFormatLocal = "0.00"
    End If
    
    ' ���[��R���ғ����ԁi��S�́j
    If tgtCols("���[��R���ғ�����") > 0 Then
        tgtTable.ListColumns("���[��R���ғ�����").DataBodyRange.NumberFormatLocal = "0.00"
    End If
    
    ' ��������
    Application.StatusBar = "���[��FR�ʓ]�L����: " & transferred & "���̃f�[�^��]�L"
    Application.Wait Now + TimeValue("0:00:01")
    Application.StatusBar = False
    
    GoTo Cleanup
    
ErrorHandler:
    ' �G���[���̏���
    MsgBox "���[��FR�ʓ]�L�����ŃG���[���������܂���" & vbCrLf & vbCrLf & _
           "�G���[���e: " & Err.Description & vbCrLf & _
           "�G���[�ԍ�: " & Err.Number, vbCritical, "�]�L�G���["
    
Cleanup:
    ' �㏈���i�������ݒ�����ɖ߂��j
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    
    Set srcCols = Nothing
    Set tgtCols = Nothing
    Set srcData = Nothing
    Set tgtData = Nothing
    Set srcTable = Nothing
    Set tgtTable = Nothing
    Set ws = Nothing
    Set wb = Nothing
    Set dateIndex = Nothing
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
