Attribute VB_Name = "m�]�L_�V�[�g���H��Ǝҕ�"
Sub �]�L_�V�[�g���H��Ǝҕ�()
    ' �ϐ��錾
    Dim ws As Worksheet
    Dim sourceTable As ListObject
    Dim TargetTable As ListObject
    Dim sourceData As Range
    Dim TargetData As Range
    Dim i As Long, j As Long
    Dim sourceRow As Long, targetRow As Long
    Dim sourceDate As Date, targetDate As Date
    Dim workerName As String
    Dim workTimeColName As String, resultColName As String
    Dim workTimeCol As ListColumn, resultCol As ListColumn
    Dim workTimeColIndex As Long, resultColIndex As Long
    Dim dateColSourceIndex As Long, dateColTargetIndex As Long
    Dim workerColIndex As Long, workTimeSourceColIndex As Long, resultSourceColIndex As Long
    Dim totalRows As Long
    Dim processedCount As Long
    
    ' �G���[�n���h�����O�ݒ�
    On Error GoTo ErrorHandler
    
    ' �i���\���J�n
    Application.StatusBar = "�]�L�������J�n���܂�..."
    Application.ScreenUpdating = False
    
    ' ���[�N�V�[�g�擾
    Set ws = ThisWorkbook.Worksheets("���H��Ǝҕ�")
    
    ' �e�[�u���擾
    Set sourceTable = ws.ListObjects("_���H��Ǝҕ�a")
    Set TargetTable = ws.ListObjects("_���H��Ǝҕ�b")
    
    ' �f�[�^�͈͎擾�i�w�b�_�[�����j
    Set sourceData = sourceTable.DataBodyRange
    Set TargetData = TargetTable.DataBodyRange
    
    ' ��C���f�b�N�X�擾
    dateColSourceIndex = sourceTable.ListColumns("���t").Index
    dateColTargetIndex = TargetTable.ListColumns("���t").Index
    workerColIndex = sourceTable.ListColumns("��Ǝ�").Index
    workTimeSourceColIndex = sourceTable.ListColumns("�ғ�����").Index
    resultSourceColIndex = sourceTable.ListColumns("����").Index
    
    ' ���s���擾
    totalRows = sourceData.Rows.Count
    processedCount = 0
    
    ' �\�[�X�e�[�u���̊e�s������
    For i = 1 To totalRows
        ' �i���\���X�V
        processedCount = processedCount + 1
        Application.StatusBar = "�]�L������... (" & processedCount & "/" & totalRows & ")"
        
        ' �\�[�X�f�[�^�擾
        sourceDate = sourceData.Cells(i, dateColSourceIndex).Value
        workerName = Trim(sourceData.Cells(i, workerColIndex).Value)
        
        ' ��ƎҖ����󔒂̏ꍇ�̓X�L�b�v
        If workerName = "" Then
            GoTo NextSourceRow
        End If
        
        ' �]�L��̑Ή����t�s������
        targetRow = 0
        For j = 1 To TargetData.Rows.Count
            If TargetData.Cells(j, dateColTargetIndex).Value = sourceDate Then
                targetRow = j
                Exit For
            End If
        Next j
        
        ' �Ή�������t��������Ȃ��ꍇ�̓X�L�b�v
        If targetRow = 0 Then
            GoTo NextSourceRow
        End If
        
        ' �ғ����ԓ]�L����
        workTimeColName = workerName & "�ғ�����"
        Set workTimeCol = Nothing
        
        ' �ғ����ԗ�̑��݊m�F
        For Each workTimeCol In TargetTable.ListColumns
            If workTimeCol.Name = workTimeColName Then
                workTimeColIndex = workTimeCol.Index
                ' �ғ����Ԓl��]�L
                TargetData.Cells(targetRow, workTimeColIndex).Value = sourceData.Cells(i, workTimeSourceColIndex).Value
                Exit For
            End If
        Next workTimeCol
        
        ' ���ѓ]�L����
        resultColName = workerName & "����"
        Set resultCol = Nothing
        
        ' ���ї�̑��݊m�F
        For Each resultCol In TargetTable.ListColumns
            If resultCol.Name = resultColName Then
                resultColIndex = resultCol.Index
                ' ���ђl��]�L
                TargetData.Cells(targetRow, resultColIndex).Value = sourceData.Cells(i, resultSourceColIndex).Value
                Exit For
            End If
        Next resultCol
        
NextSourceRow:
    Next i
    
    ' ��������
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' �G���[���̏���
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "�]�L�������ɃG���[���������܂����B" & vbCrLf & _
           "�G���[���e: " & Err.Description & vbCrLf & _
           "�G���[�ԍ�: " & Err.Number, vbCritical, "�]�L�G���["
    
End Sub

