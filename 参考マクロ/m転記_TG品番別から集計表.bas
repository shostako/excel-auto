Attribute VB_Name = "m�]�L_TG�i�ԕʂ���W�v�\"
Sub �]�L_TG�i�ԕʂ���W�v�\()
    ' �ϐ��錾
    Dim wsTarget As Worksheet
    Dim wsSource As Worksheet
    Dim targetDate As Date
    Dim sourceTable As ListObject
    Dim sourceData As Range
    Dim i As Long, j As Long, k As Long
    Dim sourceRow As Long
    Dim totalCombinations As Long
    Dim processedCombinations As Long
    
    ' �i�Ԑړ����̔z��
    Dim prefixList() As Variant
    prefixList = Array("RH", "LH", "���v")
    
    ' �]�L���񖼖����̔z��
    Dim suffixList() As Variant
    suffixList = Array("������", "���o��������", "�݌v����", "���ώ���", _
                      "�݌v�o��������", "���s�ǎ���", "�݌v�s�ǐ�", _
                      "�݌v�s�Ǘ�", "���ϕs�ǐ�")
    
    ' �]�L��s�ԍ��̔z��isuffixList�ɑΉ��j
    Dim targetRows() As Variant
    targetRows = Array(33, 34, 35, 36, 37, 39, 40, 41, 42)
    
    ' �i�ԂɑΉ�����]�L���̔z��
    Dim targetColumns() As Variant
    targetColumns = Array(12, 14, 16)  ' L, N, P��
    
    ' �G���[�n���h�����O�ݒ�
    On Error GoTo ErrorHandler
    
    ' �i���\���J�n
    Application.StatusBar = "TG�f�[�^�̓]�L�������J�n���܂�..."
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' �W�v�\�V�[�g�擾
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("�W�v�\")
    If wsTarget Is Nothing Then
        MsgBox "�u�W�v�\�v�V�[�g��������܂���B" & vbCrLf & _
               "�V�[�g�����炢�o���Ƃ���", vbCritical
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' �W�v�\��A1�Z��������t�擾
    If Not IsDate(wsTarget.Range("A1").Value) Then
        MsgBox "�W�v�\�̃Z��A1�ɗL���ȓ��t�����͂���Ă��܂���B" & vbCrLf & _
               "���t�̓��͂��ł��Ȃ��Ƃ��A���v���H", vbCritical
        GoTo CleanupAndExit
    End If
    targetDate = wsTarget.Range("A1").Value
    
    ' TG�i�ԕʃV�[�g�擾
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("TG�i�ԕ�")
    If wsSource Is Nothing Then
        MsgBox "�uTG�i�ԕʁv�V�[�g��������܂���B" & vbCrLf & _
               "�V�[�g����Ă�����s�����", vbCritical
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' �\�[�X�e�[�u���擾
    On Error Resume Next
    Set sourceTable = wsSource.ListObjects("_TG�i�ԕ�b")
    If sourceTable Is Nothing Then
        MsgBox "�uTG�i�ԕʁv�V�[�g�Ɂu_TG�i�ԕ�b�v�e�[�u����������܂���B" & vbCrLf & _
               "�e�[�u�������炢���ꂵ����", vbCritical
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' �f�[�^�͈͎擾
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "�u_TG�i�ԕ�b�v�e�[�u���Ƀf�[�^������܂���B" & vbCrLf & _
               "����ۂ̃e�[�u�����牽��]�L����C���H", vbCritical
        GoTo CleanupAndExit
    End If
    Set sourceData = sourceTable.DataBodyRange
    
    ' ���t��̃C���f�b�N�X�擾
    Dim dateColIndex As Long
    On Error Resume Next
    dateColIndex = sourceTable.ListColumns("���t").Index
    If Err.Number <> 0 Then
        MsgBox "�u_TG�i�ԕ�b�v�e�[�u���Ɂu���t�v�񂪌�����܂���B" & vbCrLf & _
               "���t����Ȃ��̂ɉ�����ɓ]�L����񂾂�", vbCritical
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' �Y�����t�̍s������
    sourceRow = 0
    For j = 1 To sourceData.Rows.Count
        If sourceData.Cells(j, dateColIndex).Value = targetDate Then
            sourceRow = j
            Exit For
        End If
    Next j
    
    If sourceRow = 0 Then
        MsgBox "���t " & Format(targetDate, "yyyy/mm/dd") & " �̃f�[�^��������܂���B" & vbCrLf & _
               "���̓��̃f�[�^�A�{���ɑ��݂���̂��H", vbCritical
        GoTo CleanupAndExit
    End If
    
    ' �e�i�ԂƖ����̑g�ݍ��킹�œ]�L����
    totalCombinations = (UBound(prefixList) + 1) * (UBound(suffixList) + 1)
    processedCombinations = 0
    
    For i = 0 To UBound(prefixList)
        Application.StatusBar = "TG�f�[�^�]�L��... (" & prefixList(i) & ")"
        
        For k = 0 To UBound(suffixList)
            processedCombinations = processedCombinations + 1
            
            ' �񖼂��\�z�i�i�Ԑړ��� + ����������j
            Dim columnName As String
            columnName = prefixList(i) & suffixList(k)
            
            ' �]�L���s
            On Error Resume Next
            Dim colIndex As Long
            colIndex = sourceTable.ListColumns(columnName).Index
            
            If Err.Number = 0 Then
                ' �]�L��̃Z���ɒl��ݒ�
                wsTarget.Cells(targetRows(k), targetColumns(i)).Value = _
                    sourceData.Cells(sourceRow, colIndex).Value
            Else
                ' �񂪌�����Ȃ��ꍇ�͌x���i�f�o�b�O�p�j
                Debug.Print "�x��: ��u" & columnName & "�v��������܂���B"
                Err.Clear
            End If
            On Error GoTo ErrorHandler
            
            ' �i���X�V�i10�����Ɓj
            If processedCombinations Mod 10 = 0 Then
                Application.StatusBar = "TG�f�[�^�]�L��... (" & _
                    processedCombinations & "/" & totalCombinations & ")"
            End If
        Next k
    Next i
    
    ' ����I�����b�Z�[�W�i�R�����g�A�E�g�ς� - �G���[���ȊO�͔�\���j
    ' MsgBox "TG�f�[�^�̓]�L���������܂����B", vbInformation
    GoTo CleanupAndExit
    
ErrorHandler:
    MsgBox "�]�L�������ɗ\�����ʃG���[���������܂����B" & vbCrLf & _
           "�G���[���e: " & Err.Description & vbCrLf & _
           "�G���[�ԍ�: " & Err.Number & vbCrLf & vbCrLf & _
           "�f�o�b�O���炢���Ă�����s�����", vbCritical, "�]�L�G���["
    
CleanupAndExit:
    ' �㏈��
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False  ' �X�e�[�^�X�o�[���N���A
End Sub

