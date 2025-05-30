Attribute VB_Name = "m�]�L_TG��Ǝҕʂ���W�v�\"
Sub �]�L_TG��Ǝҕʂ���W�v�\()
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
    
    ' �]�L���񖼖����̔z��
    Dim suffixList() As Variant
    suffixList = Array("����", "���o��������", "�����ԓ���o����", "�݌v", _
                      "�����ώ���", "���Ϗo��������", "���ώ��ԓ��萔")
    
    ' �]�L��s�ԍ��̔z��isuffixList�ɑΉ��j
    Dim targetRows() As Variant
    targetRows = Array(59, 60, 61, 62, 63, 64, 65)
    
    ' ��ƎҖ����擾�����̔z��i58�s�ځj
    Dim workerColumns() As Variant
    workerColumns = Array(4, 6, 8, 10, 12, 14, 16)  ' D, F, H, J, L, N, P��
    
    ' �G���[�n���h�����O�ݒ�
    On Error GoTo ErrorHandler
    
    ' �i���\���J�n
    Application.StatusBar = "TG��Ǝҕʃf�[�^�̓]�L�������J�n���܂�..."
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' �W�v�\�V�[�g�擾
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("�W�v�\")
    If wsTarget Is Nothing Then
        MsgBox "�u�W�v�\�v�V�[�g��������܂���B", vbCritical, "�V�[�g�G���["
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' �W�v�\��A1�Z��������t�擾
    If Not IsDate(wsTarget.Range("A1").Value) Then
        MsgBox "�W�v�\�̃Z��A1�ɗL���ȓ��t�����͂���Ă��܂���B", vbCritical, "���t�G���["
        GoTo CleanupAndExit
    End If
    targetDate = wsTarget.Range("A1").Value
    
    ' TG��ƎҕʃV�[�g�擾
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("TG��Ǝҕ�")
    If wsSource Is Nothing Then
        MsgBox "�uTG��Ǝҕʁv�V�[�g��������܂���B", vbCritical, "�V�[�g�G���["
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' �\�[�X�e�[�u���擾
    On Error Resume Next
    Set sourceTable = wsSource.ListObjects("_TG��Ǝҕ�b")
    If sourceTable Is Nothing Then
        MsgBox "�uTG��Ǝҕʁv�V�[�g�Ɂu_TG��Ǝҕ�b�v�e�[�u����������܂���B", vbCritical, "�e�[�u���G���["
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' �f�[�^�͈͎擾
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "�u_TG��Ǝҕ�b�v�e�[�u���Ƀf�[�^������܂���B", vbCritical, "�f�[�^�G���["
        GoTo CleanupAndExit
    End If
    Set sourceData = sourceTable.DataBodyRange
    
    ' ���t��̃C���f�b�N�X�擾
    Dim dateColIndex As Long
    On Error Resume Next
    dateColIndex = sourceTable.ListColumns("���t").Index
    If Err.Number <> 0 Then
        MsgBox "�u_TG��Ǝҕ�b�v�e�[�u���Ɂu���t�v�񂪌�����܂���B", vbCritical, "��G���["
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
        MsgBox "���t " & Format(targetDate, "yyyy/mm/dd") & " �̃f�[�^��������܂���B", vbCritical, "�f�[�^�G���["
        GoTo CleanupAndExit
    End If
    
    ' �e��̍�ƎҖ����擾���ē]�L����
    totalCombinations = 0
    processedCombinations = 0
    
    ' �܂������������J�E���g�i�i���\���p�j
    For i = 0 To UBound(workerColumns)
        If wsTarget.Cells(58, workerColumns(i)).Value <> "" Then
            totalCombinations = totalCombinations + (UBound(suffixList) + 1)
        End If
    Next i
    
    ' �e��̏���
    For i = 0 To UBound(workerColumns)
        ' 58�s�ڂ����ƎҖ����擾
        Dim workerName As String
        workerName = CStr(wsTarget.Cells(58, workerColumns(i)).Value)
        
        ' �󔒃Z���̓X�L�b�v
        If workerName <> "" Then
            Application.StatusBar = "TG��Ǝҕʃf�[�^�]�L��... (" & workerName & ")"
            
            ' �e�����Ƃ̑g�ݍ��킹�œ]�L
            For k = 0 To UBound(suffixList)
            processedCombinations = processedCombinations + 1
            
            ' �񖼂��\�z�i��ƎҖ� + ����������j
            Dim columnName As String
            columnName = workerName & suffixList(k)
            
            ' �]�L���s
            On Error Resume Next
            Dim colIndex As Long
            colIndex = sourceTable.ListColumns(columnName).Index
            
            If Err.Number = 0 Then
                ' �]�L��̃Z���ɒl��ݒ�
                wsTarget.Cells(targetRows(k), workerColumns(i)).Value = _
                    sourceData.Cells(sourceRow, colIndex).Value
            Else
                ' �񂪌�����Ȃ��ꍇ�͌x���i�f�o�b�O�p�j
                Debug.Print "�x��: ��u" & columnName & "�v��������܂���B"
                Err.Clear
            End If
            On Error GoTo ErrorHandler
            
                ' �i���X�V
                If processedCombinations Mod 5 = 0 Then
                    Application.StatusBar = "TG��Ǝҕʃf�[�^�]�L��... (" & _
                        processedCombinations & "/" & totalCombinations & ")"
                End If
            Next k
        End If
    Next i
    
    ' ����I���i�G���[���ȊO�̓��b�Z�[�W��\���j
    GoTo CleanupAndExit
    
ErrorHandler:
    MsgBox "�]�L�������ɃG���[���������܂����B" & vbCrLf & _
           "�G���[���e: " & Err.Description & vbCrLf & _
           "�G���[�ԍ�: " & Err.Number, vbCritical, "�]�L�G���["
    
CleanupAndExit:
    ' �㏈��
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False  ' �X�e�[�^�X�o�[���N���A
End Sub

