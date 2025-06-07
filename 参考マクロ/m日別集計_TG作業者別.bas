Attribute VB_Name = "m���ʏW�v_TG��Ǝҕ�"
Option Explicit

' TG��ƎҕʏW�v�}�N���i��ʂ�����h�~�Łj
' �u���H3�v�̃f�[�^����t�E��Ǝ҂ŃO���[�v�����ďW�v
Sub ���ʏW�v_TG��Ǝҕ�()
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim tblSource As ListObject
    Dim tblOutput As ListObject
    Dim dict As Object 'Scripting.Dictionary
    Dim outputArray() As Variant
    Dim dataArray As Variant
    Dim sortKeys() As String ' �\�[�g�p�̃L�[�z��
    
    Dim sourceSheetName As String
    Dim sourceTableName As String
    Dim outputSheetName As String
    Dim outputTableName As String
    Dim outputStartCellAddress As String
    Dim outputHeader As Range
    
    Dim i As Long, r As Long, j As Long, k As Long
    Dim colDate As Long, colProcess As Long, colWorker As Long
    Dim colJisseki As Long, colDandori As Long, colKadou As Long, colFuryo As Long
    
    Dim currentDate As Date
    Dim currentWorker As String
    Dim dictKey As String
    Dim jissekiVal As Double, dandoriVal As Double, kadouVal As Double, furyoVal As Double
    Dim item As Variant
    Dim key As Variant
    Dim tempKey As String
    
    ' ��{�ݒ�
    Set wb = ThisWorkbook
    sourceSheetName = "�S�H��"
    sourceTableName = "_�S�H��"
    outputSheetName = "TG��Ǝҕ�"
    outputTableName = "_TG��Ǝҕ�a"
    outputStartCellAddress = "A3"
    
    ' �X�e�[�^�X�o�[�\��
    Application.StatusBar = "TG��ƎҕʏW�v���J�n���܂�..."
    
    ' ������ ��ʍX�V�ݒ�͍폜�iCommandButton�ňꊇ�Ǘ��j ������
    ' Application.ScreenUpdating = False �� �폜
    ' Application.Calculation = xlCalculationManual �� �폜
    ' Application.DisplayAlerts = False �� �폜
    
    ' �G���[�n���h�����O�ݒ�
    On Error GoTo ErrorHandler
    
    ' 1. ���͌��V�[�g�E�e�[�u���̑��݊m�F�Ǝ擾
    On Error Resume Next
    Set wsSource = wb.Sheets(sourceSheetName)
    If wsSource Is Nothing Then
        MsgBox "�V�[�g�u" & sourceSheetName & "�v��������܂���B", vbCritical
        GoTo Cleanup
    End If
    
    Set tblSource = wsSource.ListObjects(sourceTableName)
    If tblSource Is Nothing Then
        MsgBox "�e�[�u���u" & sourceTableName & "�v���V�[�g�u" & sourceSheetName & "�v�Ɍ�����܂���B", vbCritical
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler
    
    ' �f�[�^���Ȃ��ꍇ�͏I��
    If tblSource.DataBodyRange Is Nothing Then
        MsgBox "�e�[�u���u" & sourceTableName & "�v�Ƀf�[�^������܂���B", vbInformation
        GoTo Cleanup
    End If
    
    ' 2. �u�S�H���v�e�[�u���̗�C���f�b�N�X�擾
    colDate = GetColumnIndex(tblSource, "���t")
    colProcess = GetColumnIndex(tblSource, "�H��")
    colWorker = GetColumnIndex(tblSource, "��Ǝ�")
    colJisseki = GetColumnIndex(tblSource, "����")
    colDandori = GetColumnIndex(tblSource, "�i�掞��")
    colKadou = GetColumnIndex(tblSource, "�ғ�����")
    colFuryo = GetColumnIndex(tblSource, "�s��")
    
    If colDate = 0 Or colProcess = 0 Or colWorker = 0 Or colJisseki = 0 Or colDandori = 0 Or colKadou = 0 Or colFuryo = 0 Then
        MsgBox "�u�S�H���v�e�[�u���ɕK�v�ȗ�i���t, �H��, ��Ǝ�, ����, �i�掞��, �ғ�����, �s�ǁj��������܂���B�񖼂��m�F���Ă��������B", vbCritical
        GoTo Cleanup
    End If
    
    ' 3. �f�[�^�W�v (Dictionary���g�p)
    Set dict = CreateObject("Scripting.Dictionary")
    dataArray = tblSource.DataBodyRange.Value2 ' �������̂��ߔz��ŏ���
    
    Application.StatusBar = "�f�[�^���W�v��..."
    
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        ' �u�H���v��̒l���u���H3�v�Ɗ��S��v����s�𒊏o
        If CStr(dataArray(i, colProcess)) = "���H3" Then
            ' ���t�̑Ó����`�F�b�N�ƕϊ�
            If IsDate(dataArray(i, colDate)) Then
                currentDate = CDate(dataArray(i, colDate))
            ElseIf IsNumeric(dataArray(i, colDate)) Then
                ' ���l�̏ꍇ�͓��t�V���A���l�Ƃ��Ĉ���
                currentDate = CDate(CLng(dataArray(i, colDate)))
            Else
                ' ���t�Ƃ��ĔF���ł��Ȃ��f�[�^�̓X�L�b�v
                Debug.Print "�x��: ���t�Ƃ��ĔF���ł��Ȃ��f�[�^��������܂����B�s " & i + tblSource.HeaderRowRange.row & ", �l: " & dataArray(i, colDate)
                GoTo NextIteration
            End If
            
            ' ��Ǝ҂̎擾
            currentWorker = CStr(dataArray(i, colWorker))
            
            ' �����L�[�̍쐬�i���t|��Ǝҁj
            dictKey = Format(currentDate, "yyyy/mm/dd") & "|" & currentWorker
            
            jissekiVal = val(dataArray(i, colJisseki))
            dandoriVal = val(dataArray(i, colDandori))
            kadouVal = val(dataArray(i, colKadou))
            furyoVal = val(dataArray(i, colFuryo))
            
            If dict.Exists(dictKey) Then
                item = dict(dictKey)
                item(0) = item(0) + jissekiVal '����
                item(1) = item(1) + furyoVal  '�s��
                item(2) = item(2) + kadouVal  '�ғ�����
                item(3) = item(3) + dandoriVal '�i�掞��
                dict(dictKey) = item
            Else
                ReDim newItem(0 To 3) As Double
                newItem(0) = jissekiVal
                newItem(1) = furyoVal
                newItem(2) = kadouVal
                newItem(3) = dandoriVal
                dict.Add dictKey, newItem
            End If
        End If
NextIteration:
    Next i
    
    If dict.Count = 0 Then
        MsgBox "�H���u���H3�v�ɊY������f�[�^���W�v����܂���ł����B", vbInformation
        ' ���̏ꍇ�ł��V�[�g�Ƌ�̃e�[�u���͍쐬�����悤�ɂ���
    End If
    
    ' 4. �o�͐�V�[�g�̏���
    Application.StatusBar = "�o�͐�V�[�g��������..."
    
    On Error Resume Next
    Set wsOutput = wb.Sheets(outputSheetName)
    If wsOutput Is Nothing Then
        Set wsOutput = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsOutput.Name = outputSheetName
    End If
    ' ������ wsOutput.Activate ���폜�i��ʐ؂�ւ��̌����j ������
    On Error GoTo ErrorHandler
    
    ' 5. �o�͐�e�[�u���̏���
    Set outputHeader = wsOutput.Range(outputStartCellAddress)
    
    ' �����e�[�u�������邩�`�F�b�N
    On Error Resume Next
    Set tblOutput = wsOutput.ListObjects(outputTableName)
    On Error GoTo ErrorHandler
    
    Dim isNewTable As Boolean
    isNewTable = (tblOutput Is Nothing)
    
    If Not isNewTable Then
        ' �����e�[�u���̏ꍇ�F�f�[�^�����̂݃N���A
        On Error Resume Next
        If Not tblOutput.DataBodyRange Is Nothing Then
            tblOutput.DataBodyRange.ClearContents
        End If
        On Error GoTo ErrorHandler
    Else
        ' �w�b�_�[�������݁i�V�K�e�[�u���̏ꍇ�̂݁j
        outputHeader.Resize(1, 6).Value = Array("���t", "��Ǝ�", "����", "�s��", "�ғ�����", "�i�掞��")
    End If
    
    ' 6. �W�v���ʂ�z��ɕϊ����A���t�E��Ǝ҂Ń\�[�g
    If dict.Count > 0 Then
        Application.StatusBar = "�f�[�^���\�[�g��..."
        
        ' �\�[�g�p�̃L�[�z����쐬
        ReDim sortKeys(1 To dict.Count)
        i = 0
        For Each key In dict.Keys
            i = i + 1
            sortKeys(i) = CStr(key)
        Next key
        
        ' �o�u���\�[�g�œ��t�E��Ǝҏ��ɕ��בւ�
        ' �i�����ł�QuickSort�Ȃǂ��g���ׂ������A�����ł͊Ȍ��Ƀo�u���\�[�g�j
        For i = 1 To dict.Count - 1
            For j = i + 1 To dict.Count
                If sortKeys(i) > sortKeys(j) Then
                    tempKey = sortKeys(i)
                    sortKeys(i) = sortKeys(j)
                    sortKeys(j) = tempKey
                End If
            Next j
        Next i
        
        ' �o�͔z��̍쐬
        ReDim outputArray(1 To dict.Count, 1 To 6)
        For r = 1 To dict.Count
            key = sortKeys(r)
            item = dict(key)
            
            ' �L�[�𕪉����ē��t�ƍ�Ǝ҂��擾
            Dim keyParts() As String
            keyParts = Split(key, "|")
            
            outputArray(r, 1) = CDate(keyParts(0)) '���t
            outputArray(r, 2) = keyParts(1)        '��Ǝ�
            outputArray(r, 3) = item(0)            '����
            outputArray(r, 4) = item(1)            '�s��
            outputArray(r, 5) = item(2)            '�ғ�����
            outputArray(r, 6) = item(3)            '�i�掞��
        Next r
        
        ' 7. �f�[�^�o��
        Application.StatusBar = "�f�[�^���o�͒�..."
        
        If Not isNewTable Then
            ' �����e�[�u���̏ꍇ�F�e�[�u���T�C�Y������Ƀf�[�^��������
            tblOutput.Resize outputHeader.Resize(UBound(outputArray, 1) + 1, 6)
        End If
        outputHeader.Offset(1, 0).Resize(UBound(outputArray, 1), UBound(outputArray, 2)).Value = outputArray
    End If
    
    ' 8. �e�[�u���쐬�i�V�K�̏ꍇ�̂݁j�܂��͍X�V
    If isNewTable Then
        ' �f�[�^���Ȃ��ꍇ�ł��w�b�_�[�݂̂̃e�[�u�����쐬
        Dim dataRangeForTable As Range
        If dict.Count > 0 Then
            Set dataRangeForTable = outputHeader.Resize(dict.Count + 1, 6)
        Else
            Set dataRangeForTable = outputHeader.Resize(1, 6) ' �w�b�_�[�̂�
        End If
        
        Set tblOutput = wsOutput.ListObjects.Add(xlSrcRange, dataRangeForTable, , xlYes)
        tblOutput.Name = outputTableName
        tblOutput.TableStyle = "TableStyleMedium9"
    ElseIf dict.Count = 0 Then
        ' �����e�[�u���Ńf�[�^���Ȃ��ꍇ�F�w�b�_�[�݂̂Ƀ��T�C�Y
        tblOutput.Resize outputHeader.Resize(1, 6)
    End If
    
    ' �e�[�u���̃t�B���^�[�{�^�����\���ɐݒ�
    tblOutput.ShowAutoFilter = False
    
    ' ���t��̏����ݒ�
    If dict.Count > 0 Then
        tblOutput.ListColumns("���t").DataBodyRange.NumberFormatLocal = "yyyy/mm/dd"
    End If
    
    ' ========== �ǉ������ݒ� ==========
    Application.StatusBar = "������ݒ蒆..."
    
    ' 1. �f�[�^�͈͂́u�k�����đS�̂�\������v�ݒ�
    If dict.Count > 0 Then
        tblOutput.DataBodyRange.ShrinkToFit = True
    End If
    
    ' 2. �S��̗񕝂�6.4�ɐݒ�
    Dim col As ListColumn
    For Each col In tblOutput.ListColumns
        col.Range.ColumnWidth = 6.4
    Next col
    
    ' 3. �u�ғ����ԁv�u�i�掞�ԁv��̏����_�ȉ�2���ݒ�
    If dict.Count > 0 Then
        On Error Resume Next
        tblOutput.ListColumns("�ғ�����").DataBodyRange.NumberFormatLocal = "0.00"
        tblOutput.ListColumns("�i�掞��").DataBodyRange.NumberFormatLocal = "0.00"
        On Error GoTo ErrorHandler
    End If
    
    ' 4. A1�Z���Ƀ^�C�g����ݒ�
    With wsOutput.Range("A1")
        .Value = "TG��Ǝҕʃf�[�^���o"
        .Font.Name = "Yu Gothic UI"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' ��������
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' �G���[���̏���
    Application.StatusBar = False
    MsgBox "�G���[���������܂���: " & Err.Description & vbCrLf & _
           "�G���[�ԍ�: " & Err.Number, vbCritical
    
Cleanup:
    ' �㏈��
    Set dict = Nothing
    Set wsSource = Nothing
    Set wsOutput = Nothing
    Set tblSource = Nothing
    Set tblOutput = Nothing
    Set wb = Nothing
    
    ' ������ ��ʍX�V�ݒ��߂��������폜�iCommandButton�ňꊇ�Ǘ��j ������
    ' Application.ScreenUpdating = True �� �폜
    ' Application.Calculation = xlCalculationAutomatic �� �폜
    ' Application.DisplayAlerts = True �� �폜
    Application.StatusBar = False
End Sub

' �e�[�u���̗񖼂����C���f�b�N�X���擾����w���p�[�֐�
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    Dim col As ListColumn
    Dim i As Long
    i = 0
    On Error Resume Next
    i = tbl.ListColumns(columnName).Index
    On Error GoTo 0
    GetColumnIndex = i
End Function

