Attribute VB_Name = "m���ʏW�v_�h���i�ԕ�"
Option Explicit

' �h���i�ԕʏW�v�}�N���i��ʂ�����h�~�Łj
' �u�h���v�̃f�[�^����t�E�i�ԂŃO���[�v�����ďW�v���A�u�ʏ́v���ǉ�
Sub ���ʏW�v_�h���i�ԕ�()
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
    Dim colDate As Long, colProcess As Long, colHinban As Long
    Dim colJisseki As Long, colDandori As Long, colKadou As Long, colFuryo As Long
    
    Dim currentDate As Date
    Dim currentHinban As String
    Dim dictKey As String
    Dim jissekiVal As Double, dandoriVal As Double, kadouVal As Double, furyoVal As Double
    Dim item As Variant
    Dim key As Variant
    Dim tempKey As String
    
    ' ��{�ݒ�
    Set wb = ThisWorkbook
    sourceSheetName = "�S�H��"
    sourceTableName = "_�S�H��"
    outputSheetName = "�h���i�ԕ�"
    outputTableName = "_�h���i�ԕ�a"
    outputStartCellAddress = "A3"
    
    ' �X�e�[�^�X�o�[�\��
    Application.StatusBar = "�h���i�ԕʏW�v���J�n���܂�..."
    
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
    colHinban = GetColumnIndex(tblSource, "�i��")
    colJisseki = GetColumnIndex(tblSource, "����")
    colDandori = GetColumnIndex(tblSource, "�i�掞��")
    colKadou = GetColumnIndex(tblSource, "�ғ�����")
    colFuryo = GetColumnIndex(tblSource, "�s��")
    
    If colDate = 0 Or colProcess = 0 Or colHinban = 0 Or colJisseki = 0 Or colDandori = 0 Or colKadou = 0 Or colFuryo = 0 Then
        MsgBox "�u�S�H���v�e�[�u���ɕK�v�ȗ�i���t, �H��, �i��, ����, �i�掞��, �ғ�����, �s�ǁj��������܂���B�񖼂��m�F���Ă��������B", vbCritical
        GoTo Cleanup
    End If
    
    ' 3. �f�[�^�W�v (Dictionary���g�p)
    Set dict = CreateObject("Scripting.Dictionary")
    dataArray = tblSource.DataBodyRange.Value2 ' �������̂��ߔz��ŏ���
    
    Application.StatusBar = "�f�[�^���W�v��..."
    
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        ' �u�H���v��̒l���u�h���v�Ɗ��S��v����s�𒊏o
        If CStr(dataArray(i, colProcess)) = "�h��" Then
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
            
            ' �i�Ԃ̎擾
            currentHinban = CStr(dataArray(i, colHinban))
            
            ' �����L�[�̍쐬�i���t|�i�ԁj
            dictKey = Format(currentDate, "yyyy/mm/dd") & "|" & currentHinban
            
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
        MsgBox "�H���u�h���v�ɊY������f�[�^���W�v����܂���ł����B", vbInformation
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
        ' �w�b�_�[�������݁i�V�K�e�[�u���̏ꍇ�̂݁j�F�u�ʏ́v���ǉ�
        outputHeader.Resize(1, 7).Value = Array("���t", "�i��", "�ʏ�", "����", "�s��", "�ғ�����", "�i�掞��")
    End If
    
    ' 6. �W�v���ʂ�z��ɕϊ����A���t�E�i�ԂŃ\�[�g
    If dict.Count > 0 Then
        Application.StatusBar = "�f�[�^���\�[�g��..."
        
        ' �\�[�g�p�̃L�[�z����쐬
        ReDim sortKeys(1 To dict.Count)
        i = 0
        For Each key In dict.Keys
            i = i + 1
            sortKeys(i) = CStr(key)
        Next key
        
        ' �o�u���\�[�g�œ��t�E�i�ԏ��ɕ��בւ�
        For i = 1 To dict.Count - 1
            For j = i + 1 To dict.Count
                If sortKeys(i) > sortKeys(j) Then
                    tempKey = sortKeys(i)
                    sortKeys(i) = sortKeys(j)
                    sortKeys(j) = tempKey
                End If
            Next j
        Next i
        
        ' �o�͔z��̍쐬�F�u�ʏ́v���ǉ����ė񐔂�7��
        ReDim outputArray(1 To dict.Count, 1 To 7)
        For r = 1 To dict.Count
            key = sortKeys(r)
            item = dict(key)
            
            ' �L�[�𕪉����ē��t�ƕi�Ԃ��擾
            Dim keyParts() As String
            keyParts = Split(key, "|")
            
            outputArray(r, 1) = CDate(keyParts(0))          '���t
            outputArray(r, 2) = keyParts(1)                 '�i��
            outputArray(r, 3) = GetTsusho(CStr(keyParts(1))) '�ʏ�
            outputArray(r, 4) = item(0)                     '����
            outputArray(r, 5) = item(1)                     '�s��
            outputArray(r, 6) = item(2)                     '�ғ�����
            outputArray(r, 7) = item(3)                     '�i�掞��
        Next r
        
        ' 7. �f�[�^�o��
        Application.StatusBar = "�f�[�^���o�͒�..."
        
        If Not isNewTable Then
            ' �����e�[�u���̏ꍇ�F�e�[�u���T�C�Y������Ƀf�[�^��������
            tblOutput.Resize outputHeader.Resize(UBound(outputArray, 1) + 1, 7) ' �񐔂�7��
        End If
        outputHeader.Offset(1, 0).Resize(UBound(outputArray, 1), UBound(outputArray, 2)).Value = outputArray
    End If
    
    ' 8. �e�[�u���쐬�i�V�K�̏ꍇ�̂݁j�܂��͍X�V
    If isNewTable Then
        ' �f�[�^���Ȃ��ꍇ�ł��w�b�_�[�݂̂̃e�[�u�����쐬
        Dim dataRangeForTable As Range
        If dict.Count > 0 Then
            Set dataRangeForTable = outputHeader.Resize(dict.Count + 1, 7) ' �񐔂�7��
        Else
            Set dataRangeForTable = outputHeader.Resize(1, 7) ' �w�b�_�[�̂݁A�񐔂�7��
        End If
        
        Set tblOutput = wsOutput.ListObjects.Add(xlSrcRange, dataRangeForTable, , xlYes)
        tblOutput.Name = outputTableName
        tblOutput.TableStyle = "TableStyleMedium9"
    ElseIf dict.Count = 0 Then
        ' �����e�[�u���Ńf�[�^���Ȃ��ꍇ�F�w�b�_�[�݂̂Ƀ��T�C�Y
        tblOutput.Resize outputHeader.Resize(1, 7) ' �񐔂�7��
    End If
    
    ' �e�[�u���̃t�B���^�[�{�^�����\���ɐݒ�
    If Not tblOutput Is Nothing Then
      tblOutput.ShowAutoFilter = False
    End If
    
    ' ���t��̏����ݒ�
    If dict.Count > 0 And Not tblOutput Is Nothing Then
        tblOutput.ListColumns("���t").DataBodyRange.NumberFormatLocal = "yyyy/mm/dd"
    End If
    
    ' ========== �ǉ������ݒ� ==========
    Application.StatusBar = "������ݒ蒆..."
    
    If Not tblOutput Is Nothing Then
        ' 1. �f�[�^�͈͂́u�k�����đS�̂�\������v�ݒ�
        If dict.Count > 0 Then
            tblOutput.DataBodyRange.ShrinkToFit = True
        End If
        
        ' 2. �S��̗񕝂�6.4�ɐݒ� (�u�ʏ́v����܂�)
        Dim col As ListColumn
        For Each col In tblOutput.ListColumns
            col.Range.ColumnWidth = 6.4
        Next col
        
        ' 3. �u�ғ����ԁv�u�i�掞�ԁv��̏����_�ȉ�2���ݒ�
        If dict.Count > 0 Then
            On Error Resume Next ' �񂪑��݂��Ȃ��ꍇ�̃G���[�����
            tblOutput.ListColumns("�ғ�����").DataBodyRange.NumberFormatLocal = "0.00"
            tblOutput.ListColumns("�i�掞��").DataBodyRange.NumberFormatLocal = "0.00"
            On Error GoTo ErrorHandler
        End If
    End If
    
    ' 4. A1�Z���Ƀ^�C�g����ݒ�
    With wsOutput.Range("A1")
        .Value = "�h���i�ԕʃf�[�^���o"
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

' �i�Ԃ���ʏ̂��擾����w���p�[�֐�
Private Function GetTsusho(hinban As String) As String
    Dim �A�����F��F_�i�� As Variant
    Dim �A�����F��R_�i�� As Variant
    Dim �m�A���H�NF_�i�� As Variant
    Dim �m�A���H�NR_�i�� As Variant
    Dim i As Long

    �A�����F��F_�i�� = Array("58020F", "58030F", "58040F", "58050F", "58060F")
    �A�����F��R_�i�� = Array("58020R", "58030R", "58040R", "58050R", "58060R")
    �m�A���H�NF_�i�� = Array("28030F", "28040F")
    �m�A���H�NR_�i�� = Array("28030R", "28040R")

    For i = LBound(�A�����F��F_�i��) To UBound(�A�����F��F_�i��)
        If InStr(1, hinban, CStr(�A�����F��F_�i��(i)), vbTextCompare) > 0 Then
            GetTsusho = "�A�����F��F"
            Exit Function
        End If
    Next i

    For i = LBound(�A�����F��R_�i��) To UBound(�A�����F��R_�i��)
        If InStr(1, hinban, CStr(�A�����F��R_�i��(i)), vbTextCompare) > 0 Then
            GetTsusho = "�A�����F��R"
            Exit Function
        End If
    Next i

    For i = LBound(�m�A���H�NF_�i��) To UBound(�m�A���H�NF_�i��)
        If InStr(1, hinban, CStr(�m�A���H�NF_�i��(i)), vbTextCompare) > 0 Then
            GetTsusho = "�m�A���H�NF"
            Exit Function
        End If
    Next i

    For i = LBound(�m�A���H�NR_�i��) To UBound(�m�A���H�NR_�i��)
        If InStr(1, hinban, CStr(�m�A���H�NR_�i��(i)), vbTextCompare) > 0 Then
            GetTsusho = "�m�A���H�NR"
            Exit Function
        End If
    Next i

    GetTsusho = "�⋋�i"
End Function

