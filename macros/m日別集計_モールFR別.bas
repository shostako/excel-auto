Attribute VB_Name = "m���ʏW�v_���[��FR��"
Option Explicit

' ���[��F/R�����W�v�}�N���i��ʍX�V�}���Łj
' �u�S�H���v�e�[�u������u���[���v�H���̃f�[�^����t�EF/R�ŃO���[�v�����ďW�v
Sub ���ʏW�v_���[��FR��()
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim tblSource As ListObject
    Dim tblOutput As ListObject
    Dim dict As Object 'Scripting.Dictionary
    Dim outputArray() As Variant
    Dim dataArray As Variant
    Dim sortKeys() As String ' �\�[�g�p�̃L�[�z��
    
    ' �Ώەi�ԃ��X�g�i���[��F�p�j
    Dim targetHinbanListF As Variant
    targetHinbanListF = Array("58020F", "58030F", "58040F", "58050F", "58060F", "58830F", _
                            "58021", "58022", "58031", "58032", "58041", "58042", _
                            "58051", "58052", "58061", "58062", "47030F", "47030R", _
                            "47031", "47032", "47035", "47036", "58221F", "58223", "58224")
    
    ' �Ώەi�ԃ��X�g�i���[��R�p�j
    Dim targetHinbanListR As Variant
    targetHinbanListR = Array("58020R", "58030R", "58040R", "58050R", "58060R", "58830R", _
                            "58025", "58026", "58035", "58036", "58045", "58046", _
                            "58055", "58056", "58065", "58066", "58221R", "58015", "58016")
    
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
    Dim currentFR As String
    Dim dictKey As String
    Dim jissekiVal As Double, dandoriVal As Double, kadouVal As Double, furyoVal As Double
    Dim item As Variant
    Dim key As Variant
    Dim tempKey As String
    Dim hinbanValue As String
    Dim isTargetHinban As Boolean
    
    ' ��{�ݒ�
    Set wb = ThisWorkbook
    sourceSheetName = "�S�H��"
    sourceTableName = "_�S�H��"
    outputSheetName = "���[��FR��"
    outputTableName = "_���[��FR��a"
    outputStartCellAddress = "A3"
    
    ' �X�e�[�^�X�o�[�\��
    Application.StatusBar = "���[��FR�ʏW�v���J�n���܂�..."
    
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
    
    ' 3. �f�[�^�W�v (Dictionary�g�p)
    Set dict = CreateObject("Scripting.Dictionary")
    dataArray = tblSource.DataBodyRange.Value2 ' �������̂��ߔz��ŏ���
    
    Application.StatusBar = "�f�[�^���W�v��..."
    
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        ' �u�H���v��̒l���u���[���v���܂ނ��`�F�b�N
        If InStr(1, CStr(dataArray(i, colProcess)), "���[��", vbTextCompare) > 0 Then
            ' �u�i�ԁv�񂪑Ώۃ��X�g�Ɋ܂܂�邩�`�F�b�N
            hinbanValue = CStr(dataArray(i, colHinban))
            isTargetHinban = False
            currentFR = ""
            
            ' ���[��F�i�ԃ`�F�b�N
            For k = LBound(targetHinbanListF) To UBound(targetHinbanListF)
                If InStr(1, hinbanValue, targetHinbanListF(k), vbTextCompare) > 0 Then
                    isTargetHinban = True
                    currentFR = "F"
                    Exit For
                End If
            Next k
            
            ' ���[��R�i�ԃ`�F�b�N�iF�Ō�����Ȃ������ꍇ�j
            If Not isTargetHinban Then
                For k = LBound(targetHinbanListR) To UBound(targetHinbanListR)
                    If InStr(1, hinbanValue, targetHinbanListR(k), vbTextCompare) > 0 Then
                        isTargetHinban = True
                        currentFR = "R"
                        Exit For
                    End If
                Next k
            End If
            
            If isTargetHinban Then
                ' ���t�̑Ó����`�F�b�N�ƕϊ�
                If IsDate(dataArray(i, colDate)) Then
                    currentDate = CDate(dataArray(i, colDate))
                ElseIf IsNumeric(dataArray(i, colDate)) Then
                    ' ���l�̏ꍇ�͓��t�V���A���l�Ƃ��Ĉ���
                    currentDate = CDate(CLng(dataArray(i, colDate)))
                Else
                    ' ���t�Ƃ��ĔF���ł��Ȃ��f�[�^�̓X�L�b�v
                    Debug.Print "�x��: ���t�Ƃ��ĔF���ł��Ȃ��f�[�^������܂����B�s " & i + tblSource.HeaderRowRange.row & ", �l: " & dataArray(i, colDate)
                    GoTo NextIteration
                End If
                
                ' �����L�[�̍쐬�i���t|F/R�j
                dictKey = Format(currentDate, "yyyy/mm/dd") & "|" & currentFR
                
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
        End If
NextIteration:
    Next i
    
    If dict.Count = 0 Then
        MsgBox "���[���H���i�w��i�ԁj�ɊY������f�[�^���W�v����܂���ł����B", vbInformation
        ' ���̏ꍇ�ł��V�[�g�Ƌ�̃e�[�u���͍쐬����悤�ɑ��s
    End If
    
    ' 4. �o�͐�V�[�g�̏���
    Application.StatusBar = "�o�͐�V�[�g��������..."
    
    On Error Resume Next
    Set wsOutput = wb.Sheets(outputSheetName)
    If wsOutput Is Nothing Then
        Set wsOutput = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsOutput.Name = outputSheetName
    End If
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
        ' �w�b�_�[�s��ݒ�i�V�K�e�[�u���̏ꍇ�̂݁j�F�uF/R�v���ǉ�
        outputHeader.Resize(1, 6).Value = Array("���t", "F/R", "����", "�s��", "�ғ�����", "�i�掞��")
    End If
    
    ' 6. �W�v���ʂ�z��ɕϊ����A���t�EF/R�Ń\�[�g
    If dict.Count > 0 Then
        Application.StatusBar = "�f�[�^���\�[�g��..."
        
        ' �\�[�g�p�̃L�[�z����쐬
        ReDim sortKeys(1 To dict.Count)
        i = 0
        For Each key In dict.Keys
            i = i + 1
            sortKeys(i) = CStr(key)
        Next key
        
        ' �o�u���\�[�g�œ��t�EF/R���ɕ��ёւ�
        For i = 1 To dict.Count - 1
            For j = i + 1 To dict.Count
                If sortKeys(i) > sortKeys(j) Then
                    tempKey = sortKeys(i)
                    sortKeys(i) = sortKeys(j)
                    sortKeys(j) = tempKey
                End If
            Next j
        Next i
        
        ' �o�͔z��̍쐬�F�uF/R�v���ǉ����ė񐔂�6��
        ReDim outputArray(1 To dict.Count, 1 To 6)
        For r = 1 To dict.Count
            key = sortKeys(r)
            item = dict(key)
            
            ' �L�[�𕪉����ē��t��F/R���擾
            Dim keyParts() As String
            keyParts = Split(key, "|")
            
            outputArray(r, 1) = CDate(keyParts(0))          '���t
            outputArray(r, 2) = keyParts(1)                 'F/R
            outputArray(r, 3) = item(0)                     '����
            outputArray(r, 4) = item(1)                     '�s��
            outputArray(r, 5) = item(2)                     '�ғ�����
            outputArray(r, 6) = item(3)                     '�i�掞��
        Next r
        
        ' 7. �f�[�^�o��
        Application.StatusBar = "�f�[�^���o�͒�..."
        
        If Not isNewTable Then
            ' �����e�[�u���̏ꍇ�F�e�[�u���T�C�Y�𒲐����ăf�[�^��}��
            tblOutput.Resize outputHeader.Resize(UBound(outputArray, 1) + 1, 6) ' �񐔂�6��
        End If
        outputHeader.Offset(1, 0).Resize(UBound(outputArray, 1), UBound(outputArray, 2)).Value = outputArray
    End If
    
    ' 8. �e�[�u���쐬�i�V�K�̏ꍇ�̂݁j�܂��͍X�V
    If isNewTable Then
        ' �f�[�^���Ȃ��ꍇ�ł��w�b�_�[�݂̂̃e�[�u�����쐬
        Dim dataRangeForTable As Range
        If dict.Count > 0 Then
            Set dataRangeForTable = outputHeader.Resize(dict.Count + 1, 6) ' �񐔂�6��
        Else
            Set dataRangeForTable = outputHeader.Resize(1, 6) ' �w�b�_�[�̂݁A�񐔂�6��
        End If
        
        Set tblOutput = wsOutput.ListObjects.Add(xlSrcRange, dataRangeForTable, , xlYes)
        tblOutput.Name = outputTableName
        tblOutput.TableStyle = "TableStyleMedium9"
    ElseIf dict.Count = 0 Then
        ' �����e�[�u���Ńf�[�^���Ȃ��ꍇ�F�w�b�_�[�݂̂Ƀ��T�C�Y
        tblOutput.Resize outputHeader.Resize(1, 6) ' �񐔂�6��
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
    Application.StatusBar = "�����ݒ蒆..."
    
    If Not tblOutput Is Nothing Then
        ' 1. �f�[�^�͈͂́u�k�����đS�̂�\������v�ݒ�
        If dict.Count > 0 Then
            tblOutput.DataBodyRange.ShrinkToFit = True
        End If
        
        ' 2. �S��̗񕝂�6.4�ɐݒ�
        Dim col As ListColumn
        For Each col In tblOutput.ListColumns
            col.Range.ColumnWidth = 6.4
        Next col
        
        ' 3. �u�ғ����ԁv�u�i�掞�ԁv��̏����F�����_�ȉ�2���ݒ�
        If dict.Count > 0 Then
            On Error Resume Next ' �񂪑��݂��Ȃ��ꍇ�̃G���[����
            tblOutput.ListColumns("�ғ�����").DataBodyRange.NumberFormatLocal = "0.00"
            tblOutput.ListColumns("�i�掞��").DataBodyRange.NumberFormatLocal = "0.00"
            On Error GoTo ErrorHandler
        End If
    End If
    
    ' 4. A1�Z���Ƀ^�C�g����ݒ�
    With wsOutput.Range("A1")
        .Value = "���[��FR�ʃf�[�^�o��"
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
    
    Application.StatusBar = False
End Sub

' �e�[�u���̗񖼂���C���f�b�N�X���擾����w���p�[�֐�
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    Dim col As ListColumn
    Dim i As Long
    i = 0
    On Error Resume Next
    i = tbl.ListColumns(columnName).Index
    On Error GoTo 0
    GetColumnIndex = i
End Function
