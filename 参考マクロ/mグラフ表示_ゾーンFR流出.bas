Attribute VB_Name = "m���ԏW�v_�ʏ̕�b"
Sub ���ԏW�v_�ʏ̕�b()
    ' �u�i�ԕʁv�V�[�g�́u_�i�ԕʁv�e�[�u�����璼�ڒʏ̕ʏW�v���s���A
    ' �u�i�ԕ�bb�v�V�[�g�́u_�i�ԕ�bb�v�e�[�u���ɏo�͂���}�N��
    ' �i�Ԃ�ʏ̂ɕϊ����Ă���O���[�v�����A�s�Ǎ��ڂ͗��Ƃ��Čv�Z����
    
    ' �X�e�[�^�X�o�[�ɏ����󋵂�\��
    Application.StatusBar = "�ʏ̕ʒ��ڏW�vb: �������J�n���܂�..."
    
    ' �ϐ��錾
    Dim srcSheet As Worksheet, destSheet As Worksheet
    Dim srcTable As ListObject, destTable As ListObject
    Dim srcData As Variant
    Dim StartDate As Double, EndDate As Double  ' ���t���V���A���l�Ƃ��Ĉ���
    Dim dictGroups As Object
    Dim dictSums As Object
    Dim dictCounts As Object
    Dim headerRow As Range
    Dim i As Long, j As Long
    Dim key As Variant
    Dim destRow As Long
    Dim tempValue As Variant
    Dim tsushoArr() As Variant
    Dim useFilter As Boolean
    Dim rowDateValue As Double
    Dim dataStartRow As Long
    Dim dataEndRow As Long
    Dim tableRange As Range
    Dim hinban As String
    Dim tsusho As String
    Dim isInDateRange As Boolean
    Dim tableFound As Boolean
    Dim dataRng As Range
    Dim lastRow As Long, lastCol As Long
    
    ' ��C���f�b�N�X�p�ϐ�
    Dim hinbanCol As Integer, dateCol As Integer
    Dim kataKaeCol As Integer, kadoCol As Integer, cycleCol As Integer
    Dim shotCol As Integer, furyoCol As Integer
    Dim uchidashiCol As Integer, shortCol As Integer, weldCol As Integer
    Dim shiwaCol As Integer, ibutsuCol As Integer, silverCol As Integer
    Dim flowCol As Integer, gomiCol As Integer, gcKasuCol As Integer
    Dim kizuCol As Integer, hikeCol As Integer, itohikiCol As Integer
    Dim kataYogoreCol As Integer, makureCol As Integer, toridashiFuryoCol As Integer
    Dim wareHakukaCol As Integer, coreKasuCol As Integer, sonotaCol As Integer
    Dim chocoCol As Integer, kensaCol As Integer, ryushutuCol As Integer
    
    ' �G���[�n���h�����O�ݒ�
    On Error GoTo ErrorHandler
    
    ' ���͌��V�[�g�̎擾
    On Error Resume Next
    Set srcSheet = ThisWorkbook.Worksheets("�i�ԕ�")
    On Error GoTo 0
    
    If srcSheet Is Nothing Then
        Application.StatusBar = "�ʏ̕ʒ��ڏW�vb: �u�i�ԕʁv�V�[�g��������܂���B"
        Exit Sub
    End If
    
    ' �o�͐�V�[�g�̐ݒ�
    On Error Resume Next
    Set destSheet = ThisWorkbook.Worksheets("�i�ԕ�bb")
    On Error GoTo 0
    
    If destSheet Is Nothing Then
        ' �V�[�g�����݂��Ȃ��ꍇ�͐V�K�쐬
        Set destSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        destSheet.Name = "�i�ԕ�bb"
    End If
    
    ' �X�e�[�^�X�o�[���X�V
    Application.StatusBar = "�ʏ̕ʒ��ڏW�vb: ���t�������擾��..."
    
    ' �J�n���ƏI�������擾�i�i�ԕ�bb�V�[�g����j
    On Error Resume Next
    StartDate = CDbl(destSheet.Range("B1").Value)
    EndDate = CDbl(destSheet.Range("B2").Value)
    On Error GoTo 0
    
    ' ���t���ݒ肳��Ă��邩�`�F�b�N
    useFilter = (StartDate > 0) And (EndDate > 0)
    
    ' �e�[�u���̌���
    tableFound = False
    On Error Resume Next
    Set srcTable = srcSheet.ListObjects("_�i�ԕ�")
    On Error GoTo 0
    
    If Not srcTable Is Nothing Then
        tableFound = True
    End If
    
    ' �e�[�u����������Ȃ��ꍇ�́A�f�[�^�͈͂��������ĕϊ�
    If Not tableFound Then
        ' �f�[�^�͈͂����
        If Not IsEmpty(srcSheet.Range("A1").Value) Then
            lastRow = srcSheet.Cells(srcSheet.Rows.Count, 1).End(xlUp).Row
            lastCol = srcSheet.Cells(1, srcSheet.Columns.Count).End(xlToLeft).Column
            
            If lastRow > 1 Then  ' �w�b�_�[�s�ƃf�[�^�����Ȃ��Ƃ�1�s����
                Set dataRng = srcSheet.Range(srcSheet.Cells(1, 1), srcSheet.Cells(lastRow, lastCol))
                
                ' �f�[�^�͈͂��e�[�u���ɕϊ�
                On Error Resume Next
                Set srcTable = srcSheet.ListObjects.Add(xlSrcRange, dataRng, , xlYes)
                If Err.Number = 0 Then
                    srcTable.Name = "_�i�ԕ�"
                    tableFound = True
                End If
                On Error GoTo 0
            End If
        End If
    End If
    
    ' ����ł��e�[�u����������Ȃ��ꍇ�͏������~
    If Not tableFound Then
        Application.StatusBar = "�ʏ̕ʒ��ڏW�vb: �e�[�u���u_�i�ԕʁv��������܂���B"
        Exit Sub
    End If
    
    ' �X�e�[�^�X�o�[���X�V
    Application.StatusBar = "�ʏ̕ʒ��ڏW�vb: �f�[�^�擾��..."
    
    ' ���f�[�^�̎擾
    srcData = srcTable.DataBodyRange.Value
    
    ' ��̃C���f�b�N�X���擾
    Set headerRow = srcTable.HeaderRowRange
    
    For i = 1 To headerRow.Cells.Count
        Select Case headerRow.Cells(1, i).Value
            Case "�i��"
                hinbanCol = i
            Case "���t"
                dateCol = i
            Case "�^��"
                kataKaeCol = i
            Case "�ғ�"
                kadoCol = i
            Case "�T�C�N��"
                cycleCol = i
            Case "�V���b�g��"
                shotCol = i
            Case "�s�ǐ�"
                furyoCol = i
            Case "�ŏo��"
                uchidashiCol = i
            Case "�V���[�g"
                shortCol = i
            Case "�E�G���h"
                weldCol = i
            Case "�V��"
                shiwaCol = i
            Case "�ٕ�"
                ibutsuCol = i
            Case "�V���o�["
                silverCol = i
            Case "�t���[�}�[�N"
                flowCol = i
            Case "�S�~����"
                gomiCol = i
            Case "GC�J�X"
                gcKasuCol = i
            Case "�L�Y"
                kizuCol = i
            Case "�q�P"
                hikeCol = i
            Case "������"
                itohikiCol = i
            Case "�^����"
                kataYogoreCol = i
            Case "�}�N��"
                makureCol = i
            Case "��o�s��"
                toridashiFuryoCol = i
            Case "���ꔒ��"
                wareHakukaCol = i
            Case "�R�A�J�X"
                coreKasuCol = i
            Case "���̑�"
                sonotaCol = i
            Case "�`���R��ŏo��"
                chocoCol = i
            Case "����"
                kensaCol = i
            Case "���o�s��"
                ryushutuCol = i
        End Select
    Next i
    
    ' �K�v�ȗ񂪌�����Ȃ��ꍇ�͏������~
    If hinbanCol = 0 Or dateCol = 0 Then
        Application.StatusBar = "�ʏ̕ʒ��ڏW�vb: �K�v�ȗ񂪌�����܂���B"
        Exit Sub
    End If
    
    ' �X�e�[�^�X�o�[���X�V
    Application.StatusBar = "�ʏ̕ʒ��ڏW�vb: �f�[�^�W�v��..."
    
    ' Dictionary�I�u�W�F�N�g���쐬
    Set dictGroups = CreateObject("Scripting.Dictionary")
    Set dictSums = CreateObject("Scripting.Dictionary")
    Set dictCounts = CreateObject("Scripting.Dictionary")
    
    ' �f�[�^���O���[�v�����ďW�v
    For i = 1 To UBound(srcData, 1)
        ' ���t�̊m�F
        isInDateRange = True
        
        If useFilter Then
            ' ���t�̃V���A���l���擾
            rowDateValue = CDbl(srcData(i, dateCol))
            
            ' ���t�͈͓����ǂ����`�F�b�N
            isInDateRange = (rowDateValue >= StartDate And rowDateValue <= EndDate)
        End If
        
        ' ���t�͈͓��̃f�[�^�̂ݏ���
        If isInDateRange Then
            ' �i�Ԃ���ʏ̂𔻒�
            hinban = srcData(i, hinbanCol)
            tsusho = �i�Ԓʏ̔���(hinban)
            
            ' �V�����ʏ̂̏ꍇ�ADictionary�ɒǉ�
            If Not dictGroups.Exists(tsusho) Then
                dictGroups.Add tsusho, tsusho
                
                ' �W�v�p��Dictionary��������
                Set dictSums(tsusho) = CreateObject("Scripting.Dictionary")
                Set dictCounts(tsusho) = CreateObject("Scripting.Dictionary")
                
                ' �e�W�v���ڂ�������
                dictSums(tsusho)("�^��") = 0
                dictSums(tsusho)("�ғ�") = 0
                dictSums(tsusho)("�T�C�N��") = 0
                dictCounts(tsusho)("�T�C�N��") = 0  ' �T�C�N�����όv�Z�p
                dictSums(tsusho)("�V���b�g��") = 0
                dictSums(tsusho)("�s�ǐ�") = 0
                dictSums(tsusho)("�ŏo��") = 0
                dictSums(tsusho)("�V���[�g") = 0
                dictSums(tsusho)("�E�G���h") = 0
                dictSums(tsusho)("�V��") = 0
                dictSums(tsusho)("�ٕ�") = 0
                dictSums(tsusho)("�V���o�[") = 0
                dictSums(tsusho)("�t���[�}�[�N") = 0
                dictSums(tsusho)("�S�~����") = 0
                dictSums(tsusho)("GC�J�X") = 0
                dictSums(tsusho)("�L�Y") = 0
                dictSums(tsusho)("�q�P") = 0
                dictSums(tsusho)("������") = 0
                dictSums(tsusho)("�^����") = 0
                dictSums(tsusho)("�}�N��") = 0
                dictSums(tsusho)("��o�s��") = 0
                dictSums(tsusho)("���ꔒ��") = 0
                dictSums(tsusho)("�R�A�J�X") = 0
                dictSums(tsusho)("���̑�") = 0
                dictSums(tsusho)("�`���R��ŏo��") = 0
                dictSums(tsusho)("����") = 0
                dictSums(tsusho)("���o�s��") = 0
            End If
            
            ' �e���ڂ̍��v�l���X�V
            If kataKaeCol > 0 And IsNumeric(srcData(i, kataKaeCol)) Then
                dictSums(tsusho)("�^��") = dictSums(tsusho)("�^��") + CDbl(srcData(i, kataKaeCol))
            End If
            
            If kadoCol > 0 And IsNumeric(srcData(i, kadoCol)) Then
                dictSums(tsusho)("�ғ�") = dictSums(tsusho)("�ғ�") + CDbl(srcData(i, kadoCol))
            End If
            
            If cycleCol > 0 And IsNumeric(srcData(i, cycleCol)) And srcData(i, cycleCol) <> 0 Then
                dictSums(tsusho)("�T�C�N��") = dictSums(tsusho)("�T�C�N��") + CDbl(srcData(i, cycleCol))
                dictCounts(tsusho)("�T�C�N��") = dictCounts(tsusho)("�T�C�N��") + 1
            End If
            
            If shotCol > 0 And IsNumeric(srcData(i, shotCol)) Then
                dictSums(tsusho)("�V���b�g��") = dictSums(tsusho)("�V���b�g��") + CDbl(srcData(i, shotCol))
            End If
            
            If furyoCol > 0 And IsNumeric(srcData(i, furyoCol)) Then
                dictSums(tsusho)("�s�ǐ�") = dictSums(tsusho)("�s�ǐ�") + CDbl(srcData(i, furyoCol))
            End If
            
            ' �s�Ǎ��ڂ̏W�v
            If uchidashiCol > 0 And IsNumeric(srcData(i, uchidashiCol)) Then
                dictSums(tsusho)("�ŏo��") = dictSums(tsusho)("�ŏo��") + CDbl(srcData(i, uchidashiCol))
            End If
            
            If shortCol > 0 And IsNumeric(srcData(i, shortCol)) Then
                dictSums(tsusho)("�V���[�g") = dictSums(tsusho)("�V���[�g") + CDbl(srcData(i, shortCol))
            End If
            
            If weldCol > 0 And IsNumeric(srcData(i, weldCol)) Then
                dictSums(tsusho)("�E�G���h") = dictSums(tsusho)("�E�G���h") + CDbl(srcData(i, weldCol))
            End If
            
            If shiwaCol > 0 And IsNumeric(srcData(i, shiwaCol)) Then
                dictSums(tsusho)("�V��") = dictSums(tsusho)("�V��") + CDbl(srcData(i, shiwaCol))
            End If
            
            If ibutsuCol > 0 And IsNumeric(srcData(i, ibutsuCol)) Then
                dictSums(tsusho)("�ٕ�") = dictSums(tsusho)("�ٕ�") + CDbl(srcData(i, ibutsuCol))
            End If
            
            If silverCol > 0 And IsNumeric(srcData(i, silverCol)) Then
                dictSums(tsusho)("�V���o�[") = dictSums(tsusho)("�V���o�[") + CDbl(srcData(i, silverCol))
            End If
            
            If flowCol > 0 And IsNumeric(srcData(i, flowCol)) Then
                dictSums(tsusho)("�t���[�}�[�N") = dictSums(tsusho)("�t���[�}�[�N") + CDbl(srcData(i, flowCol))
            End If
            
            If gomiCol > 0 And IsNumeric(srcData(i, gomiCol)) Then
                dictSums(tsusho)("�S�~����") = dictSums(tsusho)("�S�~����") + CDbl(srcData(i, gomiCol))
            End If
            
            If gcKasuCol > 0 And IsNumeric(srcData(i, gcKasuCol)) Then
                dictSums(tsusho)("GC�J�X") = dictSums(tsusho)("GC�J�X") + CDbl(srcData(i, gcKasuCol))
            End If
            
            If kizuCol > 0 And IsNumeric(srcData(i, kizuCol)) Then
                dictSums(tsusho)("�L�Y") = dictSums(tsusho)("�L�Y") + CDbl(srcData(i, kizuCol))
            End If
            
            If hikeCol > 0 And IsNumeric(srcData(i, hikeCol)) Then
                dictSums(tsusho)("�q�P") = dictSums(tsusho)("�q�P") + CDbl(srcData(i, hikeCol))
            End If
            
            If itohikiCol > 0 And IsNumeric(srcData(i, itohikiCol)) Then
                dictSums(tsusho)("������") = dictSums(tsusho)("������") + CDbl(srcData(i, itohikiCol))
            End If
            
            If kataYogoreCol > 0 And IsNumeric(srcData(i, kataYogoreCol)) Then
                dictSums(tsusho)("�^����") = dictSums(tsusho)("�^����") + CDbl(srcData(i, kataYogoreCol))
            End If
            
            If makureCol > 0 And IsNumeric(srcData(i, makureCol)) Then
                dictSums(tsusho)("�}�N��") = dictSums(tsusho)("�}�N��") + CDbl(srcData(i, makureCol))
            End If
            
            If toridashiFuryoCol > 0 And IsNumeric(srcData(i, toridashiFuryoCol)) Then
                dictSums(tsusho)("��o�s��") = dictSums(tsusho)("��o�s��") + CDbl(srcData(i, toridashiFuryoCol))
            End If
            
            If wareHakukaCol > 0 And IsNumeric(srcData(i, wareHakukaCol)) Then
                dictSums(tsusho)("���ꔒ��") = dictSums(tsusho)("���ꔒ��") + CDbl(srcData(i, wareHakukaCol))
            End If
            
            If coreKasuCol > 0 And IsNumeric(srcData(i, coreKasuCol)) Then
                dictSums(tsusho)("�R�A�J�X") = dictSums(tsusho)("�R�A�J�X") + CDbl(srcData(i, coreKasuCol))
            End If
            
            If sonotaCol > 0 And IsNumeric(srcData(i, sonotaCol)) Then
                dictSums(tsusho)("���̑�") = dictSums(tsusho)("���̑�") + CDbl(srcData(i, sonotaCol))
            End If
            
            If chocoCol > 0 And IsNumeric(srcData(i, chocoCol)) Then
                dictSums(tsusho)("�`���R��ŏo��") = dictSums(tsusho)("�`���R��ŏo��") + CDbl(srcData(i, chocoCol))
            End If
            
            If kensaCol > 0 And IsNumeric(srcData(i, kensaCol)) Then
                dictSums(tsusho)("����") = dictSums(tsusho)("����") + CDbl(srcData(i, kensaCol))
            End If
            
            If ryushutuCol > 0 And IsNumeric(srcData(i, ryushutuCol)) Then
                dictSums(tsusho)("���o�s��") = dictSums(tsusho)("���o�s��") + CDbl(srcData(i, ryushutuCol))
            End If
        End If
    Next i
    
    ' �X�e�[�^�X�o�[���X�V
    Application.StatusBar = "�ʏ̕ʒ��ڏW�vb: �f�[�^�o�͏�����..."
    
    ' �o�͐�V�[�g��4�s�ڈȍ~���N���A�i1-3�s�ڂ͎c���j
    ' �o�͐�V�[�g��4�s�ڂ���31�s�ڂ܂ł��N���A
    destSheet.Range("A4:AB31").Clear
    
    ' 4�s�ڈȍ~�̏����ݒ�
    With destSheet.Range("A4:AB" & destSheet.Rows.Count)
        .Font.Name = "Yu Gothic UI"
        .Font.Size = 11
    End With
    
    ' �^�C�g���s�̍쐬�i4�s�ځj
    destRow = 4
    If useFilter Then
        destSheet.Range("A" & destRow).Value = "���`�ʏ̕ʕs�ǏW�v�F" & Format(StartDate, "yyyy/mm/dd") & "�`" & Format(EndDate, "yyyy/mm/dd")
    Else
        destSheet.Range("A" & destRow).Value = "���`�ʏ̕ʕs�ǏW�v�F�S����"
    End If
    destSheet.Range("A" & destRow).Font.Bold = True
    
    ' �w�b�_�[�s�̍쐬�i5�s�ځj
    destRow = 5
    destSheet.Range("A" & destRow).Value = "�ʏ�"
    destSheet.Range("B" & destRow).Value = "�^��"
    destSheet.Range("C" & destRow).Value = "�ғ�"
    destSheet.Range("D" & destRow).Value = "�T�C�N��"
    destSheet.Range("E" & destRow).Value = "�V���b�g��"
    destSheet.Range("F" & destRow).Value = "�s�ǐ�"
    destSheet.Range("G" & destRow).Value = "�s�Ǘ�"
    destSheet.Range("H" & destRow).Value = "�ŏo��"
    destSheet.Range("I" & destRow).Value = "�V���[�g"
    destSheet.Range("J" & destRow).Value = "�E�G���h"
    destSheet.Range("K" & destRow).Value = "�V��"
    destSheet.Range("L" & destRow).Value = "�ٕ�"
    destSheet.Range("M" & destRow).Value = "�V���o�["
    destSheet.Range("N" & destRow).Value = "�t���[�}�[�N"
    destSheet.Range("O" & destRow).Value = "�S�~����"
    destSheet.Range("P" & destRow).Value = "GC�J�X"
    destSheet.Range("Q" & destRow).Value = "�L�Y"
    destSheet.Range("R" & destRow).Value = "�q�P"
    destSheet.Range("S" & destRow).Value = "������"
    destSheet.Range("T" & destRow).Value = "�^����"
    destSheet.Range("U" & destRow).Value = "�}�N��"
    destSheet.Range("V" & destRow).Value = "��o�s��"
    destSheet.Range("W" & destRow).Value = "���ꔒ��"
    destSheet.Range("X" & destRow).Value = "�R�A�J�X"
    destSheet.Range("Y" & destRow).Value = "���̑�"
    destSheet.Range("Z" & destRow).Value = "�`���R��ŏo��"
    destSheet.Range("AA" & destRow).Value = "����"
    destSheet.Range("AB" & destRow).Value = "���o�s��"
    
    ' �w�b�_�[�s�̏����ݒ�
    With destSheet.Range("A" & destRow & ":AB" & destRow)
        .HorizontalAlignment = xlCenter  ' ��������
        .Font.Bold = True
        .ShrinkToFit = True  ' �k�����đS�̂�\��
    End With
    
    destRow = destRow + 1
    dataStartRow = destRow
    
    ' �f�[�^���Ȃ��ꍇ�̏���
    If dictGroups.Count = 0 Then
        destSheet.Cells(dataStartRow, 1).Value = "�Y���f�[�^�Ȃ�"
        For j = 2 To 28  ' B�񂩂�AB��܂�0�Ŗ��߂�
            destSheet.Cells(dataStartRow, j).Value = 0
        Next j
        
        dataEndRow = dataStartRow
    Else
        ' �ʏ̂̔z����쐬���ă\�[�g
        ReDim tsushoArr(0 To dictGroups.Count - 1)
        i = 0
        For Each key In dictGroups.Keys
            tsushoArr(i) = key
            i = i + 1
        Next key
        
        ' �z����\�[�g�i����̏����Łj
        ' TG �� 62-28030Fr �� 62-28030Rr �� 62-58050Fr �� 62-58050Rr �� �⋋�i
        Dim sortedArr() As Variant
        ReDim sortedArr(0 To UBound(tsushoArr))
        Dim sortIdx As Integer
        sortIdx = 0
        
        ' ���Ԃɔz����쐬
        Dim orderList As Variant
        orderList = Array("TG", "62-28030Fr", "62-28030Rr", "62-58050Fr", "62-58050Rr", "�⋋�i")
        
        For i = 0 To UBound(orderList)
            For j = 0 To UBound(tsushoArr)
                If tsushoArr(j) = orderList(i) Then
                    sortedArr(sortIdx) = tsushoArr(j)
                    sortIdx = sortIdx + 1
                    Exit For
                End If
            Next j
        Next i
        
        ' �X�e�[�^�X�o�[���X�V
        Application.StatusBar = "�ʏ̕ʒ��ڏW�vb: �f�[�^�o�͒�..."
        
        ' �f�[�^�̏�������
        For i = 0 To sortIdx - 1
            key = sortedArr(i)
            
            ' ��{�f�[�^����������
            destSheet.Cells(destRow, 1).Value = key
            destSheet.Cells(destRow, 2).Value = dictSums(key)("�^��")
            destSheet.Cells(destRow, 3).Value = dictSums(key)("�ғ�")
            
            ' �T�C�N���̕��ϒl���v�Z
            If dictCounts(key)("�T�C�N��") > 0 Then
                destSheet.Cells(destRow, 4).Value = dictSums(key)("�T�C�N��") / dictCounts(key)("�T�C�N��")
            Else
                destSheet.Cells(destRow, 4).Value = 0
            End If
            
            destSheet.Cells(destRow, 5).Value = dictSums(key)("�V���b�g��")
            destSheet.Cells(destRow, 6).Value = dictSums(key)("�s�ǐ�")
            
            ' �s�Ǘ��̌v�Z�i�s�ǐ����V���b�g���j
            If dictSums(key)("�V���b�g��") > 0 Then
                destSheet.Cells(destRow, 7).Value = dictSums(key)("�s�ǐ�") / dictSums(key)("�V���b�g��")
            Else
                destSheet.Cells(destRow, 7).Value = 0
            End If
            
            ' �s�Ǎ��ڃf�[�^�𗦂Ƃ��Čv�Z���ď�������
            ' ���ׂāu���ڒl���V���b�g���v�Ōv�Z
            If dictSums(key)("�V���b�g��") > 0 Then
                destSheet.Cells(destRow, 8).Value = dictSums(key)("�ŏo��") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 9).Value = dictSums(key)("�V���[�g") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 10).Value = dictSums(key)("�E�G���h") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 11).Value = dictSums(key)("�V��") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 12).Value = dictSums(key)("�ٕ�") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 13).Value = dictSums(key)("�V���o�[") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 14).Value = dictSums(key)("�t���[�}�[�N") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 15).Value = dictSums(key)("�S�~����") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 16).Value = dictSums(key)("GC�J�X") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 17).Value = dictSums(key)("�L�Y") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 18).Value = dictSums(key)("�q�P") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 19).Value = dictSums(key)("������") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 20).Value = dictSums(key)("�^����") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 21).Value = dictSums(key)("�}�N��") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 22).Value = dictSums(key)("��o�s��") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 23).Value = dictSums(key)("���ꔒ��") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 24).Value = dictSums(key)("�R�A�J�X") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 25).Value = dictSums(key)("���̑�") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 26).Value = dictSums(key)("�`���R��ŏo��") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 27).Value = dictSums(key)("����") / dictSums(key)("�V���b�g��")
                destSheet.Cells(destRow, 28).Value = dictSums(key)("���o�s��") / dictSums(key)("�V���b�g��")
            Else
                ' �V���b�g����0�̏ꍇ�͂��ׂ�0
                For j = 8 To 28
                    destSheet.Cells(destRow, j).Value = 0
                Next j
            End If
            
            destRow = destRow + 1
        Next i
        
        dataEndRow = destRow - 1
    End If
    
    ' �e�[�u���̍쐬
    Set tableRange = destSheet.Range("A5").Resize(dataEndRow - 4, 28)
    
    ' ���łɓ����̃e�[�u�������݂���ꍇ�͍폜
    On Error Resume Next
    If Not destSheet.ListObjects("_�i�ԕ�bb") Is Nothing Then
        destSheet.ListObjects("_�i�ԕ�bb").Delete
    End If
    On Error GoTo 0
    
    Set destTable = destSheet.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    destTable.Name = "_�i�ԕ�bb"
    destTable.ShowAutoFilter = False  ' �t�B���^�[�{�^�����\��
    
    ' �X�e�[�^�X�o�[���X�V
    Application.StatusBar = "�ʏ̕ʒ��ڏW�vb: �����ݒ蒆..."
    
    ' �e�[�u�����̏����ݒ�
    With destSheet.Range("A" & dataStartRow & ":AB" & dataEndRow)
        .ShrinkToFit = True  ' �k�����đS�̂�\��
    End With
    
    ' �s�Ǘ���̃t�H�[�}�b�g�ݒ�i%�\���A�����_�ȉ�2���j
    destSheet.Range("G" & dataStartRow & ":G" & dataEndRow).NumberFormat = "0.00%"
    
    ' �T�C�N����̃t�H�[�}�b�g�ݒ�i�����_�ȉ�1���j
    destSheet.Range("D" & dataStartRow & ":D" & dataEndRow).NumberFormat = "0.0"
    
    ' �s�Ǎ��ڗ�̃t�H�[�}�b�g�ݒ�i%�\���A�����_�ȉ�2���j
    destSheet.Range("H" & dataStartRow & ":AB" & dataEndRow).NumberFormat = "0.00%"
    
    ' �񕝂̐ݒ�
    destSheet.Columns("A").ColumnWidth = 14  ' �ʏ̗�
    destSheet.Columns("B:G").ColumnWidth = 7  ' �^�ց`�s�Ǘ���
    destSheet.Columns("H:AB").ColumnWidth = 3  ' �s�Ǎ��ڗ�i�w��ʂ�5�ɐݒ�j
    
    ' 0�̒l�𔖂��O���[�ɂ�������t������
    With destSheet.Range("B" & dataStartRow & ":AB" & dataEndRow)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
        .FormatConditions(1).Font.Color = RGB(192, 192, 192)  ' �����O���[
    End With
    
    ' ��������
    Application.StatusBar = "�ʏ̕ʒ��ڏW�vb: �������������܂����B"
    
    ' 1�b�ҋ@���ăX�e�[�^�X�o�[�N���A
    Application.Wait Now + TimeValue("00:00:01")
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' �G���[����
    Application.StatusBar = False
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical
End Sub
