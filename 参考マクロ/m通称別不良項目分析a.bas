Attribute VB_Name = "m�ʏ̕ʕs�Ǎ��ڕ���a"
Sub �ʏ̕ʕs�Ǎ��ڕ���a()
    ' �u�i�ԕ�aa�v�V�[�g�́u_�i�ԕ�aa�v�e�[�u������ʏ̕ʂ�
    ' �s�Ǎ��ڃx�X�g3�Ǝc�荀�ڍ��v�̃e�[�u���쐬����}�N��
    
    ' �X�e�[�^�X�o�[�ɏ����󋵂�\��
    Application.StatusBar = "�ʏ̕ʕs�Ǎ��ڕ���a: �������J�n���܂�..."
    
    ' �ϐ��錾
    Dim srcSheet As Worksheet, srcTable As ListObject
    Dim srcData As Variant, headerRow As Range
    Dim dictTsusho As Object, dictData As Object
    Dim i As Long, j As Long, k As Long
    Dim tsusho As Variant, key As Variant
    Dim currentRow As Long, outputRow As Long
    Dim tableRange As Range, newTable As ListObject
    Dim tableName As String
    
    ' �s�Ǎ��ڗ񖼔z��i21���ځj
    Dim furyoItems As Variant
    furyoItems = Array("�ŏo��", "�V���[�g", "�E�G���h", "�V��", "�ٕ�", "�V���o�[", _
                       "�t���[�}�[�N", "�S�~����", "GC�J�X", "�L�Y", "�q�P", "������", _
                       "�^����", "�}�N��", "��o�s��", "���ꔒ��", "�R�A�J�X", "���̑�", _
                       "�`���R��ŏo��", "����", "���o�s��")
    
    ' ��C���f�b�N�X�p�ϐ�
    Dim tsushoCol As Integer
    Dim furyoColIdx As Object  ' Dictionary for �s�Ǎ��ڂ̃C���f�b�N�X
    
    ' �\�[�g�p�ϐ�
    Dim sortData() As Variant
    Dim tempName As Variant, tempValue As Variant
    
    ' �x�X�g3�Ǝc�荀�ڗp�ϐ�
    Dim best3Names(0 To 2) As String
    Dim best3Values(0 To 2) As Double
    Dim remainingNames As String
    Dim remainingValue As Double
    
    ' �G���[�n���h�����O�ݒ�
    On Error GoTo ErrorHandler
    
    ' �V�[�g�ƃe�[�u���̎擾
    On Error Resume Next
    Set srcSheet = ThisWorkbook.Worksheets("�i�ԕ�aa")
    On Error GoTo 0
    
    If srcSheet Is Nothing Then
        Application.StatusBar = "�ʏ̕ʕs�Ǎ��ڕ���a: �u�i�ԕ�aa�v�V�[�g��������܂���B"
        Exit Sub
    End If
    
    On Error Resume Next
    Set srcTable = srcSheet.ListObjects("_�i�ԕ�aa")
    On Error GoTo 0
    
    If srcTable Is Nothing Then
        Application.StatusBar = "�ʏ̕ʕs�Ǎ��ڕ���a: �e�[�u���u_�i�ԕ�aa�v��������܂���B"
        Exit Sub
    End If
    
    ' �X�e�[�^�X�o�[���X�V
    Application.StatusBar = "�ʏ̕ʕs�Ǎ��ڕ���a: �f�[�^�擾��..."
    
    ' �f�[�^�̎擾
    srcData = srcTable.DataBodyRange.Value
    Set headerRow = srcTable.HeaderRowRange
    
    ' Dictionary�I�u�W�F�N�g�̍쐬
    Set furyoColIdx = CreateObject("Scripting.Dictionary")
    
    ' ��C���f�b�N�X�̓���
    For i = 1 To headerRow.Cells.Count
        Dim colName As String
        colName = CStr(headerRow.Cells(1, i).Value)
        
        If colName = "�ʏ�" Then
            tsushoCol = i
        End If
        
        ' �s�Ǎ��ڂ̗�C���f�b�N�X���L�^
        For j = 0 To UBound(furyoItems)
            If colName = furyoItems(j) Then
                furyoColIdx.Add furyoItems(j), i
                Exit For
            End If
        Next j
    Next i
    
    ' �K�v�ȗ񂪌�����Ȃ��ꍇ�͏������~
    If tsushoCol = 0 Then
        Application.StatusBar = "�ʏ̕ʕs�Ǎ��ڕ���a: �u�ʏ́v�񂪌�����܂���B"
        Exit Sub
    End If
    
    ' �X�e�[�^�X�o�[���X�V
    Application.StatusBar = "�ʏ̕ʕs�Ǎ��ڕ���a: �ʏ̕ʃO���[�v����..."
    
    ' �ʏ̕ʂɃf�[�^���O���[�v��
    Set dictTsusho = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(srcData, 1)
        tsusho = srcData(i, tsushoCol)
        
        If Not dictTsusho.Exists(tsusho) Then
            Set dictData = CreateObject("Scripting.Dictionary")
            
            ' �s�Ǎ��ڂ̏�����
            For j = 0 To UBound(furyoItems)
                dictData.Add furyoItems(j), 0
            Next j
            
            dictTsusho.Add tsusho, dictData
        End If
        
        ' �f�[�^�̏W�v
        Set dictData = dictTsusho(tsusho)
        
        ' �s�Ǎ��ڂ̏W�v
        For j = 0 To UBound(furyoItems)
            If furyoColIdx.Exists(furyoItems(j)) Then
                Dim colIdx As Integer
                colIdx = furyoColIdx(furyoItems(j))
                If IsNumeric(srcData(i, colIdx)) Then
                    dictData(furyoItems(j)) = dictData(furyoItems(j)) + CDbl(srcData(i, colIdx))
                End If
            End If
        Next j
    Next i
    
    ' �X�e�[�^�X�o�[���X�V
    Application.StatusBar = "�ʏ̕ʕs�Ǎ��ڕ���a: �e�[�u���쐬��..."
    
    ' �o�͊J�n�ʒu���擾�i�ŏI�s����3�s�󂯂�j
    currentRow = srcTable.Range.Row + srcTable.Range.Rows.Count + 3
    
    ' �e�ʏ̂ɑ΂��ăe�[�u�����쐬
    For Each tsusho In dictTsusho.Keys
        Set dictData = dictTsusho(tsusho)
        
        ' �s�Ǎ��ڂ̃\�[�g�p�z����쐬
        ReDim sortData(0 To UBound(furyoItems), 0 To 1)
        
        For j = 0 To UBound(furyoItems)
            sortData(j, 0) = furyoItems(j)  ' ���ږ�
            sortData(j, 1) = dictData(furyoItems(j))  ' �l
        Next j
        
        ' �l�̑傫�����Ƀ\�[�g�i�o�u���\�[�g�j
        For j = 0 To UBound(furyoItems) - 1
            For k = j + 1 To UBound(furyoItems)
                If CDbl(sortData(j, 1)) < CDbl(sortData(k, 1)) Then
                    ' ���ږ��̌���
                    tempName = sortData(j, 0)
                    sortData(j, 0) = sortData(k, 0)
                    sortData(k, 0) = tempName
                    
                    ' �l�̌���
                    tempValue = sortData(j, 1)
                    sortData(j, 1) = sortData(k, 1)
                    sortData(k, 1) = tempValue
                End If
            Next k
        Next j
        
        ' �x�X�g3���擾
        For j = 0 To 2
            best3Names(j) = CStr(sortData(j, 0))
            best3Values(j) = CDbl(sortData(j, 1))
        Next j
        
        ' �c��18���ڂ̖��O�����ƒl���v�i�[���l�͏��O�j
        remainingNames = ""
        remainingValue = 0
        
        For j = 3 To UBound(furyoItems)
            If CDbl(sortData(j, 1)) <> 0 Then
                If remainingNames <> "" Then
                    remainingNames = remainingNames & "|"
                End If
                remainingNames = remainingNames & CStr(sortData(j, 0))
            End If
            remainingValue = remainingValue + CDbl(sortData(j, 1))
        Next j
        
        ' �c�荀�ڂ��Ȃ��ꍇ�̃f�t�H���g��
        If remainingNames = "" Then
            remainingNames = "���̑�"
        End If
        
        ' �w�b�_�[�s���쐬
        outputRow = currentRow
        srcSheet.Cells(outputRow, 1).Value = "�ʏ�"
        srcSheet.Cells(outputRow, 2).Value = best3Names(0)
        srcSheet.Cells(outputRow, 3).Value = best3Names(1)
        srcSheet.Cells(outputRow, 4).Value = best3Names(2)
        srcSheet.Cells(outputRow, 5).Value = remainingNames
        
        ' �f�[�^�s���쐬
        outputRow = outputRow + 1
        srcSheet.Cells(outputRow, 1).Value = CStr(tsusho)
        srcSheet.Cells(outputRow, 2).Value = best3Values(0)
        srcSheet.Cells(outputRow, 3).Value = best3Values(1)
        srcSheet.Cells(outputRow, 4).Value = best3Values(2)
        srcSheet.Cells(outputRow, 5).Value = remainingValue
        
        ' �e�[�u���̍쐬
        Set tableRange = srcSheet.Range(srcSheet.Cells(currentRow, 1), _
                                      srcSheet.Cells(outputRow, 5))
        
        tableName = "_" & CStr(tsusho) & "aa"
        
        ' �����̓����e�[�u��������ꍇ�͍폜
        On Error Resume Next
        If Not srcSheet.ListObjects(tableName) Is Nothing Then
            srcSheet.ListObjects(tableName).Delete
        End If
        On Error GoTo 0
        
        ' �V�����e�[�u�����쐬
        Set newTable = srcSheet.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        newTable.Name = tableName
        newTable.ShowAutoFilter = False
        
        ' �e�[�u���̏����ݒ�
        With tableRange
            .Font.Name = "Yu Gothic UI"
            .Font.Size = 11
            .ShrinkToFit = True
        End With
        
        ' �w�b�_�[�s�̏����ݒ�
        With srcSheet.Range(srcSheet.Cells(currentRow, 1), _
                           srcSheet.Cells(currentRow, 5))
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .ShrinkToFit = True
        End With
        
        ' �f�[�^�s�̐��l�t�H�[�}�b�g�ݒ�i�����\���j
        With srcSheet.Range(srcSheet.Cells(outputRow, 2), _
                           srcSheet.Cells(outputRow, 5))
            .NumberFormat = "0"
            .ShrinkToFit = True
        End With
        
        ' 0�̒l�𔖂��O���[�ɂ�������t������
        With srcSheet.Range(srcSheet.Cells(outputRow, 2), _
                           srcSheet.Cells(outputRow, 5))
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
            .FormatConditions(1).Font.Color = RGB(192, 192, 192)
        End With
        
        ' �񕝐ݒ�
        ' A��: 14, B, C, D��: 7�ɌŒ�
        srcSheet.Range(srcSheet.Cells(currentRow, 1), srcSheet.Cells(outputRow, 1)).ColumnWidth = 14  ' A��
        srcSheet.Range(srcSheet.Cells(currentRow, 2), srcSheet.Cells(outputRow, 2)).ColumnWidth = 7   ' B��
        srcSheet.Range(srcSheet.Cells(currentRow, 3), srcSheet.Cells(outputRow, 3)).ColumnWidth = 7   ' C��
        srcSheet.Range(srcSheet.Cells(currentRow, 4), srcSheet.Cells(outputRow, 4)).ColumnWidth = 7   ' D��
        srcSheet.Range(srcSheet.Cells(currentRow, 5), srcSheet.Cells(outputRow, 5)).ColumnWidth = 7   ' E��
        
        ' ���̃e�[�u���̈ʒu��ݒ�i2�s�󂯂�j
        currentRow = outputRow + 3
    Next tsusho
    
    ' ��������
    Application.StatusBar = "�ʏ̕ʕs�Ǎ��ڕ���a: �������������܂����B"
    
    ' 1�b�ҋ@���ăX�e�[�^�X�o�[�N���A
    Application.Wait Now + TimeValue("00:00:01")
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' �G���[����
    Application.StatusBar = False
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical
End Sub

