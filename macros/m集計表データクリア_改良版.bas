Attribute VB_Name = "m集計表データクリア_改良版"

' 集計表データクリアマクロ（改良版）
Sub 集計表データクリア()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim i As Long
    
    Set wb = ThisWorkbook
    
    ' 削除対象シート名
    sheetNames = Array("日別集計_モールFR別", "集計表_TG作業者別", "集計表_TG品番別", _
                      "集計表_モールFR別", "集計表_加工作業者別", "集計表_加工品番別", _
                      "集計表_塗装品番別", "集計表_流出廃棄")
    
    For i = 0 To UBound(sheetNames)
        On Error Resume Next
        Set ws = wb.Sheets(sheetNames(i))
        If Not ws Is Nothing Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Set ws = Nothing
        End If
        On Error GoTo 0
    Next i
    
    MsgBox "集計表データをクリアしました。", vbInformation
End Sub