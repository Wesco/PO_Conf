Attribute VB_Name = "Exports"
Option Explicit

Sub ExportPOList(Branch As String)
    Dim FileName As String
    Dim Path As String
    Dim PrevDispAlert As Boolean
    
    Path = "\\br3615gaps\gaps\PO Conf\"
    FileName = Branch & "-POList"
    PrevDispAlert = Application.DisplayAlerts
    
    Sheets("PO List").Copy
    ActiveSheet.Name = FileName
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Path & FileName, xlCSV
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert
End Sub
