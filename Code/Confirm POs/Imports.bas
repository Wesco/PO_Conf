Attribute VB_Name = "Imports"
Option Explicit

Sub ImportPOList(Branch)
    Dim Path As String
    Dim PrevDispAlert As Boolean
    
    Path = "\\br3615gaps\gaps\PO Conf\" & Branch & "-POList.csv"
    PrevDispAlert = Application.DisplayAlerts
    
    If Trim(Branch) = "" Then
        Err.Raise 18
    End If
    
    If FileExists(Path) Then
        Workbooks.Open Path
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("PO List").Range("A1")
        
        Application.DisplayAlerts = False
        ActiveWorkbook.Close
        Application.DisplayAlerts = PrevDispAlert
    Else
        Err.Raise 53
    End If
End Sub
