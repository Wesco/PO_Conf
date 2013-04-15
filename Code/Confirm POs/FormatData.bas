Attribute VB_Name = "FormatData"
Option Explicit

Sub FilterPOList()
    Dim TotalRows As Long
    Dim PrevSheet As Worksheet
    
    Set PrevSheet = ActiveSheet
    Sheets("PO List").Select
    Range("A:A").RemoveDuplicates 1, xlNo
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    
    'Add promise dates
    Range(Cells(1, 2), Cells(TotalRows, 2)).Formula = "=IFERROR(TRIM(VLOOKUP(A1,'473'!C:Z,24,FALSE)),"""")"
    [B1].AutoFill Destination:=Range(Cells(1, 2), Cells(TotalRows, 2))
    Range(Cells(1, 2), Cells(TotalRows, 2)).Value = Range(Cells(1, 2), Cells(TotalRows, 2)).Value
    Range(Cells(1, 2), Cells(TotalRows, 2)).NumberFormat = "mmm-dd"
    
    'Add column headers
    Rows(1).Insert
    [A1].Value = "PO Number"
    [B1].Value = "Promise Date"
    
    'Show only POs without promise dates
    Range("A:B").AutoFilter Field:=2, Criteria1:="="
    
    'Copy POs without promise dates
    Range("A:A").Copy Destination:=Sheets("PO Conf").Range("A1")
    
    'Remove all POs with data
    ActiveSheet.ShowAllData
    Range("A:B").AutoFilter Field:=2, Criteria1:="<>"
    ActiveSheet.Cells.Delete
    
    PrevSheet.Select
End Sub

Sub CreatePOConf()
    Dim PrevSheet As Worksheet
    Dim TotalRows As Long
    
    Set PrevSheet = ActiveSheet
    
    Sheets("PO Conf").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    
    'Add Column Headers
    [B1].Value = "Created"
    [C1].Value = "Status"
    [D1].Value = "Supplier #"
    [E1].Value = "Supplier Name"
    [F1].Value = "Contact"
    
    'Created
    [B2].Formula = "=IFERROR(TRIM(VLOOKUP(A2,'473'!C:J,8,FALSE)),"""")"
    [B2].AutoFill Destination:=Range(Cells(2, 2), Cells(TotalRows, 2))
    Range(Cells(2, 2), Cells(TotalRows, 2)).Value = Range(Cells(2, 2), Cells(TotalRows, 2)).Value
    Range(Cells(2, 2), Cells(TotalRows, 2)).NumberFormat = "mmm-dd"
    
    PrevSheet.Select
End Sub

Sub Format473()
    Dim PrevSheet As Worksheet
    Set PrevSheet = ActiveSheet
    
    Sheets("473").Select
    Rows(1).Delete
    
    PrevSheet.Select
End Sub
