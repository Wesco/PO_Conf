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
    Range(Cells(1, 2), Cells(TotalRows, 2)).Formula = "=IFERROR(TRIM(VLOOKUP(A1,'473'!C:Z,24,FALSE)),""History"")"
    '[B1].AutoFill Destination:=Range(Cells(1, 2), Cells(TotalRows, 2))
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
    [C1].Value = "Supplier #"
    [D1].Value = "Supplier Name"
    [E1].Value = "Contact"

    'Verify that PO numbers can be found on the 473 report
    If Trim(Sheets("473").Range("C1").Value) <> "PO NUMBER" Then
        Err.Raise 50000, "CreatePOConf", "Sheets(""473"").Range(""C1"").Value != ""PO NUMBER""."
    End If

    'Created
    If Trim(Sheets("473").Range("J1").Value) = "PO DATE" Then
        Range(Cells(2, 2), Cells(TotalRows, 2)).Formula = "=IFERROR(TRIM(VLOOKUP(A2,'473'!C:J,8,FALSE)),"""")"
        Range(Cells(2, 2), Cells(TotalRows, 2)).Value = Range(Cells(2, 2), Cells(TotalRows, 2)).Value
        Range(Cells(2, 2), Cells(TotalRows, 2)).NumberFormat = "mmm-dd"
    Else
        Err.Raise 50000, "CreatePOConf", "Sheets(""473"").Range(""J1"").Value != ""PO DATE""."
    End If

    'Supplier #
    If Trim(Sheets("473").Range("G1").Value) = "SUPPLIER" Then
        Range(Cells(2, 3), Cells(TotalRows, 3)).Formula = "=IFERROR(TRIM(VLOOKUP(A2,'473'!C:G,5,FALSE)),"""")"
        Range(Cells(2, 3), Cells(TotalRows, 3)).NumberFormat = "@"
        Range(Cells(2, 3), Cells(TotalRows, 3)).Value = Range(Cells(2, 3), Cells(TotalRows, 3)).Value
    Else
        Err.Raise 50000, "CreatePOConf", "Sheets(""473"").Range(""G1"").Value != ""SUPPLIER""."
    End If

    'Supplier Name
    If Trim(Sheets("473").Range("AJ1").Value) = "SUPPLIER NAME" Then
        Range(Cells(2, 4), Cells(TotalRows, 4)).Formula = "=IFERROR(TRIM(VLOOKUP(A2,'473'!C:AJ,34,FALSE)),"""")"
        Range(Cells(2, 4), Cells(TotalRows, 4)).Value = Range(Cells(2, 4), Cells(TotalRows, 4)).Value
    Else
        Err.Raise 50000, "CreatePOConf", "Sheets(""473"").Range(""AJ1"").Value != ""SUPPLIER NAME""."
    End If

    'Contact
    Range(Cells(2, 5), Cells(TotalRows, 5)).Formula = "=IFERROR(VLOOKUP(C2,Contacts!A:B,2,FALSE),"""")"
    Range(Cells(2, 5), Cells(TotalRows, 5)).Value = Range(Cells(2, 5), Cells(TotalRows, 5)).Value

    PrevSheet.Select
End Sub

Sub Format473()
    Dim PrevSheet As Worksheet
    Set PrevSheet = ActiveSheet

    Sheets("473").Select
    Rows(1).Delete

    PrevSheet.Select
End Sub
