Attribute VB_Name = "CreateReport"
Option Explicit

Sub CreatePOList()
    Dim TotalRows As Long

    Sheets("473").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Check to make sure the columns are in the correct places
    If [C1].Value <> "PO NUMBER" Then Err.Raise 50000, "CreatePOList", "473!C1 != ""PO NUMBER"""
    If [X1].Value <> "T" Then Err.Raise 50000, "CreatePOList", "473!X1 != ""T"""

    'Filter for all non-stock POs and copy them to another sheet
    ActiveSheet.UsedRange.AutoFilter Field:=24, Criteria1:="=X"
    Range("C1:C" & TotalRows).Copy Destination:=Sheets("PO Conf").Range("A1")
    ActiveSheet.AutoFilterMode = False

    'Remove duplicate PO numbers
    Sheets("PO Conf").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Range("A:A").RemoveDuplicates 1, xlYes
End Sub

Sub CreatePOConf()
    Dim PrevSheet As Worksheet
    Dim TotalRows As Long

    Set PrevSheet = ActiveSheet

    Sheets("PO Conf").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Add Column Headers
    [A1].Value = "PO #"
    [B1].Value = "Created"
    [C1].Value = "Supplier #"
    [D1].Value = "Supplier Name"
    [E1].Value = "Contact"

    'Format column headers
    With Range("A1:E1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    'Verify that PO numbers can be found on the 473 report
    If Trim(Sheets("473").Range("C1").Value) <> "PO NUMBER" Then
        Err.Raise 50000, "CreatePOConf", "Sheets(""473"").Range(""C1"").Value != ""PO NUMBER""."
    End If

    'Created
    If Trim(Sheets("473").Range("L1").Value) = "PO DATE" Then
        Range(Cells(2, 2), Cells(TotalRows, 2)).Formula = "=IFERROR(TRIM(VLOOKUP(A2,'473'!C:L,10,FALSE)),"""")"
        Range(Cells(2, 2), Cells(TotalRows, 2)).Value = Range(Cells(2, 2), Cells(TotalRows, 2)).Value
        Range(Cells(2, 2), Cells(TotalRows, 2)).NumberFormat = "mmm-dd"
    Else
        Err.Raise 50000, "CreatePOConf", "Sheets(""473"").Range(""L1"").Value != ""PO DATE""."
    End If

    'Supplier #
    If Trim(Sheets("473").Range("I1").Value) = "SUPPLIER" Then
        Range(Cells(2, 3), Cells(TotalRows, 3)).Formula = "=IFERROR(TRIM(VLOOKUP(A2,'473'!C:I,7,FALSE)),"""")"
        Range(Cells(2, 3), Cells(TotalRows, 3)).NumberFormat = "@"
        Range(Cells(2, 3), Cells(TotalRows, 3)).Value = Range(Cells(2, 3), Cells(TotalRows, 3)).Value
    Else
        Err.Raise 50000, "CreatePOConf", "Sheets(""473"").Range(""I1"").Value != ""SUPPLIER""."
    End If

    'Supplier Name
    If Trim(Sheets("473").Range("AO1").Value) = "SUPPLIER NAME" Then
        Range(Cells(2, 4), Cells(TotalRows, 4)).Formula = "=IFERROR(TRIM(VLOOKUP(A2,'473'!C:AO,39,FALSE)),"""")"
        Range(Cells(2, 4), Cells(TotalRows, 4)).Value = Range(Cells(2, 4), Cells(TotalRows, 4)).Value
    Else
        Err.Raise 50000, "CreatePOConf", "Sheets(""473"").Range(""AO1"").Value != ""SUPPLIER NAME""."
    End If

    'Contact
    Range(Cells(2, 5), Cells(TotalRows, 5)).Formula = "=IFERROR(VLOOKUP(C2,Contacts!A:B,2,FALSE),"""")"
    Range(Cells(2, 5), Cells(TotalRows, 5)).Value = Range(Cells(2, 5), Cells(TotalRows, 5)).Value

    ActiveSheet.UsedRange.Columns.AutoFit

    PrevSheet.Select
End Sub
