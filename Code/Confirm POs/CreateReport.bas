Attribute VB_Name = "CreateReport"
Option Explicit

Sub CreatePOList()
    Dim TotalRows As Long

    Sheets("473").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

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
    Dim TotalRows As Long

    Sheets("PO Conf").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    
    'Insert a column for the branch
    Columns(1).Insert
    
    'Add Column Headers
    [A1].Value = "Branch"
    [B1].Value = "PO #"
    [C1].Value = "Created"
    [D1].Value = "Promised"
    [E1].Value = "SIM"
    [F1].Value = "Description"
    [G1].Value = "Supplier Name"
    [H1].Value = "Supplier Number"
    [I1].Value = "Email"

    'Format column headers
    With Range("A1:I1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    'Branch
    Range("A2:A" & TotalRows).Value = Sheets("473").Range("A2").Value

    'Created
    Range("C2:C" & TotalRows).Formula = "=IFERROR(TRIM(VLOOKUP(B2,'473'!C:L,10,FALSE)),"""")"
    Range("C2:C" & TotalRows).Value = Range("C2:C" & TotalRows).Value
    Range("C2:C" & TotalRows).NumberFormat = "mmm dd, yyyy"
    
    'Promise Date
    Range("D2:D" & TotalRows).Formula = "=IFERROR(TRIM(VLOOKUP(B2,'473'!C:AC,27,FALSE)),"""")"
    Range("D2:D" & TotalRows).Value = Range("D2:D" & TotalRows).Value
    Range("D2:D" & TotalRows).NumberFormat = "mmm dd, yyyy"
    
    'SIM
    Range("E2:E" & TotalRows).Formula = "=IFERROR(TRIM(SUBSTITUTE(VLOOKUP(B2,'473'!C:Y,23,FALSE),""-"","""")),"""")"
    Range("E2:E" & TotalRows).NumberFormat = "@"
    Range("E2:E" & TotalRows).Value = Range("E2:E" & TotalRows).Value

    'Description
    Range("F2:F" & TotalRows).Formula = "=IFERROR(TRIM(SUBSTITUTE(VLOOKUP(B2,'473'!C:Z,24,FALSE),""***"",""*"")),"""")"
    Range("F2:F" & TotalRows).Value = Range("F2:F" & TotalRows).Value

    'Supplier Name
    Range("G2:G" & TotalRows).Formula = "=IFERROR(TRIM(VLOOKUP(B2,473!C:AO,39,FALSE)),"""")"
    Range("G2:G" & TotalRows).Value = Range("G2:G" & TotalRows).Value
    
    'Supplier Number
    Range("H2:H" & TotalRows).Formula = "=IFERROR(TRIM(VLOOKUP(B2,473!C:I,7,FALSE)),"""")"
    Range("H2:H" & TotalRows).NumberFormat = "@"
    Range("H2:H" & TotalRows).Value = Range("H2:H" & TotalRows).Value
    
    'Email
    Range("I2:I" & TotalRows).Formula = "=IFERROR(TRIM(VLOOKUP(H2,Contacts!A:B,2,FALSE)),"""")"
    Range("I2:I" & TotalRows).Value = Range("I2:I" & TotalRows).Value
    
    ActiveSheet.UsedRange.Columns.AutoFit
End Sub
