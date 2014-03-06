Attribute VB_Name = "FormatData"
Option Explicit

Sub Format473()
    Dim ColHeaders As Variant
    Dim i As Long

    'This includes every column except the last one because it is a large string of spaces
    ColHeaders = Array("BRANCH", "ERROR", "PO NUMBER", "PO TYPE", "PO CREATED USERID", "PO LAST CHANGED USERID", _
                       "SHIPPING INSTRUCTIONS 1", "SHIPPING INSTRUCTIONS 2", " SUPPLIER", "SHIP TO", "DS ORDER", _
                       "PO DATE", "PO STATUS", "REQUESTED", "ACKNOWLEDGE", "TERMS CODE", "TERMS DAYS", "REFERENCE", _
                       "DISC.%", "FOB", "BOL", "SHIPPING TERMS", "LINE", "T", "SIM", "DESCRIPTION", "UOM", "FACTOR", _
                       "PROMISED", "QTY ORD", "QTY REC", "QTY INV", "OPEN QTY", "OPEN AP QTY", "LAST REC", "PRICE", _
                       "EXTENSION", "EST", "ORDER", "LINE", "SUPPLIER NAME", "ADDRESS LINE1", "ADDRESS LINE2", "CITY", _
                       "ST", "ZIP", "SHIP TO NAME", "SHIP ADDR LN1", "SHIP ADDR LN2", "SHIP CITY", "SHIP STATE", "SHIP ZIP", _
                       "NEGNO", " COSTTYPE", "COSTDESC")

    Sheets("473").Select
    Rows(1).Delete

    For i = 1 To ActiveSheet.UsedRange.Columns.Count - 1
        If Cells(1, i).Value <> ColHeaders(i - 1) Then
            Err.Raise 50000, "Import473", "The Open PO Report (473) column order has changed."
        End If
    Next
End Sub

Sub SortPOConf()
    Dim vRng As Variant
    Dim TotalRows As Long
    Dim TotalCols As Integer
    
    Sheets("PO Conf").Select

    TotalRows = Rows(Rows.Count).End(xlUp).Row
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    'Sort by PO creation date, oldest to newest
    With ActiveWorkbook.Worksheets("PO Conf").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("H1"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SetRange Range(Cells(2, 1), Cells(TotalRows, TotalCols))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
