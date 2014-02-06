Attribute VB_Name = "Module1"
Option Explicit

Sub Test()
    Dim Body As String
    Dim PrevCell As String
    Dim CurrCell As String
    Dim NextCell As String
    Dim StartRow As Long
    Dim EndRow As Long
    Dim TotalRows As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long

    Sheets("PO Conf").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Body = SetBody

    For i = 2 To TotalRows
        PrevCell = Cells(i - 1, 3).Value
        CurrCell = Cells(i, 3).Value
        NextCell = Cells(i + 1, 3).Value

        If CurrCell <> PrevCell And CurrCell <> NextCell Then
            'Only one PO for this supplier
            Body = Body & "<tr>" & _
                   "<td>" & Cells(i, 1).Value & "</td>" & _
                   "<td>" & Format(Cells(i, 2).Value, "mmm dd, yyyy") & "</td>" & _
                   "<td>" & Cells(i, 4).Value & "</td>" & _
                   "</tr></table></body></html>"

            Email "treische@wesco.com", Subject:="TEST", Body:=Body
            Body = SetBody
        ElseIf CurrCell = NextCell And CurrCell <> PrevCell Then
            'First cell for this supplier
            StartRow = i
        ElseIf CurrCell <> NextCell And CurrCell = PrevCell Then
            'Last cell for this supplier
            EndRow = i
            'Debug.Print CurrCell & "," & Range(Cells(StartRow, 3), Cells(EndRow, 3)).Address(False, False)

            For j = StartRow To EndRow
                Body = Body & "<tr>" & _
                       "<td>" & Cells(j, 1).Value & "</td>" & _
                       "<td>" & Format(Cells(j, 2).Value, "mmm dd, yyyy") & "</td>" & _
                       "<td>" & Cells(j, 4).Value & "</td>" & _
                       "</tr>"
            Next

            Body = Body & "</table></body></html>"
            Email "treische@wesco.com", Subject:="TEST", Body:=Body
            Body = SetBody
        End If
    Next
End Sub

Private Function SetBody()
    ResetBody = "<html>" & _
                "<style>" & _
                "table{border:1px solid black; border-collapse:collapse;}" & _
                "table,th,td{border:1px solid black;}" & _
                "td{padding:5px; text-align:center;}" & _
                "th{padding:5px;}" & _
                "</style>" & _
                "<table>" & _
                "<th>PO</th><th>CREATED</th><th>SUPPLIER</th>"
End Function
