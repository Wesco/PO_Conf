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

Sub SendMail()
    'Email variables
    Dim Body As String
    Dim Subject As String
    Dim SuppName As String
    Dim Contact As String
    Dim PONumber As String
    Dim Created As String
    Dim Branch As String

    'Loop conditionals
    Dim PrevCell As String
    Dim CurrCell As String
    Dim NextCell As String
    Dim StartRow As Long
    Dim EndRow As Long
    Dim TotalRows As Long

    'Loop counters
    Dim i As Long
    Dim j As Long

    Sheets("PO Conf").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Branch = Sheets("473").Range("A2")

    For i = 2 To TotalRows
        PrevCell = Cells(i - 1, 3).Value
        CurrCell = Cells(i, 3).Value
        NextCell = Cells(i + 1, 3).Value

        If CurrCell <> PrevCell And CurrCell <> NextCell Then
            'Only one PO for this supplier
            PONumber = Cells(i, 1).Value
            Created = Format(Cells(i, 2).Value, "mmm dd, yyyy")
            SuppName = Cells(i, 4).Value

            Contact = Cells(i, 5).Value
            Subject = "Please send an estimated ship date for PO# " & Branch & "-" & PONumber
            Body = "<tr>" & _
                   "<td>" & Branch & "-" & PONumber & "</td>" & _
                   "<td>" & Created & "</td>" & _
                   "<td>" & SuppName & "</td>" & _
                   "</tr>"

            If Contact <> "" Then
                Email Contact, Subject:=Subject, Body:=EmailHeader & Body & EmailFooter
            End If

            'Reset email body
            Body = ""
        ElseIf CurrCell = NextCell And CurrCell <> PrevCell Then
            'First cell for this supplier
            StartRow = i
        ElseIf CurrCell <> NextCell And CurrCell = PrevCell Then
            'Last cell for this supplier
            EndRow = i

            'Add all rows to the email in a table
            For j = StartRow To EndRow
                PONumber = Cells(j, 1).Value
                Created = Format(Cells(j, 2).Value, "mmm dd, yyyy")
                SuppName = Cells(j, 4).Value

                Body = Body & "<tr>" & _
                       "<td>" & Branch & "-" & PONumber & "</td>" & _
                       "<td>" & Created & "</td>" & _
                       "<td>" & SuppName & "</td>" & _
                       "</tr>"
            Next
            Subject = "Please send estimated ship dates"
            Contact = Cells(i, 5).Value

            If Contact <> "" Then
                Email Contact, Subject:=Subject, Body:=EmailHeader & Body & EmailFooter
            End If

            'Reset email body
            Body = ""
        End If
    Next
End Sub

Private Function EmailHeader()
    EmailHeader = "<html>" & _
                  "<style>" & _
                  "table{border:1px solid black; border-collapse:collapse;}" & _
                  "table,th,td{border:1px solid black;}" & _
                  "td{padding:5px; text-align:center;}" & _
                  "th{padding:5px;}" & _
                  "</style>" & _
                  "Dear Supplier," & _
                  "<br>" & _
                  "<br>" & _
                  "Please review the list of orders below and confirm that they have been received and provide an estimated ship date. " & _
                  "<br>" & _
                  "If you are receiving this for a second time, we may not have received an estimated shipping date in your original response." & _
                  "<br>" & _
                  "<br>" & _
                  "<br>" & _
                  "<table>" & _
                  "<th>PO</th><th>CREATED</th><th>SUPPLIER</th>"
End Function

Private Function EmailFooter()
    EmailFooter = "</table>" & _
                  "<br>" & _
                  "<br>" & _
                  "Thanks in advance for your help!<br>" & _
                  "<br>" & _
                  "<span style='font-size:8.0pt;font-family:""Arial"",""sans-serif""'>" & _
                  Environ("username") & "@wesco.com" & " | office: 704-393-6636 | fax: 704-393-6645<br>" & _
                  "<b>WESCO Distribution<br>" & _
                  "5521 Lakeview Road, Suite W, Charlotte, NC 28269</b>" & _
                  "</span>" & _
                  "</body></html>"
End Function
