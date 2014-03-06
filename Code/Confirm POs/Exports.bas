Attribute VB_Name = "Exports"
Option Explicit

Private Const EmailHeader As String = "<html>" & _
        "<style>" & _
        "table{border:1px solid black; border-collapse:collapse}" & _
        "table,th,td{border:1px solid black}" & _
        "td{padding:5px; text-align:left}" & _
        "th{padding:5px; text-align:center}" & _
        "</style>" & _
        "Dear Supplier," & _
        "<br>" & _
        "<br>" & _
        "Please review the list of orders below and confirm that they have been received and provide an estimated ship date. " & _
        "<br>" & _
        "If you are receiving this for a second time, we may not have received an estimated shipping date in your original response or the promise date has passed." & _
        "<br>" & _
        "<br>" & _
        "<br>" & _
        "<table>" & _
        "<th>PO</th><th>CREATED</th><th>PROMISED</th><th>SIM</th><th>DESCRIPTION</th><th>SUPPLIER</th>"

Private Function EmailFooter() As String
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

Sub SendMail()
    'Email variables
    Dim Contact As String
    Dim Subject As String
    Dim Body As String
    Dim Branch As String
    Dim PONumber As String
    Dim Created As String
    Dim SuppName As String
    Dim SimNum As String
    Dim Promised As String
    Dim Desc As String

    'Loop conditions
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
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    For i = 2 To TotalRows
        'Column H contains the supplier number
        PrevCell = Range("H" & i - 1).Value
        CurrCell = Cells("H" & i).Value
        NextCell = Cells("H" & i + 1).Value

        'If there is only line for the current supplier
        If CurrCell <> PrevCell And CurrCell <> NextCell Then
            Branch = Range("A" & i).Value
            PONumber = Range("B" & i).Value
            Created = Format(Range("C" & i).Value, "mmm dd, yyyy")
            Promised = Format(Range("D" & i).Value, "mmm dd, yyyy")
            SimNum = Range("E" & i).Value
            Desc = Range("F" & i).Value
            SuppName = Range("G" & i).Value
            Contact = Range("I" & i).Value
            Subject = "Please send an estimated ship date for PO# " & Branch & "-" & PONumber
            Body = "<tr>" & _
                   "<td>" & Branch & "-" & PONumber & "</td>" & _
                   "<td>" & Created & "</td>" & _
                   "<td>" & Promised & "</td>" & _
                   "<td>" & SimNum & "</td>" & _
                   "<td>" & Desc & "</td>" & _
                   "<td>" & SuppName & "</td>" & _
                   "</tr>"

            'Send email if a contact was found
            If Contact <> "" Then
                If Promised <> "" Then
                    If CDate(Promised) <= Date Then
                        Email Contact, Subject:=Subject, Body:=EmailHeader & Body & EmailFooter
                    End If
                Else
                    Email Contact, Subject:=Subject, Body:=EmailHeader & Body & EmailFooter
                End If
            End If

            'Clear email body
            Body = ""
        ElseIf CurrCell = NextCell And CurrCell <> PrevCell Then
            'If this supplier has multiple lines
            'store the starting row number
            StartRow = i
        ElseIf CurrCell <> NextCell And CurrCell = PrevCell Then
            'If the current suplpier has multiple rows
            'store the last row number
            EndRow = i

            'Loop through each line for the current supplier
            For j = StartRow To EndRow
                Branch = Range("A" & j).Value
                PONumber = Range("B" & j).Value
                Created = Format(Range("C" & j).Value, "mmm dd, yyyy")
                Promised = Format(Range("D" & j).Value, "mmm dd, yyyy")
                SimNum = Range("E" & j).Value
                Desc = Range("F" & j).Value
                SuppName = Range("G" & j).Value
                Contact = Range("I" & j).Value
                Subject = "Please send estimated ship dates"
                Body = "<tr>" & _
                       "<td>" & Branch & "-" & PONumber & "</td>" & _
                       "<td>" & Created & "</td>" & _
                       "<td>" & Promised & "</td>" & _
                       "<td>" & SimNum & "</td>" & _
                       "<td>" & Desc & "</td>" & _
                       "<td>" & SuppName & "</td>" & _
                       "</tr>"
            Next
            
            'Send email if a contact was found
            If Contact <> "" Then
                If Promised <> "" Then
                    If CDate(Promised) <= Date Then
                        Email Contact, Subject:=Subject, Body:=EmailHeader & Body & EmailFooter
                    End If
                Else
                    Email Contact, Subject:=Subject, Body:=EmailHeader & Body & EmailFooter
                End If
            End If

            'Clear email body
            Body = ""
        End If
    Next
    
    MsgBox "Complete!"
End Sub





















