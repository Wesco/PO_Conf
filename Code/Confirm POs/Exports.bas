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
    Dim vRng As Variant
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim Branch As String
    Dim Body As String
    Dim Subject As String
    Dim PO As String
    Dim PODATE As String
    Dim UserEmail As String
    Dim SupEmail As String

    Sheets("PO Conf").Select
    
    UserEmail = Environ("username") & "@wesco.com"
    Branch = Sheets("473").Range("A2")
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    For Each vRng In Range(Cells(2, 5), Cells(TotalRows, 5))
        If vRng.Value <> "" Then
            PO = Right(String(6, "0") & vRng.Offset(0, -4).Value, "6")
            PODATE = Format(vRng.Offset(0, -3).Value, "mmm dd, yyyy")
            SupEmail = vRng.Value

            Subject = "Please Confirm PO# 3615-" & PO
            Body = "Dear Supplier,<br>" & _
                   "PO# " & Branch & "-" & PO & " was sent on " & PODATE & _
                   ". Please confirm that the order has been received and provide an estimated shipping date.<br>" & _
                   "<br>" & _
                   "<br>" & _
                   "Thanks in advance for your help!<br>" & _
                   "<br>" & _
                   "<span style='font-size:8.0pt;font-family:""Arial"",""sans-serif""'>" & _
                   UserEmail & " | office: 704-393-6629 | fax: 704-393-6645<br>" & _
                   "<b>WESCO Distribution<br>" & _
                   "5521 Lakeview Road, Suite W, Charlotte, NC 28269</b>" & _
                   "</span>"
            Email SendTo:=SupEmail, _
                  Subject:=Subject, _
                  Body:=Body
        End If
    Next

End Sub
