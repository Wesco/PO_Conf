Attribute VB_Name = "GetPO"
Option Explicit

Public Sub GetPONumber(itm As Outlook.MailItem)
    Dim sPOList As String
    Dim PO As String
    Dim Branch As String
    Dim FileNumber As Integer
    Dim subject As String

    FileNumber = FreeFile
    subject = itm.subject
    If InStr(subject, "FW: ") Then subject = Replace(subject, "FW: ", "")
    Branch = Left(Replace(subject, "WESCO International, Inc. PO #", ""), 4)
    sPOList = "\\br3615gaps\gaps\PO Conf\" & Branch & "-POList.csv"
    PO = Replace(subject, "WESCO International, Inc. PO #" & Branch & "-", "")

    'Removes FW: from forwarded emails
    If InStr(PO, "FW: ") Then PO = Replace(PO, "FW: ", "")

    On Error Resume Next
    Open sPOList For Append Shared As #FileNumber

    Select Case Err.Number
        Case 0      'Write the PO number to POList.csv
            Print #FileNumber, PO
            Close FileNumber
            
        Case Else   'Write to an error log
            Open "C:\GetPONumber_ErrorLog.csv" For Append Lock Read Write As #FileNumber
            Print #FileNumber, PO & "," & Date & "," & Time & "," & Environ("username") & "," & Err.Number & "," & Err.Description
            Close FileNumber
    End Select
    
    Err.Clear
    
    'Resume Error Checking
    On Error GoTo 0
End Sub
