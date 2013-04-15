Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    On Error GoTo Import_Err
    ImportPOList
    On Error GoTo 0

    Exit Sub

Import_Err:
    If ActiveWorkbook.Name <> ThisWorkbook.Name Then
        ActiveWorkbook.Close
    End If
    Select Case Err.Number
        Case 53:
            MsgBox Prompt:=Err.Description, Title:="Error"
        Case 18:
            MsgBox Prompt:="A branch number was not entered. Macro aborted.", Title:="Error"
        Case Else:
            MsgBox "Error " & Err.Number & vbCrLf & Err.Description
    End Select

    Exit Sub

End Sub
