Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    Dim Branch As String

    Branch = InputBox(Prompt:="Branch:", Title:="Enter your branch number")

    On Error GoTo Import_Err
    ImportPOList Branch
    Import473 Destination:=ThisWorkbook.Sheets("473").Range("A1")
    On Error GoTo 0

    Format473
    FilterPOList
    ExportPOList Branch

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

Fatal_Err:
End Sub

Sub Clean()
    Dim s As Variant

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Cells.Delete
        End If
    Next
End Sub
