Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    Dim Branch As String

    Branch = InputBox(Prompt:="Branch:", Title:="Enter your branch number")

    On Error GoTo Branch_Import_Err
    ImportPOList Branch
    On Error GoTo Fatal_Err
    Import473 Destination:=ThisWorkbook.Sheets("473").Range("A1")
    On Error GoTo 0
    
    Format473
    FilterPOList
    ExportPOList Branch
    ImportSupplierContacts ThisWorkbook.Sheets("Contacts").Range("A1")
    
    On Error GoTo Create_PO_Err
    CreatePOConf
    On Error GoTo 0
    
    Exit Sub

Branch_Import_Err:
    If ActiveWorkbook.Name <> ThisWorkbook.Name Then
        ActiveWorkbook.Close
    End If

    Select Case Err.Number
        Case 53:
            MsgBox Prompt:=Err.Description, Title:="Error"
        Case 18:
            MsgBox Prompt:="A branch number was not entered.", Title:="Error"
        Case Else:
            MsgBox "Error " & Err.Number & vbCrLf & Err.Description
    End Select
    Clean
    Exit Sub

Create_PO_Err:
    MsgBox Err.Description, vbOKOnly, "Error: " & Err.Source
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
