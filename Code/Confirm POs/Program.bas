Attribute VB_Name = "Program"
Option Explicit
Public Const VersionNumber As String = "1.0.2"
Public Const RepositoryName As String = "PO_Conf"

Sub Main()
    Dim Branch As String

    Application.ScreenUpdating = False
    Branch = InputBox(Prompt:="Branch:", Title:="Enter your branch number")
    Clean
    On Error GoTo Branch_Import_Err
    ImportPOList Branch
    On Error GoTo Fatal_Err
    Import473 ThisWorkbook.Sheets("473").Range("A1"), Branch
    On Error GoTo 0

    Format473
    FilterPOList
    ExportPOList Branch
    ImportSupplierContacts ThisWorkbook.Sheets("Contacts").Range("A1")

    On Error GoTo Create_PO_Err
    CreatePOConf
    On Error GoTo 0

    SortPOConf
    Sheets("PO Conf").Select

    Application.ScreenUpdating = True
    
    MsgBox "Complete!"
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
    Dim PrevDispAlert As Boolean
    Dim PrevScrnUpdat As Boolean
    Dim s As Variant

    PrevDispAlert = Application.DisplayAlerts
    PrevScrnUpdat = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            s.Cells.Delete
            s.Range("A1").Select
        End If
    Next

    Sheets("Macro").Select
    Range("C7").Select

    Application.DisplayAlerts = PrevDispAlert
    Application.ScreenUpdating = PrevScrnUpdat
End Sub
