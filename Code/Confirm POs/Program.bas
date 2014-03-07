Attribute VB_Name = "Program"
Option Explicit
Public Const VersionNumber As String = "1.2.0"
Public Const RepositoryName As String = "PO_Conf"

Sub Main()
    Dim Branch As String

    Application.ScreenUpdating = False

    On Error GoTo Main_Err
    'Prompt user for branch number
    Branch = InputBox(Prompt:="Branch:", Title:="Enter your branch number")

    'Get 473 report for the branch
    Import473 ThisWorkbook.Sheets("473").Range("A1"), Branch
    Format473

    'Get the supplier contact master
    ImportSupplierContacts ThisWorkbook.Sheets("Contacts").Range("A1")

    'Create the open po report
    CreatePOList
    CreatePOConf
    SortPOConf
    On Error GoTo 0

    Application.ScreenUpdating = True
    MsgBox "Complete!"
    Exit Sub

Main_Err:
    Clean
    MsgBox Err.Description, vbOKOnly, "Error: " & Err.Source
    Exit Sub
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
