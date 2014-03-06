Attribute VB_Name = "Program"
Option Explicit
Public Const VersionNumber As String = "1.1.0"
Public Const RepositoryName As String = "PO_Conf"

Sub Main()
    Dim Branch As String

    Application.ScreenUpdating = False
    Branch = InputBox(Prompt:="Branch:", Title:="Enter your branch number")
    
    On Error GoTo Main_Err
    Import473 ThisWorkbook.Sheets("473").Range("A1"), Branch
    On Error GoTo 0

    ImportSupplierContacts ThisWorkbook.Sheets("Contacts").Range("A1")

    On Error GoTo Main_Err
    CreatePOConf
    On Error GoTo 0

    SortPOConf
    Sheets("PO Conf").Select
    Clean
    Application.ScreenUpdating = True
    
    MsgBox "Complete!"
    Exit Sub

Main_Err:
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
