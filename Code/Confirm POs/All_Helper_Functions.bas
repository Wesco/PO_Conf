Attribute VB_Name = "All_Helper_Functions"
Option Explicit

'Used when importing 117 to determine the type of report to pull
Enum ReportType
    DS
    Bo
End Enum

'---------------------------------------------------------------------------------------
' Proc  : Sub ImportGaps
' Date  : 12/12/2012
' Desc  : Imports gaps to the workbook containing this macro.
' Ex    : ImportGaps
'---------------------------------------------------------------------------------------
Sub ImportGaps()
    Dim sPath As String     'Gaps file path
    Dim sName As String     'Gaps Sheet Name
    Dim iCounter As Long    'Counter to decrement the date
    Dim iRows As Long       'Total number of rows
    Dim dt As Date          'Date for gaps file name and path
    Dim Result As VbMsgBoxResult    'Yes/No to proceed with old gaps file if current one isn't found
    Dim Gaps As Worksheet           'The sheet named gaps if it exists, else this = nothing
    Dim StartTime As Double         'The time this function was started
    Dim FileFound As Boolean        'Indicates whether or not gaps was found

    StartTime = Timer
    dt = Date - iCounter
    sPath = "\\br3615gaps\gaps\3615 Gaps Download\" & Format(dt, "yyyy") & "\"
    sName = "3615 " & Format(dt, "yyyy-mm-dd") & ".xlsx"
    FileFound = False

    'This error is bypassed so you can determine whether or not the sheet exists
    On Error GoTo CREATE_GAPS
    Set Gaps = ThisWorkbook.Sheets("Gaps")
    On Error GoTo 0

    Application.DisplayAlerts = False

    'Find gaps
    For iCounter = 0 To 15
        dt = Date - iCounter
        sPath = "\\br3615gaps\gaps\3615 Gaps Download\" & Format(dt, "yyyy") & "\"
        sName = "3615 " & Format(dt, "yyyy-mm-dd") & ".xlsx"
        If FileExists(sPath & sName) Then
            FileFound = True
            Exit For
        End If
    Next

    'Make sure Gaps file was found
    If FileFound = True Then
        If dt <> Date Then
            Result = MsgBox( _
                     Prompt:="Gaps from " & Format(dt, "mmm dd, yyyy") & " was found." & vbCrLf & "Would you like to continue?", _
                     Buttons:=vbYesNo, _
                     Title:="Gaps not up to date")
        End If

        If Result <> vbNo Then
            If ThisWorkbook.Sheets("Gaps").Range("A1").Value <> "" Then
                Gaps.Cells.Delete
            End If

            Workbooks.Open sPath & sName
            ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Gaps").Range("A1")
            ActiveWorkbook.Close

            Sheets("Gaps").Select
            iRows = ActiveSheet.UsedRange.Rows.Count
            Columns(1).EntireColumn.Insert
            Range("A1").Value = "SIM"
            Range("A2").Formula = "=C2&D2"
            Range("A2").AutoFill Destination:=Range(Cells(2, 1), Cells(iRows, 1))
            Range(Cells(2, 1), Cells(iRows, 1)).Value = Range(Cells(2, 1), Cells(iRows, 1)).Value

            FillInfo FunctionName:="Gaps", _
                     FileDate:=Format(dt, "mm/dd/yy"), _
                     Parameters:="", _
                     ExecutionTime:=Timer - StartTime, _
                     Result:="Complete"
        Else
            FillInfo FunctionName:="Gaps", _
                     FileDate:=Format(dt, "mm/dd/yy"), _
                     Parameters:="", _
                     ExecutionTime:=Timer - StartTime, _
                     Result:="Failed - User Aborted"
            Err.Raise 18
        End If
    Else
        MsgBox Prompt:="Gaps could not be found.", Title:="Gaps not found"
        FillInfo FunctionName:="Gaps", _
                 FileDate:=Format(dt, "mm/dd/yy"), _
                 Parameters:="", _
                 ExecutionTime:=Timer - StartTime, _
                 Result:="Failed - Gaps not found"
        Err.Raise 53
    End If

    Application.DisplayAlerts = True
    Exit Sub

CREATE_GAPS:
    ThisWorkbook.Sheets.Add After:=Sheets(ThisWorkbook.Sheets.Count)
    ActiveSheet.Name = "Gaps"
    Resume

End Sub

'---------------------------------------------------------------------------------------
' Proc : FilterSheet
' Date : 1/29/2013
' Desc : Remove all rows that do not match a specified string
'---------------------------------------------------------------------------------------
Sub FilterSheet(sFilter As String, ColNum As Integer, Match As Boolean)
    Dim Rng As Range
    Dim aRng() As Variant
    Dim aHeaders As Variant
    Dim StartTime As Double
    Dim iCounter As Long
    Dim i As Long
    Dim y As Long

    StartTime = Timer
    Set Rng = ActiveSheet.UsedRange
    aHeaders = Range(Cells(1, 1), Cells(1, ActiveSheet.UsedRange.Columns.Count))
    iCounter = 1

    Do While iCounter <= Rng.Rows.Count
        If Match = True Then
            If Rng(iCounter, ColNum).Value = sFilter Then
                i = i + 1
            End If
        Else
            If Rng(iCounter, ColNum).Value <> sFilter Then
                i = i + 1
            End If
        End If
        iCounter = iCounter + 1
    Loop

    ReDim aRng(1 To i, 1 To Rng.Columns.Count) As Variant

    iCounter = 1
    i = 0
    Do While iCounter <= Rng.Rows.Count
        If Match = True Then
            If Rng(iCounter, ColNum).Value = sFilter Then
                i = i + 1
                For y = 1 To Rng.Columns.Count
                    aRng(i, y) = Rng(iCounter, y)
                Next
            End If
        Else
            If Rng(iCounter, ColNum).Value <> sFilter Then
                i = i + 1
                For y = 1 To Rng.Columns.Count
                    aRng(i, y) = Rng(iCounter, y)
                Next
            End If
        End If
        iCounter = iCounter + 1
    Loop

    ActiveSheet.Cells.Delete
    Range(Cells(1, 1), Cells(UBound(aRng, 1), UBound(aRng, 2))) = aRng
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, UBound(aHeaders, 2))) = aHeaders
    FillInfo "FilterSheet", _
             "", _
             "Filter: " & sFilter & vbCrLf & "Col: " & Columns(ColNum).Address(False, False) & vbCrLf & "Match: " & Match, _
             Timer - StartTime, _
             "Complete"
End Sub


'---------------------------------------------------------------------------------------
' Proc : UserImportFile
' Date : 1/29/2013
' Desc : Prompts the user to select a file for import
'---------------------------------------------------------------------------------------
Sub UserImportFile(DestRange As Range, Optional DelFile As Boolean = False, Optional ShowAllData As Boolean = False)
    Dim File As String              'Full path to user selected file
    Dim FileDate As String          'Date the file was last modified
    Dim OldDispAlert As Boolean     'Original state of Application.DisplayAlerts

    OldDispAlert = Application.DisplayAlerts
    File = Application.GetOpenFilename()

    Application.DisplayAlerts = False
    If File <> "False" Then
        FileDate = Format(FileDateTime(File), "mm/dd/yy")
        Workbooks.Open File
        If ShowAllData = True Then
            ActiveSheet.AutoFilter.ShowAllData
            ActiveSheet.UsedRange.Columns.Hidden = False
            ActiveSheet.UsedRange.Rows.Hidden = False
        End If
        ActiveSheet.UsedRange.Copy Destination:=DestRange
        ActiveWorkbook.Close
        ThisWorkbook.Activate

        If DelFile = True Then
            DeleteFile File
        End If
    Else
        Err.Raise 18
    End If
    Application.DisplayAlerts = OldDispAlert
End Sub

'---------------------------------------------------------------------------------------
' Proc : FillInfo
' Date : 1/29/2013
' Desc : Used to add a line to the Info sheet
'---------------------------------------------------------------------------------------
Sub FillInfo(Optional FunctionName As String = "", Optional Result As String = "", Optional ExecutionTime As String = "", Optional Parameters As String = "", Optional FileDate As String = "")
    Dim Info As Worksheet           'Info worksheet if it exists, else this = nothing
    Dim LastSheet As Worksheet      'The previously selected worksheet
    Dim LastWorkbook As Workbook    'The previously activated workbook
    Set LastSheet = ActiveSheet
    Set LastWorkbook = ActiveWorkbook
    Dim Row As Long

    'This error is bypassed so you can determine whether or not the sheet exists
    On Error GoTo CREATE_INFO
    Set Info = ThisWorkbook.Sheets("Info")
    On Error GoTo 0

    ThisWorkbook.Activate
    Sheets("Info").Select
    Range("A1").Value = "Function"
    Range("B1").Value = "Created"
    Range("C1").Value = "Params"
    Range("D1").Value = "Exec Time"
    Range("E1").Value = "Result"

    Row = ActiveSheet.UsedRange.Rows.Count + 1
    Cells(Row, 1).Value = FunctionName
    Cells(Row, 2).Value = FileDate
    Cells(Row, 3).Value = Parameters
    Cells(Row, 4).Value = ExecutionTime
    Cells(Row, 5).Value = Result

    ActiveSheet.UsedRange.Columns.EntireColumn.AutoFit

    LastWorkbook.Activate
    LastSheet.Select
    Exit Sub

CREATE_INFO:
    Sheets.Add After:=Sheets(ThisWorkbook.Sheets.Count)
    ActiveSheet.Name = "Info"
    Resume

End Sub

'---------------------------------------------------------------------------------------
' Proc : GetWorkbookPath
' Date : 3/19/2013
' Desc : Gets the full path of ThisWorkbook
'---------------------------------------------------------------------------------------
Function GetWorkbookPath() As String
    Dim fullName As String
    Dim wrkbookName As String
    Dim pos As Long

    wrkbookName = ThisWorkbook.Name
    fullName = ThisWorkbook.fullName

    pos = InStr(1, fullName, wrkbookName, vbTextCompare)

    GetWorkbookPath = Left$(fullName, pos - 1)
End Function

'---------------------------------------------------------------------------------------
' Proc : EndsWith
' Date : 3/19/2013
' Desc : Checks if a string ends in a specified character
'---------------------------------------------------------------------------------------
Function EndsWith(ByVal InString As String, ByVal TestString As String) As Boolean
    EndsWith = (Right$(InString, Len(TestString)) = TestString)
End Function

'---------------------------------------------------------------------------------------
' Proc : Import117byISN
' Date : 4/10/2013
' Desc : Imports the most recent 117 report for the specified sales number
'---------------------------------------------------------------------------------------
Sub Import117byISN(RepType As ReportType, Destination As Range, Optional ByVal ISN As String = "", Optional Cancel As Boolean = False)
    Dim sPath As String
    Dim FileName As String

    If ISN = "" And Cancel = False Then
        ISN = InputBox("Inside Sales Number:", "Please enter the ISN#")
    Else
        If ISN = "" Then
            FillInfo "Import117byISN", "Failed - User Aborted", Parameters:="ReportType: " & ReportTypeText(RepType)
            Err.Raise 53
        End If
    End If

    If ISN <> "" Then
        Select Case RepType
            Case ReportType.DS:
                FileName = "3615 " & Format(Date, "yyyy-mm-dd") & " DSORDERS.xlsx"

            Case ReportType.Bo:
                FileName = "3615 " & Format(Date, "yyyy-mm-dd") & " BACKORDERS.xlsx"
        End Select

        sPath = "\\br3615gaps\gaps\3615 117 Report\ByInsideSalesNumber\" & ISN & "\" & FileName

        If FileExists(sPath) Then
            Workbooks.Open sPath
            ActiveSheet.UsedRange.Copy Destination:=Destination
            Application.DisplayAlerts = False
            ActiveWorkbook.Close
            Application.DisplayAlerts = True

            FillInfo FunctionName:="Import117byISN", _
                     Parameters:="Sales #: " & ISN, _
                     Result:="Complete"
            FillInfo Parameters:="Report Type: " & ReportTypeText(RepType)
            FillInfo Parameters:="Destination: " & Destination.Address(False, False)
        Else
            FillInfo FunctionName:="Import117byISN", _
                     Parameters:="Sales #: " & ISN, _
                     Result:="Failed - File not found"
            FillInfo Parameters:="Report Type: " & ReportTypeText(RepType)
            FillInfo Parameters:="Destination: " & Destination.Address(False, False)
            MsgBox Prompt:=ReportTypeText(RepType) & " report not found.", Title:="Error 53"
        End If
    Else
        FillInfo "Import117byISN", "Failed - Missing ISN", Parameters:="ReportType: " & ReportTypeText(RepType)
        Err.Raise 18
    End If

End Sub

'---------------------------------------------------------------------------------------
' Proc : Import473
' Date : 4/11/2013
' Desc : Imports a 473 report from the current day
'---------------------------------------------------------------------------------------
Sub Import473(Destination As Range, Optional Branch As String = "3615")
    Dim sPath As String
    Dim FileName As String
    Dim AlertStatus As Boolean

    FileName = "473 " & Format(Date, "yyyy-mm-dd") & ".xlsx"
    sPath = "\\br3615gaps\gaps\" & Branch & " 473 Download\" & FileName
    AlertStatus = Application.DisplayAlerts

    If FileExists(sPath) Then
        Workbooks.Open sPath

        ActiveSheet.UsedRange.Copy Destination:=Destination

        Application.DisplayAlerts = False
        ActiveWorkbook.Close
        Application.DisplayAlerts = AlertStatus
    Else
        Err.Raise 18, "Import473", "473 report not found."
    End If

End Sub

'---------------------------------------------------------------------------------------
' Proc : ReportTypeText
' Date : 4/10/2013
' Desc : Returns the report type as a string
'---------------------------------------------------------------------------------------
Function ReportTypeText(RepType As ReportType) As String
    Select Case RepType
        Case ReportType.Bo:
            ReportTypeText = "BO"
        Case ReportType.DS:
            ReportTypeText = "DS"
    End Select
End Function

'---------------------------------------------------------------------------------------
' Proc : DeleteColumn
' Date : 4/11/2013
' Desc : Removes a column based on text in the column header
'---------------------------------------------------------------------------------------
Sub DeleteColumn(HeaderText As String)
    Dim i As Integer

    For i = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
        If Trim(Cells(1, i).Value) = HeaderText Then
            Columns(i).Delete
            Exit For
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : FindColumn
' Date : 4/11/2013
' Desc : Returns the column number if a match is found
'---------------------------------------------------------------------------------------
Function FindColumn(HeaderText As String, Optional SearchArea As Range) As Integer
    Dim i As Integer: i = 0

    If TypeName(SearchArea) = Empty Then
        SearchArea = ActiveSheet.UsedRange
    End If

    For i = 1 To SearchArea.Columns.Count
        If Trim(SearchArea.Cells(1, i).Value) = HeaderText Then
            FindColumn = i
            Exit For
        End If
    Next
End Function

'---------------------------------------------------------------------------------------
' Proc : ImportSupplierContacts
' Date : 4/22/2013
' Desc : Imports the supplier contact master list
'---------------------------------------------------------------------------------------
Sub ImportSupplierContacts(Destination As Range)
    Const sPath As String = "\\br3615gaps\gaps\Contacts\Supplier Contact Master.xlsx"
    Dim PrevDispAlerts As Boolean

    PrevDispAlerts = Application.DisplayAlerts

    Workbooks.Open sPath
    ActiveSheet.UsedRange.Copy Destination:=Destination
    ActiveSheet.AutoFilterMode = False
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlerts
End Sub
