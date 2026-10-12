Attribute VB_Name = "WB_Import"
'==============================================================================
' WB_Import  (workbook adapter, issue #71)
'------------------------------------------------------------------------------
' PURPOSE
'   Run the data import from the workbook: read the run settings on Config and
'   Segments through their names, call IRRBB_ImportData and write the result
'   to DataSummary under the status protocol of docs/WORKBOOK_CONTRACT.md.
'   Host glue only: no validation rule or calculation lives here.
'
' STATUS (DataSummary_Status)
'   RUNNING - previous results withdrawn   set first; every result area cleared
'   COMPLETE - IMPORT ACCEPTED              totals, warnings and lineage current
'   FAILED - IMPORT REJECTED (n errors)     findings listed; no totals
'   FAILED                                  the run stopped; primary error and
'                                           any cleanup error in the info block
'
' SETTINGS
'   Config_AccountsFile, Config_RatesFile: paths of the two files (read only).
'   Config_WindowStart, Config_WindowEnd: month-end dates.
'   Config_RateSeries: rate_id, may be empty.
'   Config_ClosureAttested: "Yes" to attest timely closure reporting.
'   Readme_SourceCommit: the build stamp, passed on as the toolkit commit.
'   Segments_Include: every segment must be "Yes"; what excluding a segment
'   means for an import is not yet defined (#71), so the run refuses.
'==============================================================================
Option Explicit

Private Const SOURCE As String = "WB_Import.WB_ImportData"
Private Const INFO_STARTED As Long = 1
Private Const INFO_FINISHED As Long = 2
Private Const INFO_ACCOUNTS As Long = 3
Private Const INFO_RATES As Long = 4
Private Const INFO_WINDOW As Long = 5
Private Const INFO_SERIES As Long = 6
Private Const INFO_ATTESTED As Long = 7
Private Const INFO_PRIMARY As Long = 8
Private Const INFO_CLEANUP As Long = 9


'------------------------------------------------------------------------------
' WB_ImportData: the Import data button.
'------------------------------------------------------------------------------
Public Sub WB_ImportData()
    Dim primary As String, cleanup As String, started As Date, errors As Long, cut As Boolean
    Dim accountsPath As String, ratesPath As String, first As Date, last As Date
    Dim rateSeries As String, attested As Boolean

    On Error GoTo Fail
    started = Now
    Withdraw started

    accountsPath = CStr(Setting("Config_AccountsFile"))
    ratesPath = CStr(Setting("Config_RatesFile"))
    first = DateSetting("Config_WindowStart", "Data window start")
    last = DateSetting("Config_WindowEnd", "Data window end")
    rateSeries = CStr(Setting("Config_RateSeries"))
    attested = (CStr(Setting("Config_ClosureAttested")) = "Yes")
    RequireAllSegments

    Info INFO_ACCOUNTS, FileNameOf(accountsPath)
    Info INFO_RATES, FileNameOf(ratesPath)
    Info INFO_WINDOW, IsoText(first) & " to " & IsoText(last)
    Info INFO_SERIES, rateSeries
    Info INFO_ATTESTED, IIf(attested, "Yes", "No")

    errors = IRRBB_ImportData(accountsPath, ratesPath, first, last, rateSeries, attested, _
                              CStr(Setting("Readme_SourceCommit")))

    cut = WriteTable("DataSummary_Findings", IRRBB_ImportFindings(), Array(1, 3, 4))
    cut = WriteTable("DataSummary_Warnings", IRRBB_ImportWarnings(), Array(1, 2, 3)) Or cut
    cut = WriteTable("DataSummary_Lineage", IRRBB_ImportLineage(), Array(1, 2)) Or cut
    If errors = 0 Then
        cut = WriteTable("DataSummary_Totals", IRRBB_ImportTotals(), Array(1, 2, 3)) Or cut
        Status "COMPLETE - IMPORT ACCEPTED" & IIf(cut, " (a table is cut at its last row)", "")
    Else
        Status "FAILED - IMPORT REJECTED (" & errors & " errors)" & IIf(cut, " (a table is cut at its last row)", "")
    End If
    Info INFO_FINISHED, Now
    Exit Sub

Fail:
    ' Resume clears the error, so a second failure while reporting is trapped too.
    primary = "Primary error " & Err.Number & ": " & Err.Description & " [" & Err.Source & "]"
    Resume ReportPrimary
ReportPrimary:
    On Error GoTo CleanupFail
    ClearResults
    Status "FAILED"
    Info INFO_PRIMARY, primary
    Info INFO_FINISHED, Now
    Exit Sub
CleanupFail:
    cleanup = "Cleanup error " & Err.Number & ": " & Err.Description & " [" & Err.Source & "]"
    Resume ReportBoth
ReportBoth:
    On Error Resume Next
    Status "FAILED"
    Info INFO_PRIMARY, primary
    Info INFO_CLEANUP, cleanup
End Sub


'------------------------------------------------------------------------------
' WB_ChooseAccountsFile / WB_ChooseRatesFile: file pickers that only store the
' chosen path; nothing is read until WB_ImportData.
'------------------------------------------------------------------------------
Public Sub WB_ChooseAccountsFile()
    ChooseFile "Config_AccountsFile", "Account panel CSV"
End Sub


Public Sub WB_ChooseRatesFile()
    ChooseFile "Config_RatesFile", "Market-rate CSV"
End Sub


'==============================================================================
' Helpers
'==============================================================================
Private Sub ChooseFile(ByVal settingName As String, ByVal title As String)
    Dim chosen As Variant
    chosen = Application.GetOpenFilename("CSV files (*.csv),*.csv", , title)
    If VarType(chosen) = vbString Then NamedRange(settingName).Value = chosen
End Sub


' Mark every earlier result withdrawn before anything is computed.
Private Sub Withdraw(ByVal started As Date)
    Status "RUNNING - previous results withdrawn"
    NamedRange("DataSummary_Info").ClearContents
    ClearResults
    Info INFO_STARTED, started
End Sub


Private Sub ClearResults()
    NamedRange("DataSummary_Totals").ClearContents
    NamedRange("DataSummary_Findings").ClearContents
    NamedRange("DataSummary_Warnings").ClearContents
    NamedRange("DataSummary_Lineage").ClearContents
End Sub


' Write rows 1..n of a facade table (row 0 holds the column names, already on
' the sheet) into the named area; text columns are formatted as text first so
' that dates, codes and identifiers stay as given. Returns True when the table
' is longer than the area and was cut.
Private Function WriteTable(ByVal areaName As String, ByVal table As Variant, ByVal textColumns As Variant) As Boolean
    Dim area As Range, rowCount As Long, cols As Long, i As Long, j As Long, block() As Variant
    Set area = NamedRange(areaName)
    rowCount = UBound(table, 1)
    cols = UBound(table, 2)
    If rowCount > area.Rows.Count Then
        WriteTable = True
        rowCount = area.Rows.Count
    End If
    For j = LBound(textColumns) To UBound(textColumns)
        area.Columns(textColumns(j)).NumberFormat = "@"
    Next j
    If rowCount = 0 Then Exit Function
    ReDim block(1 To rowCount, 1 To cols)
    For i = 1 To rowCount
        For j = 1 To cols
            block(i, j) = table(i, j)
        Next j
    Next i
    area.Resize(rowCount, cols).Value = block
End Function


Private Sub RequireAllSegments()
    Dim cell As Range
    For Each cell In NamedRange("Segments_Include").Cells
        If CStr(cell.Value) <> "Yes" Then
            Err.Raise 5, SOURCE, "Excluding a segment is not defined for the import yet (#71): set Include to " & _
                      "Yes for every segment."
        End If
    Next cell
End Sub


Private Function Setting(ByVal settingName As String) As Variant
    Setting = NamedRange(settingName).Value
End Function


' A setting that must hold a date (a date cell, not text).
Private Function DateSetting(ByVal settingName As String, ByVal label As String) As Date
    Dim value As Variant
    value = Setting(settingName)
    If VarType(value) <> vbDate Then Err.Raise 5, SOURCE, label & " must be a date."
    DateSetting = value
End Function


Private Sub Status(ByVal text As String)
    NamedRange("DataSummary_Status").Value = text
End Sub


Private Sub Info(ByVal item As Long, ByVal value As Variant)
    NamedRange("DataSummary_Info").Cells(item, 1).Value = value
End Sub


Private Function NamedRange(ByVal definedName As String) As Range
    Set NamedRange = ThisWorkbook.Names(definedName).RefersToRange
End Function


Private Function FileNameOf(ByVal path As String) As String
    Dim i As Long
    For i = Len(path) To 1 Step -1
        If Mid$(path, i, 1) = "\" Or Mid$(path, i, 1) = "/" Then Exit For
    Next i
    FileNameOf = Mid$(path, i + 1)
End Function


Private Function IsoText(ByVal d As Date) As String
    IsoText = Format$(Year(d), "0000") & "-" & Format$(Month(d), "00") & "-" & Format$(Day(d), "00")
End Function
