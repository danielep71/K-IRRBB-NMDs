Attribute VB_Name = "TEST_ImportCases"
'==============================================================================
' TEST_ImportCases
'------------------------------------------------------------------------------
' PURPOSE
'   VBA implementation of DATA-IMPORT-01 (docs/methodology/TEST_CASES.md): the
'   import from files through IRRBB_ImportData and WB_ImportData, end to end.
'   The fixtures of TEST_Fixtures are written to temporary files (LF and CRLF,
'   a file longer than one read chunk, a non-ASCII byte, an empty file), the
'   run settings on Config are set, the import runs, and DataSummary is read
'   back. The settings on Config and Segments are restored afterwards.
'
' FILES
'   Only in a new folder created for the run inside the system temporary
'   folder; the case deletes the files it wrote and the folder, and nothing
'   else. The import itself never writes a file.
'==============================================================================
Option Explicit
Option Private Module

Private Const HEADER As String = "as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date"
Private Const RATE_HEADER As String = "as_of_date,rate_id,currency,tenor_months,rate"
Private Const ACCEPTED As String = "COMPLETE - IMPORT ACCEPTED"

Private mSaved As Collection
Private mFiles As Collection
Private mRunFolder As String


Public Sub TEST_RunImportCase(ByVal caseId As String)
    Select Case caseId
        Case "DATA-IMPORT-01": CaseImport
        Case Else
            Err.Raise TEST_ERR_NO_CASE, "TEST_ImportCases.TEST_RunImportCase", "No VBA implementation for " & caseId & "."
    End Select
End Sub


'------------------------------------------------------------------------------
' DATA-IMPORT-01: files to DataSummary, end to end.
'------------------------------------------------------------------------------
Private Sub CaseImport()
    Const C As String = "DATA-IMPORT-01"
    Dim accounts As String, rates As String, invalid As String, window As Variant
    SaveSettings
    window = TEST_ExpectedLines("accounts_summary.window")
    accounts = WriteFixture("irrbb_accounts.csv", JoinFixture("accounts.csv", vbLf))
    rates = WriteFixture("irrbb_rates.csv", JoinFixture("market_rates.csv", vbLf))

    ' Accepted panel: totals and lineage as hand-computed.
    RunImport accounts, rates, window(1), window(2), "EUR_1M", "Yes"
    TEST_AssertText C, "LF files: status", StatusText(), ACCEPTED
    CheckTotals C, "LF files"
    TEST_AssertText C, "LF files: no finding", CellText("DataSummary_Findings", 1, 1), ""
    TEST_AssertText C, "LF files: rows read", LineageValue("Account rows read"), "51"
    TEST_AssertText C, "LF files: rows accepted", LineageValue("Account rows accepted"), "51"
    TEST_AssertText C, "LF files: warnings W04", LineageValue("Warnings W04"), "2"
    TEST_AssertText C, "LF files: warnings W05", LineageValue("Warnings W05"), "1"
    TEST_AssertText C, "LF files: rows outside the window", LineageValue("Rows outside the window (W06)"), "0"
    TEST_AssertText C, "LF files: window in lineage", LineageValue("Data window start") & "|" & _
                    LineageValue("Data window end"), window(1) & "|" & window(2)
    TEST_AssertText C, "LF files: settings in lineage", LineageValue("Rate series") & "|" & _
                    LineageValue("Closures reported by the month end (attested)"), "EUR_1M|Yes"
    TEST_AssertText C, "LF files: account panel SHA-256", LineageValue("Account panel SHA-256"), _
                    FirstExpected("sha256.accounts.csv")
    TEST_AssertText C, "LF files: market-rate SHA-256", LineageValue("Market-rate SHA-256"), _
                    FirstExpected("sha256.market_rates.csv")
    TEST_AssertText C, "LF files: toolkit commit", LineageValue("Toolkit commit"), _
                    CStr(NamedRange("Readme_SourceCommit").Value)
    TEST_AssertText C, "LF files: file name only", CellText("DataSummary_Info", 3, 1), "irrbb_accounts.csv"

    ' The same files with CRLF endings give the same result.
    accounts = WriteFixture("irrbb_accounts_crlf.csv", JoinFixture("accounts.csv", vbCrLf))
    rates = WriteFixture("irrbb_rates_crlf.csv", JoinFixture("market_rates.csv", vbCrLf))
    RunImport accounts, rates, window(1), window(2), "", "Yes"
    TEST_AssertText C, "CRLF files: status", StatusText(), ACCEPTED
    CheckTotals C, "CRLF files"
    TEST_AssertTrue C, "CRLF files: other bytes, other fingerprint", _
                    LineageValue("Account panel SHA-256") <> FirstExpected("sha256.accounts.csv")

    ' Rejected panel: findings listed, earlier totals withdrawn, no new totals.
    NamedRange("DataSummary_Totals").Cells(1, 1).Value = "STALE"
    invalid = WriteFixture("irrbb_invalid.csv", JoinFixture("accounts_invalid.csv", vbLf))
    RunImport invalid, rates, window(1), window(2), "", "Yes"
    TEST_AssertTrue C, "invalid panel: status rejected", Left$(StatusText(), 24) = "FAILED - IMPORT REJECTED"
    TEST_AssertText C, "invalid panel: totals withdrawn", CellText("DataSummary_Totals", 1, 1), ""
    TEST_AssertText C, "invalid panel: findings", FindingText(), JoinExpected("accounts_invalid_findings.errors")

    ' Settings the import refuses: no attestation, a missing file, no segment.
    RunImport accounts, rates, window(1), window(2), "", "No"
    TEST_AssertText C, "not attested: status", StatusText(), "FAILED"
    TEST_AssertTrue C, "not attested: reason", InStr(1, CellText("DataSummary_Info", 8, 1), "attested") > 0
    RunImport accounts & ".missing", rates, window(1), window(2), "", "Yes"
    TEST_AssertTrue C, "missing file: error 53", InStr(1, CellText("DataSummary_Info", 8, 1), "error 53") > 0
    NamedRange("Segments_Include").Value = "No"
    RunImport accounts, rates, window(1), window(2), "", "Yes"
    TEST_AssertTrue C, "no segment selected: refused", InStr(1, CellText("DataSummary_Info", 8, 1), "No segment") > 0

    ' A segment left out of the run is still imported and reconciled; only the
    ' lineage's segment list changes.
    NamedRange("Segments_Include").Value = "Yes"
    NamedRange("Segments_Include").Cells(2, 1).Value = "No"
    RunImport accounts, rates, window(1), window(2), "", "Yes"
    TEST_AssertText C, "segment left out: status", StatusText(), ACCEPTED
    CheckTotals C, "segment left out"
    TEST_AssertText C, "segment left out: lineage", LineageValue("Segments selected"), "RET_TX,WHS_NFC"
    NamedRange("Segments_Include").Value = "Yes"
    RunImport accounts, rates, window(1), window(2), "GBP_1M", "Yes"
    TEST_AssertTrue C, "unknown rate series: reason", InStr(1, CellText("DataSummary_Info", 8, 1), "GBP_1M") > 0
    NamedRange("Readme_SourceCommit").Value = "None: template without a VBA project"
    RunImport accounts, rates, window(1), window(2), "", "Yes"
    TEST_AssertTrue C, "no commit stamp: refused", InStr(1, CellText("DataSummary_Info", 8, 1), "commit stamp") > 0
    NamedRange("Readme_SourceCommit").Value = mSaved.Item(1)
    NamedRange("Config_WindowStart").Value = IsoDate(window(1)) + TimeSerial(12, 0, 0)
    WB_ImportData
    TEST_AssertTrue C, "window start with a time of day: refused", _
                    InStr(1, CellText("DataSummary_Info", 8, 1), "time of day") > 0

    ' File reading: more than one 64 KB chunk without a final line ending; a
    ' byte outside ASCII; an empty file.
    CheckLongFile C, rates
    accounts = WriteFixture("irrbb_nonascii.csv", HEADER & vbLf & "2025-01-31,A" & Chr$(233) & "1,RET_TX,EUR,1.00,,0,," & vbLf)
    RunImport accounts, rates, "2025-01-31", "2025-01-31", "", "Yes"
    TEST_AssertText C, "non-ASCII byte: findings", FindingText(), "2|E01"
    accounts = WriteFixture("irrbb_empty.csv", "")
    RunImport accounts, rates, "2025-01-31", "2025-01-31", "", "Yes"
    TEST_AssertText C, "empty file: findings", FindingText(), "1|E01"

    RestoreSettings
    DeleteFiles
End Sub


' 2,000 accounts in one month: about 80 KB, so lines cross a chunk boundary.
Private Sub CheckLongFile(ByVal caseId As String, ByVal rates As String)
    Dim text As String, i As Long, path As String
    text = HEADER
    For i = 1 To 2000
        text = text & vbLf & "2025-01-31,X" & Format$(i, "00000") & ",RET_TX,EUR,1.00,,0,,"
    Next i
    path = WriteFixture("irrbb_long.csv", text)
    TEST_AssertTrue caseId, "long file: larger than one chunk", Len(text) > 65536
    RunImport path, rates, "2025-01-31", "2025-01-31", "", "Yes"
    TEST_AssertText caseId, "long file: status", StatusText(), ACCEPTED
    TEST_AssertText caseId, "long file: total", CellText("DataSummary_Totals", 1, 1) & "|" & _
                    CellText("DataSummary_Totals", 1, 4) & "|" & CellText("DataSummary_Totals", 1, 5), "2025-01-31|2000|2000"
End Sub


Private Sub CheckTotals(ByVal caseId As String, ByVal label As String)
    Dim expected As Variant, fields() As String, i As Long, area As Range
    expected = TEST_ExpectedLines("accounts_summary.totals")
    Set area = NamedRange("DataSummary_Totals")
    For i = 1 To UBound(expected)
        fields = Split(expected(i), "|")
        TEST_AssertText caseId, label & ": total " & i, CellText("DataSummary_Totals", i, 1) & "|" & _
                        CellText("DataSummary_Totals", i, 2) & "|" & CellText("DataSummary_Totals", i, 3) & "|" & _
                        CellText("DataSummary_Totals", i, 4), fields(0) & "|" & fields(1) & "|" & fields(2) & "|" & fields(3)
        TEST_AssertNear caseId, label & ": total " & i & " balance", CDbl(area.Cells(i, 5).Value), _
                        DecimalValue(fields(4)), "abs", 0.005
    Next i
    TEST_AssertText caseId, label & ": no extra total", CellText("DataSummary_Totals", UBound(expected) + 1, 1), ""
End Sub


'==============================================================================
' Helpers
'==============================================================================
Private Sub RunImport(ByVal accounts As String, ByVal rates As String, ByVal first As String, _
                      ByVal last As String, ByVal rateSeries As String, ByVal attested As String)
    NamedRange("Config_AccountsFile").Value = accounts
    NamedRange("Config_RatesFile").Value = rates
    NamedRange("Config_WindowStart").Value = IsoDate(first)
    NamedRange("Config_WindowEnd").Value = IsoDate(last)
    NamedRange("Config_RateSeries").Value = rateSeries
    NamedRange("Config_ClosureAttested").Value = attested
    WB_ImportData
End Sub


Private Function StatusText() As String
    StatusText = CStr(NamedRange("DataSummary_Status").Value)
End Function


Private Function CellText(ByVal areaName As String, ByVal row As Long, ByVal col As Long) As String
    CellText = CStr(NamedRange(areaName).Cells(row, col).Value)
End Function


Private Function LineageValue(ByVal item As String) As String
    Dim i As Long
    For i = 1 To 50
        If CellText("DataSummary_Lineage", i, 1) = item Then
            LineageValue = CellText("DataSummary_Lineage", i, 2)
            Exit Function
        End If
    Next i
    LineageValue = "(no lineage item " & item & ")"
End Function


' "line|code" of every finding on DataSummary, separated by ",".
Private Function FindingText() As String
    Dim i As Long, text As String
    i = 1
    Do While Len(CellText("DataSummary_Findings", i, 3)) > 0
        If i > 1 Then text = text & ","
        text = text & CellText("DataSummary_Findings", i, 2) & "|" & CellText("DataSummary_Findings", i, 3)
        i = i + 1
    Loop
    FindingText = text
End Function


Private Function FirstExpected(ByVal name As String) As String
    Dim lines As Variant
    lines = TEST_ExpectedLines(name)
    FirstExpected = lines(LBound(lines))
End Function


Private Function JoinExpected(ByVal name As String) As String
    Dim lines As Variant, i As Long, text As String
    lines = TEST_ExpectedLines(name)
    For i = LBound(lines) To UBound(lines)
        If i > LBound(lines) Then text = text & ","
        text = text & lines(i)
    Next i
    JoinExpected = text
End Function


Private Function JoinFixture(ByVal name As String, ByVal lineEnd As String) As String
    Dim lines As Variant, i As Long, text As String
    lines = TEST_FixtureLines(name)
    For i = LBound(lines) To UBound(lines)
        text = text & lines(i) & lineEnd
    Next i
    JoinFixture = text
End Function


' Write text as bytes (one byte per character) to a new file in the temporary
' folder; returns its path.
Private Function WriteFixture(ByVal fileName As String, ByVal text As String) As String
    Dim path As String, f As Integer
    path = RunFolder() & Application.PathSeparator & fileName
    f = FreeFile
    Open path For Binary Access Write As #f
    If Len(text) > 0 Then Put #f, , text
    Close #f
    If mFiles Is Nothing Then Set mFiles = New Collection
    mFiles.Add path
    WriteFixture = path
End Function


Private Sub DeleteFiles()
    Dim path As Variant
    If Not mFiles Is Nothing Then
        For Each path In mFiles
            If Len(Dir$(CStr(path))) > 0 Then Kill CStr(path)
        Next path
    End If
    Set mFiles = Nothing
    If Len(mRunFolder) > 0 Then RmDir mRunFolder
    mRunFolder = vbNullString
End Sub


' A folder of this run's own inside the temporary folder, so no existing file
' is ever overwritten or deleted.
Private Function RunFolder() As String
    Dim attempt As Long, candidate As String
    If Len(mRunFolder) = 0 Then
        Do
            attempt = attempt + 1
            candidate = TempFolder() & Application.PathSeparator & "irrbb_test_" & Format$(Now, "yyyymmddhhnnss") & _
                        "_" & attempt
        Loop While Len(Dir$(candidate, vbDirectory)) > 0
        MkDir candidate
        mRunFolder = candidate
    End If
    RunFolder = mRunFolder
End Function


Private Function TempFolder() As String
    TempFolder = Environ$("TEMP")
    If Len(TempFolder) = 0 Then TempFolder = Environ$("TMP")
    If Len(TempFolder) = 0 Then TempFolder = Environ$("TMPDIR")
    If Len(TempFolder) = 0 Then TempFolder = "/tmp"
End Function


' Item 1 is the build stamp; the settings follow.
Private Sub SaveSettings()
    Dim cell As Range
    Set mSaved = New Collection
    mSaved.Add NamedRange("Readme_SourceCommit").Value
    For Each cell In SettingCells()
        mSaved.Add cell.Value
    Next cell
End Sub


Private Sub RestoreSettings()
    Dim cell As Range, i As Long
    NamedRange("Readme_SourceCommit").Value = mSaved.Item(1)
    i = 1
    For Each cell In SettingCells()
        i = i + 1
        cell.Value = mSaved.Item(i)
    Next cell
End Sub


Private Function SettingCells() As Collection
    Dim out As New Collection, settingName As Variant, cell As Range
    For Each settingName In Array("Config_AccountsFile", "Config_RatesFile", "Config_WindowStart", _
                                  "Config_WindowEnd", "Config_RateSeries", "Config_ClosureAttested")
        out.Add NamedRange(CStr(settingName))
    Next settingName
    For Each cell In NamedRange("Segments_Include").Cells
        out.Add cell
    Next cell
    Set SettingCells = out
End Function


Private Function NamedRange(ByVal definedName As String) As Range
    Set NamedRange = ThisWorkbook.Names(definedName).RefersToRange
End Function


Private Function IsoDate(ByVal text As String) As Date
    Dim value As Date
    If CORE_ParseIsoDate(text, value) <> CORE_PARSE_OK Then
        Err.Raise 5, "TEST_ImportCases.IsoDate", "Not an ISO date: " & text
    End If
    IsoDate = value
End Function


Private Function DecimalValue(ByVal text As String) As Double
    Dim value As Double
    If CORE_ParseDecimal(text, 8, value) <> CORE_PARSE_OK Then
        Err.Raise 5, "TEST_ImportCases.DecimalValue", "Not a decimal: " & text
    End If
    DecimalValue = value
End Function
