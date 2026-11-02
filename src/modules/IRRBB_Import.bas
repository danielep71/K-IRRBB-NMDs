Attribute VB_Name = "IRRBB_Import"
'==============================================================================
' IRRBB_Import  (public facade, issue #71)
'------------------------------------------------------------------------------
' PURPOSE
'   Import the account panel and the market-rate file of
'   docs/methodology/DATA_CONTRACT.md from CSV files and expose the result as
'   tables. Validation, warnings, totals and outcomes are CORE_Import's; this
'   module reads the files and checks the run settings at the public boundary.
'
' FILE ACCESS (SECURITY.md, security scope)
'   Read only. Each file is opened for binary read with writing by others
'   denied while it is open (a file another process is writing cannot be
'   opened: error 70), streamed in 64 KB chunks and closed, so the import
'   reads one snapshot; nothing is written next to it and no other file is
'   touched. Lines end in LF or CRLF. A file must be below 2 GB
'   (the VBA file length is a Long). Bytes outside ASCII are reported as E01
'   by the validation.
'
' RESULTS
'   IRRBB_ImportData returns the number of errors: account and market-rate
'   findings plus missing extracts (E12). Zero means the panel is accepted.
'   The import is all-or-nothing: totals and the per-period outcome lineage
'   are returned only for an accepted panel. Every IRRBB_Import* table is a Variant array (0 To n, 1 To k)
'   whose row 0 holds the column names, so a table is never empty.
'
' ERRORS
'   Error 5 for invalid settings (window not two month ends in order, empty
'   path, closure reporting not attested, no commit stamp, no or an unknown or
'   repeated segment selected, rate series not in the file); error 5 from
'   IRRBB_ImportFindings, IRRBB_ImportWarnings and IRRBB_ImportLineage before
'   any import has completed (IRRBB_ImportTotals and IRRBB_ImportOutcomes then
'   return the header row only); error 53 when a file does not exist;
'   file-system errors pass through.
'==============================================================================
Option Explicit

Private Const CHUNK_SIZE As Long = 65536
Private Const SOURCE As String = "IRRBB_Import.IRRBB_ImportData"

Private mAccountsFile As String
Private mAccountsBytes As Double
Private mRatesFile As String
Private mRatesBytes As Double
Private mAccountErrors As Long
Private mMissingExtracts As Long
Private mRateErrors As Long
Private mAccepted As Boolean
Private mImported As Boolean
Private mHeaderAccepted As Boolean
Private mFirstMonthEnd As Date
Private mLastMonthEnd As Date
Private mRateSeriesSetting As String
Private mAttested As Boolean
Private mToolkitCommit As String
Private mSegments As String
Private mImportTime As Date
Private mAccountsSha As String
Private mRatesSha As String


'------------------------------------------------------------------------------
' IRRBB_ImportData
'   accountsPath     [in]  account panel CSV (DATA_CONTRACT.md, account panel)
'   ratesPath        [in]  market-rate CSV (DATA_CONTRACT.md, market rates)
'   firstMonthEnd    [in]  first month end of the data window
'   lastMonthEnd     [in]  last month end of the data window
'   rateSeries       [in]  rate_id the rate model uses; "" when none is chosen
'   includedSegments [in]  segment codes selected for the run, comma-separated
'                          (e.g. "RET_TX,WHS_NFC"); at least one, no repeats.
'                          The import still validates and reconciles every
'                          segment; the list is recorded in the lineage and
'                          decides which transitions enter model samples
'                          (DATA_CONTRACT.md, segment selection)
'   closuresAttested [in]  True when the user attests that closures are reported
'                          by the month end in which they occur (timely closure
'                          reporting); the import refuses to run otherwise
'   toolkitCommit    [in]  commit of the toolkit code running the import, as
'                          stamped on the workbook: 40 hex digits, optionally
'                          followed by the builder's " (dirty working tree:
'                          NOT evidence-grade)"; recorded in the lineage
'   Returns the number of errors; 0 means the panel is accepted.
'------------------------------------------------------------------------------
Public Function IRRBB_ImportData(ByVal accountsPath As String, ByVal ratesPath As String, _
                                 ByVal firstMonthEnd As Date, ByVal lastMonthEnd As Date, _
                                 ByVal rateSeries As String, ByVal includedSegments As String, _
                                 ByVal closuresAttested As Boolean, ByVal toolkitCommit As String) As Long
    mImported = False
    mAccepted = False
    If Not closuresAttested Then
        Err.Raise 5, SOURCE, "Timely closure reporting is not attested; the import does not run."
    End If
    If Not IsCommitStamp(toolkitCommit) Then
        Err.Raise 5, SOURCE, "The workbook has no source commit stamp (Readme_SourceCommit): build it with " & _
                  "tools/devbuild, or enter the 40-digit commit of the imported source."
    End If
    RequireSegmentList includedSegments
    If Len(accountsPath) = 0 Or Len(ratesPath) = 0 Then
        Err.Raise 5, SOURCE, "Both the account panel file and the market-rate file are required."
    End If
    If Not CORE_IsMonthEnd(firstMonthEnd) Or Not CORE_IsMonthEnd(lastMonthEnd) _
            Or firstMonthEnd > lastMonthEnd Then
        Err.Raise 5, SOURCE, "The data window must be two month ends in order, without a time of day."
    End If
    RequireFile accountsPath
    RequireFile ratesPath

    mFirstMonthEnd = firstMonthEnd
    mLastMonthEnd = lastMonthEnd
    mRateSeriesSetting = rateSeries
    mAttested = closuresAttested
    mToolkitCommit = toolkitCommit
    mSegments = includedSegments
    mImportTime = Now
    mAccountsFile = FileNameOf(accountsPath)
    mRatesFile = FileNameOf(ratesPath)
    mAccountsBytes = StreamFile(accountsPath, True, mAccountsSha)
    mAccountErrors = CORE_AccountFindingCount()
    mMissingExtracts = 0
    If mAccountErrors = 0 Then mMissingExtracts = CORE_AccountsSummarize(firstMonthEnd, lastMonthEnd)
    mRatesBytes = StreamFile(ratesPath, False, mRatesSha)
    mRateErrors = CORE_RateFindingCount()

    If Len(rateSeries) > 0 And mRateErrors = 0 Then
        If Not CORE_RateSeriesExists(rateSeries) Then
            Err.Raise 5, SOURCE, "Rate series " & rateSeries & " is not in the market-rate file."
        End If
    End If
    mImported = True
    mAccepted = (mAccountErrors + mMissingExtracts + mRateErrors = 0)
    IRRBB_ImportData = mAccountErrors + mMissingExtracts + mRateErrors
End Function


'------------------------------------------------------------------------------
' IRRBB_ImportTotals: (0 To n, 1 To 5) Date, Segment, Currency, Accounts,
' Balance (currency units, at input precision) per date, segment and currency.
' n = 0 unless the last import was accepted.
'------------------------------------------------------------------------------
Public Function IRRBB_ImportTotals() As Variant
    Dim out() As Variant, i As Long, n As Long
    If mAccepted Then n = CORE_TotalCount()
    ReDim out(0 To n, 1 To 5)
    out(0, 1) = "Date": out(0, 2) = "Segment": out(0, 3) = "Currency": out(0, 4) = "Accounts"
    out(0, 5) = "Balance"
    For i = 1 To n
        out(i, 1) = CORE_TotalDate(i)
        out(i, 2) = CORE_TotalSegment(i)
        out(i, 3) = CORE_TotalCurrency(i)
        out(i, 4) = CORE_TotalAccounts(i)
        out(i, 5) = CORE_TotalBalance(i)
    Next i
    IRRBB_ImportTotals = out
End Function


'------------------------------------------------------------------------------
' IRRBB_ImportOutcomes: (0 To n, 1 To 10) Month end, Segment, Currency,
' Observed, Verified closures, Gaps, Censored, Censored last balance (currency
' units, at input precision), End of sample, Close dates masked: the
' per-period lineage of DATA_CONTRACT.md (outcomes). The outcome columns count
' the transitions from the month end under the segment at that month end; an
' account whose close_date the point-in-time view at the month end masks is
' counted under the segment and currency of its latest row on or before it.
' n = 0 unless the last import was accepted.
'------------------------------------------------------------------------------
Public Function IRRBB_ImportOutcomes() As Variant
    Dim out() As Variant, i As Long, n As Long
    If mAccepted Then n = CORE_PeriodCount()
    ReDim out(0 To n, 1 To 10)
    out(0, 1) = "Month end": out(0, 2) = "Segment": out(0, 3) = "Currency": out(0, 4) = "Observed"
    out(0, 5) = "Verified closures": out(0, 6) = "Gaps": out(0, 7) = "Censored"
    out(0, 8) = "Censored last balance": out(0, 9) = "End of sample": out(0, 10) = "Close dates masked"
    For i = 1 To n
        out(i, 1) = CORE_PeriodDate(i)
        out(i, 2) = CORE_PeriodSegment(i)
        out(i, 3) = CORE_PeriodCurrency(i)
        out(i, 4) = CORE_PeriodObserved(i)
        out(i, 5) = CORE_PeriodClosures(i)
        out(i, 6) = CORE_PeriodGaps(i)
        out(i, 7) = CORE_PeriodCensored(i)
        out(i, 8) = CORE_PeriodCensoredBalance(i)
        out(i, 9) = CORE_PeriodEndOfSample(i)
        out(i, 10) = CORE_PeriodMasked(i)
    Next i
    IRRBB_ImportOutcomes = out
End Function


'------------------------------------------------------------------------------
' IRRBB_ImportFindings: (0 To n, 1 To 4) File, Line, Code, Date: every error of
' the last import, account panel first, then missing extracts (Line 0, the
' month end in Date), then market rates.
'------------------------------------------------------------------------------
Public Function IRRBB_ImportFindings() As Variant
    Dim out() As Variant, i As Long, r As Long
    RequireImport
    ReDim out(0 To mAccountErrors + mMissingExtracts + mRateErrors, 1 To 4)
    out(0, 1) = "File": out(0, 2) = "Line": out(0, 3) = "Code": out(0, 4) = "Date"
    For i = 1 To mAccountErrors
        r = r + 1
        out(r, 1) = mAccountsFile: out(r, 2) = CORE_AccountFindingLine(i): out(r, 3) = CORE_AccountFindingCode(i)
        out(r, 4) = vbNullString
    Next i
    For i = 1 To mMissingExtracts
        r = r + 1
        out(r, 1) = mAccountsFile: out(r, 2) = 0: out(r, 3) = "E12": out(r, 4) = CORE_MissingExtract(i)
    Next i
    For i = 1 To mRateErrors
        r = r + 1
        out(r, 1) = mRatesFile: out(r, 2) = CORE_RateFindingLine(i): out(r, 3) = CORE_RateFindingCode(i)
        out(r, 4) = vbNullString
    Next i
    IRRBB_ImportFindings = out
End Function


'------------------------------------------------------------------------------
' IRRBB_ImportWarnings: (0 To n, 1 To 3) Code, Account, Date, in the contract's
' order. n = 0 unless the account panel had no finding.
'------------------------------------------------------------------------------
Public Function IRRBB_ImportWarnings() As Variant
    Dim out() As Variant, i As Long, n As Long
    RequireImport
    If mAccountErrors = 0 Then n = CORE_WarningCount()
    ReDim out(0 To n, 1 To 3)
    out(0, 1) = "Code": out(0, 2) = "Account": out(0, 3) = "Date"
    For i = 1 To n
        out(i, 1) = CORE_WarningCode(i)
        out(i, 2) = CORE_WarningAccount(i)
        out(i, 3) = CORE_WarningDate(i)
    Next i
    IRRBB_ImportWarnings = out
End Function


'------------------------------------------------------------------------------
' IRRBB_ImportLineage: (0 To n, 1 To 2) Item, Value: file names (no directory)
' and sizes in bytes, rows read and accepted, rows outside the window (W06),
' accounts, warnings by code, market-rate rows read, the run settings (data
' window, rate series, selected segments, attestation), the SHA-256 of each file's bytes, the
' toolkit commit and the local import time (DATA_CONTRACT.md, lineage).
'------------------------------------------------------------------------------
Public Function IRRBB_ImportLineage() As Variant
    Dim out() As Variant, codes As Variant, i As Long, k As Long, counts(1 To 6) As Long
    RequireImport
    codes = Array("W01", "W02", "W03", "W04", "W05", "W06")
    If mAccountErrors = 0 Then
        For i = 1 To CORE_WarningCount()
            For k = 0 To 5
                If CORE_WarningCode(i) = codes(k) Then counts(k + 1) = counts(k + 1) + 1
            Next k
        Next i
    End If
    ReDim out(0 To 23, 1 To 2)
    out(0, 1) = "Item": out(0, 2) = "Value"
    out(1, 1) = "Account panel file": out(1, 2) = mAccountsFile
    out(2, 1) = "Account panel bytes": out(2, 2) = mAccountsBytes
    out(3, 1) = "Account rows read": out(3, 2) = CORE_AccountLinesRead()
    out(4, 1) = "Account rows accepted": out(4, 2) = CORE_AcceptedRowCount()
    out(5, 1) = "Accounts": out(5, 2) = CORE_AccountCount()
    out(6, 1) = "Market-rate file": out(6, 2) = mRatesFile
    out(7, 1) = "Market-rate bytes": out(7, 2) = mRatesBytes
    out(8, 1) = "Market-rate rows read": out(8, 2) = CORE_RateLinesRead()
    For k = 1 To 5
        out(8 + k, 1) = "Warnings " & codes(k - 1): out(8 + k, 2) = counts(k)
    Next k
    out(14, 1) = "Rows outside the window (W06)": out(14, 2) = counts(6)
    out(15, 1) = "Data window start": out(15, 2) = IsoText(mFirstMonthEnd)
    out(16, 1) = "Data window end": out(16, 2) = IsoText(mLastMonthEnd)
    out(17, 1) = "Rate series": out(17, 2) = mRateSeriesSetting
    out(18, 1) = "Closures reported by the month end (attested)": out(18, 2) = IIf(mAttested, "Yes", "No")
    out(19, 1) = "Account panel SHA-256": out(19, 2) = mAccountsSha
    out(20, 1) = "Market-rate SHA-256": out(20, 2) = mRatesSha
    out(21, 1) = "Toolkit commit": out(21, 2) = mToolkitCommit
    out(22, 1) = "Import time (local; UTC offset not recorded)": out(22, 2) = Format$(mImportTime, "yyyy-mm-dd hh:nn:ss")
    out(23, 1) = "Segments selected": out(23, 2) = mSegments
    IRRBB_ImportLineage = out
End Function


'==============================================================================
' File input
'==============================================================================

' Stream one file into CORE_Import, line by line, and its bytes into the
' SHA-256 fingerprint (one pass); returns its size in bytes.
Private Function StreamFile(ByVal path As String, ByVal accounts As Boolean, ByRef sha As String) As Double
    Dim f As Integer, size As Double, remaining As Double, want As Long, bytes() As Byte, carry As String
    Dim lineStart As Long, lf As Long, started As Boolean, opened As Boolean

    f = FreeFile
    Open path For Binary Access Read Lock Write As #f
    opened = True
    On Error GoTo Failed
    size = LOF(f)
    remaining = size
    CORE_Sha256Begin
    Do While remaining > 0
        ' Progress is counted in bytes requested, never in what a read returned,
        ' so the loop ends on every host.
        If remaining >= CHUNK_SIZE Then
            want = CHUNK_SIZE
        Else
            want = CLng(remaining)
        End If
        ReDim bytes(0 To want - 1)
        Get #f, , bytes
        remaining = remaining - want
        CORE_Sha256Update bytes, want
        ' One character per byte; ASCII bytes map to themselves under every code
        ' page, and any other byte becomes a character above 127 (E01).
        carry = carry & StrConv(bytes, vbUnicode)
        lineStart = 1
        lf = InStr(lineStart, carry, vbLf, vbBinaryCompare)
        Do While lf > 0
            EmitLine Mid$(carry, lineStart, lf - lineStart), accounts, started
            lineStart = lf + 1
            lf = InStr(lineStart, carry, vbLf, vbBinaryCompare)
        Loop
        carry = Mid$(carry, lineStart)
    Loop
    Close #f
    opened = False
    sha = CORE_Sha256End()

    ' A last line without a line ending still counts; an empty file is a
    ' missing header (E01).
    If Len(carry) > 0 Or Not started Then EmitLine carry, accounts, started
    If mHeaderAccepted Then
        If accounts Then
            CORE_AccountsEnd
        Else
            CORE_RatesEnd
        End If
    End If
    StreamFile = size
    Exit Function

Failed:
    Dim number As Long, description As String
    number = Err.Number
    description = Err.Description
    If opened Then Close #f
    Err.Raise number, SOURCE, description
End Function


' The first line of a file is its header; a rejected header ends the import of
' that file (only finding: line 1, E01).
Private Sub EmitLine(ByVal text As String, ByVal accounts As Boolean, ByRef started As Boolean)
    If Not started Then
        started = True
        If accounts Then
            mHeaderAccepted = CORE_AccountsBegin(text)
        Else
            mHeaderAccepted = CORE_RatesBegin(text)
        End If
        Exit Sub
    End If
    If Not mHeaderAccepted Then Exit Sub
    If accounts Then
        CORE_AccountsAddLine text
    Else
        CORE_RatesAddLine text
    End If
End Sub


Private Sub RequireFile(ByVal path As String)
    If Len(Dir$(path)) = 0 Then Err.Raise 53, SOURCE, "File not found: " & FileNameOf(path)
End Sub


Private Sub RequireImport()
    If Not mImported Then Err.Raise 5, "IRRBB_Import", "No import has completed."
End Sub


' A comma-separated list of known segment codes, at least one, none repeated.
Private Sub RequireSegmentList(ByVal list As String)
    Dim codes() As String, i As Long, j As Long
    If Len(list) = 0 Then Err.Raise 5, SOURCE, "No segment is selected for the run."
    codes = Split(list, ",")
    For i = LBound(codes) To UBound(codes)
        If Not CORE_IsSegmentCode(codes(i)) Then
            Err.Raise 5, SOURCE, "Unknown segment in the selection: " & codes(i) & "."
        End If
        For j = LBound(codes) To i - 1
            If codes(j) = codes(i) Then Err.Raise 5, SOURCE, "Segment selected twice: " & codes(i) & "."
        Next j
    Next i
End Sub


' 40 lower-case hex digits, alone or followed by the builder's dirty-tree note.
Private Function IsCommitStamp(ByVal stamp As String) As Boolean
    Dim i As Long, ch As String
    If Len(stamp) < 40 Then Exit Function
    For i = 1 To 40
        ch = Mid$(stamp, i, 1)
        If Not ((ch >= "0" And ch <= "9") Or (ch >= "a" And ch <= "f")) Then Exit Function
    Next i
    IsCommitStamp = (Len(stamp) = 40 Or Mid$(stamp, 41) = " (dirty working tree: NOT evidence-grade)")
End Function


Private Function IsoText(ByVal d As Date) As String
    IsoText = Format$(Year(d), "0000") & "-" & Format$(Month(d), "00") & "-" & Format$(Day(d), "00")
End Function


' File name without its directory (lineage records no path).
Private Function FileNameOf(ByVal path As String) As String
    Dim i As Long
    For i = Len(path) To 1 Step -1
        If Mid$(path, i, 1) = "\" Or Mid$(path, i, 1) = "/" Then Exit For
    Next i
    FileNameOf = Mid$(path, i + 1)
End Function
