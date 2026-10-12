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
'   The import is all-or-nothing: totals are returned only for an accepted
'   panel. Every IRRBB_Import* table is a Variant array (0 To n, 1 To k)
'   whose row 0 holds the column names, so a table is never empty.
'
' ERRORS
'   Error 5 for invalid settings (window not two month ends in order, empty
'   path, closure reporting not attested, rate series not in the file);
'   error 53 when a file does not exist; file-system errors pass through.
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
'   closuresAttested [in]  True when the user attests that closures are reported
'                          by the month end in which they occur (timely closure
'                          reporting); the import refuses to run otherwise
'   toolkitCommit    [in]  commit of the toolkit code running the import, as
'                          stamped on the workbook; recorded in the lineage
'   Returns the number of errors; 0 means the panel is accepted.
'------------------------------------------------------------------------------
Public Function IRRBB_ImportData(ByVal accountsPath As String, ByVal ratesPath As String, _
                                 ByVal firstMonthEnd As Date, ByVal lastMonthEnd As Date, _
                                 ByVal rateSeries As String, ByVal closuresAttested As Boolean, _
                                 ByVal toolkitCommit As String) As Long
    mImported = False
    mAccepted = False
    If Not closuresAttested Then
        Err.Raise 5, SOURCE, "Timely closure reporting is not attested; the import does not run."
    End If
    If Len(accountsPath) = 0 Or Len(ratesPath) = 0 Then
        Err.Raise 5, SOURCE, "Both the account panel file and the market-rate file are required."
    End If
    If Not CORE_IsMonthEnd(firstMonthEnd) Or Not CORE_IsMonthEnd(lastMonthEnd) _
            Or firstMonthEnd > lastMonthEnd Then
        Err.Raise 5, SOURCE, "The data window must be two month ends in order."
    End If
    RequireFile accountsPath
    RequireFile ratesPath

    mFirstMonthEnd = firstMonthEnd
    mLastMonthEnd = lastMonthEnd
    mRateSeriesSetting = rateSeries
    mAttested = closuresAttested
    mToolkitCommit = toolkitCommit
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
' window, rate series, attestation), the SHA-256 of each file's bytes, the
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
    ReDim out(0 To 22, 1 To 2)
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
