Attribute VB_Name = "TEST_DataCases"
'==============================================================================
' TEST_DataCases
'------------------------------------------------------------------------------
' PURPOSE
'   VBA implementations of DATA-VALID-01, DATA-INVALID-01, DATA-PIT-01 and
'   DATA-PIT-02 (docs/methodology/TEST_CASES.md) against CORE_Import. The fixtures and
'   their hand-computed expectations come from TEST_Fixtures, generated from
'   tests/fixtures and tests/expected. The literal cases below are re-derived
'   from the Python reading of the contract by tools/test_vba_harness.py.
'
' FORMAT
'   One Expect* call per line, arguments as literals. In a literal file "|"
'   separates lines; expected findings are "line:code" separated by ",".
'==============================================================================
Option Explicit
Option Private Module


Public Sub TEST_RunDataCase(ByVal caseId As String)
    Select Case caseId
        Case "DATA-VALID-01": CaseValid
        Case "DATA-INVALID-01": CaseInvalid
        Case "DATA-PIT-01": CasePitOutcomes
        Case "DATA-PIT-02": CasePitLeakage
        Case Else
            Err.Raise TEST_ERR_NO_CASE, "TEST_DataCases.TEST_RunDataCase", "No VBA implementation for " & caseId & "."
    End Select
End Sub


'------------------------------------------------------------------------------
' DATA-VALID-01: the valid panel is accepted whole; warnings and totals per
' date, segment and currency equal the hand computation; the market-rate
' fixture is accepted.
'------------------------------------------------------------------------------
Private Sub CaseValid()
    Const C As String = "DATA-VALID-01"
    Dim window As Variant, counts As Variant, totals As Variant, limit As Variant, fields() As String
    Dim first As Date, last As Date, tolerance As Double, i As Long

    ImportAccounts TEST_FixtureLines("accounts.csv")
    CompareFindings C, "accounts.csv", TEST_ExpectedLines("accounts_summary.errors")
    counts = TEST_ExpectedLines("accounts_summary.counts")
    TEST_AssertText C, "accepted rows", CStr(CORE_AcceptedRowCount()), counts(1)
    TEST_AssertText C, "accounts", CStr(CORE_AccountCount()), counts(2)

    window = TEST_ExpectedLines("accounts_summary.window")
    first = IsoDate(window(1))
    last = IsoDate(window(2))
    TEST_AssertText C, "missing extracts", CStr(CORE_AccountsSummarize(first, last)), "0"
    CompareList C, "warning", WarningLines(), TEST_ExpectedLines("accounts_summary.warnings")

    totals = TEST_ExpectedLines("accounts_summary.totals")
    limit = TEST_ExpectedLines("accounts_summary.tolerance")
    tolerance = DecimalValue(limit(1))
    TEST_AssertText C, "total count", CStr(CORE_TotalCount()), CStr(UBound(totals) - LBound(totals) + 1)
    For i = 1 To CORE_TotalCount()
        If i > UBound(totals) Then Exit For
        fields = Split(totals(i), "|")
        TEST_AssertText C, "total " & i, CORE_TotalDate(i) & "|" & CORE_TotalSegment(i) & "|" & _
                        CORE_TotalCurrency(i) & "|" & CORE_TotalAccounts(i), _
                        fields(0) & "|" & fields(1) & "|" & fields(2) & "|" & fields(3)
        TEST_AssertNear C, "total " & i & " balance", CORE_TotalBalance(i), DecimalValue(fields(4)), "abs", tolerance
    Next i

    ImportRates TEST_FixtureLines("market_rates.csv")
    TEST_AssertText C, "market rate findings", RateFindingText(), ""
End Sub


'------------------------------------------------------------------------------
' DATA-INVALID-01: first error per line of the invalid panel (E02-E11); E01 for
' the header and the field count; cross-row precedence; market-rate rules; a
' missing month end inside the window (E12).
'------------------------------------------------------------------------------
Private Sub CaseInvalid()
    Const C As String = "DATA-INVALID-01"
    ImportAccounts TEST_FixtureLines("accounts_invalid.csv")
    CompareFindings C, "accounts_invalid.csv", TEST_ExpectedLines("accounts_invalid_findings.errors")

    ExpectAccountFindings C, "as_of_date,account_id|2025-01-31,A001", "1:E01"
    ExpectAccountFindings C, "as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date|2025-01-31,A001,RET_TX,EUR,1.00,,0,", "2:E01"
    ExpectAccountFindings C, "as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date|2025-01-31,C001,RET_TX,EUR,10.00,,0,2020-01-01,|2025-01-31,C001,RET_TX,EUR,10.00,,0,2021-01-01,", "2:E11,3:E11"
    ExpectAccountFindings C, "as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date|2025-01-31,C002,RET_TX,EUR,10.00,,0,2020-01-01,|2025-02-28,C002,RET_TX,USD,10.00,,0,2021-01-01,", "2:E11,3:E11"
    ExpectAccountFindings C, "as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date|2025-03-31,C003,RET_TX,EUR,1.00,,0,,|2025-01-31,C003,RET_TX,EUR,1.00,,0,,2025-02-15", "2:E10"
    ExpectAccountFindings C, "as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date|2025-01-31,a001,RET_TX,EUR,1.00,,0,,|2025-01-31,A001,RET_TX,EUR,1.00,,0,,", ""
    ExpectAccountFindings C, "as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date|2025-01-31,C004,RET_TX,EUR,1e3,,0,,|2025-01-31,C005,RET_TX,EUR,1.00,0.25000001,0,,|2025-01-31,C006,RET_TX,EUR,-0.00,,1,,", "2:E04,3:E04"
    ExpectRateFindings C, "as_of_date,rate_id,currency,tenor_months,rate|2025-01-31,EUR_1M,EUR,1,0.0290|2025-01-31,EUR_1M,EUR,1,0.0300|2025-02-28,EUR_1M,EUR,3,0.0300|2025-02-28,EUR_3M,EUR,0,0.0300|2025-02-27,EUR_6M,EUR,6,0.0100|2025-02-28,GBP_1M,GBP,1,0.0100|2025-03-31,EUR_1M,EUR,1,", "3:E05,4:E08,5:E04,6:E02,7:E07,8:E03"
    ExpectRateFindings C, "as_of_date,rate_id,currency,tenor_months|2025-01-31,EUR_1M,EUR,1", "1:E01"
    ExpectMissingExtract C, "accounts.csv", "2025-03-31"
    ExpectWindowGaps C, "as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date|9999-10-31,Z001,RET_TX,EUR,1.00,,0,,", "9999-10-31", "9999-12-31", "9999-11-30,9999-12-31"
End Sub


'------------------------------------------------------------------------------
' DATA-PIT-01: transition outcome of every observed month, warnings, the
' close_date values visible at each origin and the per-period lineage equal the
' hand assignment in availability_outcomes.json; a month-end closure has no row
' in its month.
'------------------------------------------------------------------------------
Private Sub CasePitOutcomes()
    Const C As String = "DATA-PIT-01"
    Dim window As Variant, counts As Variant
    ImportAccounts TEST_FixtureLines("availability.csv")
    CompareFindings C, "availability.csv", Array()
    counts = TEST_ExpectedLines("availability_outcomes.counts")
    TEST_AssertText C, "accepted rows", CStr(CORE_AcceptedRowCount()), counts(1)
    window = TEST_ExpectedLines("availability_outcomes.window")
    TEST_AssertText C, "missing extracts", CStr(CORE_AccountsSummarize(IsoDate(window(1)), IsoDate(window(2)))), "0"
    CompareList C, "outcome", OutcomeLines(), TEST_ExpectedLines("availability_outcomes.outcomes")
    CompareList C, "warning", WarningLines(), TEST_ExpectedLines("availability_outcomes.warnings")
    CompareList C, "visible close_date", VisibleCloseLines(TEST_ExpectedLines("availability_outcomes.origins")), _
                TEST_ExpectedLines("availability_outcomes.visible")
    ComparePeriods C, TEST_ExpectedLines("availability_outcomes.lineage")

    ExpectAccountFindings C, "as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date|2025-05-31,M001,RET_TX,EUR,10.00,,0,2020-01-01,2025-06-30|2025-06-30,M001,RET_TX,EUR,0.00,,0,2020-01-01,2025-06-30", "3:E10"
    ExpectOutcomes C, "as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date|2025-05-31,M001,RET_TX,EUR,10.00,,0,2020-01-01,2025-06-30|2025-06-30,M001,RET_TX,EUR,0.00,,0,2020-01-01,2025-06-30", "2025-05-31", "2025-07-31", "M001|2025-05-31|RET_TX|closure"
    ExpectOutcomes C, "as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date|2025-01-31,G001,RET_TX,EUR,5.00,,0,,2025-09-30|2025-03-31,G001,RET_TX,EUR,5.00,,0,,2025-09-30|2025-01-31,G002,RET_TX,EUR,5.00,,0,,|2025-02-28,G002,WHS_NFC,EUR,5.00,,0,,", "2025-01-31", "2025-04-30", "G001|2025-01-31|RET_TX|gap,G001|2025-03-31|RET_TX|censored,G002|2025-01-31|RET_TX|observed,G002|2025-02-28|WHS_NFC|censored"
End Sub


'------------------------------------------------------------------------------
' DATA-PIT-02: at every origin of both panels, altering, adding or removing
' everything learned after the origin leaves the point-in-time view unchanged;
' no account with a row at the origin shows a closure in the view; a missing
' month end rejects the import (E12).
'------------------------------------------------------------------------------
Private Sub CasePitLeakage()
    Const C As String = "DATA-PIT-02"
    CheckNoLookAhead C, "availability.csv", "availability_outcomes.window"
    CheckNoLookAhead C, "accounts.csv", "accounts_summary.window"
    ExpectMissingExtract C, "availability.csv", "2025-03-31"
    CheckViewScope C
End Sub


'==============================================================================
' Expectation helpers
'==============================================================================
Private Sub ExpectAccountFindings(ByVal caseId As String, ByVal fileText As String, ByVal expected As String)
    ImportAccounts Split(fileText, "|")
    TEST_AssertText caseId, "account findings of """ & Left$(fileText, 60) & "...""", AccountFindingText(), expected
End Sub


Private Sub ExpectRateFindings(ByVal caseId As String, ByVal fileText As String, ByVal expected As String)
    ImportRates Split(fileText, "|")
    TEST_AssertText caseId, "rate findings of """ & Left$(fileText, 60) & "...""", RateFindingText(), expected
End Sub


' A valid panel without the rows of one month end: E12 for that month only.
Private Sub ExpectMissingExtract(ByVal caseId As String, ByVal fixture As String, ByVal monthEnd As String)
    Dim source As Variant, kept() As String, window As Variant, i As Long, n As Long, missing As Long
    source = TEST_FixtureLines(fixture)
    ReDim kept(1 To UBound(source))
    For i = LBound(source) To UBound(source)
        If Left$(source(i), Len(monthEnd) + 1) <> monthEnd & "," Then
            n = n + 1
            kept(n) = source(i)
        End If
    Next i
    ReDim Preserve kept(1 To n)
    ImportAccounts kept
    window = TEST_ExpectedLines(WindowKey(fixture))
    missing = CORE_AccountsSummarize(IsoDate(window(1)), IsoDate(window(2)))
    TEST_AssertText caseId, fixture & " missing extracts without " & monthEnd, CStr(missing), "1"
    If missing >= 1 Then TEST_AssertText caseId, "missing extract", CORE_MissingExtract(1), monthEnd
End Sub


' Missing month ends (E12) of a literal panel over a literal window.
Private Sub ExpectWindowGaps(ByVal caseId As String, ByVal fileText As String, ByVal first As String, _
                             ByVal last As String, ByVal expected As String)
    Dim i As Long, text As String
    ImportAccounts Split(fileText, "|")
    CORE_AccountsSummarize IsoDate(first), IsoDate(last)
    For i = 1 To CORE_MissingExtractCount()
        If i > 1 Then text = text & ","
        text = text & CORE_MissingExtract(i)
    Next i
    TEST_AssertText caseId, "missing extracts " & first & " to " & last, text, expected
End Sub


' Transition outcomes "account|origin|segment|outcome" of a literal panel.
Private Sub ExpectOutcomes(ByVal caseId As String, ByVal fileText As String, ByVal first As String, _
                           ByVal last As String, ByVal expected As String)
    Dim lines As Variant
    ImportAccounts Split(fileText, "|")
    CORE_AccountsSummarize IsoDate(first), IsoDate(last)
    lines = OutcomeLines()
    TEST_AssertText caseId, "outcomes " & first & " to " & last, JoinLines(lines, ","), expected
End Sub


' The view at every origin of the fixture's window is the same for the panel
' and for the panel with everything learned after that origin changed.
Private Sub CheckNoLookAhead(ByVal caseId As String, ByVal fixture As String, ByVal windowKey As String)
    Dim source As Variant, window As Variant, origin As Date, last As Date, before As String, after As String
    Dim altered As Variant, label As String, n As Long
    source = TEST_FixtureLines(fixture)
    window = TEST_ExpectedLines(windowKey)
    origin = IsoDate(window(1))
    last = IsoDate(window(2))
    Do
        label = fixture & " at " & IsoText(origin)
        ImportAccounts source
        CORE_AccountsSummarize IsoDate(window(1)), last
        n = CORE_ViewBuild(origin)
        before = ViewText()
        TEST_AssertTrue caseId, label & ": view not empty", n > 0
        TEST_AssertText caseId, label & ": closures in rows at the origin", CStr(ClosuresAtOrigin(origin)), "0"

        altered = AlteredPanel(source, IsoText(origin))
        If origin < last Then TEST_AssertTrue caseId, label & ": later information changed", _
                                              Join(altered, "|") <> Join(source, "|")
        ImportAccounts altered
        TEST_AssertText caseId, label & ": altered panel findings", AccountFindingText(), ""
        CORE_AccountsSummarize IsoDate(window(1)), last
        CORE_ViewBuild origin
        after = ViewText()
        TEST_AssertText caseId, label & ": view unchanged", after, before
        If origin >= last Then Exit Do
        origin = IsoDate(NextMonthEndText(IsoText(origin)))
    Loop
End Sub


' The view needs an accepted import and keeps to the data window: refused after
' a finding or a missing extract, and rows before the window are not in it.
Private Sub CheckViewScope(ByVal caseId As String)
    Const HEADER As String = "as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date"
    ImportAccounts Split(HEADER & "|2025-01-31,V001,RET_TX,EUR,1.00,,0,,|2025-02-28,V001,RET_TX,EUR,-1.00,,0,,", "|")
    CORE_AccountsSummarize DateSerial(2025, 1, 31), DateSerial(2025, 1, 31)
    TEST_AssertText caseId, "view after a finding", CStr(ViewError(DateSerial(2025, 1, 31))), "5"

    ImportAccounts Split(HEADER & "|2025-01-31,V001,RET_TX,EUR,1.00,,0,,|2025-03-31,V001,RET_TX,EUR,1.00,,0,,", "|")
    CORE_AccountsSummarize DateSerial(2025, 1, 31), DateSerial(2025, 3, 31)
    TEST_AssertText caseId, "view after a missing extract", CStr(ViewError(DateSerial(2025, 3, 31))), "5"

    ImportAccounts Split(HEADER & "|2024-12-31,V001,RET_TX,EUR,1.00,,0,,|2025-01-31,V001,RET_TX,EUR,1.00,,0,,", "|")
    CORE_AccountsSummarize DateSerial(2025, 1, 31), DateSerial(2025, 1, 31)
    TEST_AssertText caseId, "view of a window with an earlier row", CStr(CORE_ViewBuild(DateSerial(2025, 1, 31))), "1"
    TEST_AssertText caseId, "view origin after the window", CStr(ViewError(DateSerial(2025, 2, 28))), "5"
End Sub


' Error number raised by CORE_ViewBuild at the origin, or 0.
Private Function ViewError(ByVal origin As Date) As Long
    On Error GoTo Failed
    CORE_ViewBuild origin
    Exit Function
Failed:
    ViewError = Err.Number
End Function


' The panel with everything after originText changed: later rows of accounts
' ending in an odd digit removed, other later rows altered, close dates after
' the origin moved or erased, empty close dates set in the future, and a new
' account added in the following month.
Private Function AlteredPanel(ByVal source As Variant, ByVal originText As String) As Variant
    Dim out() As String, f() As String, i As Long, n As Long, nextText As String
    ReDim out(1 To UBound(source) - LBound(source) + 2)
    n = 1
    out(1) = source(LBound(source))
    For i = LBound(source) + 1 To UBound(source)
        f = Split(source(i), ",")
        If f(0) > originText Then
            If InStr(1, "13579", Right$(f(1), 1)) > 0 Then GoTo NextLine
            f(4) = "7777.77"
            If f(2) = "WHS_NFC" Then
                f(2) = "RET_TX"
            Else
                f(2) = "WHS_NFC"
            End If
            f(5) = "0.2000"
        End If
        If Len(f(8)) = 0 Then
            f(8) = "2026-12-31"
        ElseIf f(8) > originText Then
            If CLng(Right$(f(8), 2)) Mod 2 = 1 Then
                f(8) = vbNullString
            Else
                f(8) = "2027-01-31"
            End If
        End If
        n = n + 1
        out(n) = Join(f, ",")
NextLine:
    Next i
    nextText = NextMonthEndText(originText)
    n = n + 1
    out(n) = nextText & ",Z99,RET_TX,EUR,1.00,,0," & nextText & ","
    ReDim Preserve out(1 To n)
    AlteredPanel = out
End Function


' View rows dated at the origin whose account shows a closure in the view.
Private Function ClosuresAtOrigin(ByVal origin As Date) As Long
    Dim i As Long, j As Long, count As Long
    For i = 1 To CORE_ViewCount()
        If CORE_ViewAsOf(i) = IsoText(origin) Then
            For j = 1 To CORE_ViewCount()
                If CORE_ViewAccount(j) = CORE_ViewAccount(i) And Len(CORE_ViewCloseDate(j)) > 0 Then count = count + 1
            Next j
        End If
    Next i
    ClosuresAtOrigin = count
End Function


Private Function ViewText() As String
    Dim i As Long, text As String
    For i = 1 To CORE_ViewCount()
        text = text & CORE_ViewRecord(i) & ";"
    Next i
    ViewText = text
End Function


Private Sub CompareFindings(ByVal caseId As String, ByVal fixture As String, ByVal expected As Variant)
    CompareList caseId, fixture & " finding", FindingLines(), expected
End Sub


' Equal length, then equal items in order.
Private Sub CompareList(ByVal caseId As String, ByVal label As String, ByVal actual As Variant, _
                        ByVal expected As Variant)
    Dim i As Long, nActual As Long, nExpected As Long
    nActual = UBound(actual) - LBound(actual) + 1
    nExpected = UBound(expected) - LBound(expected) + 1
    TEST_AssertText caseId, label & " count", CStr(nActual), CStr(nExpected)
    For i = 0 To nActual - 1
        If i >= nExpected Then Exit For
        TEST_AssertText caseId, label & " " & (i + 1), CStr(actual(LBound(actual) + i)), _
                        CStr(expected(LBound(expected) + i))
    Next i
End Sub


' Per-period lineage against "month|segment|currency|observed|closures|gaps|
' censored|censored balance|end of sample|masked" lines; the balance is compared
' numerically, the other fields as text.
Private Sub ComparePeriods(ByVal caseId As String, ByVal expected As Variant)
    Dim i As Long, n As Long, fields() As String, balance As Double
    n = UBound(expected) - LBound(expected) + 1
    TEST_AssertText caseId, "period count", CStr(CORE_PeriodCount()), CStr(n)
    For i = 1 To CORE_PeriodCount()
        If i > n Then Exit For
        fields = Split(expected(LBound(expected) + i - 1), "|")
        TEST_AssertText caseId, "period " & i, CORE_PeriodDate(i) & "|" & CORE_PeriodSegment(i) & "|" & _
                        CORE_PeriodCurrency(i) & "|" & CORE_PeriodObserved(i) & "|" & CORE_PeriodClosures(i) & "|" & _
                        CORE_PeriodGaps(i) & "|" & CORE_PeriodCensored(i) & "|" & CORE_PeriodEndOfSample(i) & "|" & _
                        CORE_PeriodMasked(i), fields(0) & "|" & fields(1) & "|" & fields(2) & "|" & fields(3) & "|" & _
                        fields(4) & "|" & fields(5) & "|" & fields(6) & "|" & fields(8) & "|" & fields(9)
        If CORE_ParseDecimal(fields(7), 4, balance) <> CORE_PARSE_OK Then
            Err.Raise 5, "TEST_DataCases.ComparePeriods", "Not a decimal: " & fields(7)
        End If
        TEST_AssertNear caseId, "period " & i & " censored balance", CORE_PeriodCensoredBalance(i), balance, _
                        "abs", 0.005
    Next i
End Sub


Private Sub ImportAccounts(ByVal lines As Variant)
    Dim i As Long
    If CORE_AccountsBegin(lines(LBound(lines))) Then
        For i = LBound(lines) + 1 To UBound(lines)
            CORE_AccountsAddLine lines(i)
        Next i
        CORE_AccountsEnd
    End If
End Sub


Private Sub ImportRates(ByVal lines As Variant)
    Dim i As Long
    If CORE_RatesBegin(lines(LBound(lines))) Then
        For i = LBound(lines) + 1 To UBound(lines)
            CORE_RatesAddLine lines(i)
        Next i
        CORE_RatesEnd
    End If
End Sub


Private Function FindingLines() As Variant
    Dim out() As String, i As Long
    If CORE_AccountFindingCount() = 0 Then
        FindingLines = Array()
        Exit Function
    End If
    ReDim out(1 To CORE_AccountFindingCount())
    For i = 1 To CORE_AccountFindingCount()
        out(i) = CORE_AccountFindingLine(i) & "|" & CORE_AccountFindingCode(i)
    Next i
    FindingLines = out
End Function


Private Function WarningLines() As Variant
    Dim out() As String, i As Long
    If CORE_WarningCount() = 0 Then
        WarningLines = Array()
        Exit Function
    End If
    ReDim out(1 To CORE_WarningCount())
    For i = 1 To CORE_WarningCount()
        out(i) = CORE_WarningCode(i) & "|" & CORE_WarningAccount(i) & "|" & CORE_WarningDate(i)
    Next i
    WarningLines = out
End Function


Private Function AccountFindingText() As String
    Dim i As Long, text As String
    For i = 1 To CORE_AccountFindingCount()
        If i > 1 Then text = text & ","
        text = text & CORE_AccountFindingLine(i) & ":" & CORE_AccountFindingCode(i)
    Next i
    AccountFindingText = text
End Function


Private Function RateFindingText() As String
    Dim i As Long, text As String
    For i = 1 To CORE_RateFindingCount()
        If i > 1 Then text = text & ","
        text = text & CORE_RateFindingLine(i) & ":" & CORE_RateFindingCode(i)
    Next i
    RateFindingText = text
End Function


Private Function OutcomeLines() As Variant
    Dim out() As String, i As Long
    If CORE_OutcomeCount() = 0 Then
        OutcomeLines = Array()
        Exit Function
    End If
    ReDim out(1 To CORE_OutcomeCount())
    For i = 1 To CORE_OutcomeCount()
        out(i) = CORE_OutcomeAccount(i) & "|" & CORE_OutcomeOrigin(i) & "|" & CORE_OutcomeSegment(i) & "|" & _
                 CORE_OutcomeCode(i)
    Next i
    OutcomeLines = out
End Function


' "origin|account|close_date" for every account showing a closure in the view
' at each origin (one line per account; the view is ordered by account).
Private Function VisibleCloseLines(ByVal origins As Variant) As Variant
    Dim out() As String, n As Long, k As Long, i As Long, line As String, previous As String
    ReDim out(1 To 64)
    For k = LBound(origins) To UBound(origins)
        CORE_ViewBuild IsoDate(origins(k))
        previous = vbNullString
        For i = 1 To CORE_ViewCount()
            If Len(CORE_ViewCloseDate(i)) > 0 Then
                line = origins(k) & "|" & CORE_ViewAccount(i) & "|" & CORE_ViewCloseDate(i)
                If line <> previous Then
                    n = n + 1
                    If n > UBound(out) Then
                        ReDim Preserve out(1 To 2 * UBound(out))
                    End If
                    out(n) = line
                    previous = line
                End If
            End If
        Next i
    Next k
    If n = 0 Then
        VisibleCloseLines = Array()
    Else
        ReDim Preserve out(1 To n)
        VisibleCloseLines = out
    End If
End Function


Private Function JoinLines(ByVal lines As Variant, ByVal separator As String) As String
    Dim i As Long, text As String
    For i = LBound(lines) To UBound(lines)
        If i > LBound(lines) Then text = text & separator
        text = text & lines(i)
    Next i
    JoinLines = text
End Function


Private Function WindowKey(ByVal fixture As String) As String
    Select Case fixture
        Case "accounts.csv": WindowKey = "accounts_summary.window"
        Case "availability.csv": WindowKey = "availability_outcomes.window"
        Case Else: Err.Raise 5, "TEST_DataCases.WindowKey", "No window for " & fixture
    End Select
End Function


' ISO text of the month end after an ISO month end.
Private Function NextMonthEndText(ByVal monthEndText As String) As String
    Dim y As Long, m As Long, lastDay As Date
    y = CLng(Left$(monthEndText, 4))
    m = CLng(Mid$(monthEndText, 6, 2)) + 1
    If m = 13 Then
        y = y + 1
        m = 1
    End If
    If m = 12 Then
        lastDay = DateSerial(y, 12, 31)
    Else
        lastDay = DateSerial(y, m + 1, 1) - 1
    End If
    NextMonthEndText = IsoText(lastDay)
End Function


Private Function IsoText(ByVal d As Date) As String
    IsoText = Format$(Year(d), "0000") & "-" & Format$(Month(d), "00") & "-" & Format$(Day(d), "00")
End Function


Private Function IsoDate(ByVal text As String) As Date
    Dim value As Date
    If CORE_ParseIsoDate(text, value) <> CORE_PARSE_OK Then
        Err.Raise 5, "TEST_DataCases.IsoDate", "Not an ISO date: " & text
    End If
    IsoDate = value
End Function


Private Function DecimalValue(ByVal text As String) As Double
    Dim value As Double
    If CORE_ParseDecimal(text, 8, value) <> CORE_PARSE_OK Then
        Err.Raise 5, "TEST_DataCases.DecimalValue", "Not a decimal: " & text
    End If
    DecimalValue = value
End Function
