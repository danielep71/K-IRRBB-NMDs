Attribute VB_Name = "TEST_DataCases"
'==============================================================================
' TEST_DataCases
'------------------------------------------------------------------------------
' PURPOSE
'   VBA implementations of DATA-VALID-01 and DATA-INVALID-01
'   (docs/methodology/TEST_CASES.md) against CORE_Import. The fixtures and
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
    ExpectMissingExtract C, "2025-03-31"
    ExpectWindowGaps C, "as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date|9999-10-31,Z001,RET_TX,EUR,1.00,,0,,", "9999-10-31", "9999-12-31", "9999-11-30,9999-12-31"
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


' The valid panel without the rows of one month end: E12 for that month only.
Private Sub ExpectMissingExtract(ByVal caseId As String, ByVal monthEnd As String)
    Dim source As Variant, kept() As String, window As Variant, i As Long, n As Long, missing As Long
    source = TEST_FixtureLines("accounts.csv")
    ReDim kept(1 To UBound(source))
    For i = LBound(source) To UBound(source)
        If Left$(source(i), Len(monthEnd) + 1) <> monthEnd & "," Then
            n = n + 1
            kept(n) = source(i)
        End If
    Next i
    ReDim Preserve kept(1 To n)
    ImportAccounts kept
    window = TEST_ExpectedLines("accounts_summary.window")
    missing = CORE_AccountsSummarize(IsoDate(window(1)), IsoDate(window(2)))
    TEST_AssertText caseId, "missing extracts without " & monthEnd, CStr(missing), "1"
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
