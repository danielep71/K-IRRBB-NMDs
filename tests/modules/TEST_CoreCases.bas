Attribute VB_Name = "TEST_CoreCases"
'==============================================================================
' TEST_CoreCases
'------------------------------------------------------------------------------
' PURPOSE
'   VBA implementations of the harness and data-format cases registered in
'   docs/methodology/TEST_CASES.md: HARN-OUT-01, DATA-PARSE-01, DATA-PARSE-02
'   and DATA-CODES-01. Expected values are written by hand from
'   docs/methodology/DATA_CONTRACT.md and the harness rules; every Expect* call
'   below is re-checked against the Python reference reading of the contract
'   by tools/test_vba_harness.py, so a wrong expectation fails without Excel.
'
' FORMAT
'   One Expect* call per line, arguments as literals, so the Python check can
'   read them. Do not compute expected values here.
'==============================================================================
Option Explicit
Option Private Module


Public Sub TEST_RunCoreCase(ByVal caseId As String)
    Select Case caseId
        Case "HARN-OUT-01": CaseHarnessRules
        Case "DATA-PARSE-01": CaseDates
        Case "DATA-PARSE-02": CaseDecimals
        Case "DATA-CODES-01": CaseCodes
        Case Else
            Err.Raise TEST_ERR_NO_CASE, "TEST_CoreCases.TEST_RunCoreCase", "No VBA implementation for " & caseId & "."
    End Select
End Sub


'------------------------------------------------------------------------------
' HARN-OUT-01: outcome and tolerance rules of the harness itself. A case with no
' assertion, a failed assertion or an error is never PASS; unbuilt is NOT RUN.
'------------------------------------------------------------------------------
Private Sub CaseHarnessRules()
    Const C As String = "HARN-OUT-01"
    ExpectOutcome C, False, False, 0, 0, "NOT RUN"
    ExpectOutcome C, False, True, 1, 0, "NOT RUN"
    ExpectOutcome C, True, False, 0, 0, "FAIL"
    ExpectOutcome C, True, False, 3, 3, "PASS"
    ExpectOutcome C, True, False, 3, 2, "FAIL"
    ExpectOutcome C, True, True, 3, 3, "ERROR"
    ExpectOutcome C, True, True, 1, 0, "ERROR"

    ExpectWithin C, 1#, 1.05, "abs", 0.1, True
    ExpectWithin C, 1#, 1.2, "abs", 0.1, False
    ExpectWithin C, 1000#, 1000.5, "rel", 0.001, True
    ExpectWithin C, 1000#, 1002#, "rel", 0.001, False
    ExpectWithin C, 0.5, 0.5005, "rel", 0.001, True
    ExpectWithin C, 0.5, 0.502, "rel", 0.001, False
    ExpectWithin C, 2#, 2#, "exact", 0#, True
    ExpectWithin C, 2#, 2.000001, "exact", 0#, False
    ExpectWithin C, 2#, 2#, "fuzzy", 1#, False
End Sub


'------------------------------------------------------------------------------
' DATA-PARSE-01: ISO dates, missing versus invalid, leap years, month ends.
'------------------------------------------------------------------------------
Private Sub CaseDates()
    Const C As String = "DATA-PARSE-01"
    ExpectDate C, "2025-01-31", CORE_PARSE_OK, 2025, 1, 31
    ExpectDate C, "2024-02-29", CORE_PARSE_OK, 2024, 2, 29
    ExpectDate C, "2000-02-29", CORE_PARSE_OK, 2000, 2, 29
    ExpectDate C, "1900-01-01", CORE_PARSE_OK, 1900, 1, 1
    ExpectDate C, "9999-12-31", CORE_PARSE_OK, 9999, 12, 31
    ExpectDate C, "", CORE_PARSE_MISSING, 0, 0, 0
    ExpectDate C, "2025-02-29", CORE_PARSE_INVALID, 0, 0, 0
    ExpectDate C, "1900-02-29", CORE_PARSE_INVALID, 0, 0, 0
    ExpectDate C, "2025-04-31", CORE_PARSE_INVALID, 0, 0, 0
    ExpectDate C, "2025-13-01", CORE_PARSE_INVALID, 0, 0, 0
    ExpectDate C, "2025-00-10", CORE_PARSE_INVALID, 0, 0, 0
    ExpectDate C, "2025-01-00", CORE_PARSE_INVALID, 0, 0, 0
    ExpectDate C, "2025-1-31", CORE_PARSE_INVALID, 0, 0, 0
    ExpectDate C, "31/01/2025", CORE_PARSE_INVALID, 0, 0, 0
    ExpectDate C, "2025/01/31", CORE_PARSE_INVALID, 0, 0, 0
    ExpectDate C, "20250131", CORE_PARSE_INVALID, 0, 0, 0
    ExpectDate C, " 2025-01-31", CORE_PARSE_INVALID, 0, 0, 0
    ExpectDate C, "2025-01-31 ", CORE_PARSE_INVALID, 0, 0, 0
    ExpectDate C, "2025-01-31T00:00", CORE_PARSE_INVALID, 0, 0, 0
    ExpectDate C, "1899-12-31", CORE_PARSE_INVALID, 0, 0, 0

    ExpectMonthEnd C, 2025, 1, 31, True
    ExpectMonthEnd C, 2024, 2, 29, True
    ExpectMonthEnd C, 2025, 2, 28, True
    ExpectMonthEnd C, 2024, 2, 28, False
    ExpectMonthEnd C, 2025, 4, 30, True
    ExpectMonthEnd C, 2025, 6, 15, False
    ExpectMonthEnd C, 2025, 12, 31, True
End Sub


'------------------------------------------------------------------------------
' DATA-PARSE-02: fixed-point decimals (balances: 4 places) and rates (8 places,
' -0.05 to 0.25), missing versus invalid, locale-like forms rejected.
'------------------------------------------------------------------------------
Private Sub CaseDecimals()
    Const C As String = "DATA-PARSE-02"
    ExpectDecimal C, "1200.50", 4, CORE_PARSE_OK, 1200.5
    ExpectDecimal C, "0", 4, CORE_PARSE_OK, 0#
    ExpectDecimal C, "-0.0001", 4, CORE_PARSE_OK, -0.0001
    ExpectDecimal C, "007.25", 4, CORE_PARSE_OK, 7.25
    ExpectDecimal C, "1234567890.1234", 4, CORE_PARSE_OK, 1234567890.1234
    ExpectDecimal C, "", 4, CORE_PARSE_MISSING, 0#
    ExpectDecimal C, "1.23456", 4, CORE_PARSE_INVALID, 0#
    ExpectDecimal C, "1,5", 4, CORE_PARSE_INVALID, 0#
    ExpectDecimal C, "1.000,50", 4, CORE_PARSE_INVALID, 0#
    ExpectDecimal C, "1e-3", 4, CORE_PARSE_INVALID, 0#
    ExpectDecimal C, "+1.5", 4, CORE_PARSE_INVALID, 0#
    ExpectDecimal C, ".5", 4, CORE_PARSE_INVALID, 0#
    ExpectDecimal C, "5.", 4, CORE_PARSE_INVALID, 0#
    ExpectDecimal C, "-", 4, CORE_PARSE_INVALID, 0#
    ExpectDecimal C, " 1.5", 4, CORE_PARSE_INVALID, 0#
    ExpectDecimal C, "1.5 ", 4, CORE_PARSE_INVALID, 0#
    ExpectDecimal C, "1..5", 4, CORE_PARSE_INVALID, 0#

    ExpectRate C, "0.0125", CORE_PARSE_OK, 0.0125
    ExpectRate C, "-0.05", CORE_PARSE_OK, -0.05
    ExpectRate C, "0.25", CORE_PARSE_OK, 0.25
    ExpectRate C, "0.00000001", CORE_PARSE_OK, 0.00000001
    ExpectRate C, "", CORE_PARSE_MISSING, 0#
    ExpectRate C, "0.25000001", CORE_PARSE_INVALID, 0#
    ExpectRate C, "-0.05000001", CORE_PARSE_INVALID, 0#
    ExpectRate C, "0.000000001", CORE_PARSE_INVALID, 0#
    ExpectRate C, "1.25", CORE_PARSE_INVALID, 0#
    ExpectRate C, "0,0125", CORE_PARSE_INVALID, 0#
End Sub


'------------------------------------------------------------------------------
' DATA-CODES-01: segment codes and the currency allowlist reject other input.
'------------------------------------------------------------------------------
Private Sub CaseCodes()
    Const C As String = "DATA-CODES-01"
    ExpectSegment C, "RET_TX", True
    ExpectSegment C, "RET_NTX", True
    ExpectSegment C, "WHS_NFC", True
    ExpectSegment C, "ret_tx", False
    ExpectSegment C, "RET_TX ", False
    ExpectSegment C, "", False
    ExpectSegment C, "WHS_FIN", False

    ExpectCurrency C, "EUR", True
    ExpectCurrency C, "USD", True
    ExpectCurrency C, "GBP", False
    ExpectCurrency C, "eur", False
    ExpectCurrency C, "EU", False
    ExpectCurrency C, "EURO", False
    ExpectCurrency C, "EUR,", False
    ExpectCurrency C, "", False
End Sub


'==============================================================================
' Expectation helpers: one assertion each, labelled with the input.
'==============================================================================
Private Sub ExpectOutcome(ByVal caseId As String, ByVal built As Boolean, ByVal errored As Boolean, _
                          ByVal asserts As Long, ByVal passed As Long, ByVal expected As String)
    Dim actual As String
    actual = TEST_CaseOutcome(built, errored, asserts, passed)
    TEST_AssertText caseId, "outcome(built=" & built & ", error=" & errored & ", " & passed & "/" & asserts & ")", _
                    actual, expected
End Sub


Private Sub ExpectWithin(ByVal caseId As String, ByVal expected As Double, ByVal actual As Double, _
                         ByVal kind As String, ByVal eps As Double, ByVal want As Boolean)
    TEST_AssertText caseId, "within(" & actual & " vs " & expected & ", " & kind & " " & eps & ")", _
                    CStr(TEST_WithinTolerance(actual, expected, kind, eps)), CStr(want)
End Sub


Private Sub ExpectDate(ByVal caseId As String, ByVal text As String, ByVal status As Long, _
                       ByVal y As Long, ByVal m As Long, ByVal d As Long)
    Dim value As Date, got As Long
    got = CORE_ParseIsoDate(text, value)
    TEST_AssertText caseId, "date status """ & text & """", CStr(got), CStr(status)
    If status = CORE_PARSE_OK And got = CORE_PARSE_OK Then
        TEST_AssertTrue caseId, "date value """ & text & """", value = DateSerial(y, m, d)
    End If
End Sub


Private Sub ExpectMonthEnd(ByVal caseId As String, ByVal y As Long, ByVal m As Long, ByVal d As Long, _
                           ByVal want As Boolean)
    TEST_AssertText caseId, "month end " & y & "-" & m & "-" & d, CStr(CORE_IsMonthEnd(DateSerial(y, m, d))), _
                    CStr(want)
End Sub


Private Sub ExpectDecimal(ByVal caseId As String, ByVal text As String, ByVal places As Long, _
                          ByVal status As Long, ByVal expected As Double)
    Dim value As Double, got As Long
    got = CORE_ParseDecimal(text, places, value)
    TEST_AssertText caseId, "decimal(" & places & ") status """ & text & """", CStr(got), CStr(status)
    If status = CORE_PARSE_OK And got = CORE_PARSE_OK Then
        TEST_AssertNear caseId, "decimal value """ & text & """", value, expected, "rel", 0.000000000000001
    End If
End Sub


Private Sub ExpectRate(ByVal caseId As String, ByVal text As String, ByVal status As Long, ByVal expected As Double)
    Dim value As Double, got As Long
    got = CORE_ParseRate(text, value)
    TEST_AssertText caseId, "rate status """ & text & """", CStr(got), CStr(status)
    If status = CORE_PARSE_OK And got = CORE_PARSE_OK Then
        TEST_AssertNear caseId, "rate value """ & text & """", value, expected, "exact", 0#
    End If
End Sub


Private Sub ExpectSegment(ByVal caseId As String, ByVal code As String, ByVal want As Boolean)
    TEST_AssertText caseId, "segment """ & code & """", CStr(CORE_IsSegmentCode(code)), CStr(want)
End Sub


Private Sub ExpectCurrency(ByVal caseId As String, ByVal code As String, ByVal want As Boolean)
    TEST_AssertText caseId, "currency """ & code & """", CStr(CORE_IsCurrencyAllowed(code)), CStr(want)
End Sub
