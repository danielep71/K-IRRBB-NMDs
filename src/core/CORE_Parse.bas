Attribute VB_Name = "CORE_Parse"
'==============================================================================
' CORE_Parse
'------------------------------------------------------------------------------
' PURPOSE
'   Locale-independent parsing of the CSV field formats in
'   docs/methodology/DATA_CONTRACT.md ("File format"): ISO dates, month ends,
'   fixed-point decimals and customer or market rates. Excel's locale-dependent
'   conversions (CDate, CDbl or CDec on text, Val, sheet auto-formatting) are
'   never used, so the same text gives the same value on every host.
'
' RESULT CODES
'   Every parser returns one of the codes below and writes its value only on
'   CORE_PARSE_OK. An empty field is *missing*, never zero or a sentinel; the
'   caller decides whether a missing field is allowed (E03) or invalid.
'     CORE_PARSE_OK       the text is valid; value written
'     CORE_PARSE_MISSING  the text is empty; value untouched
'     CORE_PARSE_INVALID  the text breaks the format or range; value untouched
'
' CONTRACT
'   - Dates: exactly "YYYY-MM-DD" with ASCII digits, a real Gregorian date and
'     a year from 1900 to 9999 (the Excel date system); no time part.
'   - Decimals: an optional "-", one or more digits, then optionally "." and
'     1 to maxDecimals digits. No "+", exponent, thousands separator, leading
'     or trailing ".", or spaces. A value too large for a Double is invalid.
'   - Rates: decimals per annum with up to 8 places, from -0.05 to 0.25
'     inclusive.
'   No Excel object model, worksheet function or UI (src/core boundary).
'==============================================================================
Option Explicit
Option Private Module

Public Const CORE_PARSE_OK As Long = 0
Public Const CORE_PARSE_MISSING As Long = 1
Public Const CORE_PARSE_INVALID As Long = 2

Public Const CORE_RATE_DECIMALS As Long = 8
Public Const CORE_RATE_MIN As Double = -0.05
Public Const CORE_RATE_MAX As Double = 0.25

Private Const MIN_YEAR As Long = 1900
Private Const MAX_YEAR As Long = 9999
Private Const MAX_MAGNITUDE As Double = 1E+300


'------------------------------------------------------------------------------
' CORE_ParseIsoDate
'   text   [in]   field text, e.g. "2025-01-31"
'   value  [out]  the date at midnight, written only on CORE_PARSE_OK
'   Returns CORE_PARSE_OK, CORE_PARSE_MISSING (empty) or CORE_PARSE_INVALID.
'------------------------------------------------------------------------------
Public Function CORE_ParseIsoDate(ByVal text As String, ByRef value As Date) As Long
    Dim y As Long, m As Long, d As Long

    If Len(text) = 0 Then
        CORE_ParseIsoDate = CORE_PARSE_MISSING
        Exit Function
    End If
    CORE_ParseIsoDate = CORE_PARSE_INVALID
    If Len(text) <> 10 Then Exit Function
    If Mid$(text, 5, 1) <> "-" Or Mid$(text, 8, 1) <> "-" Then Exit Function
    If Not AllDigits(Mid$(text, 1, 4)) Or Not AllDigits(Mid$(text, 6, 2)) Or Not AllDigits(Mid$(text, 9, 2)) Then
        Exit Function
    End If

    y = DigitsValue(Mid$(text, 1, 4))
    m = DigitsValue(Mid$(text, 6, 2))
    d = DigitsValue(Mid$(text, 9, 2))
    If y < MIN_YEAR Or y > MAX_YEAR Then Exit Function
    If m < 1 Or m > 12 Then Exit Function
    If d < 1 Or d > DaysInMonth(y, m) Then Exit Function

    value = DateSerial(y, m, d)
    CORE_ParseIsoDate = CORE_PARSE_OK
End Function


'------------------------------------------------------------------------------
' CORE_IsMonthEnd: True when d is the last calendar day of its month.
'------------------------------------------------------------------------------
Public Function CORE_IsMonthEnd(ByVal d As Date) As Boolean
    CORE_IsMonthEnd = (Day(d) = DaysInMonth(Year(d), Month(d)))
End Function


'------------------------------------------------------------------------------
' CORE_ParseDecimal
'   text         [in]   field text, e.g. "1200.50"
'   maxDecimals  [in]   largest number of digits allowed after "." (>= 1)
'   value        [out]  the value as a Double, written only on CORE_PARSE_OK
'   All digits are accumulated as one integer and divided once by an exact
'   power of ten, so up to 15 significant digits give the correctly rounded
'   Double (the value of the same literal); longer inputs are within a few
'   units in the last place. The result is the same on every host.
'------------------------------------------------------------------------------
Public Function CORE_ParseDecimal(ByVal text As String, ByVal maxDecimals As Long, ByRef value As Double) As Long
    Dim pos As Long, dot As Long, digits As Long, decimals As Long
    Dim acc As Double, factor As Double, ch As String, negative As Boolean

    If Len(text) = 0 Then
        CORE_ParseDecimal = CORE_PARSE_MISSING
        Exit Function
    End If
    CORE_ParseDecimal = CORE_PARSE_INVALID
    If maxDecimals < 1 Then Exit Function

    pos = 1
    If Left$(text, 1) = "-" Then
        negative = True
        pos = 2
    End If
    If pos > Len(text) Then Exit Function

    For pos = pos To Len(text)
        ch = Mid$(text, pos, 1)
        If ch = "." Then
            If dot > 0 Or digits = 0 Then Exit Function
            dot = pos
        ElseIf ch >= "0" And ch <= "9" Then
            digits = digits + 1
            If dot > 0 Then decimals = decimals + 1
            If decimals > maxDecimals Then Exit Function
            acc = acc * 10 + (Asc(ch) - 48)
            If acc > MAX_MAGNITUDE Then Exit Function
        Else
            Exit Function
        End If
    Next pos
    If dot > 0 And decimals = 0 Then Exit Function

    factor = 1
    Do While decimals > 0
        factor = factor * 10
        decimals = decimals - 1
    Loop
    acc = acc / factor
    If negative Then acc = -acc
    value = acc
    CORE_ParseDecimal = CORE_PARSE_OK
End Function


'------------------------------------------------------------------------------
' CORE_ParseRate: a decimal rate per annum with up to 8 places, accepted only
' within CORE_RATE_MIN to CORE_RATE_MAX inclusive.
'------------------------------------------------------------------------------
Public Function CORE_ParseRate(ByVal text As String, ByRef value As Double) As Long
    Dim parsed As Double, status As Long

    status = CORE_ParseDecimal(text, CORE_RATE_DECIMALS, parsed)
    If status = CORE_PARSE_OK Then
        If parsed < CORE_RATE_MIN Or parsed > CORE_RATE_MAX Then
            status = CORE_PARSE_INVALID
        Else
            value = parsed
        End If
    End If
    CORE_ParseRate = status
End Function


'==============================================================================
' Helpers
'==============================================================================
Private Function AllDigits(ByVal text As String) As Boolean
    Dim i As Long, ch As String
    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    AllDigits = True
End Function


Private Function DigitsValue(ByVal text As String) As Long
    Dim i As Long
    For i = 1 To Len(text)
        DigitsValue = DigitsValue * 10 + (Asc(Mid$(text, i, 1)) - 48)
    Next i
End Function


Private Function DaysInMonth(ByVal y As Long, ByVal m As Long) As Long
    Select Case m
        Case 4, 6, 9, 11
            DaysInMonth = 30
        Case 2
            If (y Mod 4 = 0 And y Mod 100 <> 0) Or y Mod 400 = 0 Then DaysInMonth = 29 Else DaysInMonth = 28
        Case Else
            DaysInMonth = 31
    End Select
End Function
