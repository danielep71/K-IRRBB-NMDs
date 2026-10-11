Attribute VB_Name = "CORE_Import"
'==============================================================================
' CORE_Import
'------------------------------------------------------------------------------
' PURPOSE
'   Validation and summary of the account panel and the market-rate file of
'   docs/methodology/DATA_CONTRACT.md, line by line, with no file access and no
'   Excel object model: the caller streams the lines (issue #71, stage 1).
'
' ACCOUNT PANEL
'   CORE_AccountsBegin header      False (finding 1, E01) if the header differs;
'                                  the import then stops.
'   CORE_AccountsAddLine text      once per data line, in file order; one
'                                  trailing CR is removed (LF or CRLF files).
'   CORE_AccountsEnd               applies the cross-row rules; returns the
'                                  number of findings.
'   CORE_AccountsSummarize a, b    data window a..b (month ends): warnings
'                                  W01-W06, totals per date, segment and
'                                  currency, and missing extracts (E12);
'                                  returns the number of E12 findings.
'   Results are read through the CORE_Account*, CORE_Warning*, CORE_Total*
'   and CORE_MissingExtract* accessors (1-based, in the contract's order).
'
' MARKET RATES
'   CORE_RatesBegin, CORE_RatesAddLine, CORE_RatesEnd and CORE_RateFinding*,
'   as above.
'
' RULES
'   Each line reports its first failing rule in the contract's order: E01
'   (field count, or a character outside ASCII), E03, E02, E04, E06, E07, E09,
'   E11, E10; then, across the locally valid rows of the file, E11 (known
'   opening or closure dates disagree), E10, E05 and E08. Findings are sorted
'   by line. The import is all-or-nothing: the caller keeps no dataset when
'   any finding exists. Warnings never stop the import.
'
' TOTALS
'   Balances are summed with compensated (Neumaier) summation, in currency
'   units, never across currencies. A total above 1E+307 raises error 6
'   rather than produce an infinite result.
'
' ERRORS
'   Calls out of order, or a window that is not two month ends in order,
'   raise error 5 with the procedure as Source.
'==============================================================================
Option Explicit
Option Private Module

Private Const ACCOUNT_HEADER As String = _
    "as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date"
Private Const RATE_HEADER As String = "as_of_date,rate_id,currency,tenor_months,rate"
Private Const BALANCE_DECIMALS As Long = 4
Private Const OUTLIER_RATIO As Double = 10
Private Const MAX_TOTAL As Double = 1E+307

Private Const STATE_NONE As Long = 0
Private Const STATE_READING As Long = 1
Private Const STATE_ENDED As Long = 2
Private Const STATE_REJECTED As Long = 3

' Account panel state.
Private mState As Long
Private mLine As Long
' One entry per locally valid row, in file order. Parallel arrays rather than a
' user-defined type: LibreOffice, the early-warning proxy, does not copy
' user-defined types into arrays reliably.
Private mRowLine() As Long
Private mRowAsOf() As Date
Private mRowAsOfText() As String
Private mRowAccount() As String
Private mRowSegment() As String
Private mRowCurrency() As String
Private mRowBalance() As Double
Private mRowOpen() As String
Private mRowClose() As String
Private mRowAccepted() As Boolean
Private mRowCount As Long
Private mFindLine() As Long
Private mFindCode() As String
Private mFindCount As Long
Private mFindOrder() As Long
Private mAccepted As Long
Private mAccounts As Long

' Summary state.
Private mSummarized As Boolean
Private mWarnCode() As String
Private mWarnAccount() As String
Private mWarnDate() As String
Private mWarnCount As Long
Private mWarnOrder() As Long
Private mTotDate() As String
Private mTotSegment() As String
Private mTotCurrency() As String
Private mTotAccounts() As Long
Private mTotSum() As Double
Private mTotComp() As Double
Private mTotCount As Long
Private mTotOrder() As Long
Private mGaps() As String
Private mGapCount As Long

' Market rate state.
Private mRateState As Long
Private mRateLine As Long
Private mRateFindLine() As Long
Private mRateFindCode() As String
Private mRateFindCount As Long
Private mRateSeen As Collection
Private mRateSeries As Collection


'==============================================================================
' Account panel
'==============================================================================

'------------------------------------------------------------------------------
' CORE_AccountsBegin: start a new account import. Returns False, with the
' single finding (1, E01), when header is not the contract header exactly.
'------------------------------------------------------------------------------
Public Function CORE_AccountsBegin(ByVal header As String) As Boolean
    mRowCount = 0
    mFindCount = 0
    mAccepted = 0
    mAccounts = 0
    mSummarized = False
    GrowRows 64
    ReDim mFindLine(1 To 16)
    ReDim mFindCode(1 To 16)
    mLine = 1
    If StripCr(header) <> ACCOUNT_HEADER Then
        AddFinding 1, "E01"
        ReDim mFindOrder(1 To 1)
        mFindOrder(1) = 1
        mState = STATE_REJECTED
        Exit Function
    End If
    mState = STATE_READING
    CORE_AccountsBegin = True
End Function


'------------------------------------------------------------------------------
' CORE_AccountsAddLine: validate one data line (the next line number) on its
' own and keep it for the cross-row rules when it passes.
'------------------------------------------------------------------------------
Public Sub CORE_AccountsAddLine(ByVal text As String)
    Dim fields() As String, code As String, asOf As Date, balance As Double
    RequireState mState, STATE_READING, "CORE_Import.CORE_AccountsAddLine"
    mLine = mLine + 1
    text = StripCr(text)
    If Not IsAscii(text) Then
        AddFinding mLine, "E01"
        Exit Sub
    End If
    fields = Split(text, ",")
    If UBound(fields) - LBound(fields) + 1 <> 9 Then
        AddFinding mLine, "E01"
        Exit Sub
    End If
    code = AccountRowError(fields, asOf, balance)
    If Len(code) > 0 Then
        AddFinding mLine, code
        Exit Sub
    End If
    mRowCount = mRowCount + 1
    If mRowCount > UBound(mRowLine) Then GrowRows 2 * UBound(mRowLine)
    mRowLine(mRowCount) = mLine
    mRowAsOf(mRowCount) = asOf
    mRowAsOfText(mRowCount) = fields(0)
    mRowAccount(mRowCount) = fields(1)
    mRowSegment(mRowCount) = fields(2)
    mRowCurrency(mRowCount) = fields(3)
    mRowBalance(mRowCount) = balance
    mRowOpen(mRowCount) = fields(7)
    mRowClose(mRowCount) = fields(8)
    mRowAccepted(mRowCount) = False
End Sub


'------------------------------------------------------------------------------
' CORE_AccountsEnd: apply the cross-row rules (E11, E10, E05, E08) to the
' locally valid rows in file order and sort the findings by line. Returns the
' number of findings; zero means the panel is accepted.
'------------------------------------------------------------------------------
Public Function CORE_AccountsEnd() As Long
    Dim accounts As New Collection, seen As New Collection
    Dim accOpen() As String, accClose() As String, accCurrency() As String
    Dim accOpenMulti() As Boolean, accCloseMulti() As Boolean
    Dim i As Long, a As Long, code As String, key As String, keys() As String

    RequireState mState, STATE_READING, "CORE_Import.CORE_AccountsEnd"
    ReDim accOpen(1 To mRowCount + 1)
    ReDim accClose(1 To mRowCount + 1)
    ReDim accCurrency(1 To mRowCount + 1)
    ReDim accOpenMulti(1 To mRowCount + 1)
    ReDim accCloseMulti(1 To mRowCount + 1)

    ' Known opening and closure dates of every account, from all its rows.
    For i = 1 To mRowCount
        a = IndexOf(accounts, CaseKey(mRowAccount(i)), mAccounts)
        If Len(mRowOpen(i)) > 0 Then
            If Len(accOpen(a)) = 0 Then
                accOpen(a) = mRowOpen(i)
            ElseIf accOpen(a) <> mRowOpen(i) Then
                accOpenMulti(a) = True
            End If
        End If
        If Len(mRowClose(i)) > 0 Then
            If Len(accClose(a)) = 0 Then
                accClose(a) = mRowClose(i)
            ElseIf accClose(a) <> mRowClose(i) Then
                accCloseMulti(a) = True
            End If
        End If
    Next i

    ' Cross-row rules in file order; a row counts as seen even when it fails.
    For i = 1 To mRowCount
        a = accounts.Item(CaseKey(mRowAccount(i)))
        If Len(accCurrency(a)) = 0 Then accCurrency(a) = mRowCurrency(i)
        key = CaseKey(mRowAsOfText(i) & "|" & mRowAccount(i))
        code = vbNullString
        If accOpenMulti(a) Or accCloseMulti(a) Then
            code = "E11"
        ElseIf Len(accOpen(a)) > 0 And (accOpen(a) > mRowAsOfText(i) _
                Or (Len(accClose(a)) > 0 And accClose(a) < accOpen(a))) Then
            code = "E11"
        ElseIf Len(accClose(a)) > 0 And mRowAsOfText(i) >= accClose(a) Then
            code = "E10"
        ElseIf HasKey(seen, key) Then
            code = "E05"
        ElseIf accCurrency(a) <> mRowCurrency(i) Then
            code = "E08"
        End If
        If Not HasKey(seen, key) Then seen.Add i, key
        If Len(code) > 0 Then
            AddFinding mRowLine(i), code
        Else
            mRowAccepted(i) = True
            mAccepted = mAccepted + 1
        End If
    Next i

    ' Accounts with at least one accepted row.
    Set accounts = New Collection
    mAccounts = 0
    For i = 1 To mRowCount
        If mRowAccepted(i) Then a = IndexOf(accounts, CaseKey(mRowAccount(i)), mAccounts)
    Next i

    ReDim keys(1 To mFindCount + 1)
    For i = 1 To mFindCount
        keys(i) = Format$(mFindLine(i), "0000000000")
    Next i
    SortOrder keys, mFindCount, mFindOrder
    mState = STATE_ENDED
    CORE_AccountsEnd = mFindCount
End Function


Public Function CORE_AccountFindingCount() As Long
    CORE_AccountFindingCount = mFindCount
End Function


Public Function CORE_AccountFindingLine(ByVal index As Long) As Long
    CORE_AccountFindingLine = mFindLine(mFindOrder(index))
End Function


Public Function CORE_AccountFindingCode(ByVal index As Long) As String
    CORE_AccountFindingCode = mFindCode(mFindOrder(index))
End Function


'------------------------------------------------------------------------------
' CORE_AcceptedRowCount / CORE_AccountCount: rows that passed every rule and
' the distinct accounts among them (after CORE_AccountsEnd).
'------------------------------------------------------------------------------
Public Function CORE_AcceptedRowCount() As Long
    CORE_AcceptedRowCount = mAccepted
End Function


Public Function CORE_AccountCount() As Long
    CORE_AccountCount = mAccounts
End Function


'------------------------------------------------------------------------------
' CORE_AccountsSummarize: warnings, totals and missing extracts of the accepted
' rows for the data window firstMonthEnd..lastMonthEnd (both month ends).
' Rows outside the window are excluded and reported (W06). Returns the number
' of month ends in the window with no accepted row (E12); any rejects the
' import.
'------------------------------------------------------------------------------
Public Function CORE_AccountsSummarize(ByVal firstMonthEnd As Date, ByVal lastMonthEnd As Date) As Long
    Dim present As New Collection, totals As New Collection
    Dim inWindow() As Long, keys() As String, order() As Long, n As Long
    Dim i As Long, j As Long, k As Long, t As Long, gap As Long
    Dim previous As Long, closeText As String, ratio As Double, monthEndText As String

    RequireState mState, STATE_ENDED, "CORE_Import.CORE_AccountsSummarize"
    If Not CORE_IsMonthEnd(firstMonthEnd) Or Not CORE_IsMonthEnd(lastMonthEnd) _
            Or firstMonthEnd > lastMonthEnd Then
        Err.Raise 5, "CORE_Import.CORE_AccountsSummarize", "The data window must be two month ends in order."
    End If

    mWarnCount = 0
    mTotCount = 0
    mGapCount = 0
    ReDim mWarnCode(1 To 16)
    ReDim mWarnAccount(1 To 16)
    ReDim mWarnDate(1 To 16)
    ReDim mTotDate(1 To 16)
    ReDim mTotSegment(1 To 16)
    ReDim mTotCurrency(1 To 16)
    ReDim mTotAccounts(1 To 16)
    ReDim mTotSum(1 To 16)
    ReDim mTotComp(1 To 16)
    ReDim mGaps(1 To 16)
    ReDim inWindow(1 To mRowCount + 1)

    For i = 1 To mRowCount
        If mRowAccepted(i) Then
            If Not HasKey(present, mRowAsOfText(i)) Then present.Add i, mRowAsOfText(i)
            If mRowAsOf(i) < firstMonthEnd Or mRowAsOf(i) > lastMonthEnd Then
                AddWarning "W06", mRowAccount(i), mRowAsOfText(i)
            Else
                n = n + 1
                inWindow(n) = i
            End If
        End If
    Next i

    ' Rows of each account in date order.
    ReDim keys(1 To n + 1)
    For j = 1 To n
        keys(j) = mRowAccount(inWindow(j)) & vbTab & mRowAsOfText(inWindow(j))
    Next j
    SortOrder keys, n, order

    For j = 1 To n
        i = inWindow(order(j))
        previous = 0
        If j > 1 Then
            If mRowAccount(inWindow(order(j - 1))) = mRowAccount(i) Then previous = inWindow(order(j - 1))
        End If
        If previous = 0 Then closeText = vbNullString

        If Len(mRowClose(i)) > 0 Then
            If Len(closeText) = 0 Or mRowClose(i) < closeText Then closeText = mRowClose(i)
        End If
        t = IndexOfTotal(totals, mRowAsOfText(i), mRowSegment(i), mRowCurrency(i))
        mTotAccounts(t) = mTotAccounts(t) + 1
        AddCompensated t, mRowBalance(i)
        If mRowBalance(i) = 0 Then AddWarning "W02", mRowAccount(i), mRowAsOfText(i)
        If previous > 0 Then
            gap = MonthIndex(mRowAsOf(i)) - MonthIndex(mRowAsOf(previous))
            For k = 1 To gap - 1
                AddWarning "W01", mRowAccount(i), IsoText(MonthEnd(MonthIndex(mRowAsOf(previous)) + k))
            Next k
            If mRowSegment(i) <> mRowSegment(previous) Then AddWarning "W03", mRowAccount(i), mRowAsOfText(i)
            If gap = 1 And mRowBalance(previous) > 0 And mRowBalance(i) > 0 Then
                ratio = mRowBalance(i) / mRowBalance(previous)
                If ratio < 1 / OUTLIER_RATIO Or ratio > OUTLIER_RATIO Then
                    AddWarning "W04", mRowAccount(i), mRowAsOfText(i)
                End If
            End If
        End If

        ' Last row of the account: a stop before the window end without a closure
        ' in the following month is a censored disappearance (W05).
        If j = n Then
            CheckDisappearance i, closeText, lastMonthEnd
        ElseIf mRowAccount(inWindow(order(j + 1))) <> mRowAccount(i) Then
            CheckDisappearance i, closeText, lastMonthEnd
        End If
    Next j

    ' By month index, so a window ending in December 9999 never builds a later date.
    For k = MonthIndex(firstMonthEnd) To MonthIndex(lastMonthEnd)
        monthEndText = IsoText(MonthEnd(k))
        If Not HasKey(present, monthEndText) Then AddMissing monthEndText
    Next k

    ReDim keys(1 To mWarnCount + 1)
    For j = 1 To mWarnCount
        keys(j) = mWarnCode(j) & vbTab & mWarnAccount(j) & vbTab & mWarnDate(j)
    Next j
    SortOrder keys, mWarnCount, mWarnOrder
    ReDim keys(1 To mTotCount + 1)
    For j = 1 To mTotCount
        keys(j) = mTotDate(j) & vbTab & mTotSegment(j) & vbTab & mTotCurrency(j)
    Next j
    SortOrder keys, mTotCount, mTotOrder
    mSummarized = True
    CORE_AccountsSummarize = mGapCount
End Function


Public Function CORE_WarningCount() As Long
    CORE_WarningCount = mWarnCount
End Function


Public Function CORE_WarningCode(ByVal index As Long) As String
    CORE_WarningCode = mWarnCode(mWarnOrder(index))
End Function


Public Function CORE_WarningAccount(ByVal index As Long) As String
    CORE_WarningAccount = mWarnAccount(mWarnOrder(index))
End Function


Public Function CORE_WarningDate(ByVal index As Long) As String
    CORE_WarningDate = mWarnDate(mWarnOrder(index))
End Function


Public Function CORE_TotalCount() As Long
    CORE_TotalCount = mTotCount
End Function


Public Function CORE_TotalDate(ByVal index As Long) As String
    CORE_TotalDate = mTotDate(mTotOrder(index))
End Function


Public Function CORE_TotalSegment(ByVal index As Long) As String
    CORE_TotalSegment = mTotSegment(mTotOrder(index))
End Function


Public Function CORE_TotalCurrency(ByVal index As Long) As String
    CORE_TotalCurrency = mTotCurrency(mTotOrder(index))
End Function


Public Function CORE_TotalAccounts(ByVal index As Long) As Long
    CORE_TotalAccounts = mTotAccounts(mTotOrder(index))
End Function


'------------------------------------------------------------------------------
' CORE_TotalBalance: total balance in currency units, at input precision
' (compensated sum); round only for presentation.
'------------------------------------------------------------------------------
Public Function CORE_TotalBalance(ByVal index As Long) As Double
    CORE_TotalBalance = mTotSum(mTotOrder(index)) + mTotComp(mTotOrder(index))
End Function


Public Function CORE_MissingExtractCount() As Long
    CORE_MissingExtractCount = mGapCount
End Function


Public Function CORE_MissingExtract(ByVal index As Long) As String
    CORE_MissingExtract = mGaps(index)
End Function


'==============================================================================
' Market rates
'==============================================================================

'------------------------------------------------------------------------------
' CORE_RatesBegin: start a new market-rate import. Returns False, with the
' single finding (1, E01), when header is not the contract header exactly.
'------------------------------------------------------------------------------
Public Function CORE_RatesBegin(ByVal header As String) As Boolean
    mRateFindCount = 0
    ReDim mRateFindLine(1 To 16)
    ReDim mRateFindCode(1 To 16)
    Set mRateSeen = New Collection
    Set mRateSeries = New Collection
    mRateLine = 1
    If StripCr(header) <> RATE_HEADER Then
        AddRateFinding 1, "E01"
        mRateState = STATE_REJECTED
        Exit Function
    End If
    mRateState = STATE_READING
    CORE_RatesBegin = True
End Function


'------------------------------------------------------------------------------
' CORE_RatesAddLine: validate one market-rate line: E01, E03, E02, E04, E07,
' then E05 (repeated date and series) and E08 (series currency or tenor
' differs from its earlier accepted rows).
'------------------------------------------------------------------------------
Public Sub CORE_RatesAddLine(ByVal text As String)
    Dim fields() As String, asOf As Date, rate As Double, i As Long
    Dim code As String, key As String, seriesKey As String, properties As String, missing As Boolean

    RequireState mRateState, STATE_READING, "CORE_Import.CORE_RatesAddLine"
    mRateLine = mRateLine + 1
    text = StripCr(text)
    If Not IsAscii(text) Then
        AddRateFinding mRateLine, "E01"
        Exit Sub
    End If
    fields = Split(text, ",")
    If UBound(fields) - LBound(fields) + 1 <> 5 Then
        AddRateFinding mRateLine, "E01"
        Exit Sub
    End If

    key = CaseKey(fields(0) & "|" & fields(1))
    seriesKey = CaseKey(fields(1))
    properties = fields(2) & "|" & fields(3)
    For i = 0 To 4
        If Len(fields(i)) = 0 Then missing = True
    Next i
    If missing Then
        code = "E03"
    ElseIf CORE_ParseIsoDate(fields(0), asOf) <> CORE_PARSE_OK Then
        code = "E02"
    ElseIf Not CORE_IsMonthEnd(asOf) Then
        code = "E02"
    ElseIf Not CORE_IsIdentifier(fields(1)) Or Not IsPositiveInteger(fields(3)) _
            Or CORE_ParseRate(fields(4), rate) <> CORE_PARSE_OK Then
        code = "E04"
    ElseIf Not CORE_IsCurrencyAllowed(fields(2)) Then
        code = "E07"
    ElseIf HasKey(mRateSeen, key) Then
        code = "E05"
    ElseIf HasKey(mRateSeries, seriesKey) Then
        If mRateSeries.Item(seriesKey) <> properties Then code = "E08"
    End If

    If Len(code) > 0 Then
        AddRateFinding mRateLine, code
    Else
        mRateSeen.Add mRateLine, key
        If Not HasKey(mRateSeries, seriesKey) Then mRateSeries.Add properties, seriesKey
    End If
End Sub


Public Function CORE_RatesEnd() As Long
    RequireState mRateState, STATE_READING, "CORE_Import.CORE_RatesEnd"
    mRateState = STATE_ENDED
    CORE_RatesEnd = mRateFindCount
End Function


Public Function CORE_RateFindingCount() As Long
    CORE_RateFindingCount = mRateFindCount
End Function


Public Function CORE_RateFindingLine(ByVal index As Long) As Long
    CORE_RateFindingLine = mRateFindLine(index)
End Function


Public Function CORE_RateFindingCode(ByVal index As Long) As String
    CORE_RateFindingCode = mRateFindCode(index)
End Function


'==============================================================================
' Row rules
'==============================================================================

' First failing single-row rule of an account line, or "" with the parsed
' as-of date and balance.
Private Function AccountRowError(ByRef fields() As String, ByRef asOf As Date, ByRef balance As Double) As String
    Dim opened As Date, closed As Date, rate As Double
    Dim hasOpen As Boolean, hasClose As Boolean

    If Len(fields(0)) = 0 Or Len(fields(1)) = 0 Or Len(fields(2)) = 0 Or Len(fields(3)) = 0 _
            Or Len(fields(4)) = 0 Or Len(fields(6)) = 0 Then
        AccountRowError = "E03"
        Exit Function
    End If

    If CORE_ParseIsoDate(fields(0), asOf) <> CORE_PARSE_OK Then
        AccountRowError = "E02"
        Exit Function
    End If
    If Not CORE_IsMonthEnd(asOf) Then
        AccountRowError = "E02"
        Exit Function
    End If
    If Len(fields(7)) > 0 Then
        If CORE_ParseIsoDate(fields(7), opened) <> CORE_PARSE_OK Then
            AccountRowError = "E02"
            Exit Function
        End If
        hasOpen = True
    End If
    If Len(fields(8)) > 0 Then
        If CORE_ParseIsoDate(fields(8), closed) <> CORE_PARSE_OK Then
            AccountRowError = "E02"
            Exit Function
        End If
        hasClose = True
    End If

    If Not CORE_IsIdentifier(fields(1)) Then
        AccountRowError = "E04"
    ElseIf CORE_ParseDecimal(fields(4), BALANCE_DECIMALS, balance) <> CORE_PARSE_OK Then
        AccountRowError = "E04"
    ElseIf Len(fields(5)) > 0 And CORE_ParseRate(fields(5), rate) <> CORE_PARSE_OK Then
        AccountRowError = "E04"
    ElseIf fields(6) <> "0" And fields(6) <> "1" Then
        AccountRowError = "E04"
    ElseIf Not CORE_IsSegmentCode(fields(2)) Then
        AccountRowError = "E06"
    ElseIf Not CORE_IsCurrencyAllowed(fields(3)) Then
        AccountRowError = "E07"
    ElseIf balance < 0 Then
        AccountRowError = "E09"
    ElseIf hasOpen And (opened > asOf Or (hasClose And closed < opened)) Then
        AccountRowError = "E11"
    ElseIf hasClose And asOf >= closed Then
        AccountRowError = "E10"
    End If
End Function


' W05 for the account whose last in-window row is row i.
Private Sub CheckDisappearance(ByVal i As Long, ByVal closeText As String, ByVal lastMonthEnd As Date)
    Dim closed As Date, verified As Boolean
    If mRowAsOf(i) >= lastMonthEnd Then Exit Sub
    If Len(closeText) > 0 Then
        If CORE_ParseIsoDate(closeText, closed) = CORE_PARSE_OK Then
            verified = (mRowAsOf(i) < closed And closed <= MonthEnd(MonthIndex(mRowAsOf(i)) + 1))
        End If
    End If
    If Not verified Then AddWarning "W05", mRowAccount(i), mRowAsOfText(i)
End Sub


'==============================================================================
' Helpers
'==============================================================================
Private Sub GrowRows(ByVal size As Long)
    If mRowCount = 0 And size <= 64 Then
        ReDim mRowLine(1 To size)
        ReDim mRowAsOf(1 To size)
        ReDim mRowAsOfText(1 To size)
        ReDim mRowAccount(1 To size)
        ReDim mRowSegment(1 To size)
        ReDim mRowCurrency(1 To size)
        ReDim mRowBalance(1 To size)
        ReDim mRowOpen(1 To size)
        ReDim mRowClose(1 To size)
        ReDim mRowAccepted(1 To size)
        Exit Sub
    End If
    ReDim Preserve mRowLine(1 To size)
    ReDim Preserve mRowAsOf(1 To size)
    ReDim Preserve mRowAsOfText(1 To size)
    ReDim Preserve mRowAccount(1 To size)
    ReDim Preserve mRowSegment(1 To size)
    ReDim Preserve mRowCurrency(1 To size)
    ReDim Preserve mRowBalance(1 To size)
    ReDim Preserve mRowOpen(1 To size)
    ReDim Preserve mRowClose(1 To size)
    ReDim Preserve mRowAccepted(1 To size)
End Sub


Private Sub RequireState(ByVal actual As Long, ByVal expected As Long, ByVal source As String)
    If actual <> expected Then Err.Raise 5, source, "Import procedures were called out of order."
End Sub


Private Sub AddFinding(ByVal lineNo As Long, ByVal code As String)
    mFindCount = mFindCount + 1
    If mFindCount > UBound(mFindLine) Then
        ReDim Preserve mFindLine(1 To 2 * UBound(mFindLine))
        ReDim Preserve mFindCode(1 To 2 * UBound(mFindCode))
    End If
    mFindLine(mFindCount) = lineNo
    mFindCode(mFindCount) = code
End Sub


Private Sub AddRateFinding(ByVal lineNo As Long, ByVal code As String)
    mRateFindCount = mRateFindCount + 1
    If mRateFindCount > UBound(mRateFindLine) Then
        ReDim Preserve mRateFindLine(1 To 2 * UBound(mRateFindLine))
        ReDim Preserve mRateFindCode(1 To 2 * UBound(mRateFindCode))
    End If
    mRateFindLine(mRateFindCount) = lineNo
    mRateFindCode(mRateFindCount) = code
End Sub


Private Sub AddWarning(ByVal code As String, ByVal accountId As String, ByVal asOfText As String)
    mWarnCount = mWarnCount + 1
    If mWarnCount > UBound(mWarnCode) Then
        ReDim Preserve mWarnCode(1 To 2 * UBound(mWarnCode))
        ReDim Preserve mWarnAccount(1 To 2 * UBound(mWarnAccount))
        ReDim Preserve mWarnDate(1 To 2 * UBound(mWarnDate))
    End If
    mWarnCode(mWarnCount) = code
    mWarnAccount(mWarnCount) = accountId
    mWarnDate(mWarnCount) = asOfText
End Sub


Private Sub AddMissing(ByVal monthEndText As String)
    mGapCount = mGapCount + 1
    If mGapCount > UBound(mGaps) Then
        ReDim Preserve mGaps(1 To 2 * UBound(mGaps))
    End If
    mGaps(mGapCount) = monthEndText
End Sub


' Index of the (date, segment, currency) total, created at zero when new.
Private Function IndexOfTotal(ByVal totals As Collection, ByVal asOfText As String, ByVal segment As String, _
                              ByVal currency As String) As Long
    Dim key As String
    key = asOfText & "|" & segment & "|" & currency
    If HasKey(totals, key) Then
        IndexOfTotal = totals.Item(key)
        Exit Function
    End If
    mTotCount = mTotCount + 1
    If mTotCount > UBound(mTotDate) Then
        ReDim Preserve mTotDate(1 To 2 * UBound(mTotDate))
        ReDim Preserve mTotSegment(1 To 2 * UBound(mTotSegment))
        ReDim Preserve mTotCurrency(1 To 2 * UBound(mTotCurrency))
        ReDim Preserve mTotAccounts(1 To 2 * UBound(mTotAccounts))
        ReDim Preserve mTotSum(1 To 2 * UBound(mTotSum))
        ReDim Preserve mTotComp(1 To 2 * UBound(mTotComp))
    End If
    mTotDate(mTotCount) = asOfText
    mTotSegment(mTotCount) = segment
    mTotCurrency(mTotCount) = currency
    totals.Add mTotCount, key
    IndexOfTotal = mTotCount
End Function


' Neumaier compensated summation into total t; a total above MAX_TOTAL raises
' error 6 rather than produce an infinite result.
Private Sub AddCompensated(ByVal t As Long, ByVal x As Double)
    Dim s As Double, total As Double
    s = mTotSum(t)
    If Abs(s) > MAX_TOTAL - Abs(x) Then
        Err.Raise 6, "CORE_Import.CORE_AccountsSummarize", "Total balance for " & mTotDate(t) & " " & _
                  mTotSegment(t) & " " & mTotCurrency(t) & " exceeds 1E+307."
    End If
    total = s + x
    If Abs(s) >= Abs(x) Then
        mTotComp(t) = mTotComp(t) + ((s - total) + x)
    Else
        mTotComp(t) = mTotComp(t) + ((x - total) + s)
    End If
    mTotSum(t) = total
End Sub


' Index of key in c, added with the next number (counter) when new.
Private Function IndexOf(ByVal c As Collection, ByVal key As String, ByRef counter As Long) As Long
    If HasKey(c, key) Then
        IndexOf = c.Item(key)
    Else
        counter = counter + 1
        c.Add counter, key
        IndexOf = counter
    End If
End Function


Private Function HasKey(ByVal c As Collection, ByVal key As String) As Boolean
    Dim probe As Variant
    On Error GoTo NotFound
    probe = c.Item(key)
    HasKey = True
    Exit Function
NotFound:
End Function


' Collection keys ignore case; mark each upper-case letter so that keys built
' from identifiers stay case-sensitive. "^" is not an identifier character.
Private Function CaseKey(ByVal text As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "A" And ch <= "Z" Then out = out & "^"
        out = out & ch
    Next i
    CaseKey = out
End Function


' Stable merge sort of keys(1..n) (binary comparison) into order(1..n). A Sub
' with a ByRef array rather than a function returning Long(): LibreOffice,
' the early-warning proxy, does not compile array return types.
Private Sub SortOrder(ByRef keys() As String, ByVal n As Long, ByRef order() As Long)
    Dim work() As Long, i As Long, width As Long
    Dim lo As Long, middle As Long, hi As Long, a As Long, b As Long, k As Long, takeA As Boolean
    ReDim order(1 To n + 1)
    ReDim work(1 To n + 1)
    For i = 1 To n
        order(i) = i
    Next i
    width = 1
    Do While width < n
        lo = 1
        Do While lo <= n
            middle = lo + width - 1
            hi = lo + 2 * width - 1
            If middle > n Then middle = n
            If hi > n Then hi = n
            a = lo
            b = middle + 1
            For k = lo To hi
                If a > middle Then
                    takeA = False
                ElseIf b > hi Then
                    takeA = True
                Else
                    takeA = (StrComp(keys(order(a)), keys(order(b)), vbBinaryCompare) <= 0)
                End If
                If takeA Then
                    work(k) = order(a)
                    a = a + 1
                Else
                    work(k) = order(b)
                    b = b + 1
                End If
            Next k
            lo = lo + 2 * width
        Loop
        For i = 1 To n
            order(i) = work(i)
        Next i
        width = 2 * width
    Loop
End Sub


Private Function StripCr(ByVal text As String) As String
    If Right$(text, 1) = vbCr Then text = Left$(text, Len(text) - 1)
    StripCr = text
End Function


' The file format is ASCII (DATA_CONTRACT.md); any other character is E01.
Private Function IsAscii(ByVal text As String) As Boolean
    Dim i As Long
    For i = 1 To Len(text)
        If AscW(Mid$(text, i, 1)) > 127 Or AscW(Mid$(text, i, 1)) < 0 Then Exit Function
    Next i
    IsAscii = True
End Function


Private Function IsPositiveInteger(ByVal text As String) As Boolean
    Dim i As Long, ch As String
    If Len(text) = 0 Then Exit Function
    If Left$(text, 1) < "1" Or Left$(text, 1) > "9" Then Exit Function
    For i = 2 To Len(text)
        ch = Mid$(text, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsPositiveInteger = True
End Function


Private Function MonthIndex(ByVal d As Date) As Long
    MonthIndex = Year(d) * 12 + Month(d) - 1
End Function


' Last day of the month with the given index (year * 12 + month - 1). December
' is built directly: the first of the next month does not exist after 9999.
Private Function MonthEnd(ByVal index As Long) As Date
    Dim y As Long, m As Long
    y = index \ 12
    m = index Mod 12 + 1
    If m = 12 Then
        MonthEnd = DateSerial(y, 12, 31)
    Else
        MonthEnd = DateSerial(y, m + 1, 1) - 1
    End If
End Function


Private Function IsoText(ByVal d As Date) As String
    IsoText = Format$(Year(d), "0000") & "-" & Format$(Month(d), "00") & "-" & Format$(Day(d), "00")
End Function
