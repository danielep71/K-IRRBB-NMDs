Attribute VB_Name = "TEST_DecayCases"
'==============================================================================
' TEST_DecayCases
'------------------------------------------------------------------------------
' PURPOSE
'   VBA implementations of the decay cases in docs/methodology/TEST_CASES.md.
'   Inputs and expected values come from TEST_DecayData, generated from
'   tests/fixtures/decay and tests/expected/decay (independent references).
'   Cases may call src/core directly (docs/REPOSITORY_STRUCTURE.md) and use
'   the facade where the contract is about the public boundary.
'==============================================================================
Option Explicit
Option Private Module

Private mExpected As Variant
Private mInputs As Variant


Public Sub TEST_RunDecayCase(ByVal caseId As String)
    If IsEmpty(mExpected) Then
        mExpected = TEST_DecayExpected()
        mInputs = TEST_DecayInputs()
    End If
    Select Case caseId
        Case "DEC-LOG-01": CaseLog
        Case "DEC-MPA-01": CaseMpa
        Case "DEC-PROF-01": CaseProfile
        Case "DEC-LIFE-01": CaseLife
        Case "DEC-CAP-01": CaseCap
        Case "DEC-E2E-01": CaseEndToEnd
        Case "DEC-ERR-01": CaseErrors
        Case "OOS-FREEZE-01": CaseFreeze
        Case "OOS-LEAK-01": CaseLeak
        Case "OOS-SCORE-01": CaseScore
        Case Else
            Err.Raise TEST_ERR_NO_CASE, "TEST_DecayCases.TEST_RunDecayCase", "No VBA implementation for " & caseId & "."
    End Select
End Sub


'==============================================================================
' Cases
'==============================================================================
Private Sub CaseLog()
    Const C As String = "DEC-LOG-01"
    Dim dates() As Date, bal() As Double, g() As Double, i As Long, pos() As Long
    Dim mu As Double, sigma As Double, kappa() As Double, resid() As Double

    TEST_FixtureShort dates, bal
    ReDim g(1 To UBound(bal) - 1)
    For i = 2 To UBound(bal)
        g(i - 1) = Log(bal(i)) - Log(bal(i - 1))
        Check C, "g[" & (i - 2) & "]", g(i - 1)
    Next i
    ReDim pos(1 To 1)
    CORE_DecayEstimate g, pos, 0, mu, sigma, kappa, resid
    Check C, "nobreak_mu", mu
    Check C, "nobreak_sigma", sigma

    pos(1) = 0
    For i = 2 To UBound(dates)
        If dates(i) = InputValue(C, "break_dates[0]") Then pos(1) = i - 1
    Next i
    TEST_AssertTrue C, "break date maps to a log change", pos(1) > 0
    CORE_DecayEstimate g, pos, 1, mu, sigma, kappa, resid
    Check C, "break_mu", mu
    Check C, "break_sigma", sigma
    Check C, "break_kappa", kappa(1)
    For i = 1 To UBound(resid)
        Check C, "break_resid[" & (i - 1) & "]", resid(i)
    Next i
End Sub


Private Sub CaseMpa()
    Const C As String = "DEC-MPA-01"
    Dim z As Double, mpa() As Double, runoff() As Double, median() As Double, h As Long
    Dim nonCore As Double, life As Double, coreLife As Double
    z = CORE_NormSInv(1 - InputValue(C, "confidence"))
    Check C, "z", z
    CORE_DecayProfile InputValue(C, "b0"), InputValue(C, "mu_star"), InputValue(C, "sigma"), z, _
                      CLng(InputValue(C, "cutoff")), mpa, runoff, median, nonCore, life, coreLife
    For h = 0 To UBound(mpa)
        Check C, "mpa[" & h & "]", mpa(h)
    Next h
End Sub


Private Sub CaseProfile()
    Const C As String = "DEC-PROF-01"
    Dim mpa() As Double, runoff() As Double, median() As Double, h As Long, total As Double, mono As Boolean
    Dim nonCore As Double, life As Double, coreLife As Double
    HandProfile C, mpa, runoff, median, nonCore, life, coreLife
    mono = True
    For h = 0 To UBound(mpa)
        Check C, "mpa[" & h & "]", mpa(h)
        If h > 0 Then
            If mpa(h) > mpa(h - 1) Then mono = False
        End If
    Next h
    For h = 1 To UBound(runoff)
        Check C, "runoff[" & (h - 1) & "]", runoff(h)
        total = total + runoff(h)
    Next h
    Check C, "remainder", mpa(UBound(mpa))
    Check C, "sum", total + mpa(UBound(mpa))
    TEST_AssertTrue C, "MPA non-increasing", mono
End Sub


Private Sub CaseLife()
    Const C As String = "DEC-LIFE-01"
    Dim mpa() As Double, runoff() As Double, median() As Double, nonCore As Double, life As Double, coreLife As Double
    HandProfile C, mpa, runoff, median, nonCore, life, coreLife
    Check C, "mean_life", life
    Check C, "core_mean_life", coreLife
    Check C, "noncore_share", nonCore
End Sub


Private Sub CaseCap()
    Const C As String = "DEC-CAP-01"
    Dim k As Long, p As String, mpa() As Double, runoff() As Double, median() As Double
    Dim nonCore As Double, life As Double, coreLife As Double, ov() As Double, ovInfo() As Double
    Dim ovLabel As String, h As Long, total As Double
    For k = 1 To 3
        p = "case" & k & "."
        CORE_DecayProfile InputValue(C, p & "b0"), InputValue(C, p & "mu_star"), InputValue(C, p & "sigma"), _
                          InputValue(C, p & "z"), CLng(InputValue(C, p & "cutoff")), mpa, runoff, median, _
                          nonCore, life, coreLife
        IRRBB_DecayOverlay CStr(InputValue(C, p & "segment")), mpa, ov, ovInfo, ovLabel
        total = 0
        For h = 0 To UBound(ov)
            Check C, p & "buckets[" & h & "]", ov(h)
            total = total + ov(h)
        Next h
        Check C, p & "scale", ovInfo(IRRBB_OV_SCALE)
        Check C, p & "share_applied", ovInfo(IRRBB_OV_SHARE_APPLIED)
        Check C, p & "core_share", ovInfo(IRRBB_OV_CORE_SHARE)
        Check C, p & "core_mean_life_in", ovInfo(IRRBB_OV_CORE_MEANLIFE_IN)
        Check C, p & "cutoff", ovInfo(IRRBB_OV_CUTOFF)
        Check C, p & "mat_applied", ovInfo(IRRBB_OV_MAT_APPLIED)
        Check C, p & "core_mean_life", ovInfo(IRRBB_OV_CORE_MEANLIFE)
        Check C, p & "overnight", ovInfo(IRRBB_OV_OVERNIGHT)
        Check C, p & "share_cap", ovInfo(IRRBB_OV_CORE_SHARE_CAP)
        Check C, p & "mat_cap", ovInfo(IRRBB_OV_MAT_CAP)
        TEST_AssertNear C, p & "overlay sums to B0", total, mpa(0), "abs", 0.00000001
        TEST_AssertTrue C, p & "bucket 1 empty (non-core overnight)", ov(1) = 0
        TEST_AssertTrue C, p & "label marks REG-2 provisional", InStr(1, ovLabel, "REG-2, unverified]: provisional") > 0
    Next k
End Sub


Private Sub CaseEndToEnd()
    Const C As String = "DEC-E2E-01"
    Dim dates() As Date, bal() As Double, fit() As Double, mpa() As Double, runoff() As Double
    Dim median() As Double, kappa() As Double, resid() As Double, diag() As Double
    Dim ov() As Double, ovInfo() As Double, ovLabel As String, pairs() As Double, exceed() As Double, sm() As Double
    Dim h As Long

    FitRetTx dates, bal, fit, mpa, runoff, median, kappa, resid, diag
    Check C, "n", fit(IRRBB_DF_N)
    Check C, "mu_hat", fit(IRRBB_DF_MU_HAT)
    Check C, "kappa", kappa(1)
    Check C, "sigma", fit(IRRBB_DF_SIGMA)
    Check C, "z", fit(IRRBB_DF_Z)
    Check C, "mu_star", fit(IRRBB_DF_MU_STAR)
    Check C, "b0", fit(IRRBB_DF_B0)
    Check C, "mpa1", mpa(1)
    Check C, "mpa12", mpa(12)
    Check C, "mpa60", mpa(60)
    Check C, "mpa120", mpa(120)
    Check C, "noncore_share", fit(IRRBB_DF_NONCORE_SHARE)
    Check C, "mean_life", fit(IRRBB_DF_MEANLIFE)
    Check C, "core_mean_life", fit(IRRBB_DF_CORE_MEANLIFE)
    Check C, "lb_q", diag(IRRBB_DD_LB_Q)
    Check C, "lb_p", diag(IRRBB_DD_LB_P)
    Check C, "jb", diag(IRRBB_DD_JB)
    Check C, "jb_p", diag(IRRBB_DD_JB_P)
    Check C, "skew", diag(IRRBB_DD_SKEW)
    Check C, "kurt", diag(IRRBB_DD_KURT)
    Check C, "emp_q", diag(IRRBB_DD_EMP_Q)
    TEST_AssertTrue C, "fit origin is the estimation end", CDate(fit(IRRBB_DF_ORIGIN)) = InputValue(C, "est_end")

    IRRBB_DecayOverlay "RET_TX", mpa, ov, ovInfo, ovLabel
    Check C, "ov_scale", ovInfo(IRRBB_OV_SCALE)
    Check C, "ov_cutoff", ovInfo(IRRBB_OV_CUTOFF)
    Check C, "ov_overnight", ovInfo(IRRBB_OV_OVERNIGHT)
    Check C, "ov_core_mean_life", ovInfo(IRRBB_OV_CORE_MEANLIFE)

    IRRBB_DecayBacktest dates, bal, fit, InputValue(C, "test_end"), CLng(InputValue(C, "max_horizon")), _
                        pairs, exceed, sm
    Check C, "bt_origins", sm(IRRBB_BT_ORIGINS)
    Check C, "bt_pairs", sm(IRRBB_BT_PAIRS)
    Check C, "bt_exceed", sm(IRRBB_BT_EXCEED)
    Check C, "bt_rate", sm(IRRBB_BT_RATE)
    For h = 1 To UBound(pairs)
        Check C, "bt_pairs_by_h[" & (h - 1) & "]", pairs(h)
        Check C, "bt_exceed_by_h[" & (h - 1) & "]", exceed(h)
    Next h

    ' The facade layouts must match the core layouts they publish.
    TEST_AssertTrue C, "fit layout matches core", IRRBB_DF_FIRST_DATE = CORE_DF_FIRST_DATE And _
                    IRRBB_DF_CUTOFF = CORE_DF_CUTOFF And IRRBB_DF_MU_STAR = CORE_DF_MU_STAR And _
                    UBound(fit) = CORE_DF_COUNT
    TEST_AssertTrue C, "diagnostics layout matches core", IRRBB_DD_NORM_Q = CORE_DD_NORM_Q And _
                    IRRBB_DD_EMP_Q = CORE_DD_EMP_Q And UBound(diag) = CORE_DD_COUNT
    TEST_AssertTrue C, "overlay layout matches core", IRRBB_OV_OVERNIGHT = CORE_OV_OVERNIGHT And _
                    IRRBB_OV_SCALE = CORE_OV_SCALE And UBound(ovInfo) = CORE_OV_COUNT
    TEST_AssertTrue C, "backtest layout matches core", IRRBB_BT_MAX_H = CORE_BT_MAX_H And _
                    IRRBB_BT_RATE = CORE_BT_RATE And UBound(sm) = CORE_BT_COUNT
End Sub


Private Sub CaseErrors()
    Const C As String = "DEC-ERR-01"
    Dim dates() As Date, bal() As Double, d2() As Date, b2() As Double, i As Long, n As Long
    Dim mpa() As Double, ov() As Double, ovInfo() As Double, ovLabel As String
    Dim fit() As Double, runoff() As Double, median() As Double, kappa() As Double
    Dim resid() As Double, diag() As Double

    TEST_FixtureRetTx dates, bal
    TEST_AssertError C, "35 log changes are below the minimum", _
                     FitError("RET_TX", "EUR", dates, bal, DateSerial(2012, 1, 31), DateSerial(2014, 12, 31), Empty, 0.99, 120), _
                     CORE_ERR_SAMPLE
    TEST_AssertError C, "36 log changes are accepted", _
                     FitError("RET_TX", "EUR", dates, bal, DateSerial(2012, 1, 31), DateSerial(2015, 1, 31), Empty, 0.99, 120), 0
    TEST_AssertError C, "unknown segment", _
                     FitError("RET_XX", "EUR", dates, bal, DateSerial(2012, 1, 31), DateSerial(2022, 12, 31), Empty, 0.99, 120), _
                     CORE_ERR_INPUT
    TEST_AssertError C, "currency outside the run allowlist", _
                     FitError("RET_TX", "GBP", dates, bal, DateSerial(2012, 1, 31), DateSerial(2022, 12, 31), Empty, 0.99, 120), _
                     CORE_ERR_INPUT
    TEST_AssertError C, "currency not ISO upper case", _
                     FitError("RET_TX", "eur", dates, bal, DateSerial(2012, 1, 31), DateSerial(2022, 12, 31), Empty, 0.99, 120), _
                     CORE_ERR_INPUT
    TEST_AssertError C, "confidence of 1", _
                     FitError("RET_TX", "EUR", dates, bal, DateSerial(2012, 1, 31), DateSerial(2022, 12, 31), Empty, 1, 120), _
                     CORE_ERR_DOMAIN
    TEST_AssertError C, "cutoff of 1 month", _
                     FitError("RET_TX", "EUR", dates, bal, DateSerial(2012, 1, 31), DateSerial(2022, 12, 31), Empty, 0.99, 1), _
                     CORE_ERR_INPUT
    TEST_AssertError C, "break outside the window", _
                     FitError("RET_TX", "EUR", dates, bal, DateSerial(2012, 1, 31), DateSerial(2022, 12, 31), _
                              Array(DateSerial(2024, 3, 31)), 0.99, 120), CORE_ERR_INPUT
    TEST_AssertError C, "break given as text is rejected", _
                     FitError("RET_TX", "EUR", dates, bal, DateSerial(2012, 1, 31), DateSerial(2022, 12, 31), _
                              Array("2020-03-31"), 0.99, 120), CORE_ERR_INPUT

    d2 = dates
    b2 = bal
    b2(50) = 0
    TEST_AssertError C, "zero balance in the window", _
                     FitError("RET_TX", "EUR", d2, b2, DateSerial(2012, 1, 31), DateSerial(2022, 12, 31), Empty, 0.99, 120), _
                     CORE_ERR_DOMAIN
    d2 = dates
    b2 = bal
    d2(50) = d2(50) - 1
    TEST_AssertError C, "date that is not a month end", _
                     FitError("RET_TX", "EUR", d2, b2, DateSerial(2012, 1, 31), DateSerial(2022, 12, 31), Empty, 0.99, 120), _
                     CORE_ERR_INPUT
    ReDim d2(1 To UBound(dates) - 1)
    ReDim b2(1 To UBound(dates) - 1)
    For i = 1 To UBound(d2)
        If i < 60 Then
            d2(i) = dates(i)
            b2(i) = bal(i)
        Else
            d2(i) = dates(i + 1)
            b2(i) = bal(i + 1)
        End If
    Next i
    TEST_AssertError C, "missing month (gap) in the window", _
                     FitError("RET_TX", "EUR", d2, b2, DateSerial(2012, 1, 31), DateSerial(2022, 12, 31), Empty, 0.99, 120), _
                     CORE_ERR_INPUT

    ReDim mpa(0 To 3)
    mpa(0) = 100: mpa(1) = 90: mpa(2) = 95: mpa(3) = 80
    TEST_AssertError C, "overlay rejects a non-monotone path", OverlayError("RET_TX", mpa), CORE_ERR_INVARIANT
    mpa(2) = 85
    TEST_AssertError C, "overlay rejects an unknown segment", OverlayError("RET", mpa), CORE_ERR_INPUT

    FitRetTx dates, bal, fit, mpa, runoff, median, kappa, resid, diag
    n = UBound(dates)
    TEST_FixtureRetTx d2, b2
    b2(n - 1) = 0
    TEST_AssertError C, "backtest rejects a zero balance at a forecast origin", _
                     BacktestError(dates, b2, fit, dates(n)), CORE_ERR_DOMAIN
    TEST_FixtureRetTx d2, b2
    b2(n) = 0
    TEST_AssertError C, "backtest accepts a zero last realized balance", BacktestError(dates, b2, fit, dates(n)), 0
End Sub


Private Sub CaseFreeze()
    Const C As String = "OOS-FREEZE-01"
    Dim dates() As Date, bal() As Double, fit() As Double, mpa() As Double, runoff() As Double
    Dim median() As Double, kappa() As Double, resid() As Double, diag() As Double
    Dim fit2() As Double, keep() As Double, pairs() As Double, exceed() As Double, sm() As Double
    Dim i As Long, same As Boolean

    FitRetTx dates, bal, fit, mpa, runoff, median, kappa, resid, diag
    keep = fit
    IRRBB_DecayBacktest dates, bal, fit, DateSerial(2025, 12, 31), 12, pairs, exceed, sm
    same = True
    For i = 1 To UBound(fit)
        If fit(i) <> keep(i) Then same = False
    Next i
    TEST_AssertTrue C, "backtest leaves the frozen parameters unchanged", same

    For i = 1 To UBound(dates)
        If dates(i) > DateSerial(2022, 12, 31) Then bal(i) = bal(i) * 0.5
    Next i
    IRRBB_DecayFit "RET_TX", "EUR", dates, bal, DateSerial(2012, 1, 31), DateSerial(2022, 12, 31), _
                   Array(DateSerial(2020, 3, 31)), 0.99, 120, fit2, mpa, runoff, median, kappa, resid, diag
    same = True
    For i = 1 To UBound(fit)
        If fit2(i) <> keep(i) Then same = False
    Next i
    TEST_AssertTrue C, "test-window data do not change the estimated parameters", same
End Sub


Private Sub CaseLeak()
    Const C As String = "OOS-LEAK-01"
    Dim dates() As Date, bal() As Double, fit() As Double, mpa() As Double, runoff() As Double
    Dim median() As Double, kappa() As Double, resid() As Double, diag() As Double
    Dim p1() As Double, e1() As Double, s1() As Double, p2() As Double, e2() As Double, s2() As Double
    Dim pr() As Double, er() As Double, before As Double, after As Double
    Dim i As Long, same As Boolean

    FitRetTx dates, bal, fit, mpa, runoff, median, kappa, resid, diag
    IRRBB_DecayBacktest dates, bal, fit, DateSerial(2024, 12, 31), 12, p1, e1, s1
    For i = 1 To UBound(dates)
        If dates(i) > DateSerial(2024, 12, 31) Then bal(i) = bal(i) * 0.1
    Next i
    IRRBB_DecayBacktest dates, bal, fit, DateSerial(2024, 12, 31), 12, p2, e2, s2
    same = True
    For i = 1 To UBound(p1)
        If p1(i) <> p2(i) Or e1(i) <> e2(i) Then same = False
    Next i
    For i = 1 To UBound(s1)
        If s1(i) <> s2(i) Then same = False
    Next i
    TEST_AssertTrue C, "scores unchanged when data after the test window change", same

    ' Data after an origin inside the window: every forecast must still come from its own
    ' origin balance and the frozen parameters. The fixture is reloaded because LibreOffice
    ' shares an array on assignment.
    TEST_FixtureRetTx dates, bal
    For i = 1 To UBound(dates)
        If dates(i) > DateSerial(2023, 12, 31) Then bal(i) = bal(i) * 0.1
    Next i
    IRRBB_DecayBacktest dates, bal, fit, DateSerial(2024, 12, 31), 12, p2, e2, s2
    ReferenceScores dates, bal, fit, DateSerial(2024, 12, 31), 12, pr, er
    same = True
    For i = 1 To 12
        If p2(i) <> pr(i) Or e2(i) <> er(i) Then same = False
        before = before + e1(i)
        after = after + e2(i)
    Next i
    TEST_AssertTrue C, "the change after the origin alters the realized exceedances", after > before
    TEST_AssertTrue C, "scores equal a reference built from each origin balance only", same
End Sub


Private Sub ReferenceScores(ByRef dates() As Date, ByRef bal() As Double, ByRef fit() As Double, _
                            ByVal testEnd As Date, ByVal maxH As Long, ByRef pairs() As Double, _
                            ByRef exceed() As Double)
    ' Exceedance counts per horizon from the origin balance B(u) and the frozen fit only.
    Dim u As Long, h As Long, threshold As Double
    ReDim pairs(1 To maxH)
    ReDim exceed(1 To maxH)
    For u = 1 To UBound(dates)
        If dates(u) > CDate(fit(CORE_DF_ORIGIN)) And dates(u) < testEnd Then
            For h = 1 To maxH
                If u + h > UBound(dates) Then Exit For
                If dates(u + h) > testEnd Then Exit For
                threshold = bal(u) * Exp(fit(CORE_DF_MU_STAR) * h + fit(CORE_DF_Z) * fit(CORE_DF_SIGMA) * Sqr(h))
                pairs(h) = pairs(h) + 1
                If bal(u + h) < threshold Then exceed(h) = exceed(h) + 1
            Next h
        End If
    Next u
End Sub


Private Sub CaseScore()
    Const C As String = "OOS-SCORE-01"
    Dim dates() As Date, bal() As Double, fit() As Double, i As Long, n As Long
    Dim pairs() As Double, exceed() As Double, sm() As Double
    n = 7
    ReDim dates(1 To n)
    ReDim bal(1 To n)
    For i = 1 To n
        dates(i) = InputValue(C, "dates[" & (i - 1) & "]")
        bal(i) = InputValue(C, "balances[" & (i - 1) & "]")
    Next i
    ReDim fit(1 To CORE_DF_COUNT)
    fit(CORE_DF_MU_STAR) = InputValue(C, "mu_star")
    fit(CORE_DF_SIGMA) = InputValue(C, "sigma")
    fit(CORE_DF_Z) = InputValue(C, "z")
    fit(CORE_DF_CONFIDENCE) = InputValue(C, "confidence")
    fit(CORE_DF_ORIGIN) = CDbl(CDate(InputValue(C, "origin")))
    CORE_DecayBacktest dates, bal, fit, dates(n), CLng(InputValue(C, "max_horizon")), pairs, exceed, sm
    For i = 1 To UBound(pairs)
        Check C, "pairs_by_h[" & (i - 1) & "]", pairs(i)
        Check C, "exceed_by_h[" & (i - 1) & "]", exceed(i)
    Next i
    Check C, "origins", sm(CORE_BT_ORIGINS)
    Check C, "pairs", sm(CORE_BT_PAIRS)
    Check C, "exceed", sm(CORE_BT_EXCEED)
    Check C, "rate", sm(CORE_BT_RATE)
End Sub


'==============================================================================
' Helpers
'==============================================================================
Private Sub FitRetTx(ByRef dates() As Date, ByRef bal() As Double, ByRef fit() As Double, ByRef mpa() As Double, _
                     ByRef runoff() As Double, ByRef median() As Double, ByRef kappa() As Double, _
                     ByRef resid() As Double, ByRef diag() As Double)
    Const C As String = "DEC-E2E-01"
    TEST_FixtureRetTx dates, bal
    IRRBB_DecayFit CStr(InputValue(C, "segment")), CStr(InputValue(C, "currency")), dates, bal, _
                   InputValue(C, "est_start"), InputValue(C, "est_end"), Array(InputValue(C, "break_dates[0]")), _
                   InputValue(C, "confidence"), CLng(InputValue(C, "cutoff")), fit, mpa, runoff, median, _
                   kappa, resid, diag
End Sub


Private Sub HandProfile(ByVal caseId As String, ByRef mpa() As Double, ByRef runoff() As Double, _
                        ByRef median() As Double, ByRef nonCore As Double, ByRef life As Double, _
                        ByRef coreLife As Double)
    CORE_DecayProfile InputValue(caseId, "b0"), InputValue(caseId, "mu_star"), InputValue(caseId, "sigma"), _
                      InputValue(caseId, "z"), CLng(InputValue(caseId, "cutoff")), mpa, runoff, median, _
                      nonCore, life, coreLife
End Sub


Private Function FitError(ByVal segment As String, ByVal currencyCode As String, ByRef dates() As Date, _
                          ByRef bal() As Double, ByVal estStart As Date, ByVal estEnd As Date, _
                          ByVal breaks As Variant, ByVal confidence As Double, ByVal cutoff As Long) As Long
    Dim fit() As Double, mpa() As Double, runoff() As Double, median() As Double
    Dim kappa() As Double, resid() As Double, diag() As Double
    On Error Resume Next
    IRRBB_DecayFit segment, currencyCode, dates, bal, estStart, estEnd, breaks, confidence, cutoff, _
                   fit, mpa, runoff, median, kappa, resid, diag
    FitError = Err.Number
    On Error GoTo 0
End Function


Private Function BacktestError(ByRef dates() As Date, ByRef bal() As Double, ByRef fit() As Double, _
                               ByVal testEnd As Date) As Long
    Dim pairs() As Double, exceed() As Double, sm() As Double
    On Error Resume Next
    IRRBB_DecayBacktest dates, bal, fit, testEnd, 12, pairs, exceed, sm
    BacktestError = Err.Number
    On Error GoTo 0
End Function


Private Function OverlayError(ByVal segment As String, ByRef mpa() As Double) As Long
    Dim ov() As Double, ovInfo() As Double, ovLabel As String
    On Error Resume Next
    IRRBB_DecayOverlay segment, mpa, ov, ovInfo, ovLabel
    OverlayError = Err.Number
    On Error GoTo 0
End Function


Private Sub Check(ByVal caseId As String, ByVal key As String, ByVal actual As Double)
    Dim i As Long
    For i = LBound(mExpected, 1) To UBound(mExpected, 1)
        If mExpected(i, 1) = caseId And mExpected(i, 2) = key Then
            TEST_AssertNear caseId, key, actual, CDbl(mExpected(i, 3)), CStr(mExpected(i, 4)), CDbl(mExpected(i, 5))
            Exit Sub
        End If
    Next i
    TEST_AssertTrue caseId, key & " has no registered expected value", False
End Sub


Private Function InputValue(ByVal caseId As String, ByVal key As String) As Variant
    Dim i As Long
    For i = LBound(mInputs, 1) To UBound(mInputs, 1)
        If mInputs(i, 1) = caseId And mInputs(i, 2) = key Then
            InputValue = mInputs(i, 3)
            Exit Function
        End If
    Next i
    Err.Raise CORE_ERR_INPUT, "TEST_DecayCases.InputValue", "No input " & key & " for " & caseId & "."
End Function
