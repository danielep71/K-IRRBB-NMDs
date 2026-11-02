Attribute VB_Name = "CORE_Decay"
'==============================================================================
' CORE_Decay
'------------------------------------------------------------------------------
' PURPOSE
'   Decay model of docs/methodology/MODEL_CONTRACTS.md (accepted in #9):
'   log balance changes with declared impulse breaks, prudential drift,
'   minimum probable amount (MPA), conserved runoff profile, core / non-core
'   split, mean lives, diagnostics, and the out-of-sample exceedance score
'   with frozen parameters.
'
' METHOD
'   g(t)   = ln B(t) - ln B(t-1), t = s+1..e (consecutive month ends, B > 0)
'   g(t)   = mu + sum_k kappa_k D_k(t) + e(t)                      (OLS)
'   sigma  = sqrt(SSR / (n - 1 - K)); mu* = min(mu, 0); z = Phi^-1(1 - c)
'   MPA(h) = B(e) exp(mu* h + z sigma sqrt(h)),  h = 0..H
'   o(h)   = MPA(h-1) - MPA(h), h = 1..H; remainder MPA(H) at month H
'   L      = (sum_h h o(h) + H MPA(H)) / B(e)                     (months)
'   Lcore  = (sum_{h>=2} h o(h) + H MPA(H)) / MPA(1)              (months)
'
' IMPLEMENTATION CHOICES NOT FIXED BY THE CONTRACT (to be confirmed in #36)
'   - A break date is the month end at which the affected log change ends.
'   - The median path reported alongside is B(e) exp(mu* h), i.e. the same
'     prudential drift with z = 0.
'   - Residual autocorrelation: Ljung-Box Q with 12 lags, chi-square(12).
'     Normality: Jarque-Bera with population moments, chi-square(2).
'     Empirical quantile: Hyndman-Fan type 7 of e / sigma at 1 - c.
'   - Backtest pairs use horizons 1..maxHorizon with both origin and
'     realisation inside the test window (T_e, T_f].
'
' CONTRACT
'   Dates are VBA Date values at month end, ascending; balances in currency
'   units; shares in [0, 1]; horizons and mean lives in whole/fractional
'   months. The model has no rate input: rate scenarios cannot reach it.
'==============================================================================
Option Explicit
Option Private Module

Public Const CORE_DEC_MIN_CHANGES As Long = 36
Public Const CORE_DEC_LB_LAGS As Long = 12

' Fit vector layout (1 To CORE_DF_COUNT).
Public Const CORE_DF_MU_HAT As Long = 1
Public Const CORE_DF_MU_STAR As Long = 2
Public Const CORE_DF_SIGMA As Long = 3
Public Const CORE_DF_Z As Long = 4
Public Const CORE_DF_CONFIDENCE As Long = 5
Public Const CORE_DF_N As Long = 6
Public Const CORE_DF_K As Long = 7
Public Const CORE_DF_ORIGIN As Long = 8
Public Const CORE_DF_B0 As Long = 9
Public Const CORE_DF_CORE As Long = 10
Public Const CORE_DF_NONCORE_SHARE As Long = 11
Public Const CORE_DF_MEANLIFE As Long = 12
Public Const CORE_DF_CORE_MEANLIFE As Long = 13
Public Const CORE_DF_CUTOFF As Long = 14
Public Const CORE_DF_FIRST_DATE As Long = 15
Public Const CORE_DF_COUNT As Long = 15

' Diagnostics vector layout (1 To CORE_DD_COUNT).
Public Const CORE_DD_LB_LAGS As Long = 1
Public Const CORE_DD_LB_Q As Long = 2
Public Const CORE_DD_LB_P As Long = 3
Public Const CORE_DD_JB As Long = 4
Public Const CORE_DD_JB_P As Long = 5
Public Const CORE_DD_SKEW As Long = 6
Public Const CORE_DD_KURT As Long = 7
Public Const CORE_DD_EMP_Q As Long = 8
Public Const CORE_DD_NORM_Q As Long = 9
Public Const CORE_DD_COUNT As Long = 9

' Backtest summary layout (1 To CORE_BT_COUNT).
Public Const CORE_BT_ORIGINS As Long = 1
Public Const CORE_BT_PAIRS As Long = 2
Public Const CORE_BT_EXCEED As Long = 3
Public Const CORE_BT_RATE As Long = 4
Public Const CORE_BT_TAIL As Long = 5
Public Const CORE_BT_MAX_H As Long = 6
Public Const CORE_BT_COUNT As Long = 6


'------------------------------------------------------------------------------
' CORE_DecayFit
'   dates(), bal()       aligned 1-based series                              [in]
'   estStart, estEnd     estimation window; the origin t0 is the last month
'                        end <= estEnd                                       [in]
'   breakDates()         impulse-break dates, 1-based, first breakCount used [in]
'   confidence           c, 0.5 < c < 1                                       [in]
'   cutoff               H, 2 <= H                                            [in]
'   fit(1..15), mpa(0..H), runoff(1..H), median(0..H)                        [out]
'   kappa(1..K) or (0 To 0) when K = 0, resid(1..n), diag(1..9)              [out]
'------------------------------------------------------------------------------
Public Sub CORE_DecayFit(ByRef dates() As Date, ByRef bal() As Double, ByVal estStart As Date, _
                         ByVal estEnd As Date, ByRef breakDates() As Date, ByVal breakCount As Long, _
                         ByVal confidence As Double, ByVal cutoff As Long, ByRef fit() As Double, _
                         ByRef mpa() As Double, ByRef runoff() As Double, ByRef median() As Double, _
                         ByRef kappa() As Double, ByRef resid() As Double, ByRef diag() As Double)
    Dim s As Long, e As Long, n As Long, t As Long, i As Long, j As Long, k As Long
    Dim breakIdx() As Long, y() As Double, muHat As Double, z As Double, sigma As Double
    Dim muStar As Double, b0 As Double, nonCore As Double, meanLife As Double, coreLife As Double
    Dim std() As Double, q As Double, p As Double, skew As Double, kurt As Double, jb As Double, jbP As Double

    If confidence <= 0.5 Or confidence >= 1 Then
        Err.Raise CORE_ERR_DOMAIN, "CORE_Decay.CORE_DecayFit", "Confidence must lie in (0.5, 1)."
    End If
    If cutoff < 2 Then
        Err.Raise CORE_ERR_INPUT, "CORE_Decay.CORE_DecayFit", "The cutoff must be at least 2 months."
    End If
    WindowIndex dates, estStart, estEnd, s, e, "CORE_DecayFit"
    CheckMonthly dates, s, e, "CORE_DecayFit"
    For t = s To e
        If Not (bal(t) > 0) Then
            Err.Raise CORE_ERR_DOMAIN, "CORE_Decay.CORE_DecayFit", "Balance must be positive in the window (" & _
                      Format$(dates(t), "yyyy-mm-dd") & ")."
        End If
    Next t
    n = e - s
    If n < CORE_DEC_MIN_CHANGES Then
        Err.Raise CORE_ERR_SAMPLE, "CORE_Decay.CORE_DecayFit", "Only " & n & " log changes; the minimum is " & _
                  CORE_DEC_MIN_CHANGES & "."
    End If

    ' Map break dates to observation indices of the log changes they end.
    k = breakCount
    If k < 0 Then Err.Raise CORE_ERR_INPUT, "CORE_Decay.CORE_DecayFit", "Negative break count."
    ReDim breakIdx(1 To 1)
    If k > 0 Then
        ReDim breakIdx(1 To k)
        For j = 1 To k
            breakIdx(j) = 0
            For t = s + 1 To e
                If dates(t) = breakDates(LBound(breakDates) + j - 1) Then breakIdx(j) = t
            Next t
            If breakIdx(j) = 0 Then
                Err.Raise CORE_ERR_INPUT, "CORE_Decay.CORE_DecayFit", "Break date " & _
                          Format$(breakDates(LBound(breakDates) + j - 1), "yyyy-mm-dd") & _
                          " does not end a log change inside the estimation window."
            End If
            For i = 1 To j - 1
                If breakIdx(i) = breakIdx(j) Then
                    Err.Raise CORE_ERR_INPUT, "CORE_Decay.CORE_DecayFit", "Duplicate break date."
                End If
            Next i
        Next j
    End If

    ReDim y(1 To n)
    For t = s + 1 To e
        y(t - s) = Log(bal(t)) - Log(bal(t - 1))
    Next t
    If k > 0 Then
        For j = 1 To k
            breakIdx(j) = breakIdx(j) - s
        Next j
    End If
    CORE_DecayEstimate y, breakIdx, k, muHat, sigma, kappa, resid
    If Not (sigma > 0) Then
        Err.Raise CORE_ERR_DOMAIN, "CORE_Decay.CORE_DecayFit", "Residual standard deviation is zero."
    End If

    muStar = CORE_Min(muHat, 0)
    z = CORE_NormSInv(1 - confidence)
    b0 = bal(e)
    CORE_DecayProfile b0, muStar, sigma, z, cutoff, mpa, runoff, median, nonCore, meanLife, coreLife

    ReDim fit(1 To CORE_DF_COUNT)
    fit(CORE_DF_MU_HAT) = muHat
    fit(CORE_DF_MU_STAR) = muStar
    fit(CORE_DF_SIGMA) = sigma
    fit(CORE_DF_Z) = z
    fit(CORE_DF_CONFIDENCE) = confidence
    fit(CORE_DF_N) = n
    fit(CORE_DF_K) = k
    fit(CORE_DF_ORIGIN) = CDbl(dates(e))
    fit(CORE_DF_B0) = b0
    fit(CORE_DF_CORE) = mpa(1)
    fit(CORE_DF_NONCORE_SHARE) = nonCore
    fit(CORE_DF_MEANLIFE) = meanLife
    fit(CORE_DF_CORE_MEANLIFE) = coreLife
    fit(CORE_DF_CUTOFF) = cutoff
    fit(CORE_DF_FIRST_DATE) = CDbl(dates(s))

    ' Diagnostics on the residuals.
    ReDim diag(1 To CORE_DD_COUNT)
    CORE_LjungBox resid, 1, n, CORE_DEC_LB_LAGS, q, p
    diag(CORE_DD_LB_LAGS) = CORE_DEC_LB_LAGS
    diag(CORE_DD_LB_Q) = q
    diag(CORE_DD_LB_P) = p
    CORE_JarqueBera resid, 1, n, skew, kurt, jb, jbP
    diag(CORE_DD_JB) = jb
    diag(CORE_DD_JB_P) = jbP
    diag(CORE_DD_SKEW) = skew
    diag(CORE_DD_KURT) = kurt
    ReDim std(1 To n)
    For i = 1 To n
        std(i) = resid(i) / sigma
    Next i
    diag(CORE_DD_EMP_Q) = CORE_QuantileType7(std, 1, n, 1 - confidence)
    diag(CORE_DD_NORM_Q) = z
End Sub


'------------------------------------------------------------------------------
' CORE_DecayEstimate
'   OLS of the log changes g(1 To n) on a constant and K impulse dummies, the
'   j-th equal to 1 at position breakPos(j) of g. No minimum-sample rule here
'   (CORE_DecayFit applies it). sigma uses n - 1 - K degrees of freedom.
'   muHat, sigma, kappa(1 To K) or (0 To 0), resid(1 To n)                 [out]
'------------------------------------------------------------------------------
Public Sub CORE_DecayEstimate(ByRef g() As Double, ByRef breakPos() As Long, ByVal k As Long, _
                              ByRef muHat As Double, ByRef sigma As Double, ByRef kappa() As Double, _
                              ByRef resid() As Double)
    Dim n As Long, i As Long, j As Long, x() As Double, beta() As Double, se() As Double, st() As Double
    If LBound(g) <> 1 Then Err.Raise CORE_ERR_INPUT, "CORE_Decay.CORE_DecayEstimate", "g must be 1-based."
    n = UBound(g)
    If k < 0 Or n - 1 - k < 1 Then
        Err.Raise CORE_ERR_SAMPLE, "CORE_Decay.CORE_DecayEstimate", "Too few log changes for the declared breaks."
    End If
    ReDim x(1 To n, 1 To 1 + k)
    For i = 1 To n
        x(i, 1) = 1
    Next i
    For j = 1 To k
        If breakPos(j) < 1 Or breakPos(j) > n Then
            Err.Raise CORE_ERR_INPUT, "CORE_Decay.CORE_DecayEstimate", "Break position outside the sample."
        End If
        x(breakPos(j), 1 + j) = 1
    Next j
    CORE_Ols x, g, beta, se, st, resid
    muHat = beta(1)
    sigma = st(CORE_OLS_SIGMA)
    If k > 0 Then
        ReDim kappa(1 To k)
        For j = 1 To k
            kappa(j) = beta(1 + j)
        Next j
    Else
        ReDim kappa(0 To 0)
    End If
End Sub


'------------------------------------------------------------------------------
' CORE_DecayProfile
'   MPA(h) = b0 exp(muStar h + z sigma sqrt(h)), h = 0..cutoff, with
'   muStar <= 0 and z < 0 required. Runoff o(h) = MPA(h-1) - MPA(h); the
'   remainder MPA(H) is placed at month H. Checks monotonicity and
'   conservation of the notional (relative 1E-9) and raises on violation.
'   mpa(0..H), runoff(1..H), median(0..H), nonCoreShare, meanLife,
'   coreMeanLife (months)                                                  [out]
'------------------------------------------------------------------------------
Public Sub CORE_DecayProfile(ByVal b0 As Double, ByVal muStar As Double, ByVal sigma As Double, _
                             ByVal z As Double, ByVal cutoff As Long, ByRef mpa() As Double, _
                             ByRef runoff() As Double, ByRef median() As Double, ByRef nonCoreShare As Double, _
                             ByRef meanLife As Double, ByRef coreMeanLife As Double)
    Dim h As Long, total As Double, sumLife As Double
    If Not (b0 > 0) Then Err.Raise CORE_ERR_DOMAIN, "CORE_Decay.CORE_DecayProfile", "b0 must be positive."
    If muStar > 0 Or z >= 0 Or sigma < 0 Then
        Err.Raise CORE_ERR_DOMAIN, "CORE_Decay.CORE_DecayProfile", "Requires muStar <= 0, z < 0 and sigma >= 0."
    End If
    If cutoff < 2 Then Err.Raise CORE_ERR_INPUT, "CORE_Decay.CORE_DecayProfile", "The cutoff must be at least 2 months."
    ReDim mpa(0 To cutoff)
    ReDim median(0 To cutoff)
    ReDim runoff(1 To cutoff)
    mpa(0) = b0
    median(0) = b0
    For h = 1 To cutoff
        mpa(h) = b0 * Exp(muStar * h + z * sigma * Sqr(h))
        median(h) = b0 * Exp(muStar * h)
        If mpa(h) > mpa(h - 1) Then
            Err.Raise CORE_ERR_INVARIANT, "CORE_Decay.CORE_DecayProfile", "MPA is not non-increasing at h = " & h & "."
        End If
        runoff(h) = mpa(h - 1) - mpa(h)
    Next h
    total = mpa(cutoff)
    sumLife = cutoff * mpa(cutoff)
    For h = 1 To cutoff
        total = total + runoff(h)
        sumLife = sumLife + h * runoff(h)
    Next h
    If Abs(total - b0) > 0.000000001 * b0 Then
        Err.Raise CORE_ERR_INVARIANT, "CORE_Decay.CORE_DecayProfile", "Profile does not conserve the notional."
    End If
    If Not (mpa(1) > 0) Then
        Err.Raise CORE_ERR_DOMAIN, "CORE_Decay.CORE_DecayProfile", "MPA(1) underflows to zero."
    End If
    nonCoreShare = 1 - mpa(1) / b0
    meanLife = sumLife / b0
    coreMeanLife = (sumLife - runoff(1)) / mpa(1)
End Sub


'------------------------------------------------------------------------------
' CORE_DecayBacktest
'   Scores frozen fit parameters on the test window (T_e, T_f], where T_e is
'   the fit origin. Every month end u in the window is an origin; for
'   h = 1..maxHorizon with u + h still in the window the pair is scored:
'   exceedance when B(u + h) < B(u) exp(mu* h + z sigma sqrt(h)). A zero
'   balance at an origin would make every threshold zero and every pair a
'   non-exceedance, so it is rejected; zero is accepted only as the last
'   realized balance T_f.
'   pairs(1..maxH), exceed(1..maxH), summary(1..6)                          [out]
'------------------------------------------------------------------------------
Public Sub CORE_DecayBacktest(ByRef dates() As Date, ByRef bal() As Double, ByRef fit() As Double, _
                              ByVal testEnd As Date, ByVal maxHorizon As Long, ByRef pairs() As Double, _
                              ByRef exceed() As Double, ByRef summary() As Double)
    Dim o As Long, f As Long, u As Long, h As Long, origins As Long, scored As Boolean
    Dim origin As Date, level As Double, totalPairs As Double, totalExceed As Double

    If maxHorizon < 1 Then Err.Raise CORE_ERR_INPUT, "CORE_Decay.CORE_DecayBacktest", "maxHorizon must be positive."
    origin = CDate(fit(CORE_DF_ORIGIN))
    o = 0
    f = 0
    For u = LBound(dates) To UBound(dates)
        If dates(u) = origin Then o = u
        If dates(u) <= testEnd Then f = u
    Next u
    If o = 0 Then
        Err.Raise CORE_ERR_INPUT, "CORE_Decay.CORE_DecayBacktest", "The fit origin is not in the date series."
    End If
    If f <= o + 1 Then
        Err.Raise CORE_ERR_SAMPLE, "CORE_Decay.CORE_DecayBacktest", "The test window holds no scorable pair."
    End If
    CheckMonthly dates, o, f, "CORE_DecayBacktest"
    For u = o + 1 To f
        If bal(u) < 0 Then
            Err.Raise CORE_ERR_DOMAIN, "CORE_Decay.CORE_DecayBacktest", "Negative balance in the test window."
        End If
        If bal(u) = 0 And u < f Then
            Err.Raise CORE_ERR_DOMAIN, "CORE_Decay.CORE_DecayBacktest", "Zero balance at a forecast origin in the test window."
        End If
    Next u

    ReDim pairs(1 To maxHorizon)
    ReDim exceed(1 To maxHorizon)
    For u = o + 1 To f - 1
        scored = False
        For h = 1 To maxHorizon
            If u + h > f Then Exit For
            level = bal(u) * Exp(fit(CORE_DF_MU_STAR) * h + fit(CORE_DF_Z) * fit(CORE_DF_SIGMA) * Sqr(h))
            pairs(h) = pairs(h) + 1
            If bal(u + h) < level Then exceed(h) = exceed(h) + 1
            scored = True
        Next h
        If scored Then origins = origins + 1
    Next u
    For h = 1 To maxHorizon
        totalPairs = totalPairs + pairs(h)
        totalExceed = totalExceed + exceed(h)
    Next h

    ReDim summary(1 To CORE_BT_COUNT)
    summary(CORE_BT_ORIGINS) = origins
    summary(CORE_BT_PAIRS) = totalPairs
    summary(CORE_BT_EXCEED) = totalExceed
    summary(CORE_BT_RATE) = totalExceed / totalPairs
    summary(CORE_BT_TAIL) = 1 - fit(CORE_DF_CONFIDENCE)
    summary(CORE_BT_MAX_H) = maxHorizon
End Sub


Private Sub WindowIndex(ByRef dates() As Date, ByVal fromDate As Date, ByVal toDate As Date, _
                        ByRef s As Long, ByRef e As Long, ByVal caller As String)
    Dim i As Long
    If toDate < fromDate Then
        Err.Raise CORE_ERR_INPUT, "CORE_Decay." & caller, "The estimation window ends before it starts."
    End If
    s = 0
    e = 0
    For i = LBound(dates) To UBound(dates)
        If s = 0 And dates(i) >= fromDate Then s = i
        If dates(i) <= toDate Then e = i
    Next i
    If s = 0 Or e = 0 Or e <= s Then
        Err.Raise CORE_ERR_SAMPLE, "CORE_Decay." & caller, "The estimation window holds fewer than two observations."
    End If
End Sub


Private Sub CheckMonthly(ByRef dates() As Date, ByVal s As Long, ByVal e As Long, ByVal caller As String)
    Dim t As Long
    For t = s To e
        If Not CORE_IsMonthEnd(dates(t)) Then
            Err.Raise CORE_ERR_INPUT, "CORE_Decay." & caller, "Not a month end: " & Format$(dates(t), "yyyy-mm-dd") & "."
        End If
        If t > s Then
            If dates(t) <> DateSerial(Year(dates(t - 1)), Month(dates(t - 1)) + 2, 0) Then
                Err.Raise CORE_ERR_INPUT, "CORE_Decay." & caller, "Month ends are not consecutive at " & _
                          Format$(dates(t), "yyyy-mm-dd") & "."
            End If
        End If
    Next t
End Sub
