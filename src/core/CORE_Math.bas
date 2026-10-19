Attribute VB_Name = "CORE_Math"
'==============================================================================
' CORE_Math
'------------------------------------------------------------------------------
' PURPOSE
'   Host-independent numerical primitives: ordinary least squares by
'   Householder QR, the standard normal CDF and its inverse, sample moments,
'   an empirical quantile, and the Ljung-Box and Jarque-Bera statistics.
'
' CONTRACT
'   - Arrays are 1-based; matrices are (1 To n, 1 To k).
'   - No Excel object model, worksheet function or UI (src/core boundary).
'   - Failures raise CORE_ERR_* errors and return no result.
'
' NUMERICAL NOTES
'   - OLS never forms the normal equations.
'   - CORE_NormSDist: power series for |x| < 3, Lentz continued fraction for
'     the tail otherwise; relative error near machine precision on |x| <= 37.
'   - CORE_NormSInv: Acklam's rational start (relative error < 1.15E-09)
'     followed by one Halley step on CORE_NormSDist.
'==============================================================================
Option Explicit
Option Private Module

' Error numbers: vbObjectError (-2147221504) + 512 + n, written as literals.
Public Const CORE_ERR_BASE As Long = -2147220992
Public Const CORE_ERR_INPUT As Long = -2147220991
Public Const CORE_ERR_SINGULAR As Long = -2147220990
Public Const CORE_ERR_DOMAIN As Long = -2147220989
Public Const CORE_ERR_SAMPLE As Long = -2147220988
Public Const CORE_ERR_INVARIANT As Long = -2147220987

' Layout of the OLS statistics vector returned by CORE_Ols.
Public Const CORE_OLS_R2 As Long = 1
Public Const CORE_OLS_SIGMA As Long = 2
Public Const CORE_OLS_DW As Long = 3
Public Const CORE_OLS_N As Long = 4
Public Const CORE_OLS_K As Long = 5
Public Const CORE_OLS_SSR As Long = 6
Public Const CORE_OLS_STATS As Long = 6

Private Const RANK_TOLERANCE As Double = 0.000000000001
Private Const SQRT_2PI As Double = 2.506628274631
Private Const LN_SQRT_2PI As Double = 0.918938533204673


'------------------------------------------------------------------------------
' CORE_Ols
'   Ordinary least squares of y on the columns of X (include a column of ones
'   for an intercept). Coefficient standard errors use sigma^2 = SSR / (n - k).
'
'   x(1 To n, 1 To k), y(1 To n)                         [in]
'   beta(1 To k), se(1 To k), stats(1 To 6), resid(1 To n) [out, ReDim'd]
'   stats: R2 (centred), sigma, Durbin-Watson, n, k, SSR.
'   Raises CORE_ERR_INPUT if n <= k or shapes disagree; CORE_ERR_SINGULAR if
'   X is (numerically) rank deficient.
'------------------------------------------------------------------------------
Public Sub CORE_Ols(ByRef x() As Double, ByRef y() As Double, ByRef beta() As Double, _
                    ByRef se() As Double, ByRef stats() As Double, ByRef resid() As Double)
    Dim n As Long, k As Long, i As Long, j As Long, c As Long
    Dim a() As Double, b() As Double, v() As Double
    Dim colNorm As Double, alpha As Double, vNorm2 As Double, s As Double
    Dim rMax As Double, rInv() As Double, ssr As Double, sst As Double
    Dim yMean As Double, sigma2 As Double, dwNum As Double

    n = UBound(y) - LBound(y) + 1
    k = UBound(x, 2) - LBound(x, 2) + 1
    If LBound(y) <> 1 Or LBound(x, 1) <> 1 Or LBound(x, 2) <> 1 Then
        Err.Raise CORE_ERR_INPUT, "CORE_Math.CORE_Ols", "Arrays must be 1-based."
    End If
    If UBound(x, 1) <> n Then
        Err.Raise CORE_ERR_INPUT, "CORE_Math.CORE_Ols", "X and y have different row counts."
    End If
    If n <= k Then
        Err.Raise CORE_ERR_INPUT, "CORE_Math.CORE_Ols", "Need more observations than regressors."
    End If

    ReDim a(1 To n, 1 To k)
    ReDim b(1 To n)
    ReDim v(1 To n)
    For i = 1 To n
        For j = 1 To k
            a(i, j) = x(i, j)
        Next j
        b(i) = y(i)
    Next i

    ' Householder triangularisation: A = Q R, b <- Q' b.
    For j = 1 To k
        colNorm = 0
        For i = j To n
            colNorm = colNorm + a(i, j) * a(i, j)
        Next i
        colNorm = Sqr(colNorm)
        If colNorm = 0 Then
            Err.Raise CORE_ERR_SINGULAR, "CORE_Math.CORE_Ols", "Regressor column " & j & " is zero."
        End If
        If a(j, j) > 0 Then alpha = -colNorm Else alpha = colNorm
        vNorm2 = 0
        For i = j To n
            v(i) = a(i, j)
            If i = j Then v(i) = v(i) - alpha
            vNorm2 = vNorm2 + v(i) * v(i)
        Next i
        If vNorm2 > 0 Then
            For c = j To k
                s = 0
                For i = j To n
                    s = s + v(i) * a(i, c)
                Next i
                s = 2 * s / vNorm2
                For i = j To n
                    a(i, c) = a(i, c) - s * v(i)
                Next i
            Next c
            s = 0
            For i = j To n
                s = s + v(i) * b(i)
            Next i
            s = 2 * s / vNorm2
            For i = j To n
                b(i) = b(i) - s * v(i)
            Next i
        End If
    Next j

    rMax = 0
    For j = 1 To k
        If Abs(a(j, j)) > rMax Then rMax = Abs(a(j, j))
    Next j
    For j = 1 To k
        If Abs(a(j, j)) <= RANK_TOLERANCE * rMax Then
            Err.Raise CORE_ERR_SINGULAR, "CORE_Math.CORE_Ols", "Design matrix is rank deficient."
        End If
    Next j

    ' Back substitution R beta = Q'y.
    ReDim beta(1 To k)
    For j = k To 1 Step -1
        s = b(j)
        For c = j + 1 To k
            s = s - a(j, c) * beta(c)
        Next c
        beta(j) = s / a(j, j)
    Next j

    ' Residuals from the original data.
    ReDim resid(1 To n)
    ssr = 0
    yMean = 0
    For i = 1 To n
        s = y(i)
        For j = 1 To k
            s = s - x(i, j) * beta(j)
        Next j
        resid(i) = s
        ssr = ssr + s * s
        yMean = yMean + y(i)
    Next i
    yMean = yMean / n
    sst = 0
    dwNum = 0
    For i = 1 To n
        sst = sst + (y(i) - yMean) * (y(i) - yMean)
        If i > 1 Then dwNum = dwNum + (resid(i) - resid(i - 1)) * (resid(i) - resid(i - 1))
    Next i
    sigma2 = ssr / (n - k)

    ' Cov(beta) = sigma^2 * R^-1 * R^-T.
    ReDim rInv(1 To k, 1 To k)
    For j = 1 To k
        rInv(j, j) = 1 / a(j, j)
        For i = j - 1 To 1 Step -1
            s = 0
            For c = i + 1 To j
                s = s + a(i, c) * rInv(c, j)
            Next c
            rInv(i, j) = -s / a(i, i)
        Next i
    Next j
    ReDim se(1 To k)
    For i = 1 To k
        s = 0
        For c = i To k
            s = s + rInv(i, c) * rInv(i, c)
        Next c
        se(i) = Sqr(sigma2 * s)
    Next i

    ReDim stats(1 To CORE_OLS_STATS)
    If sst > 0 Then stats(CORE_OLS_R2) = 1 - ssr / sst Else stats(CORE_OLS_R2) = 0
    stats(CORE_OLS_SIGMA) = Sqr(sigma2)
    If ssr > 0 Then stats(CORE_OLS_DW) = dwNum / ssr Else stats(CORE_OLS_DW) = 0
    stats(CORE_OLS_N) = n
    stats(CORE_OLS_K) = k
    stats(CORE_OLS_SSR) = ssr
End Sub


'------------------------------------------------------------------------------
' CORE_Mean: arithmetic mean of v(lo To hi).
'------------------------------------------------------------------------------
Public Function CORE_Mean(ByRef v() As Double, ByVal lo As Long, ByVal hi As Long) As Double
    Dim i As Long, s As Double
    If hi < lo Then Err.Raise CORE_ERR_INPUT, "CORE_Math.CORE_Mean", "Empty range."
    For i = lo To hi
        s = s + v(i)
    Next i
    CORE_Mean = s / (hi - lo + 1)
End Function


'------------------------------------------------------------------------------
' CORE_NormPdf: standard normal density.
'------------------------------------------------------------------------------
Public Function CORE_NormPdf(ByVal x As Double) As Double
    CORE_NormPdf = Exp(-0.5 * x * x - LN_SQRT_2PI)
End Function


'------------------------------------------------------------------------------
' CORE_NormSDist: standard normal CDF Phi(x).
'------------------------------------------------------------------------------
Public Function CORE_NormSDist(ByVal x As Double) As Double
    Dim t As Double, term As Double, s As Double, k As Long
    If x > 37 Then
        CORE_NormSDist = 1
        Exit Function
    End If
    If x < -37 Then
        CORE_NormSDist = 0
        Exit Function
    End If
    If Abs(x) < 3 Then
        ' Phi(x) = 1/2 + phi(x) * sum_{k>=0} x^(2k+1) / (1*3*...*(2k+1))
        term = x
        s = x
        k = 0
        Do
            k = k + 1
            term = term * x * x / (2 * k + 1)
            s = s + term
        Loop While Abs(term) > 0.0000000000000000001 * Abs(s) And k < 500
        CORE_NormSDist = 0.5 + CORE_NormPdf(x) * s
    Else
        t = Abs(x)
        If x < 0 Then
            CORE_NormSDist = UpperTail(t)
        Else
            CORE_NormSDist = 1 - UpperTail(t)
        End If
    End If
End Function


'------------------------------------------------------------------------------
' UpperTail: Q(t) = 1 - Phi(t) for t >= 3 by the Laplace continued fraction
'   Q(t) = phi(t) / (t + 1/(t + 2/(t + 3/(t + ...)))), evaluated with Lentz.
'------------------------------------------------------------------------------
Private Function UpperTail(ByVal t As Double) As Double
    Const TINY As Double = 1E-300
    Dim f As Double, c As Double, d As Double, delta As Double, k As Long
    f = t
    c = t
    d = 0
    For k = 1 To 1000
        d = t + k * d
        If Abs(d) < TINY Then d = TINY
        c = t + k / c
        If Abs(c) < TINY Then c = TINY
        d = 1 / d
        delta = c * d
        f = f * delta
        If Abs(delta - 1) < 0.0000000000000001 Then Exit For
    Next k
    UpperTail = CORE_NormPdf(t) / f
End Function


'------------------------------------------------------------------------------
' CORE_NormSInv: inverse standard normal CDF for p in the open interval (0, 1).
'   Raises CORE_ERR_DOMAIN outside that interval.
'------------------------------------------------------------------------------
Public Function CORE_NormSInv(ByVal p As Double) As Double
    Const P_LOW As Double = 0.02425
    Dim q As Double, r As Double, x As Double, e As Double, u As Double

    If p <= 0 Or p >= 1 Then
        Err.Raise CORE_ERR_DOMAIN, "CORE_Math.CORE_NormSInv", "Probability must lie in (0, 1)."
    End If
    If p < P_LOW Then
        q = Sqr(-2 * Log(p))
        x = (((((-0.007784894002430293 * q - 0.3223964580411365) * q - 2.400758277161838) * q _
            - 2.549732539343734) * q + 4.374664141464968) * q + 2.938163982698783) / _
            ((((0.007784695709041462 * q + 0.3224671290700398) * q + 2.445134137142996) * q _
            + 3.754408661907416) * q + 1)
    ElseIf p <= 1 - P_LOW Then
        q = p - 0.5
        r = q * q
        x = (((((-39.69683028665376 * r + 220.9460984245205) * r - 275.9285104469687) * r _
            + 138.357751867269) * r - 30.66479806614716) * r + 2.506628277459239) * q / _
            (((((-54.47609879822406 * r + 161.5858368580409) * r - 155.6989798598866) * r _
            + 66.80131188771972) * r - 13.28068155288572) * r + 1)
    Else
        q = Sqr(-2 * Log(1 - p))
        x = -(((((-0.007784894002430293 * q - 0.3223964580411365) * q - 2.400758277161838) * q _
             - 2.549732539343734) * q + 4.374664141464968) * q + 2.938163982698783) / _
             ((((0.007784695709041462 * q + 0.3224671290700398) * q + 2.445134137142996) * q _
             + 3.754408661907416) * q + 1)
    End If

    ' One Halley refinement step on the accurate CDF.
    If Abs(x) < 37 Then
        e = CORE_NormSDist(x) - p
        u = e * SQRT_2PI * Exp(0.5 * x * x)
        x = x - u / (1 + 0.5 * x * u)
    End If
    CORE_NormSInv = x
End Function


'------------------------------------------------------------------------------
' CORE_QuantileType7: sample quantile of v(lo To hi) at probability p in
'   [0, 1], Hyndman-Fan definition 7 (linear interpolation between order
'   statistics at position (n - 1) * p). v is not modified.
'------------------------------------------------------------------------------
Public Function CORE_QuantileType7(ByRef v() As Double, ByVal lo As Long, ByVal hi As Long, _
                                   ByVal p As Double) As Double
    Dim n As Long, i As Long, j As Long, w() As Double, key As Double, pos As Double, k As Long
    n = hi - lo + 1
    If n < 1 Then Err.Raise CORE_ERR_INPUT, "CORE_Math.CORE_QuantileType7", "Empty range."
    If p < 0 Or p > 1 Then Err.Raise CORE_ERR_DOMAIN, "CORE_Math.CORE_QuantileType7", "p must lie in [0, 1]."
    ReDim w(1 To n)
    For i = 1 To n
        w(i) = v(lo + i - 1)
    Next i
    For i = 2 To n
        key = w(i)
        j = i - 1
        Do While j >= 1
            If w(j) <= key Then Exit Do
            w(j + 1) = w(j)
            j = j - 1
        Loop
        w(j + 1) = key
    Next i
    pos = (n - 1) * p
    k = Int(pos)
    If k >= n - 1 Then
        CORE_QuantileType7 = w(n)
    Else
        CORE_QuantileType7 = w(k + 1) + (pos - k) * (w(k + 2) - w(k + 1))
    End If
End Function


'------------------------------------------------------------------------------
' CORE_LjungBox: Q = n (n + 2) sum_{k=1..m} r_k^2 / (n - k) on e(lo To hi),
'   r_k the lag-k sample autocorrelation about the mean. lags must be even
'   and below n; pValue uses the chi-square(m) survival function.
'------------------------------------------------------------------------------
Public Sub CORE_LjungBox(ByRef e() As Double, ByVal lo As Long, ByVal hi As Long, ByVal lags As Long, _
                         ByRef q As Double, ByRef pValue As Double)
    Dim n As Long, k As Long, t As Long, m As Double, den As Double, num As Double
    n = hi - lo + 1
    If lags < 2 Or lags Mod 2 <> 0 Or lags >= n Then
        Err.Raise CORE_ERR_INPUT, "CORE_Math.CORE_LjungBox", "Lags must be even and below the sample size."
    End If
    m = CORE_Mean(e, lo, hi)
    For t = lo To hi
        den = den + (e(t) - m) * (e(t) - m)
    Next t
    If den <= 0 Then Err.Raise CORE_ERR_DOMAIN, "CORE_Math.CORE_LjungBox", "Series has zero variance."
    q = 0
    For k = 1 To lags
        num = 0
        For t = lo + k To hi
            num = num + (e(t) - m) * (e(t - k) - m)
        Next t
        q = q + (num / den) * (num / den) / (n - k)
    Next k
    q = n * (n + 2) * q
    pValue = ChiSquareSurvivalEven(q, lags)
End Sub


'------------------------------------------------------------------------------
' CORE_JarqueBera: JB = n/6 (S^2 + (K - 3)^2 / 4) with population moments
'   about the mean; pValue = exp(-JB / 2) (chi-square, 2 degrees of freedom).
'------------------------------------------------------------------------------
Public Sub CORE_JarqueBera(ByRef e() As Double, ByVal lo As Long, ByVal hi As Long, ByRef skew As Double, _
                           ByRef kurt As Double, ByRef jb As Double, ByRef pValue As Double)
    Dim n As Long, t As Long, m As Double, d As Double, m2 As Double, m3 As Double, m4 As Double
    n = hi - lo + 1
    If n < 3 Then Err.Raise CORE_ERR_INPUT, "CORE_Math.CORE_JarqueBera", "Need at least three values."
    m = CORE_Mean(e, lo, hi)
    For t = lo To hi
        d = e(t) - m
        m2 = m2 + d * d
        m3 = m3 + d * d * d
        m4 = m4 + d * d * d * d
    Next t
    m2 = m2 / n
    m3 = m3 / n
    m4 = m4 / n
    If m2 <= 0 Then Err.Raise CORE_ERR_DOMAIN, "CORE_Math.CORE_JarqueBera", "Series has zero variance."
    skew = m3 / (m2 * Sqr(m2))
    kurt = m4 / (m2 * m2)
    jb = n / 6 * (skew * skew + (kurt - 3) * (kurt - 3) / 4)
    pValue = Exp(-jb / 2)
End Sub


'------------------------------------------------------------------------------
' ChiSquareSurvivalEven: P(X > x) for X ~ chi-square with an even number of
'   degrees of freedom: exp(-x/2) * sum_{j < dof/2} (x/2)^j / j!.
'------------------------------------------------------------------------------
Private Function ChiSquareSurvivalEven(ByVal x As Double, ByVal dof As Long) As Double
    Dim j As Long, term As Double, s As Double, h As Double
    If x <= 0 Then
        ChiSquareSurvivalEven = 1
        Exit Function
    End If
    h = x / 2
    term = 1
    s = 1
    For j = 1 To dof \ 2 - 1
        term = term * h / j
        s = s + term
    Next j
    ChiSquareSurvivalEven = Exp(-h) * s
End Function


Public Function CORE_Max(ByVal a As Double, ByVal b As Double) As Double
    If a >= b Then CORE_Max = a Else CORE_Max = b
End Function

Public Function CORE_Min(ByVal a As Double, ByVal b As Double) As Double
    If a <= b Then CORE_Min = a Else CORE_Min = b
End Function
