Attribute VB_Name = "IRRBB_Decay"
'==============================================================================
' IRRBB_Decay  (public facade: decay model)
'------------------------------------------------------------------------------
' PURPOSE
'   Supported entry points for the decay model of
'   docs/methodology/MODEL_CONTRACTS.md: fit, regulatory overlay and
'   out-of-sample backtest. Validates at the public boundary, then delegates
'   to src/core. No model formula lives here.
'
' UNITS (docs/REPOSITORY_STRUCTURE.md, parameter and units boundary)
'   Dates: VBA Date at month end, ascending. Balances: currency units of one
'   currency (ISO code carried alongside, never converted). Shares and the
'   confidence level: decimals. Horizons, cutoffs and mean lives: months.
'
' ARRAYS
'   Series are 1-based and aligned. Profiles are (0 To H). Output vectors
'   are 1-based and indexed by the enums below.
'
' ERRORS
'   Invalid input raises CORE_ERR_INPUT / _DOMAIN / _SAMPLE; a broken model
'   invariant raises CORE_ERR_INVARIANT. Outputs are undefined after an error
'   and must be discarded by the caller.
'
' LIMITATION
'   The decay model has no rate input: rate scenarios do not reach it. This
'   limitation applies to every decay result (MODEL_CONTRACTS.md).
'==============================================================================
Option Explicit

Public Const IRRBB_DECAY_MODEL_VERSION As String = "DECAY-1.0-draft"

Public Enum IRRBB_DecayFitField
    IRRBB_DF_MU_HAT = 1
    IRRBB_DF_MU_STAR = 2
    IRRBB_DF_SIGMA = 3
    IRRBB_DF_Z = 4
    IRRBB_DF_CONFIDENCE = 5
    IRRBB_DF_N = 6
    IRRBB_DF_K = 7
    IRRBB_DF_ORIGIN = 8
    IRRBB_DF_B0 = 9
    IRRBB_DF_CORE = 10
    IRRBB_DF_NONCORE_SHARE = 11
    IRRBB_DF_MEANLIFE = 12
    IRRBB_DF_CORE_MEANLIFE = 13
    IRRBB_DF_CUTOFF = 14
    IRRBB_DF_FIRST_DATE = 15
End Enum

Public Enum IRRBB_DecayDiagField
    IRRBB_DD_LB_LAGS = 1
    IRRBB_DD_LB_Q = 2
    IRRBB_DD_LB_P = 3
    IRRBB_DD_JB = 4
    IRRBB_DD_JB_P = 5
    IRRBB_DD_SKEW = 6
    IRRBB_DD_KURT = 7
    IRRBB_DD_EMP_Q = 8
    IRRBB_DD_NORM_Q = 9
End Enum

Public Enum IRRBB_DecayOverlayField
    IRRBB_OV_CORE_SHARE_FIT = 1
    IRRBB_OV_CORE_SHARE_CAP = 2
    IRRBB_OV_SCALE = 3
    IRRBB_OV_SHARE_APPLIED = 4
    IRRBB_OV_CORE_SHARE = 5
    IRRBB_OV_MAT_CAP = 6
    IRRBB_OV_CORE_MEANLIFE_IN = 7
    IRRBB_OV_CUTOFF = 8
    IRRBB_OV_MAT_APPLIED = 9
    IRRBB_OV_CORE_MEANLIFE = 10
    IRRBB_OV_OVERNIGHT = 11
End Enum

Public Enum IRRBB_DecayBacktestField
    IRRBB_BT_ORIGINS = 1
    IRRBB_BT_PAIRS = 2
    IRRBB_BT_EXCEED = 3
    IRRBB_BT_RATE = 4
    IRRBB_BT_TAIL = 5
    IRRBB_BT_MAX_H = 6
End Enum


Public Function IRRBB_DecayVersion() As String
    IRRBB_DecayVersion = IRRBB_DECAY_MODEL_VERSION
End Function


'------------------------------------------------------------------------------
' IRRBB_DecayFit
'   segment, currency    data-contract segment code; ISO 4217 code from the
'                        versioned run allowlist (CORE_Codes)              [in]
'   dates(), balances()  aggregate month-end balance of that segment and
'                        currency, derived by the caller from the panel     [in]
'   estStart, estEnd     estimation window [T_s, T_e]                        [in]
'   breakDates           Empty, or a 1-D array of Date values, each the month
'                        end at which a declared break log change ends       [in]
'   confidence, cutoff   c in (0.5, 1); H in months, 2..1200                 [in]
'   fit(), mpa(0..H), runoff(1..H), median(0..H), kappa(), resid(), diag()  [out]
'------------------------------------------------------------------------------
Public Sub IRRBB_DecayFit(ByVal segment As String, ByVal currencyCode As String, ByRef dates() As Date, _
                          ByRef balances() As Double, ByVal estStart As Date, ByVal estEnd As Date, _
                          ByVal breakDates As Variant, ByVal confidence As Double, ByVal cutoff As Long, _
                          ByRef fit() As Double, ByRef mpa() As Double, ByRef runoff() As Double, _
                          ByRef median() As Double, ByRef kappa() As Double, ByRef resid() As Double, _
                          ByRef diag() As Double)
    Dim breaks() As Date, k As Long
    CheckSeries dates, balances, "IRRBB_DecayFit"
    If Not CORE_IsSegmentCode(segment) Then
        Err.Raise CORE_ERR_INPUT, "IRRBB_Decay.IRRBB_DecayFit", "Unknown segment code '" & segment & "'."
    End If
    CheckCurrency currencyCode, "IRRBB_DecayFit"
    If cutoff < 2 Or cutoff > 1200 Then
        Err.Raise CORE_ERR_INPUT, "IRRBB_Decay.IRRBB_DecayFit", "Cutoff must be between 2 and 1200 months."
    End If
    ToDateArray breakDates, breaks, k
    CORE_DecayFit dates, balances, estStart, estEnd, breaks, k, confidence, cutoff, fit, mpa, runoff, _
                  median, kappa, resid, diag
End Sub


'------------------------------------------------------------------------------
' IRRBB_DecayOverlay: REG-1 constrained profile from a fitted MPA path.
'   overlay(0..H) with bucket 0 = overnight; ovInfo() per IRRBB_DecayOverlayField;
'   ovLabel: applied constraints with source IDs and verification status.
'------------------------------------------------------------------------------
Public Sub IRRBB_DecayOverlay(ByVal segment As String, ByRef mpa() As Double, ByRef overlay() As Double, _
                              ByRef ovInfo() As Double, ByRef ovLabel As String)
    CORE_DecayOverlay segment, mpa, overlay, ovInfo, ovLabel
End Sub


'------------------------------------------------------------------------------
' IRRBB_DecayBacktest: exceedance score of frozen fit parameters on the test
'   window (fit origin, testEnd]; horizons 1..maxHorizon (months).
'------------------------------------------------------------------------------
Public Sub IRRBB_DecayBacktest(ByRef dates() As Date, ByRef balances() As Double, ByRef fit() As Double, _
                               ByVal testEnd As Date, ByVal maxHorizon As Long, ByRef pairs() As Double, _
                               ByRef exceed() As Double, ByRef summary() As Double)
    CheckSeries dates, balances, "IRRBB_DecayBacktest"
    If LBound(fit) <> 1 Or UBound(fit) <> CORE_DF_COUNT Then
        Err.Raise CORE_ERR_INPUT, "IRRBB_Decay.IRRBB_DecayBacktest", "fit must come from IRRBB_DecayFit."
    End If
    If maxHorizon < 1 Or maxHorizon > 120 Then
        Err.Raise CORE_ERR_INPUT, "IRRBB_Decay.IRRBB_DecayBacktest", "maxHorizon must be between 1 and 120 months."
    End If
    CORE_DecayBacktest dates, balances, fit, testEnd, maxHorizon, pairs, exceed, summary
End Sub


Private Sub CheckSeries(ByRef dates() As Date, ByRef balances() As Double, ByVal caller As String)
    Dim i As Long
    If LBound(dates) <> 1 Or LBound(balances) <> 1 Or UBound(dates) <> UBound(balances) Then
        Err.Raise CORE_ERR_INPUT, "IRRBB_Decay." & caller, "dates and balances must be 1-based and aligned."
    End If
    For i = 2 To UBound(dates)
        If dates(i) <= dates(i - 1) Then
            Err.Raise CORE_ERR_INPUT, "IRRBB_Decay." & caller, "Dates must be strictly ascending."
        End If
    Next i
End Sub


Private Sub CheckCurrency(ByVal currencyCode As String, ByVal caller As String)
    If Not CORE_IsCurrencyAllowed(currencyCode) Then
        Err.Raise CORE_ERR_INPUT, "IRRBB_Decay." & caller, "Currency '" & currencyCode & _
                  "' is not in the run allowlist " & CORE_CURRENCY_ALLOWLIST & " (" & CORE_CURRENCY_ALLOWLIST_VERSION & ")."
    End If
End Sub


Private Sub ToDateArray(ByVal v As Variant, ByRef out() As Date, ByRef k As Long)
    Dim i As Long
    k = 0
    ReDim out(1 To 1)
    If IsEmpty(v) Then Exit Sub
    If Not IsArray(v) Then
        Err.Raise CORE_ERR_INPUT, "IRRBB_Decay.IRRBB_DecayFit", "breakDates must be Empty or an array of Date."
    End If
    If UBound(v) < LBound(v) Then Exit Sub
    ReDim out(1 To UBound(v) - LBound(v) + 1)
    For i = LBound(v) To UBound(v)
        If VarType(v(i)) <> vbDate Then
            Err.Raise CORE_ERR_INPUT, "IRRBB_Decay.IRRBB_DecayFit", "breakDates elements must be Date values."
        End If
        k = k + 1
        out(k) = v(i)
    Next i
End Sub
