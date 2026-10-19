Attribute VB_Name = "EX_DecayExample"
'==============================================================================
' EX_DecayExample
'------------------------------------------------------------------------------
' PURPOSE
'   Example of the supported decay API (IRRBB_Decay only): reads a synthetic
'   aggregate balance series and run settings from the sheet "DecayExample",
'   fits the decay model, applies the REG-1 overlay, scores the out-of-sample
'   window and writes the results next to the inputs.
'
' NOT PRODUCTION
'   The production workflow derives segment balances from the account panel
'   (DATA_CONTRACT.md) and uses the workbook template of issue #5. This sheet
'   holds a synthetic summary series only, for demonstration.
'
' FAILURE POLICY
'   Outputs are withdrawn before the run; on failure they stay withdrawn and
'   the status shows FAILED with the primary error (and any cleanup error).
'==============================================================================
Option Explicit

Private Const SH As String = "DecayExample"
Private Const DATA_ROW As Long = 22
Private Const PROFILE_COL As Long = 5
Private Const EX_ERR As Long = -2147220480   ' vbObjectError + 1024


Public Sub RunDecayExample()
    Dim ws As Worksheet, primary As String, cleanup As String
    Dim dates() As Date, bal() As Double, n As Long, i As Long, raw As Variant, breaks As Variant
    Dim fit() As Double, mpa() As Double, runoff() As Double, median() As Double, kappa() As Double
    Dim resid() As Double, diag() As Double, ov() As Double, ovInfo() As Double, ovLabel As String
    Dim pairs() As Double, exceed() As Double, sm() As Double, block() As Variant, h As Long

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(SH)
    ClearOutputs
    ws.Range("B14").Value = "RUNNING"

    n = 0
    Do While Not IsEmpty(ws.Cells(DATA_ROW + n, 1).Value)
        n = n + 1
    Loop
    If n < 2 Then Err.Raise EX_ERR, "EX_DecayExample", "No data below row " & (DATA_ROW - 1) & "."
    ReDim dates(1 To n)
    ReDim bal(1 To n)
    For i = 1 To n
        raw = ws.Cells(DATA_ROW + i - 1, 1).Value
        If VarType(raw) <> vbDate Then Err.Raise EX_ERR + 1, "EX_DecayExample", "Row " & _
                                                 (DATA_ROW + i - 1) & ": the date cell must hold a date."
        dates(i) = raw
        bal(i) = ws.Cells(DATA_ROW + i - 1, 2).Value
    Next i

    raw = ws.Range("B10").Value
    If IsEmpty(raw) Then
        breaks = Empty
    ElseIf VarType(raw) = vbDate Then
        breaks = Array(raw)
    Else
        Err.Raise EX_ERR + 2, "EX_DecayExample", "Break date (B10) must be empty or a date."
    End If

    IRRBB_DecayFit CStr(ws.Range("B3").Value), CStr(ws.Range("B4").Value), dates, bal, ws.Range("B5").Value, _
                   ws.Range("B6").Value, breaks, ws.Range("B8").Value, CLng(ws.Range("B9").Value), fit, mpa, _
                   runoff, median, kappa, resid, diag
    IRRBB_DecayOverlay CStr(ws.Range("B3").Value), mpa, ov, ovInfo, ovLabel
    IRRBB_DecayBacktest dates, bal, fit, ws.Range("B7").Value, CLng(ws.Range("B11").Value), pairs, exceed, sm

    ' Fit, diagnostics, overlay and backtest summaries (columns E:F and H:I).
    ReDim block(1 To 15, 1 To 2)
    block(1, 1) = "mu (log/month)": block(1, 2) = fit(IRRBB_DF_MU_HAT)
    block(2, 1) = "mu* used": block(2, 2) = fit(IRRBB_DF_MU_STAR)
    block(3, 1) = "sigma": block(3, 2) = fit(IRRBB_DF_SIGMA)
    block(4, 1) = "z = Phi^-1(1-c)": block(4, 2) = fit(IRRBB_DF_Z)
    block(5, 1) = "Log changes (n)": block(5, 2) = fit(IRRBB_DF_N)
    block(6, 1) = "Break dummies (K)": block(6, 2) = fit(IRRBB_DF_K)
    If fit(IRRBB_DF_K) > 0 Then
        block(7, 1) = "kappa (first break)": block(7, 2) = kappa(1)
    Else
        block(7, 1) = "kappa": block(7, 2) = "-"
    End If
    block(8, 1) = "Origin B0": block(8, 2) = fit(IRRBB_DF_B0)
    block(9, 1) = "Core MPA(1)": block(9, 2) = fit(IRRBB_DF_CORE)
    block(10, 1) = "Non-core share": block(10, 2) = fit(IRRBB_DF_NONCORE_SHARE)
    block(11, 1) = "Mean life (months)": block(11, 2) = fit(IRRBB_DF_MEANLIFE)
    block(12, 1) = "Core mean life (months)": block(12, 2) = fit(IRRBB_DF_CORE_MEANLIFE)
    block(13, 1) = "Ljung-Box Q(12) p-value": block(13, 2) = diag(IRRBB_DD_LB_P)
    block(14, 1) = "Jarque-Bera p-value": block(14, 2) = diag(IRRBB_DD_JB_P)
    block(15, 1) = "Empirical vs normal quantile": block(15, 2) = Format$(diag(IRRBB_DD_EMP_Q), "0.000") & _
                                                                   " vs " & Format$(diag(IRRBB_DD_NORM_Q), "0.000")
    ws.Range("F3:G17").Value = block

    ReDim block(1 To 11, 1 To 2)
    block(1, 1) = "Core share (fit)": block(1, 2) = ovInfo(IRRBB_OV_CORE_SHARE_FIT)
    block(2, 1) = "Core-share cap": block(2, 2) = ovInfo(IRRBB_OV_CORE_SHARE_CAP)
    block(3, 1) = "Core share (overlay)": block(3, 2) = ovInfo(IRRBB_OV_CORE_SHARE)
    block(4, 1) = "Average-maturity cap (months)": block(4, 2) = ovInfo(IRRBB_OV_MAT_CAP)
    block(5, 1) = "Core mean life before cap": block(5, 2) = ovInfo(IRRBB_OV_CORE_MEANLIFE_IN)
    block(6, 1) = "Overlay cutoff H'": block(6, 2) = ovInfo(IRRBB_OV_CUTOFF)
    block(7, 1) = "Core mean life (overlay)": block(7, 2) = ovInfo(IRRBB_OV_CORE_MEANLIFE)
    block(8, 1) = "Backtest origins": block(8, 2) = sm(IRRBB_BT_ORIGINS)
    block(9, 1) = "Backtest pairs": block(9, 2) = sm(IRRBB_BT_PAIRS)
    block(10, 1) = "Exceedances": block(10, 2) = sm(IRRBB_BT_EXCEED)
    block(11, 1) = "Exceedance rate vs tail": block(11, 2) = Format$(sm(IRRBB_BT_RATE), "0.00%") & " vs " & _
                                                            Format$(sm(IRRBB_BT_TAIL), "0.00%")
    ws.Range("I3:J13").Value = block
    ws.Range("B15").Value = ovLabel
    ws.Range("B16").Value = "Decay results have no rate input: rate scenarios do not reach this profile."

    ReDim block(1 To UBound(mpa) + 1, 1 To 5)
    For h = 0 To UBound(mpa)
        block(h + 1, 1) = h
        block(h + 1, 2) = mpa(h)
        block(h + 1, 3) = median(h)
        If h > 0 Then block(h + 1, 4) = runoff(h)
        block(h + 1, 5) = ov(h)
    Next h
    ws.Range(ws.Cells(DATA_ROW, PROFILE_COL), ws.Cells(DATA_ROW + UBound(mpa), PROFILE_COL + 4)).Value = block

    ws.Range("B14").Value = "OK - " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & ", model " & IRRBB_DecayVersion()
    Exit Sub

Fail:
    primary = "FAILED - " & Err.Description & " [" & Err.Source & ", " & Err.Number & "]"
    On Error GoTo CleanupFail
    ClearOutputs
    ThisWorkbook.Worksheets(SH).Range("B14").Value = primary
    Exit Sub
CleanupFail:
    cleanup = " | cleanup error: " & Err.Description
    On Error Resume Next
    ThisWorkbook.Worksheets(SH).Range("B14").Value = primary & cleanup
End Sub


Public Sub ClearDecayExample()
    ClearOutputs
    ThisWorkbook.Worksheets(SH).Range("B14").Value = "NOT RUN"
End Sub


Private Sub ClearOutputs()
    With ThisWorkbook.Worksheets(SH)
        .Range("F3:G17").ClearContents
        .Range("I3:J13").ClearContents
        .Range("B15:B16").ClearContents
        .Range(.Cells(DATA_ROW, PROFILE_COL), .Cells(DATA_ROW + 1300, PROFILE_COL + 4)).ClearContents
    End With
End Sub
