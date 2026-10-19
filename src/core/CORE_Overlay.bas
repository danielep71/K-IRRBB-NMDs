Attribute VB_Name = "CORE_Overlay"
'==============================================================================
' CORE_Overlay
'------------------------------------------------------------------------------
' PURPOSE
'   Regulatory overlay on a fitted decay profile, kept separate from the fit
'   (MODEL_CONTRACTS.md, "Econometric fit vs regulatory overlay"). The overlay
'   never changes a fitted parameter; it returns a separate, constrained
'   profile labelled with source IDs and verification status.
'
' STEPS (in contract order)
'   1. Non-core amount B0 - MPA(1) in the overnight bucket, not in month 1
'      (REG-1 para. 114, verified).
'   2. Core-share cap k: if MPA(1)/B0 > k, scale MPA(h), h >= 1, by
'      k B0 / MPA(1); the difference moves to the overnight bucket
'      (REG-1 para. 113, Table 2, verified).
'   3. Average-maturity cap: if the core mean life exceeds the cap, reduce the
'      cutoff to the largest H' <= H that meets it, remainder at H'
'      (REG-1 para. 115, Table 2, verified).
'   The EU 5-year cap on NMD average repricing maturity (REG-2) is NOT
'   applied: its source is unverified, so it is reported as provisional.
'
' OUTPUT
'   overlay(0 To H): bucket 0 = overnight, bucket h = month h; buckets after
'   H' are zero; the profile sums to B0. Core mean life in months as in the
'   fit, computed on the overlay core path C(h) = factor * MPA(h).
'==============================================================================
Option Explicit
Option Private Module

Public Const CORE_OV_CORE_SHARE_FIT As Long = 1
Public Const CORE_OV_CORE_SHARE_CAP As Long = 2
Public Const CORE_OV_SCALE As Long = 3
Public Const CORE_OV_SHARE_APPLIED As Long = 4
Public Const CORE_OV_CORE_SHARE As Long = 5
Public Const CORE_OV_MAT_CAP As Long = 6
Public Const CORE_OV_CORE_MEANLIFE_IN As Long = 7
Public Const CORE_OV_CUTOFF As Long = 8
Public Const CORE_OV_MAT_APPLIED As Long = 9
Public Const CORE_OV_CORE_MEANLIFE As Long = 10
Public Const CORE_OV_OVERNIGHT As Long = 11
Public Const CORE_OV_COUNT As Long = 11


'------------------------------------------------------------------------------
' CORE_Reg1ShareCap / CORE_Reg1MaturityCap
'   REG-1 Table 2 caps for a data-contract segment code: the core-share cap
'   as a decimal share and the average-maturity cap in months. Unknown codes
'   raise CORE_ERR_INPUT.
'------------------------------------------------------------------------------
Public Function CORE_Reg1ShareCap(ByVal segment As String) As Double
    Select Case segment
        Case "RET_TX": CORE_Reg1ShareCap = 0.9
        Case "RET_NTX": CORE_Reg1ShareCap = 0.7
        Case "WHS_NFC": CORE_Reg1ShareCap = 0.5
        Case Else
            Err.Raise CORE_ERR_INPUT, "CORE_Overlay.CORE_Reg1ShareCap", "Unknown segment code '" & segment & "'."
    End Select
End Function

Public Function CORE_Reg1MaturityCap(ByVal segment As String) As Double
    Select Case segment
        Case "RET_TX": CORE_Reg1MaturityCap = 60
        Case "RET_NTX": CORE_Reg1MaturityCap = 54
        Case "WHS_NFC": CORE_Reg1MaturityCap = 48
        Case Else
            Err.Raise CORE_ERR_INPUT, "CORE_Overlay.CORE_Reg1MaturityCap", "Unknown segment code '" & segment & "'."
    End Select
End Function


'------------------------------------------------------------------------------
' CORE_Overlay
'   segment        RET_TX, RET_NTX or WHS_NFC                               [in]
'   mpa(0 To H)    fitted MPA path, mpa(0) = B0 > 0                         [in]
'   overlay(0..H), ovInfo(1..11), ovLabel                                      [out]
'------------------------------------------------------------------------------
Public Sub CORE_DecayOverlay(ByVal segment As String, ByRef mpa() As Double, ByRef overlay() As Double, _
                             ByRef ovInfo() As Double, ByRef ovLabel As String)
    Dim cutoff As Long, h As Long, hc As Long, shareCap As Double, matCap As Double
    Dim b0 As Double, factor As Double, c() As Double, partial As Double, lifeIn As Double, lifeOut As Double

    If LBound(mpa) <> 0 Or UBound(mpa) < 2 Then
        Err.Raise CORE_ERR_INPUT, "CORE_Overlay.CORE_DecayOverlay", "MPA path must be (0 To H) with H >= 2."
    End If
    cutoff = UBound(mpa)
    b0 = mpa(0)
    If Not (b0 > 0) Or Not (mpa(1) > 0) Then
        Err.Raise CORE_ERR_DOMAIN, "CORE_Overlay.CORE_DecayOverlay", "B0 and MPA(1) must be positive."
    End If
    For h = 1 To cutoff
        If mpa(h) > mpa(h - 1) Or mpa(h) < 0 Then
            Err.Raise CORE_ERR_INVARIANT, "CORE_Overlay.CORE_DecayOverlay", "MPA path is not non-increasing."
        End If
    Next h
    shareCap = CORE_Reg1ShareCap(segment)
    matCap = CORE_Reg1MaturityCap(segment)

    ReDim ovInfo(1 To CORE_OV_COUNT)
    ovInfo(CORE_OV_CORE_SHARE_FIT) = mpa(1) / b0
    ovInfo(CORE_OV_CORE_SHARE_CAP) = shareCap
    ovInfo(CORE_OV_MAT_CAP) = matCap

    ' Step 2: core-share cap.
    factor = 1
    If mpa(1) / b0 > shareCap Then
        factor = shareCap * b0 / mpa(1)
        ovInfo(CORE_OV_SHARE_APPLIED) = 1
    End If
    ReDim c(0 To cutoff)
    For h = 1 To cutoff
        c(h) = factor * mpa(h)
    Next h
    ovInfo(CORE_OV_SCALE) = factor
    ovInfo(CORE_OV_CORE_SHARE) = c(1) / b0

    ' Step 3: average-maturity cap on the core mean life.
    lifeIn = CoreMeanLife(c, cutoff)
    ovInfo(CORE_OV_CORE_MEANLIFE_IN) = lifeIn
    hc = cutoff
    If lifeIn > matCap Then
        ovInfo(CORE_OV_MAT_APPLIED) = 1
        Do While hc > 1
            hc = hc - 1
            If CoreMeanLife(c, hc) <= matCap Then Exit Do
        Loop
    End If
    lifeOut = CoreMeanLife(c, hc)
    ovInfo(CORE_OV_CUTOFF) = hc
    ovInfo(CORE_OV_CORE_MEANLIFE) = lifeOut

    ' Profile: step 1 puts the non-core (and scaled-away) amount overnight.
    ReDim overlay(0 To cutoff)
    overlay(0) = b0 - c(1)
    For h = 2 To hc
        overlay(h) = c(h - 1) - c(h)
    Next h
    overlay(hc) = overlay(hc) + c(hc)
    ovInfo(CORE_OV_OVERNIGHT) = overlay(0)

    partial = 0
    For h = 0 To cutoff
        partial = partial + overlay(h)
    Next h
    If Abs(partial - b0) > 0.000000001 * b0 Then
        Err.Raise CORE_ERR_INVARIANT, "CORE_Overlay.CORE_DecayOverlay", "Overlay profile does not conserve the notional."
    End If

    ovLabel = "Overlay " & segment & ": non-core overnight [REG-1 para.114, verified]; core-share cap " & _
            Format$(shareCap * 100, "0") & "% [REG-1 para.113 Table 2, verified, " & _
            IIf(ovInfo(CORE_OV_SHARE_APPLIED) = 1, "applied", "not binding") & "]; average-maturity cap " & _
            Format$(matCap, "0") & " months [REG-1 para.115 Table 2, verified, " & _
            IIf(ovInfo(CORE_OV_MAT_APPLIED) = 1, "applied", "not binding") & _
            "]; EU 5-year NMD repricing cap [REG-2, unverified]: provisional, not applied."
End Sub


'------------------------------------------------------------------------------
' CoreMeanLife: (sum_{h=2..hc} h (c(h-1) - c(h)) + hc c(hc)) / c(1), months.
'------------------------------------------------------------------------------
Private Function CoreMeanLife(ByRef c() As Double, ByVal hc As Long) As Double
    Dim h As Long, s As Double
    For h = 2 To hc
        s = s + h * (c(h - 1) - c(h))
    Next h
    CoreMeanLife = (s + hc * c(hc)) / c(1)
End Function
