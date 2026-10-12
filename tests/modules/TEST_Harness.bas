Attribute VB_Name = "TEST_Harness"
'==============================================================================
' TEST_Harness  (in-Excel regression runner, issue #8)
'------------------------------------------------------------------------------
' PURPOSE
'   Runs every case registered in docs/methodology/TEST_CASES.md that has a VBA
'   implementation, with deterministic case IDs and assertion counts, and lists
'   every other registered case as NOT RUN. A failure is never shown as PASS.
'   The registry below mirrors the register in order; tools/test_vba_harness.py
'   fails if they differ. Case modules are reached through DispatchCase.
'
' OUTCOMES
'   PASS     every assertion of the case passed (and at least one ran)
'   FAIL     at least one assertion failed, or the case made no assertion
'   ERROR    the case raised an unexpected error (recorded as an assertion)
'   NOT RUN  registered, but no VBA implementation in this build
'
' STALE RESULTS AND ERRORS
'   Previous results are withdrawn before the run (status RUNNING). If the run
'   itself fails, the primary error and any cleanup error are both kept in the
'   status block (B10, B11) with status FAILED. TEST_RunTestsWithWriteFault
'   exercises that path by making every results write fail.
'
' OUTPUT
'   Sheet "Checks": status block (rows 3-11), case table (A13:F), assertion
'   table (H13:N). Tolerances follow TEST_CASES.md: abs |x - r| <= eps,
'   rel |x - r| <= eps * max(1, |r|), exact x = r.
'==============================================================================
Option Explicit

' Raised by a case module asked for a case it does not implement:
' vbObjectError + 1000, written as a literal because LibreOffice, the early
' warning proxy, does not compile the constant expression.
Public Const TEST_ERR_NO_CASE As Long = -2147220504

Private Const SH_CHECKS As String = "Checks"
Private Const SH_README As String = "Readme"
Private Const CASE_ROW As Long = 14
Private Const ASSERT_COL As Long = 8
' vbObjectError + 1001, a literal for the same reason as TEST_ERR_NO_CASE.
Private Const ERR_INJECTED_WRITE As Long = -2147220503

Private mCaseId() As String
Private mCaseOutcome() As String
Private mCaseNote() As String
Private mCaseAsserts() As Long
Private mCasePassed() As Long
Private mCaseCount As Long
Private mCurrent As Long

Private mRows() As Variant
Private mRowCount As Long
Private mStarted As Date
Private mInjectWriteFault As Boolean


'==============================================================================
' Entry point
'==============================================================================
Public Sub RunTests()
    Dim primary As String, cleanup As String, ids As Variant, i As Long
    On Error GoTo Fail
    mStarted = Now
    MarkStale
    ResetState

    ids = Registry()
    For i = LBound(ids) To UBound(ids)
        RunOne CStr(ids(i)(0)), CBool(ids(i)(1)), CStr(ids(i)(2))
    Next i
    WriteResults ""
    Exit Sub

Fail:
    ' Resume clears the error first: an error raised while a handler is still
    ' active is not trapped in this procedure and would escape with the status
    ' left at RUNNING and neither error recorded.
    primary = "Primary error " & Err.Number & ": " & Err.Description & " [" & Err.Source & "]"
    Resume ReportPrimary
ReportPrimary:
    On Error GoTo CleanupFail
    WriteResults primary
    Exit Sub
CleanupFail:
    cleanup = "Cleanup error " & Err.Number & ": " & Err.Description & " [" & Err.Source & "]"
    Resume ReportBoth
ReportBoth:
    On Error Resume Next
    With ThisWorkbook.Worksheets(SH_CHECKS)
        .Cells(3, 2).Value = "FAILED"
        .Cells(10, 2).Value = primary
        .Cells(11, 2).Value = cleanup
    End With
    Debug.Print primary
    Debug.Print cleanup
End Sub


'------------------------------------------------------------------------------
' TEST_RunTestsWithWriteFault: run the harness with every results write failing,
' to show that the primary and the cleanup error are both reported and that no
' earlier result is left looking current. Leaves status FAILED.
'------------------------------------------------------------------------------
Public Sub TEST_RunTestsWithWriteFault()
    mInjectWriteFault = True
    On Error GoTo Done          ' RunTests reports its own errors; this only resets the flag
    RunTests
Done:
    mInjectWriteFault = False
End Sub


'------------------------------------------------------------------------------
' Registry: (case ID, has VBA implementation, note). Order is the register's.
'------------------------------------------------------------------------------
Private Function Registry() As Variant
    ' Built as a Collection so the list never needs more than 24 line continuations.
    Dim c As New Collection, out() As Variant, i As Long
    Const DECAY_NOTE As String = "Decay model not on this branch (#36, v0.3.0)"
    c.Add Array("HARN-OUT-01", True, "")
    c.Add Array("WB-LAYOUT-01", True, "")
    c.Add Array("DATA-VALID-01", True, "")
    c.Add Array("DATA-INVALID-01", True, "")
    c.Add Array("DATA-PIT-01", True, "")
    c.Add Array("DATA-PIT-02", True, "")
    c.Add Array("DATA-IMPORT-01", True, "")
    c.Add Array("DATA-PARSE-01", True, "")
    c.Add Array("DATA-PARSE-02", True, "")
    c.Add Array("DATA-CODES-01", True, "")
    c.Add Array("RATE-OLS-01", False, "Rate model not built")
    c.Add Array("RATE-ADF-01", False, "Rate model not built")
    c.Add Array("RATE-EG-01", False, "Rate model not built")
    c.Add Array("RATE-ECM-01", False, "Rate model not built")
    c.Add Array("RATE-ASYM-01", False, "Rate model not built")
    c.Add Array("RATE-HAC-01", False, "Rate model not built")
    c.Add Array("RATE-CONS-01", False, "Rate model not built")
    c.Add Array("RATE-SIM-01", False, "Rate model not built")
    c.Add Array("STAB-CASH-01", False, "Stable-amount model not built")
    c.Add Array("STAB-LOGIT-01", False, "Stable-amount model not built")
    c.Add Array("STAB-ROBUST-01", False, "Stable-amount model not built")
    c.Add Array("STAB-SEP-01", False, "Stable-amount model not built")
    c.Add Array("STAB-CONV-01", False, "Stable-amount model not built")
    c.Add Array("STAB-LEAK-01", False, "Stable-amount model not built")
    c.Add Array("DEC-LOG-01", False, DECAY_NOTE)
    c.Add Array("DEC-MPA-01", False, DECAY_NOTE)
    c.Add Array("DEC-PROF-01", False, DECAY_NOTE)
    c.Add Array("DEC-LIFE-01", False, DECAY_NOTE)
    c.Add Array("DEC-CAP-01", False, DECAY_NOTE)
    c.Add Array("OOS-FREEZE-01", False, DECAY_NOTE)
    c.Add Array("OOS-LEAK-01", False, DECAY_NOTE)
    c.Add Array("OOS-SCORE-01", False, DECAY_NOTE)
    ReDim out(0 To c.Count - 1)
    For i = 1 To c.Count
        out(i - 1) = c(i)
    Next i
    Registry = out
End Function


Private Sub RunOne(ByVal caseId As String, ByVal built As Boolean, ByVal note As String)
    Dim errored As Boolean
    mCaseCount = mCaseCount + 1
    If mCaseCount > UBound(mCaseId) Then
        GrowCases
    End If
    mCurrent = mCaseCount
    mCaseId(mCurrent) = caseId
    mCaseNote(mCurrent) = note
    If built Then
        On Error GoTo CaseErr
        DispatchCase caseId
        On Error GoTo 0
    End If

Finish:
    mCaseOutcome(mCurrent) = TEST_CaseOutcome(built, errored, mCaseAsserts(mCurrent), mCasePassed(mCurrent))
    If built And Not errored And mCaseAsserts(mCurrent) = 0 Then
        mCaseNote(mCurrent) = Trim$(note & " No assertion was executed.")
    End If
    Exit Sub

CaseErr:
    AddRow caseId, "Unexpected error " & Err.Number & ": " & Err.Description & " [" & Err.Source & "]", _
           Empty, Empty, Empty, "", "ERROR"
    mCaseAsserts(mCurrent) = mCaseAsserts(mCurrent) + 1
    errored = True
    Resume Finish
End Sub


'------------------------------------------------------------------------------
' DispatchCase: route a built case to the module that implements it.
'------------------------------------------------------------------------------
Private Sub DispatchCase(ByVal caseId As String)
    Select Case caseId
        Case "HARN-OUT-01", "DATA-PARSE-01", "DATA-PARSE-02", "DATA-CODES-01"
            TEST_CoreCases.TEST_RunCoreCase caseId
        Case "WB-LAYOUT-01"
            TEST_WorkbookCases.TEST_RunWorkbookCase caseId
        Case "DATA-VALID-01", "DATA-INVALID-01", "DATA-PIT-01", "DATA-PIT-02"
            TEST_DataCases.TEST_RunDataCase caseId
        Case "DATA-IMPORT-01"
            TEST_ImportCases.TEST_RunImportCase caseId
        Case Else
            Err.Raise TEST_ERR_NO_CASE, "TEST_Harness.DispatchCase", "No VBA implementation for " & caseId & "."
    End Select
End Sub


'==============================================================================
' Outcome and tolerance rules (pure; tested by HARN-OUT-01)
'==============================================================================

'------------------------------------------------------------------------------
' TEST_CaseOutcome: NOT RUN if the case is not built; ERROR if it raised an
' unexpected error; FAIL if it made no assertion or any assertion failed;
' PASS only if every one of at least one assertion passed.
'------------------------------------------------------------------------------
Public Function TEST_CaseOutcome(ByVal built As Boolean, ByVal errored As Boolean, _
                                 ByVal asserts As Long, ByVal passed As Long) As String
    If Not built Then
        TEST_CaseOutcome = "NOT RUN"
    ElseIf errored Then
        TEST_CaseOutcome = "ERROR"
    ElseIf asserts = 0 Then
        TEST_CaseOutcome = "FAIL"
    ElseIf passed = asserts Then
        TEST_CaseOutcome = "PASS"
    Else
        TEST_CaseOutcome = "FAIL"
    End If
End Function


'------------------------------------------------------------------------------
' TEST_WithinTolerance (TEST_CASES.md): abs |x - r| <= eps;
' rel |x - r| <= eps * max(1, |r|); exact x = r. Any other kind is False.
'------------------------------------------------------------------------------
Public Function TEST_WithinTolerance(ByVal actual As Double, ByVal expected As Double, _
                                     ByVal kind As String, ByVal eps As Double) As Boolean
    Dim diff As Double
    diff = Abs(actual - expected)
    Select Case kind
        Case "abs"
            TEST_WithinTolerance = (diff <= eps)
        Case "rel"
            If Abs(expected) > 1 Then
                TEST_WithinTolerance = (diff <= eps * Abs(expected))
            Else
                TEST_WithinTolerance = (diff <= eps)
            End If
        Case "exact"
            TEST_WithinTolerance = (actual = expected)
        Case Else
            TEST_WithinTolerance = False
    End Select
End Function


'==============================================================================
' Assertions (called by TEST_* case modules)
'==============================================================================
Public Sub TEST_AssertNear(ByVal caseId As String, ByVal label As String, ByVal actual As Double, _
                           ByVal expected As Double, ByVal kind As String, ByVal eps As Double)
    Dim ok As Boolean
    ok = TEST_WithinTolerance(actual, expected, kind, eps)
    If kind <> "abs" And kind <> "rel" And kind <> "exact" Then
        label = label & " (unknown tolerance kind '" & kind & "')"
    End If
    Record caseId, label, expected, actual, Abs(actual - expected), kind & " " & Format$(eps, "0.0E+00"), ok
End Sub


Public Sub TEST_AssertText(ByVal caseId As String, ByVal label As String, ByVal actual As String, _
                           ByVal expected As String)
    Record caseId, label, expected, actual, Empty, "exact", actual = expected
End Sub


Public Sub TEST_AssertTrue(ByVal caseId As String, ByVal label As String, ByVal condition As Boolean)
    Record caseId, label, True, condition, Empty, "exact", condition
End Sub


Public Sub TEST_AssertError(ByVal caseId As String, ByVal label As String, ByVal actualNumber As Long, _
                            ByVal expectedNumber As Long)
    Record caseId, label, expectedNumber, actualNumber, Empty, "error number", actualNumber = expectedNumber
End Sub


Private Sub Record(ByVal caseId As String, ByVal label As String, ByVal expected As Variant, _
                   ByVal actual As Variant, ByVal diff As Variant, ByVal tolText As String, ByVal ok As Boolean)
    If mCurrent < 1 Then Exit Sub
    If caseId <> mCaseId(mCurrent) Then
        label = label & " (asserted for " & caseId & " while running " & mCaseId(mCurrent) & ")"
        ok = False
    End If
    mCaseAsserts(mCurrent) = mCaseAsserts(mCurrent) + 1
    If ok Then mCasePassed(mCurrent) = mCasePassed(mCurrent) + 1
    If ok Then
        AddRow caseId, label, expected, actual, diff, tolText, "PASS"
    Else
        AddRow caseId, label, expected, actual, diff, tolText, "FAIL"
    End If
End Sub


'==============================================================================
' State and output
'==============================================================================
Private Sub ResetState()
    ReDim mCaseId(1 To 64)
    ReDim mCaseOutcome(1 To 64)
    ReDim mCaseNote(1 To 64)
    ReDim mCaseAsserts(1 To 64)
    ReDim mCasePassed(1 To 64)
    mCaseCount = 0
    mCurrent = 0
    ReDim mRows(1 To 7, 1 To 1024)
    mRowCount = 0
End Sub


Private Sub GrowCases()
    Dim n As Long
    n = 2 * UBound(mCaseId)
    ReDim Preserve mCaseId(1 To n)
    ReDim Preserve mCaseOutcome(1 To n)
    ReDim Preserve mCaseNote(1 To n)
    ReDim Preserve mCaseAsserts(1 To n)
    ReDim Preserve mCasePassed(1 To n)
End Sub


Private Sub AddRow(ByVal caseId As String, ByVal label As String, ByVal expected As Variant, _
                   ByVal actual As Variant, ByVal diff As Variant, ByVal tolText As String, ByVal outcome As String)
    mRowCount = mRowCount + 1
    If mRowCount > UBound(mRows, 2) Then
        ReDim Preserve mRows(1 To 7, 1 To 2 * UBound(mRows, 2))
    End If
    mRows(1, mRowCount) = caseId
    mRows(2, mRowCount) = label
    mRows(3, mRowCount) = expected
    mRows(4, mRowCount) = actual
    mRows(5, mRowCount) = diff
    mRows(6, mRowCount) = tolText
    mRows(7, mRowCount) = outcome
End Sub


Private Sub MarkStale()
    With ThisWorkbook.Worksheets(SH_CHECKS)
        .Range(.Cells(CASE_ROW, 1), .Cells(CASE_ROW + 500, 6)).ClearContents
        .Range(.Cells(CASE_ROW, ASSERT_COL), .Cells(CASE_ROW + 20000, ASSERT_COL + 6)).ClearContents
        .Range("B3:B11").ClearContents
        .Cells(3, 2).Value = "RUNNING - previous results withdrawn"
        .Cells(4, 2).Value = mStarted
    End With
End Sub


Private Sub WriteResults(ByVal primary As String)
    Dim ws As Worksheet, i As Long, j As Long, block() As Variant
    Dim nPass As Long, nFail As Long, nErr As Long, nNot As Long, aPass As Long, aTotal As Long
    Set ws = ThisWorkbook.Worksheets(SH_CHECKS)
    If mInjectWriteFault Then
        Err.Raise ERR_INJECTED_WRITE, "TEST_Harness.WriteResults", "Injected results write fault"
    End If

    If mCaseCount > 0 Then
        ReDim block(1 To mCaseCount, 1 To 6)
        For i = 1 To mCaseCount
            block(i, 1) = mCaseId(i)
            block(i, 2) = mCaseOutcome(i)
            block(i, 3) = mCaseAsserts(i)
            block(i, 4) = mCasePassed(i)
            block(i, 5) = mCaseAsserts(i) - mCasePassed(i)
            block(i, 6) = mCaseNote(i)
            Select Case mCaseOutcome(i)
                Case "PASS": nPass = nPass + 1
                Case "FAIL": nFail = nFail + 1
                Case "ERROR": nErr = nErr + 1
                Case Else: nNot = nNot + 1
            End Select
            aPass = aPass + mCasePassed(i)
            aTotal = aTotal + mCaseAsserts(i)
        Next i
        ws.Range(ws.Cells(CASE_ROW, 1), ws.Cells(CASE_ROW + mCaseCount - 1, 6)).Value = block
    End If
    If mRowCount > 0 Then
        ReDim block(1 To mRowCount, 1 To 7)
        For i = 1 To mRowCount
            For j = 1 To 7
                block(i, j) = mRows(j, i)
            Next j
        Next i
        ws.Range(ws.Cells(CASE_ROW, ASSERT_COL), ws.Cells(CASE_ROW + mRowCount - 1, ASSERT_COL + 6)).Value = block
    End If

    If Len(primary) > 0 Then
        ws.Cells(3, 2).Value = "FAILED"
    ElseIf nFail + nErr > 0 Then
        ws.Cells(3, 2).Value = "COMPLETE - FAILURES"
    Else
        ws.Cells(3, 2).Value = "COMPLETE - ALL EXECUTED CASES PASSED"
    End If
    ws.Cells(4, 2).Value = mStarted
    ws.Cells(5, 2).Value = Now
    ws.Cells(6, 2).Value = SourceCommit()
    ws.Cells(7, 2).Value = HostText()
    ws.Cells(8, 2).Value = nPass & " PASS, " & nFail & " FAIL, " & nErr & " ERROR, " & nNot & " NOT RUN (of " & _
                           mCaseCount & " registered)"
    ws.Cells(9, 2).Value = aPass & " of " & aTotal & " assertions passed"
    ws.Cells(10, 2).Value = primary
End Sub


Private Function SourceCommit() As String
    Dim r As Long
    SourceCommit = "unknown"
    On Error GoTo Done
    With ThisWorkbook.Worksheets(SH_README)
        For r = 1 To 40
            If CStr(.Cells(r, 1).Value) = "Source commit" Then
                SourceCommit = CStr(.Cells(r, 2).Value)
                Exit For
            End If
        Next r
    End With
Done:
End Function


Private Function HostText() As String
    HostText = "unknown"
    On Error Resume Next
    HostText = Application.Name & " " & Application.Version & " (build " & Application.Build & "), " & _
               Application.OperatingSystem
End Function


'------------------------------------------------------------------------------
' TEST_Summary: "cases passed / executed" of the last run, for automation.
'------------------------------------------------------------------------------
Public Function TEST_Summary() As String
    Dim i As Long, nPass As Long, nRun As Long
    For i = 1 To mCaseCount
        If mCaseOutcome(i) <> "NOT RUN" Then nRun = nRun + 1
        If mCaseOutcome(i) = "PASS" Then nPass = nPass + 1
    Next i
    TEST_Summary = nPass & "/" & nRun
End Function
