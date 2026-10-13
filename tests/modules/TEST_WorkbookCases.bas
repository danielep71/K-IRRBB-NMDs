Attribute VB_Name = "TEST_WorkbookCases"
'==============================================================================
' TEST_WorkbookCases
'------------------------------------------------------------------------------
' PURPOSE
'   VBA implementation of WB-LAYOUT-01 (docs/methodology/TEST_CASES.md): the
'   workbook built from src/workbook/IRRBB_Template.xlsx has the sheets,
'   CodeNames and defined names of docs/WORKBOOK_CONTRACT.md, read through the
'   Excel object model of the host that runs it. The literals below are checked
'   against tools/workbook/build_template.py by tools/test_vba_harness.py.
'
' FORMAT
'   One Expect* call per line, arguments as literals.
'==============================================================================
Option Explicit
Option Private Module


Public Sub TEST_RunWorkbookCase(ByVal caseId As String)
    Select Case caseId
        Case "WB-LAYOUT-01": CaseLayout
        Case Else
            Err.Raise TEST_ERR_NO_CASE, "TEST_WorkbookCases.TEST_RunWorkbookCase", "No VBA implementation for " & caseId & "."
    End Select
End Sub


'------------------------------------------------------------------------------
' WB-LAYOUT-01: sheet order, names and CodeNames; every defined name refers to
' its contract address; nothing more is defined.
'------------------------------------------------------------------------------
Private Sub CaseLayout()
    Const C As String = "WB-LAYOUT-01"
    ExpectCounts C, 10, 23

    ExpectSheet C, 1, "Readme", "shReadme"
    ExpectSheet C, 2, "Config", "shConfig"
    ExpectSheet C, 3, "Segments", "shSegments"
    ExpectSheet C, 4, "DataSummary", "shDataSummary"
    ExpectSheet C, 5, "ModelParams", "shModelParams"
    ExpectSheet C, 6, "Calibration", "shCalibration"
    ExpectSheet C, 7, "Backtesting", "shBacktesting"
    ExpectSheet C, 8, "Scenarios", "shScenarios"
    ExpectSheet C, 9, "Results", "shResults"
    ExpectSheet C, 10, "Checks", "shChecks"

    ExpectName C, "Backtesting_Status", "Backtesting", "$B$3"
    ExpectName C, "Calibration_Status", "Calibration", "$B$3"
    ExpectName C, "Checks_Assertions", "Checks", "$H$14:$N$20014"
    ExpectName C, "Checks_Cases", "Checks", "$A$14:$F$514"
    ExpectName C, "Checks_Status", "Checks", "$B$3"
    ExpectName C, "Config_AccountsFile", "Config", "$B$6"
    ExpectName C, "Config_ClosureAttested", "Config", "$B$11"
    ExpectName C, "Config_RateSeries", "Config", "$B$10"
    ExpectName C, "Config_RatesFile", "Config", "$B$7"
    ExpectName C, "Config_WindowEnd", "Config", "$B$9"
    ExpectName C, "Config_WindowStart", "Config", "$B$8"
    ExpectName C, "DataSummary_Findings", "DataSummary", "$G$14:$J$20013"
    ExpectName C, "DataSummary_Info", "DataSummary", "$B$4:$B$12"
    ExpectName C, "DataSummary_Lineage", "DataSummary", "$P$14:$Q$63"
    ExpectName C, "DataSummary_Status", "DataSummary", "$B$3"
    ExpectName C, "DataSummary_Totals", "DataSummary", "$A$14:$E$20013"
    ExpectName C, "DataSummary_Warnings", "DataSummary", "$L$14:$N$20013"
    ExpectName C, "ModelParams_Status", "ModelParams", "$B$3"
    ExpectName C, "Readme_SourceCommit", "Readme", "$B$5"
    ExpectName C, "Results_Status", "Results", "$B$3"
    ExpectName C, "Scenarios_Status", "Scenarios", "$B$3"
    ExpectName C, "Segments_Codes", "Segments", "$A$6:$A$8"
    ExpectName C, "Segments_Include", "Segments", "$C$6:$C$8"
End Sub


'==============================================================================
' Expectation helpers: one assertion each, labelled with the item checked.
'==============================================================================
Private Sub ExpectCounts(ByVal caseId As String, ByVal sheets As Long, ByVal names As Long)
    TEST_AssertText caseId, "worksheet count", CStr(ThisWorkbook.Worksheets.Count), CStr(sheets)
    TEST_AssertText caseId, "defined name count", CStr(ThisWorkbook.Names.Count), CStr(names)
End Sub


Private Sub ExpectSheet(ByVal caseId As String, ByVal index As Long, ByVal sheetName As String, _
                        ByVal codeName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(index)
    TEST_AssertText caseId, "sheet " & index & " name", ws.Name, sheetName
    TEST_AssertText caseId, "sheet " & sheetName & " CodeName", ws.CodeName, codeName
End Sub


Private Sub ExpectName(ByVal caseId As String, ByVal definedName As String, ByVal sheetName As String, _
                       ByVal address As String)
    Dim target As Range
    Set target = ThisWorkbook.Names(definedName).RefersToRange
    TEST_AssertText caseId, "name " & definedName, target.Worksheet.Name & "!" & target.Address, _
                    sheetName & "!" & address
End Sub
