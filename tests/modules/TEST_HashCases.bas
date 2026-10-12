Attribute VB_Name = "TEST_HashCases"
'==============================================================================
' TEST_HashCases
'------------------------------------------------------------------------------
' PURPOSE
'   VBA implementation of DATA-SHA-01 (docs/methodology/TEST_CASES.md):
'   CORE_Sha256 against the published FIPS 180-2 examples ("abc", the 448- and
'   896-bit messages), the empty message, the padding
'   boundaries (55, 56, 63, 64, 65, 119, 120 bytes) and the same message fed in
'   pieces of several sizes. tools/test_vba_harness.py re-derives every digest
'   with Python's hashlib. The one-million-"a" example is left out: it takes
'   minutes in the LibreOffice proxy run on every harness run.
'
' FORMAT
'   One Expect* call per line, arguments as literals.
'==============================================================================
Option Explicit
Option Private Module


Public Sub TEST_RunHashCase(ByVal caseId As String)
    Select Case caseId
        Case "DATA-SHA-01": CaseSha256
        Case Else
            Err.Raise TEST_ERR_NO_CASE, "TEST_HashCases.TEST_RunHashCase", "No VBA implementation for " & caseId & "."
    End Select
End Sub


Private Sub CaseSha256()
    Const C As String = "DATA-SHA-01"
    ExpectSha C, "", "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
    ExpectSha C, "abc", "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    ExpectSha C, "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq", "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1"
    ExpectSha C, "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu", "cf5b16a778af8380036ce59e7b0492370b249b11e8f07a51afac45037afee9d1"
    ExpectShaRepeated C, "a", 55, "9f4390f8d30c2dd92ec9f095b65e2b9ae9b0a925a5258e241c9f1e910f734318"
    ExpectShaRepeated C, "a", 56, "b35439a4ac6f0948b6d6f9e3c6af0f5f590ce20f1bde7090ef7970686ec6738a"
    ExpectShaRepeated C, "a", 63, "7d3e74a05d7db15bce4ad9ec0658ea98e3f06eeecf16b4c6fff2da457ddc2f34"
    ExpectShaRepeated C, "a", 64, "ffe054fe7ae0cb6dc65c3af9b61d5209f439851db43d0ba5997337df154668eb"
    ExpectShaRepeated C, "a", 65, "635361c48bb9eab14198e76ea8ab7f1a41685d6ad62aa9146d301d4f17eb0ae0"
    ExpectShaRepeated C, "a", 119, "31eba51c313a5c08226adf18d4a359cfdfd8d2e816b13f4af952f7ea6584dcfb"
    ExpectShaRepeated C, "a", 120, "2f3d335432c70b580af0e8e1b3674a7c020d683aa5f73aaaedfdc55af904c21c"
    ExpectShaPieces C, "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu", 1, "cf5b16a778af8380036ce59e7b0492370b249b11e8f07a51afac45037afee9d1"
    ExpectShaPieces C, "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu", 7, "cf5b16a778af8380036ce59e7b0492370b249b11e8f07a51afac45037afee9d1"
    ExpectShaPieces C, "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu", 64, "cf5b16a778af8380036ce59e7b0492370b249b11e8f07a51afac45037afee9d1"
    ExpectShaPieces C, "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu", 65, "cf5b16a778af8380036ce59e7b0492370b249b11e8f07a51afac45037afee9d1"
End Sub


'==============================================================================
' Expectation helpers
'==============================================================================
Private Sub ExpectSha(ByVal caseId As String, ByVal message As String, ByVal expected As String)
    TEST_AssertText caseId, "SHA-256 of a " & Len(message) & "-byte message", CORE_Sha256Text(message), expected
End Sub


' A message of count copies of one character, fed in 4,096-byte pieces.
Private Sub ExpectShaRepeated(ByVal caseId As String, ByVal unit As String, ByVal count As Long, _
                              ByVal expected As String)
    Dim piece() As Byte, i As Long, remaining As Long, size As Long
    ReDim piece(0 To 4095)
    For i = 0 To 4095
        piece(i) = CByte(AscW(unit))
    Next i
    CORE_Sha256Begin
    remaining = count
    Do While remaining > 0
        size = 4096
        If remaining < size Then size = remaining
        CORE_Sha256Update piece, size
        remaining = remaining - size
    Loop
    TEST_AssertText caseId, "SHA-256 of " & count & " x """ & unit & """", CORE_Sha256End(), expected
End Sub


' The message fed in pieces of pieceSize bytes gives the same digest.
Private Sub ExpectShaPieces(ByVal caseId As String, ByVal message As String, ByVal pieceSize As Long, _
                            ByVal expected As String)
    Dim piece() As Byte, start As Long, size As Long, i As Long
    CORE_Sha256Begin
    start = 1
    Do While start <= Len(message)
        size = pieceSize
        If start + size - 1 > Len(message) Then size = Len(message) - start + 1
        ReDim piece(0 To size - 1)
        For i = 0 To size - 1
            piece(i) = CByte(AscW(Mid$(message, start + i, 1)))
        Next i
        CORE_Sha256Update piece, size
        start = start + size
    Loop
    TEST_AssertText caseId, "SHA-256 in " & pieceSize & "-byte pieces", CORE_Sha256End(), expected
End Sub
