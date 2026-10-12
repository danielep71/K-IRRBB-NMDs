Attribute VB_Name = "CORE_Sha256"
'==============================================================================
' CORE_Sha256
'------------------------------------------------------------------------------
' PURPOSE
'   SHA-256 (FIPS 180-4) of a byte stream, for the data fingerprint of the
'   import lineage (DATA_CONTRACT.md, decision 5: implemented in VBA, tested
'   against published test vectors). No library, COM object or declaration.
'
' USE
'   CORE_Sha256Begin, then CORE_Sha256Update for each piece of the stream in
'   order, then CORE_Sha256End for the digest as 64 lower-case hex digits.
'   Pieces may have any length; the digest depends only on the bytes.
'   CORE_Sha256Text hashes a string of characters 0-255, one byte each.
'
' ARITHMETIC
'   32-bit words are held in Long (two's complement). Shifts and rotations are
'   written out as mask expressions (the sign bit handled explicitly), and each
'   sum of up to five words is formed once in Double, exact below 2^53, and
'   reduced modulo 2^32 once: far fewer calls than one function per operation. The round constants are the first
'   32 bits of the fractional parts of the cube roots of the first 64 primes,
'   the initial hash those of the square roots of the first 8 primes.
'==============================================================================
Option Explicit
Option Private Module

Private mInitialized As Boolean
Private K(0 To 63) As Long
Private mH(0 To 7) As Long
Private mBlock(0 To 63) As Byte
Private mFill As Long
Private mBytes As Double
Private mActive As Boolean


'------------------------------------------------------------------------------
' CORE_Sha256Begin: start a new digest.
'------------------------------------------------------------------------------
Public Sub CORE_Sha256Begin()
    Initialize
    mH(0) = &H6A09E667: mH(1) = &HBB67AE85: mH(2) = &H3C6EF372: mH(3) = &HA54FF53A
    mH(4) = &H510E527F: mH(5) = &H9B05688C: mH(6) = &H1F83D9AB: mH(7) = &H5BE0CD19
    mFill = 0
    mBytes = 0
    mActive = True
End Sub


'------------------------------------------------------------------------------
' CORE_Sha256Update: add data(LBound(data)) .. data(LBound(data) + count - 1).
'------------------------------------------------------------------------------
Public Sub CORE_Sha256Update(ByRef data() As Byte, ByVal count As Long)
    Dim i As Long, last As Long
    If Not mActive Then Err.Raise 5, "CORE_Sha256.CORE_Sha256Update", "No digest has begun."
    If count < 0 Or count > UBound(data) - LBound(data) + 1 Then
        Err.Raise 5, "CORE_Sha256.CORE_Sha256Update", "count is outside the array."
    End If
    last = LBound(data) + count - 1
    For i = LBound(data) To last
        mBlock(mFill) = data(i)
        mFill = mFill + 1
        If mFill = 64 Then
            Compress
            mFill = 0
        End If
    Next i
    mBytes = mBytes + count
End Sub


'------------------------------------------------------------------------------
' CORE_Sha256End: pad, finish and return the digest (64 lower-case hex digits).
'------------------------------------------------------------------------------
Public Function CORE_Sha256End() As String
    Dim bitLength As Double, i As Long, text As String
    If Not mActive Then Err.Raise 5, "CORE_Sha256.CORE_Sha256End", "No digest has begun."
    bitLength = mBytes * 8
    mBlock(mFill) = &H80
    mFill = mFill + 1
    If mFill > 56 Then
        Do While mFill < 64
            mBlock(mFill) = 0
            mFill = mFill + 1
        Loop
        Compress
        mFill = 0
    End If
    Do While mFill < 56
        mBlock(mFill) = 0
        mFill = mFill + 1
    Loop
    ' Message length in bits, 64-bit big-endian; Double is exact below 2^53.
    For i = 63 To 56 Step -1
        mBlock(i) = CByte(bitLength - Int(bitLength / 256) * 256)
        bitLength = Int(bitLength / 256)
    Next i
    Compress
    For i = 0 To 7
        text = text & Right$("0000000" & Hex$(mH(i)), 8)
    Next i
    mActive = False
    CORE_Sha256End = LCase$(text)
End Function


'------------------------------------------------------------------------------
' CORE_Sha256Text: digest of a string whose characters are all 0-255 (one
' byte each); error 5 for any other character.
'------------------------------------------------------------------------------
Public Function CORE_Sha256Text(ByVal text As String) As String
    Dim bytes() As Byte, i As Long, code As Long
    CORE_Sha256Begin
    If Len(text) > 0 Then
        ReDim bytes(0 To Len(text) - 1)
        For i = 1 To Len(text)
            code = AscW(Mid$(text, i, 1))
            If code < 0 Or code > 255 Then
                Err.Raise 5, "CORE_Sha256.CORE_Sha256Text", "Character outside 0-255 at position " & i & "."
            End If
            bytes(i - 1) = CByte(code)
        Next i
        CORE_Sha256Update bytes, Len(text)
    End If
    CORE_Sha256Text = CORE_Sha256End()
End Function


'==============================================================================
' Compression of one 64-byte block
'==============================================================================
Private Sub Compress()
    Dim W(0 To 63) As Long, t As Long, a As Long, b As Long, c As Long, d As Long
    Dim e As Long, f As Long, g As Long, h As Long, t1 As Long, t2 As Long, s0 As Long, s1 As Long
    Dim x As Long, s As Double

    For t = 0 To 15
        W(t) = WordAt(4 * t)
    Next t
    For t = 16 To 63
        x = W(t - 15)
        s0 = (((((x) And &H7FFFFFFF) \ &H80&) Or (((x) < 0) And &H1000000&)) Or ((((x) And &H3F&) * &H2000000&) Or ((((x) And &H40&) <> 0) And &H80000000))) _
             Xor (((((x) And &H7FFFFFFF) \ &H40000&) Or (((x) < 0) And &H2000&)) Or ((((x) And &H1FFFF&) * &H4000&) Or ((((x) And &H20000&) <> 0) And &H80000000))) _
             Xor ((((x) And &H7FFFFFFF) \ &H8&) Or (((x) < 0) And &H10000000&))
        x = W(t - 2)
        s1 = (((((x) And &H7FFFFFFF) \ &H20000&) Or (((x) < 0) And &H4000&)) Or ((((x) And &HFFFF&) * &H8000&) Or ((((x) And &H10000&) <> 0) And &H80000000))) _
             Xor (((((x) And &H7FFFFFFF) \ &H80000&) Or (((x) < 0) And &H1000&)) Or ((((x) And &H3FFFF&) * &H2000&) Or ((((x) And &H40000&) <> 0) And &H80000000))) _
             Xor ((((x) And &H7FFFFFFF) \ &H400&) Or (((x) < 0) And &H200000&))
        s = CDbl(W(t - 16)) + s0 + W(t - 7) + s1
        s = s - Int((s + 2147483648#) / 4294967296#) * 4294967296#
        W(t) = CLng(s)
    Next t

    a = mH(0): b = mH(1): c = mH(2): d = mH(3)
    e = mH(4): f = mH(5): g = mH(6): h = mH(7)
    For t = 0 To 63
        s1 = (((((e) And &H7FFFFFFF) \ &H40&) Or (((e) < 0) And &H2000000&)) Or ((((e) And &H1F&) * &H4000000&) Or ((((e) And &H20&) <> 0) And &H80000000))) _
             Xor (((((e) And &H7FFFFFFF) \ &H800&) Or (((e) < 0) And &H100000&)) Or ((((e) And &H3FF&) * &H200000&) Or ((((e) And &H400&) <> 0) And &H80000000))) _
             Xor (((((e) And &H7FFFFFFF) \ &H2000000&) Or (((e) < 0) And &H40&)) Or ((((e) And &HFFFFFF&) * &H80&) Or ((((e) And &H1000000&) <> 0) And &H80000000)))
        s = CDbl(h) + s1 + ((e And f) Xor ((Not e) And g)) + K(t) + W(t)
        s = s - Int((s + 2147483648#) / 4294967296#) * 4294967296#
        t1 = CLng(s)
        s0 = (((((a) And &H7FFFFFFF) \ &H4&) Or (((a) < 0) And &H20000000&)) Or ((((a) And &H1&) * &H40000000&) Or ((((a) And &H2&) <> 0) And &H80000000))) _
             Xor (((((a) And &H7FFFFFFF) \ &H2000&) Or (((a) < 0) And &H40000&)) Or ((((a) And &HFFF&) * &H80000&) Or ((((a) And &H1000&) <> 0) And &H80000000))) _
             Xor (((((a) And &H7FFFFFFF) \ &H400000&) Or (((a) < 0) And &H200&)) Or ((((a) And &H1FFFFF&) * &H400&) Or ((((a) And &H200000&) <> 0) And &H80000000)))
        t2 = (a And b) Xor (a And c) Xor (b And c)
        h = g: g = f: f = e
        s = CDbl(d) + t1
        s = s - Int((s + 2147483648#) / 4294967296#) * 4294967296#
        e = CLng(s)
        d = c: c = b: b = a
        s = CDbl(t1) + s0 + t2
        s = s - Int((s + 2147483648#) / 4294967296#) * 4294967296#
        a = CLng(s)
    Next t
    s = CDbl(mH(0)) + a
    s = s - Int((s + 2147483648#) / 4294967296#) * 4294967296#
    mH(0) = CLng(s)
    s = CDbl(mH(1)) + b
    s = s - Int((s + 2147483648#) / 4294967296#) * 4294967296#
    mH(1) = CLng(s)
    s = CDbl(mH(2)) + c
    s = s - Int((s + 2147483648#) / 4294967296#) * 4294967296#
    mH(2) = CLng(s)
    s = CDbl(mH(3)) + d
    s = s - Int((s + 2147483648#) / 4294967296#) * 4294967296#
    mH(3) = CLng(s)
    s = CDbl(mH(4)) + e
    s = s - Int((s + 2147483648#) / 4294967296#) * 4294967296#
    mH(4) = CLng(s)
    s = CDbl(mH(5)) + f
    s = s - Int((s + 2147483648#) / 4294967296#) * 4294967296#
    mH(5) = CLng(s)
    s = CDbl(mH(6)) + g
    s = s - Int((s + 2147483648#) / 4294967296#) * 4294967296#
    mH(6) = CLng(s)
    s = CDbl(mH(7)) + h
    s = s - Int((s + 2147483648#) / 4294967296#) * 4294967296#
    mH(7) = CLng(s)
End Sub


' Big-endian 32-bit word of mBlock(i .. i + 3).
Private Function WordAt(ByVal i As Long) As Long
    Dim top As Long
    top = mBlock(i)
    WordAt = ((top And &H7F&) * &H1000000) Or (CLng(mBlock(i + 1)) * &H10000) Or (CLng(mBlock(i + 2)) * &H100&) _
             Or CLng(mBlock(i + 3))
    If top >= 128 Then WordAt = WordAt Or &H80000000
End Function


Private Sub Initialize()
    If mInitialized Then Exit Sub
    K(0) = &H428A2F98&: K(1) = &H71374491&: K(2) = &HB5C0FBCF&: K(3) = &HE9B5DBA5&
    K(4) = &H3956C25B&: K(5) = &H59F111F1&: K(6) = &H923F82A4&: K(7) = &HAB1C5ED5&
    K(8) = &HD807AA98&: K(9) = &H12835B01&: K(10) = &H243185BE&: K(11) = &H550C7DC3&
    K(12) = &H72BE5D74&: K(13) = &H80DEB1FE&: K(14) = &H9BDC06A7&: K(15) = &HC19BF174&
    K(16) = &HE49B69C1&: K(17) = &HEFBE4786&: K(18) = &H0FC19DC6&: K(19) = &H240CA1CC&
    K(20) = &H2DE92C6F&: K(21) = &H4A7484AA&: K(22) = &H5CB0A9DC&: K(23) = &H76F988DA&
    K(24) = &H983E5152&: K(25) = &HA831C66D&: K(26) = &HB00327C8&: K(27) = &HBF597FC7&
    K(28) = &HC6E00BF3&: K(29) = &HD5A79147&: K(30) = &H06CA6351&: K(31) = &H14292967&
    K(32) = &H27B70A85&: K(33) = &H2E1B2138&: K(34) = &H4D2C6DFC&: K(35) = &H53380D13&
    K(36) = &H650A7354&: K(37) = &H766A0ABB&: K(38) = &H81C2C92E&: K(39) = &H92722C85&
    K(40) = &HA2BFE8A1&: K(41) = &HA81A664B&: K(42) = &HC24B8B70&: K(43) = &HC76C51A3&
    K(44) = &HD192E819&: K(45) = &HD6990624&: K(46) = &HF40E3585&: K(47) = &H106AA070&
    K(48) = &H19A4C116&: K(49) = &H1E376C08&: K(50) = &H2748774C&: K(51) = &H34B0BCB5&
    K(52) = &H391C0CB3&: K(53) = &H4ED8AA4A&: K(54) = &H5B9CCA4F&: K(55) = &H682E6FF3&
    K(56) = &H748F82EE&: K(57) = &H78A5636F&: K(58) = &H84C87814&: K(59) = &H8CC70208&
    K(60) = &H90BEFFFA&: K(61) = &HA4506CEB&: K(62) = &HBEF9A3F7&: K(63) = &HC67178F2&
    mInitialized = True
End Sub
