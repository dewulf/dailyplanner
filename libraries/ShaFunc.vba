'Taken from here
'stackoverflow.com/questions/14717526/vba-hash-string

Private Type FourBytes
    a As Byte
    b As Byte
    c As Byte
    d As Byte
End Type
Private Type OneLong
    l As Long
End Type

Function HexDefaultSHA1(message() As Byte) As String
 Dim h1 As Long, h2 As Long, H3 As Long, H4 As Long, H5 As Long
 DefaultSHA1 message, h1, h2, H3, H4, H5
 HexDefaultSHA1 = DecToHex5(h1, h2, H3, H4, H5)
End Function

Function HexSHA1(message() As Byte, ByVal Key1 As Long, ByVal Key2 As Long, ByVal Key3 As Long, ByVal Key4 As Long) As String
 Dim h1 As Long, h2 As Long, H3 As Long, H4 As Long, H5 As Long
 xSHA1 message, Key1, Key2, Key3, Key4, h1, h2, H3, H4, H5
 HexSHA1 = DecToHex5(h1, h2, H3, H4, H5)
End Function

Sub DefaultSHA1(message() As Byte, h1 As Long, h2 As Long, H3 As Long, H4 As Long, H5 As Long)
 xSHA1 message, &H5A827999, &H6ED9EBA1, &H8F1BBCDC, &HCA62C1D6, h1, h2, H3, H4, H5
End Sub

Sub xSHA1(message() As Byte, ByVal Key1 As Long, ByVal Key2 As Long, ByVal Key3 As Long, ByVal Key4 As Long, h1 As Long, h2 As Long, H3 As Long, H4 As Long, H5 As Long)

 Dim u As Long, p As Long
 Dim FB As FourBytes, OL As OneLong
 Dim i As Integer
 Dim w(80) As Long
 Dim a As Long, b As Long, c As Long, d As Long, e As Long
 Dim T As Long

 h1 = &H67452301: h2 = &HEFCDAB89: H3 = &H98BADCFE: H4 = &H10325476: H5 = &HC3D2E1F0

 u = UBound(message) + 1: OL.l = U32ShiftLeft3(u): a = u \ &H20000000: LSet FB = OL 'U32ShiftRight29(U)

 ReDim Preserve message(0 To (u + 8 And -64) + 63)
 message(u) = 128

 u = UBound(message)
 message(u - 4) = a
 message(u - 3) = FB.d
 message(u - 2) = FB.c
 message(u - 1) = FB.b
 message(u) = FB.a

 While p < u
     For i = 0 To 15
         FB.d = message(p)
         FB.c = message(p + 1)
         FB.b = message(p + 2)
         FB.a = message(p + 3)
         LSet OL = FB
         w(i) = OL.l
         p = p + 4
     Next i

     For i = 16 To 79
         w(i) = U32RotateLeft1(w(i - 3) Xor w(i - 8) Xor w(i - 14) Xor w(i - 16))
     Next i

     a = h1: b = h2: c = H3: d = H4: e = H5

     For i = 0 To 19
         T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), w(i)), Key1), ((b And c) Or ((Not b) And d)))
         e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
     Next i
     For i = 20 To 39
         T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), w(i)), Key2), (b Xor c Xor d))
         e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
     Next i
     For i = 40 To 59
         T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), w(i)), Key3), ((b And c) Or (b And d) Or (c And d)))
         e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
     Next i
     For i = 60 To 79
         T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), w(i)), Key4), (b Xor c Xor d))
         e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
     Next i

     h1 = U32Add(h1, a): h2 = U32Add(h2, b): H3 = U32Add(H3, c): H4 = U32Add(H4, d): H5 = U32Add(H5, e)
 Wend
End Sub

Function U32Add(ByVal a As Long, ByVal b As Long) As Long
 If (a Xor b) < 0 Then
     U32Add = a + b
 Else
     U32Add = (a Xor &H80000000) + b Xor &H80000000
 End If
End Function

Function U32ShiftLeft3(ByVal a As Long) As Long
 U32ShiftLeft3 = (a And &HFFFFFFF) * 8
 If a And &H10000000 Then U32ShiftLeft3 = U32ShiftLeft3 Or &H80000000
End Function

Function U32ShiftRight29(ByVal a As Long) As Long
 U32ShiftRight29 = (a And &HE0000000) \ &H20000000 And 7
End Function

Function U32RotateLeft1(ByVal a As Long) As Long
 U32RotateLeft1 = (a And &H3FFFFFFF) * 2
 If a And &H40000000 Then U32RotateLeft1 = U32RotateLeft1 Or &H80000000
 If a And &H80000000 Then U32RotateLeft1 = U32RotateLeft1 Or 1
End Function
Function U32RotateLeft5(ByVal a As Long) As Long
 U32RotateLeft5 = (a And &H3FFFFFF) * 32 Or (a And &HF8000000) \ &H8000000 And 31
 If a And &H4000000 Then U32RotateLeft5 = U32RotateLeft5 Or &H80000000
End Function
Function U32RotateLeft30(ByVal a As Long) As Long
 U32RotateLeft30 = (a And 1) * &H40000000 Or (a And &HFFFC) \ 4 And &H3FFFFFFF
 If a And 2 Then U32RotateLeft30 = U32RotateLeft30 Or &H80000000
End Function

Function DecToHex5(ByVal h1 As Long, ByVal h2 As Long, ByVal H3 As Long, ByVal H4 As Long, ByVal H5 As Long) As String
 Dim h As String, l As Long
 DecToHex5 = "00000000 00000000 00000000 00000000 00000000"
 h = Hex(h1): l = Len(h): Mid(DecToHex5, 9 - l, l) = h
 h = Hex(h2): l = Len(h): Mid(DecToHex5, 18 - l, l) = h
 h = Hex(H3): l = Len(h): Mid(DecToHex5, 27 - l, l) = h
 h = Hex(H4): l = Len(h): Mid(DecToHex5, 36 - l, l) = h
 h = Hex(H5): l = Len(h): Mid(DecToHex5, 45 - l, l) = h
End Function

Public Function SHA1TRUNC(str)
  Dim i As Integer
  Dim arr() As Byte
  ReDim arr(0 To Len(str) - 1) As Byte
  Const cutoff As Integer = 6
  
  For i = 0 To Len(str) - 1
   arr(i) = Asc(Mid(str, i + 1, 1))
  Next i
  
  SHA1TRUNC = Replace(LCase(HexDefaultSHA1(arr)), " ", "")
  SHA1TRUNC = Left(SHA1TRUNC, cutoff)
  
End Function