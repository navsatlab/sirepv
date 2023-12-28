Attribute VB_Name = "crypto_sha1"
' ###################################################################
' # NAVSATLAB                                                       #
' # Application deployed over VBA                                   #
' # https://github.com/navsatlab/sirepv                             #
' # Some rights reserved                                            #
' ###################################################################

' This is all as per: http://stackoverflow.com/a/482150/5738 - License CC-BY 3.0
' ############################## Content from StackOverflow Answer START

' Based on: http://vb.wikia.com/wiki/SHA-1.bas
Option Explicit

Private Type FourBytes
    a As Byte
    B As Byte
    C As Byte
    D As Byte
End Type

Private Type OneLong
    L As Long
End Type

Function HexDefaultSHA1(message() As Byte) As String
    Dim H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long
    DefaultSHA1 message, H1, H2, H3, H4, H5
    HexDefaultSHA1 = DecToHex5(H1, H2, H3, H4, H5)
End Function

Function HexSHA1(message() As Byte, ByVal Key1 As Long, ByVal Key2 As Long, ByVal Key3 As Long, ByVal Key4 As Long) As String
    Dim H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long
    xSHA1 message, Key1, Key2, Key3, Key4, H1, H2, H3, H4, H5
    HexSHA1 = DecToHex5(H1, H2, H3, H4, H5)
End Function

Sub DefaultSHA1(message() As Byte, H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long)
    xSHA1 message, &H5A827999, &H6ED9EBA1, &H8F1BBCDC, &HCA62C1D6, H1, H2, H3, H4, H5
End Sub

Sub xSHA1(message() As Byte, ByVal Key1 As Long, ByVal Key2 As Long, ByVal Key3 As Long, ByVal Key4 As Long, H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long)
     'CA62C1D68F1BBCDC6ED9EBA15A827999 + "abc" = "A9993E36 4706816A BA3E2571 7850C26C 9CD0D89D"
     '"abc" = "A9993E36 4706816A BA3E2571 7850C26C 9CD0D89D"
    
     Dim U As Long, P As Long
     Dim FB As FourBytes, OL As OneLong
     Dim i As Integer
     Dim W(80) As Long
     Dim a As Long, B As Long, C As Long, D As Long, E As Long
     Dim T As Long
    
     H1 = &H67452301: H2 = &HEFCDAB89: H3 = &H98BADCFE: H4 = &H10325476: H5 = &HC3D2E1F0
    
     U = UBound(message) + 1: OL.L = U32ShiftLeft3(U): a = U \ &H20000000: LSet FB = OL 'U32ShiftRight29(U)
    
     ReDim Preserve message(0 To (U + 8 And -64) + 63)
     message(U) = 128
    
     U = UBound(message)
     message(U - 4) = a
     message(U - 3) = FB.D
     message(U - 2) = FB.C
     message(U - 1) = FB.B
     message(U) = FB.a
    
     While P < U
         For i = 0 To 15
             FB.D = message(P)
             FB.C = message(P + 1)
             FB.B = message(P + 2)
             FB.a = message(P + 3)
             LSet OL = FB
             W(i) = OL.L
             P = P + 4
         Next i
    
         For i = 16 To 79
             W(i) = U32RotateLeft1(W(i - 3) Xor W(i - 8) Xor W(i - 14) Xor W(i - 16))
         Next i
    
         a = H1: B = H2: C = H3: D = H4: E = H5
    
         For i = 0 To 19
             T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), E), W(i)), Key1), ((B And C) Or ((Not B) And D)))
             E = D: D = C: C = U32RotateLeft30(B): B = a: a = T
         Next i
         For i = 20 To 39
             T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), E), W(i)), Key2), (B Xor C Xor D))
             E = D: D = C: C = U32RotateLeft30(B): B = a: a = T
         Next i
         For i = 40 To 59
             T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), E), W(i)), Key3), ((B And C) Or (B And D) Or (C And D)))
             E = D: D = C: C = U32RotateLeft30(B): B = a: a = T
         Next i
         For i = 60 To 79
             T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), E), W(i)), Key4), (B Xor C Xor D))
             E = D: D = C: C = U32RotateLeft30(B): B = a: a = T
         Next i
    
         H1 = U32Add(H1, a): H2 = U32Add(H2, B): H3 = U32Add(H3, C): H4 = U32Add(H4, D): H5 = U32Add(H5, E)
     Wend
End Sub

Function U32Add(ByVal a As Long, ByVal B As Long) As Long
    If (a Xor B) < 0 Then
        U32Add = a + B
    Else
        U32Add = (a Xor &H80000000) + B Xor &H80000000
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

Function DecToHex5(ByVal H1 As Long, ByVal H2 As Long, ByVal H3 As Long, ByVal H4 As Long, ByVal H5 As Long) As String
    Dim H As String, L As Long
    DecToHex5 = "00000000 00000000 00000000 00000000 00000000"
    H = Hex(H1): L = Len(H): Mid(DecToHex5, 9 - L, L) = H
    H = Hex(H2): L = Len(H): Mid(DecToHex5, 18 - L, L) = H
    H = Hex(H3): L = Len(H): Mid(DecToHex5, 27 - L, L) = H
    H = Hex(H4): L = Len(H): Mid(DecToHex5, 36 - L, L) = H
    H = Hex(H5): L = Len(H): Mid(DecToHex5, 45 - L, L) = H
End Function

' Convert the string into bytes so we can use the above functions
' From Chris Hulbert: http://splinter.com.au/blog

Public Function SHA1HASH(str)
    Dim i As Integer
    Dim arr() As Byte
    If (Len(str) > 0) Then
      ReDim arr(0 To Len(str) - 1) As Byte
      For i = 0 To Len(str) - 1
       arr(i) = Asc(Mid(str, i + 1, 1))
      Next i
      SHA1HASH = Replace(LCase(HexDefaultSHA1(arr)), " ", "")
    Else
      SHA1HASH = ""
    End If
End Function
