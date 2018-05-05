Attribute VB_Name = "Math"
Public Function tolong(STR As String) As Long
Dim I As Integer
Dim S1 As String
Dim S2(1 To 4) As String
tolong = 0
S1 = STR
S2(1) = Right(S1, 1)
S2(4) = Left(S1, 1)
S2(2) = Right(S1, 2)
S2(2) = Left(S2(2), 1)
S2(3) = Left(S1, 2)
S2(3) = Right(S2(3), 1)
For I = 1 To 4
Select Case S2(I)
Case "0"
     tolong = tolong + 0 * 16 ^ (I - 1)
Case "1"
     tolong = tolong + 1 * 16 ^ (I - 1)
Case "2"
     tolong = tolong + 2 * 16 ^ (I - 1)
Case "3"
     tolong = tolong + 3 * 16 ^ (I - 1)
Case "4"
     tolong = tolong + 4 * 16 ^ (I - 1)
Case "5"
     tolong = tolong + 5 * 16 ^ (I - 1)
Case "6"
     tolong = tolong + 6 * 16 ^ (I - 1)
Case "7"
     tolong = tolong + 7 * 16 ^ (I - 1)
Case "8"
     tolong = tolong + 8 * 16 ^ (I - 1)
Case "9"
     tolong = tolong + 9 * 16 ^ (I - 1)
Case "A"
     tolong = tolong + 10 * 16 ^ (I - 1)
Case "B"
     tolong = tolong + 11 * 16 ^ (I - 1)
Case "C"
     tolong = tolong + 12 * 16 ^ (I - 1)
Case "D"
     tolong = tolong + 13 * 16 ^ (I - 1)
Case "E"
     tolong = tolong + 14 * 16 ^ (I - 1)
Case "F"
     tolong = tolong + 15 * 16 ^ (I - 1)
     End Select
  Next I
    End Function
Public Function toInt(STR As String) As Integer
Dim I As Integer
Dim S1 As String
Dim S2(1 To 2) As String
toInt = 0
S1 = STR
S2(1) = Right(S1, 1)
S2(2) = Left(S1, 1)
For I = 1 To 2
Select Case S2(I)
Case "0"
     toInt = toInt + 0 * 16 ^ (I - 1)
Case "1"
     toInt = toInt + 1 * 16 ^ (I - 1)
Case "2"
     toInt = toInt + 2 * 16 ^ (I - 1)
Case "3"
     toInt = toInt + 3 * 16 ^ (I - 1)
Case "4"
     toInt = toInt + 4 * 16 ^ (I - 1)
Case "5"
     toInt = toInt + 5 * 16 ^ (I - 1)
Case "6"
     toInt = toInt + 6 * 16 ^ (I - 1)
Case "7"
     toInt = toInt + 7 * 16 ^ (I - 1)
Case "8"
     toInt = toInt + 8 * 16 ^ (I - 1)
Case "9"
     toInt = toInt + 9 * 16 ^ (I - 1)
Case "A"
     toInt = toInt + 10 * 16 ^ (I - 1)
Case "B"
     toInt = toInt + 11 * 16 ^ (I - 1)
Case "C"
     toInt = toInt + 12 * 16 ^ (I - 1)
Case "D"
     toInt = toInt + 13 * 16 ^ (I - 1)
Case "E"
     toInt = toInt + 14 * 16 ^ (I - 1)
Case "F"
     toInt = toInt + 15 * 16 ^ (I - 1)
     End Select
  Next I
    End Function

Function AddHexLong(str1 As String, Optional str2 As String, Optional c = 0) As String
     Dim I As Long
     Dim j As Long
     Dim k As Long
     I = tolong(str1)
     j = tolong(str2)
     k = I + j + c
     AddHexLong = LongtoHex(k)
End Function
Function AddHexInt(str1 As String, Optional str2 As String, Optional c = 0) As String
     Dim I As Integer
     Dim j As Integer
     Dim k As Integer
     Dim l As Integer
     Dim M As Integer
     AC = 0
     l = valueof(Right(HextoBin(str1), 4))
     M = valueof(Right(HextoBin(str2), 4))
     If ((M + l) > 15) Then AC = 1
     I = toInt(str1)
     j = toInt(str2)
     k = I + j + c
     cy = 0
     If k > 255 Then
     cy = 1
     k = k - 256
     End If
     AddHexInt = InttoHex(k)
End Function

Function SubHexLong(str1 As String, Optional str2 As String, Optional c = 0) As String
     Dim I As Long
     Dim j As Long
     Dim k As Long
     I = tolong(str1)
     j = tolong(str2)
     k = I - j - c
     k = Abs(k)
     SubHexLong = Hex(k)
     If (Len(SubHexLong) = 1) Then SubHexLong = "000" + SubHexLong
     If (Len(SubHexLong) = 2) Then SubHexLong = "00" + SubHexLong
     If (Len(SubHexLong) = 3) Then SubHexLong = "0" + SubHexLong
End Function
Function SubHexInt(str1 As String, Optional str2 As String, Optional c = 0) As String
     Dim I As Integer
     Dim j As Integer
     Dim k As Integer
     Dim ts1 As String
     Dim ts2 As String
     ts2 = str1
     I = toInt(str1)
     j = toInt(str2)
     If (I >= (j + c)) Then
     k = I - j - c
     k = Abs(k)
     SubHexInt = InttoHex(k)
     cy = 0
     Else
     cy = 1
     j = j + c
     ts1 = CmpHex(InttoHex(j))
     ts1 = AddInr(ts1, , 1)
     SubHexInt = AddInr(ts2, ts1, 0)
     End If
End Function

Function HextoBin(STR As String) As String
Dim I As Integer
Dim S1 As String
Dim S2(1 To 2) As String
S1 = STR
S2(1) = Left(S1, 1)
S2(2) = Right(S1, 1)
For I = 1 To 2
Select Case S2(I)
Case "0"
     HextoBin = HextoBin + "0000"
Case "1"
     HextoBin = HextoBin + "0001"
Case "2"
     HextoBin = HextoBin + "0010"
Case "3"
     HextoBin = HextoBin + "0011"
Case "4"
     HextoBin = HextoBin + "0100"
Case "5"
     HextoBin = HextoBin + "0101"
Case "6"
     HextoBin = HextoBin + "0110"
Case "7"
     HextoBin = HextoBin + "0111"
Case "8"
     HextoBin = HextoBin + "1000"
Case "9"
     HextoBin = HextoBin + "1001"
Case "A"
     HextoBin = HextoBin + "1010"
Case "B"
     HextoBin = HextoBin + "1011"
Case "C"
     HextoBin = HextoBin + "1100"
Case "D"
     HextoBin = HextoBin + "1101"
Case "E"
     HextoBin = HextoBin + "1110"
Case "F"
     HextoBin = HextoBin + "1111"
     End Select
  Next I

End Function
Function CmpHex(STR As String) As String
Dim I As Integer
Dim S1 As String
Dim s As String
S1 = HextoBin(STR)

     For I = 1 To 8
     s = Mid(S1, I, 1)
     If (s = "1") Then
     CmpHex = CmpHex + "0"
     Else
     CmpHex = CmpHex + "1"
     End If
     Next I
     CmpHex = BintoHex(CmpHex)
End Function
Function BintoHex(STR As String) As String
Dim I As Integer
Dim S1 As String
Dim S2(1 To 2) As String
S1 = STR
S2(1) = Left(S1, 4)
S2(2) = Right(S1, 4)
For I = 1 To 2
Select Case S2(I)
Case "0000"
     BintoHex = BintoHex + "0"
Case "0001"
     BintoHex = BintoHex + "1"
Case "0010"
     BintoHex = BintoHex + "2"
Case "0011"
     BintoHex = BintoHex + "3"
Case "0100"
     BintoHex = BintoHex + "4"
Case "0101"
     BintoHex = BintoHex + "5"
Case "0110"
     BintoHex = BintoHex + "6"
Case "0111"
     BintoHex = BintoHex + "7"
Case "1000"
     BintoHex = BintoHex + "8"
Case "1001"
     BintoHex = BintoHex + "9"
Case "1010"
     BintoHex = BintoHex + "A"
Case "1011"
     BintoHex = BintoHex + "B"
Case "1100"
     BintoHex = BintoHex + "C"
Case "1101"
     BintoHex = BintoHex + "D"
Case "1110"
     BintoHex = BintoHex + "E"
Case "1111"
     BintoHex = BintoHex + "F"
     End Select
  Next I

End Function
Function ANDHex(str1 As String, str2 As String) As String
Dim i1 As Integer
Dim i2 As Integer
Dim i3 As Integer
i1 = toInt(str1)
i2 = toInt(str2)
i3 = i1 And i2
ANDHex = Hex(i3)
If (Len(ANDHex) = 1) Then ANDHex = "0" + ANDHex
End Function
Function ORHex(str1 As String, str2 As String) As String
Dim i1 As Integer
Dim i2 As Integer
Dim i3 As Integer
i1 = toInt(str1)
i2 = toInt(str2)
i3 = i1 Or i2
ORHex = Hex(i3)
If (Len(ORHex) = 1) Then ORHex = "0" + ORHex
End Function
Function XORHex(str1 As String, str2 As String) As String
Dim i1 As Integer
Dim i2 As Integer
Dim i3 As Integer
i1 = toInt(str1)
i2 = toInt(str2)
i3 = i1 Xor i2
XORHex = Hex(i3)
If (Len(XORHex) = 1) Then XORHex = "0" + XORHex
End Function
Function LongtoHex(I As Long) As String
LongtoHex = Hex(I)
If (Len(LongtoHex) = 1) Then LongtoHex = "000" + LongtoHex
If (Len(LongtoHex) = 2) Then LongtoHex = "00" + LongtoHex
If (Len(LongtoHex) = 3) Then LongtoHex = "0" + LongtoHex
End Function

Function InttoHex(I As Integer) As String
InttoHex = Hex(I)
If (Len(InttoHex) > 2) Then InttoHex = Right(InttoHex, 2)
If (Len(InttoHex) = 1) Then InttoHex = "0" + InttoHex
End Function

Function AddInr(str1 As String, Optional str2 As String, Optional c = 0) As String
     Dim I As Integer
     Dim j As Integer
     Dim k As Integer
     I = toInt(str1)
     j = toInt(str2)
     k = I + j + c
     If k > 255 Then
     k = k - 256
     End If
     AddInr = InttoHex(k)
     End Function
Function SubInr(str1 As String, Optional str2 As String, Optional c = 0) As String
     Dim I As Integer
     Dim j As Integer
     Dim k As Integer
     Dim ts1 As String
     Dim ts2 As String
     I = toInt(str1)
     j = toInt(str2)
     ts2 = str1
     If (I >= (j + c)) Then
     k = I - j - c
     SubInr = InttoHex(k)
     Else
     j = j + c
     ts1 = CmpHex(InttoHex(j))
     ts1 = AddInr(ts1, , 1)
     SubInr = AddInr(ts2, ts1, 0)
     End If
     End Function
Function valueof(STR As String) As Integer
Dim I As Integer
Dim S1 As String
Dim S2 As String
Dim k As Integer
S1 = STR
valueof = 0
k = 0
For I = 4 To 1 Step -1
S2 = Mid(S1, I, 1)
Select Case S2
Case "0"
     valueof = valueof + 0 * 2 ^ k
Case "1"
     valueof = valueof + 1 * 2 ^ k
     End Select
     k = k + 1
  Next I
  
End Function



Function SubHex(str1 As String, Optional str2 As String, Optional c = 0)
Dim I As Integer
Dim ts1 As String
Dim ts2 As String
Dim j As Integer
Dim k As Integer
ts1 = str1
ts2 = str2
I = toInt(ts1)
j = toInt(ts2)
If (I >= (j + c)) Then
cy = 0
k = I - (j + c)
SubHex = InttoHex(k)
Else
     cy = 1
     j = j + c
     ts2 = CmpHex(InttoHex(j))
     ts2 = AddInr(ts2, , 1)
     SubHex = AddInr(ts1, ts2, 0)
End If
End Function
Function DaaIt(STR As String) As String
Dim S1 As String
Dim ts1 As String
Dim ts2 As String
Dim stradd As String
Dim Carry As Integer
Carry = cy
stradd = STR
If (AC = 1) Then stradd = AddHexInt(stradd, "06")
ts1 = Right(HextoBin(stradd), 4)
If (valueof(ts1) > 9) Then stradd = AddHexInt(stradd, "06")
ts2 = Left(HextoBin(STR), 4)
If (valueof(ts2) > 9) Then stradd = AddHexInt(stradd, "60")
If (Carry = 1) Then stradd = AddHexInt(stradd, "60")
AC = 0
If (Carry = 1) Then cy = 1
DaaIt = stradd
End Function

