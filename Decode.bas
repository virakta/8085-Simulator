Attribute VB_Name = "Decode"
Option Explicit


Sub Initialise()
A = "  "
B = "  "
c = "  "
D = "  "
E = "  "
F = "  "
H = "  "
l = "  "
M = "  "
SP = "    "
displayadd = "FFF9"
displaydata = "CC"
delaytime = 0
Sub2Add = "04BE"
Sub1Add = "06D3"

Dim Strdir As String
Dim I As Integer
Dim CD As Integer
Dim Temps As String
For I = 0 To 8191
block(I).Data = "ZZ"
block(I).Adress = LongtoHex(tolong("C000") + I)
Next I
If (regs <> "") Then
A = Mid(regs, 1, 2)
B = Mid(regs, 3, 2)
c = Mid(regs, 5, 2)
D = Mid(regs, 7, 2)
E = Mid(regs, 9, 2)
F = Mid(regs, 11, 2)
H = Mid(regs, 13, 2)
l = Mid(regs, 15, 2)
M = Mid(regs, 17, 2)
End If
PCounter = "C000"
If (IniSP <> "") Then SP = IniSP
 Sign = 0
 Parity = 0
 AC = 0
 Zero = 0
 cy = 0
 fileno = FreeFile
 Temps = String(2, " ")
 If (startfile <> "") Then
 On Error GoTo ED1
Open startfile For Binary As #fileno
For CD = 0 To 8191
Get #fileno, , Temps
If (CD = 0 And Temps = "") Then Exit For
block(CD).Data = Temps
Next CD
Form2.Caption = frmcap + file1 + "  )"
Else
Form2.Caption = frmcap + "tempcode  )"
End If
ED1:
End Sub
Sub Randamise()
Dim I As Integer
Randomize
For I = 0 To 8191
If block(I).Data = "ZZ" Then
block(I).Data = InttoHex(Int(255 * Rnd))
End If
Next I
If (A = "  ") Then A = InttoHex(Int(255 * Rnd))
If (B = "  ") Then B = InttoHex(Int(255 * Rnd))
If (c = "  ") Then c = InttoHex(Int(255 * Rnd))
If (D = "  ") Then D = InttoHex(Int(255 * Rnd))
If (E = "  ") Then E = InttoHex(Int(255 * Rnd))
If (F = "  ") Then F = InttoHex(Int(255 * Rnd))
If (H = "  ") Then H = InttoHex(Int(255 * Rnd))
If (l = "  ") Then l = InttoHex(Int(255 * Rnd))
If (M = "  ") Then M = InttoHex(Int(255 * Rnd))
If (PCounter = "") Then PCounter = "C000"
If (SP = "    ") Then SP = "DFFF"
 Sign = 0
 Parity = 0
 AC = 0
 Zero = 0
 cy = 0

End Sub

Function GetData(STR As String) As String
Dim I As Long
 AddHexPc STR, 0
If (ErrPgm = 1) Then
Exit Function
End If
I = tolong(STR) - tolong("C000")
GetData = block(I).Data
End Function
Sub SetData(stradd As String, strdata As String)
Dim I As Long
 AddHexPc stradd, 0
If (ErrPgm = 1) Then
Exit Sub
End If
I = tolong(stradd) - tolong("C000")
 block(I).Data = strdata
End Sub
Function GetKeyValue(I As Integer) As String
If (I = 1) Then GetKeyValue = "1"
If (I = 2) Then GetKeyValue = "2"
If (I = 3) Then GetKeyValue = "3"
If (I = 4) Then GetKeyValue = "4"
If (I = 5) Then GetKeyValue = "5"
If (I = 6) Then GetKeyValue = "6"
If (I = 7) Then GetKeyValue = "7"
If (I = 8) Then GetKeyValue = "8"
If (I = 9) Then GetKeyValue = "9"
If (I = 0) Then GetKeyValue = "0"
If (I = 10) Then GetKeyValue = "A"
If (I = 11) Then GetKeyValue = "B"
If (I = 12) Then GetKeyValue = "C"
If (I = 13) Then GetKeyValue = "D"
If (I = 14) Then GetKeyValue = "E"
If (I = 15) Then GetKeyValue = "F"
End Function
Function GetRegData(I As Integer) As String
Select Case I
Case 4
     GetRegData = HByte("SP")
Case 5
     GetRegData = LByte("SP")
Case 6
     GetRegData = HByte("PCounter")
Case 7
     GetRegData = LByte("PCounter")
Case 8
     GetRegData = H
Case 9
     GetRegData = l
Case 10
     GetRegData = A
Case 11
     GetRegData = B
Case 12
     GetRegData = c
Case 13
     GetRegData = D
Case 14
     GetRegData = E
Case 15
     GetRegData = F
End Select
End Function
Function GetRegCaption(I As Integer) As String
Select Case I
Case 4
     GetRegCaption = " SPH"
Case 5
     GetRegCaption = " SPL"
Case 6
     GetRegCaption = " PCH"
Case 7
     GetRegCaption = " PCL"
Case 8
     GetRegCaption = "   H"
Case 9
     GetRegCaption = "   L"
Case 10
     GetRegCaption = "   A"
Case 11
     GetRegCaption = "   B"
Case 12
     GetRegCaption = "   C"
Case 13
     GetRegCaption = "   D"
Case 14
     GetRegCaption = "   E"
Case 15
     GetRegCaption = "   F"
    
End Select
End Function
Function HByte(STR As String) As String
If (STR = "SP") Then HByte = Left(SP, 2)
If (STR = "PCounter") Then HByte = Left(PCounter, 2)
If (STR = "PSW") Then HByte = Left(PSW, 2)
End Function
Function LByte(STR As String) As String
If (STR = "SP") Then LByte = Right(SP, 2)
If (STR = "PCounter") Then LByte = Right(PCounter, 2)
If (STR = "PSW") Then LByte = Right(PSW, 2)
End Function
Sub SetRegData(I As Integer, STR As String)
Select Case I
Case 4
      SP = STR + Right(SP, 2)
Case 5
     SP = Left(SP, 2) + STR
Case 6
      PCounter = STR + Right(PCounter, 2)
Case 7
      PCounter = Left(PCounter, 2) + STR
Case 8
      H = STR
Case 9
      l = STR
Case 10
     A = STR
Case 11
     B = STR
Case 12
     c = STR
Case 13
     D = STR
Case 14
     E = STR
Case 15
     F = STR
End Select

End Sub

Function GetFlagReg() As String
Dim S1 As String
If (Sign = 1) Then
     S1 = "1"
     Else
     S1 = "0"
End If
If (Zero = 1) Then
     S1 = S1 + "1000"
     Else
     S1 = S1 + "0000"
End If

If (Parity = 1) Then
     S1 = S1 + "10"
     Else
     S1 = S1 + "00"
End If
If (cy = 1) Then
     S1 = S1 + "1"
     Else
     S1 = S1 + "0"
End If
     GetFlagReg = BintoHex(S1)
End Function

Function RALCY(S1 As String) As String
Dim tempi As Integer
Dim Temps As String
Dim S2 As String
S2 = HextoBin(S2)
tempi = cy
Temps = Left(S2, 1)
If (Temps = "1") Then
cy = 1
Else
cy = 0
End If
Temps = Right(S2, 7)
If (tempi = 1) Then
Temps = Temps + "1"
Else
Temps = Temps + "0"
End If
RALCY = BintoHex(Temps)
End Function
Function RARCY(S1 As String) As String
Dim tempi As Integer
Dim Temps As String
Dim S2 As String
S2 = HextoBin(S1)
tempi = cy
Temps = Right(S2, 1)
If (Temps = "1") Then
cy = 1
Else
cy = 0
End If
Temps = Left(S2, 7)
If (tempi = 1) Then
Temps = "1" + Temps
Else
Temps = "0" + Temps
End If
RARCY = BintoHex(Temps)
End Function

Function RLC(S1 As String) As String
Dim Temps As String
Dim S2 As String
S2 = HextoBin(S1)
Temps = Left(S2, 1)
If (Temps = "1") Then
cy = 1
Else
cy = 0
End If
Temps = Right(S2, 7)
If (cy = 1) Then
Temps = Temps + "1"
Else
Temps = Temps + "0"
End If
RLC = BintoHex(Temps)
End Function
Function RRC(S1 As String) As String
Dim Temps As String
Dim S2 As String
S2 = HextoBin(S1)
Temps = Right(S2, 1)
If (Temps = "1") Then
cy = 1
Else
cy = 0
End If
Temps = Left(S2, 7)
If (cy = 1) Then
Temps = "1" + Temps
Else
Temps = "0" + Temps
End If
RRC = BintoHex(Temps)
End Function

Sub SetFlags(STR As String)
Dim S1 As String
Dim I As Integer
Dim Count As Integer
Dim sr As String
S1 = HextoBin(STR)
sr = Left(S1, 1)
If (sr = "1") Then
Sign = 1
Else
Sign = 0
End If
I = toInt(STR)
If (I = 0) Then
Zero = 1
Else
Zero = 0
End If
Count = 0
For I = 1 To 8
sr = Mid(S1, I, 1)
If (sr = "1") Then Count = Count + 1
Next I
If ((Count Mod 2) = 0) Then
Parity = 1
Else
Parity = 0
End If
End Sub
Sub SetPSW(STR As String)
Dim S1 As String
Dim S2 As String
Sign = 0
Zero = 0
Parity = 0
cy = 0
A = Left(PSW, 8)
S1 = Right(PSW, 8)
S2 = HextoBin(S1)
S1 = Left(S2, 1)
If (S1 = "1") Then Sign = 1
S1 = Mid(S2, 2, 1)
If (S1 = "1") Then Zero = 1
S1 = Mid(S2, 6, 1)
If (S1 = "1") Then Parity = 1
S1 = Mid(S2, 8, 1)
If (S1 = "1") Then cy = 1

End Sub
Sub PushIT(STR As String)
Dim S1 As String
Dim sh As String
Dim sl As String
S1 = STR
sh = Left(S1, 2)
sl = Right(S1, 2)
SP = SubHexLong(SP, , 1)
SetData SP, sh
SP = SubHexLong(SP, , 1)
SetData SP, sl
StackCase = 5
End Sub
Sub ADD16BIT(str1 As String, str2 As String)
     Dim I As Long
     Dim S1 As String
     Dim j As Long
     Dim k As Long
     I = tolong(str1)
     j = tolong(str2)
     k = I + j
     If (k > 65535) Then
     cy = 1
     k = k - 65536
     End If
     S1 = LongtoHex(k)
     l = Right(S1, 2)
     H = Left(S1, 2)
End Sub

