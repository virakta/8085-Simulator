Attribute VB_Name = "Exec"
Option Explicit

Function ExecPgm() As Integer
Dim byte1 As String
Dim byte2 As String
Dim ts1 As String
Dim ts2 As String
Dim Ti1 As Integer
Dim Ti2 As Integer

Dim j As Integer
Dim strnemo As String
Dim strbyte1 As String
Dim strbyte2 As String
Dim strdisply As String
Dim byted As Integer
Dim startdata As String
Dim stradd As String
Dim tempsp As String
StackCase = 0
ErrPgm = 0
ExecPgm = 0
BugAdd(1) = ""
BugData(1) = ""
BugAdd(2) = ""
BugData(2) = ""

BugCount = 1
Do
   DoEvents
   If (break = 0) Then
     If (step1 = 1) Then
          If (step2 = 1) Then
          stradd = PCounter
          startdata = GetData(stradd)
            For j = 1 To 246
               If (startdata = InSet(j).OpCode) Then
               strnemo = InSet(j).Nemo
               byted = InSet(j).ByteCount
               Exit For
               End If
            Next j
            If (byted = 1) Then
            strbyte1 = ""
            End If
          If (byted = 2) Then
          strbyte1 = GetData(AddHexPc(PCounter, 1))
          PCounter = AddHexPc(PCounter, 0)
          End If
          If (byted = 3) Then
          strbyte1 = GetData(AddHexPc(PCounter, 1))
          PCounter = AddHexPc(PCounter, 0)
          strbyte2 = GetData(AddHexPc(AddHexPc(PCounter, 1), 1))
          PCounter = AddHexPc(PCounter, 0)
          strbyte1 = strbyte2 + strbyte1
          End If
          strdisply = strnemo + strbyte1
         End If
       
Select Case GetData(PCounter)
Case "CE" 'ACI
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     A = AddHexInt(A, byte1, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "8F" 'ADC A
     A = AddHexInt(A, A, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "88" 'ADC B
     A = AddHexInt(A, B, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "89" 'ADC C
     A = AddHexInt(A, c, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
 Case "8A" 'ADC D
     A = AddHexInt(A, D, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
 Case "8B" 'ADC E
     A = AddHexInt(A, E, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "8C" 'ADC H
     A = AddHexInt(A, H, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "8D" 'ADC L
     A = AddHexInt(A, l, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
     
Case "8E" 'ADC M
     ts1 = H + l
     M = GetData(ts1)
     A = AddHexInt(A, M, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
 Case "87" 'ADD A
     A = AddHexInt(A, A, 0)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "80" 'ADD B
     A = AddHexInt(A, B, 0)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "81" 'ADD C
     A = AddHexInt(A, c, 0)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
 Case "82" 'ADD D
     A = AddHexInt(A, D, 0)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
 Case "83" 'ADD E
     A = AddHexInt(A, E, 0)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "84" 'ADD H
     A = AddHexInt(A, H, 0)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "85" 'ADD L
     A = AddHexInt(A, l, 0)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "86" 'ADD M
     ts1 = H + l
     M = GetData(ts1)
     A = AddHexInt(A, M, 0)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
     
Case "C6" 'ADI
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     A = AddHexInt(A, byte1, 0)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
     
Case "A7" 'ANA
     A = ANDHex(A, A)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "A0" 'ANA B
     A = ANDHex(A, B)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "A1" 'ANA C
     A = ANDHex(A, c)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "A2" 'ANA D
     A = ANDHex(A, D)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "A3" 'ANA E
     A = ANDHex(A, E)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "A4" 'ANA  H
     A = ANDHex(A, H)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
Case "A5" 'ANA L
     A = ANDHex(A, l)
     cy = 0
     SetFlags (A)
     
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
     
Case "A6" 'ANA M
     ts1 = H + l
     M = GetData(ts1)
     A = ANDHex(A, M)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
     
 Case "E6" 'AN1 8-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     A = ANDHex(A, byte1)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
   
   Case "CD" 'CALL
     PCounter = AddHexPc(PCounter, 1)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
       
     If ((byte2 + byte1) = Sub1Add) Then
    
       If (step2 <> 1) Then
               Form2.DISPLAY = Space(2) + displaydata + strDot + Space(2)
               Displayed = True
       End If
       
           Else
    If ((byte2 + byte1) = Sub2Add) Then
    
               Form2.T1.Interval = Int(0.00785 * tolong(D + E))
               If (Form2.T1.Interval <= 10) Then Form2.T1.Interval = 10
               YES = "YES"
               Form2.T1.Enabled = True
                  While (YES = "YES")
                  DoEvents
                  Wend
                  
                   Else
     PCounter = AddHexPc(PCounter, 0)
     PushIT PCounter
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
    End If
   End If
   
 Case "DC" 'CC
     PCounter = AddHexPc(PCounter, 1)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
    
     
     If (cy = 1) Then
      
     If ((byte2 + byte1) = Sub1Add) Then
    
       If (step2 <> 1) Then
               Form2.DISPLAY = Space(2) + displaydata + strDot + Space(2)
               Displayed = True
       End If
       PCounter = AddHexPc(PCounter, 0)
    Else
    If ((byte2 + byte1) = Sub2Add) Then
               Form2.T1.Interval = Int(0.00785 * tolong(D + E))
               If (Form2.T1.Interval <= 10) Then Form2.T1.Interval = 10
               YES = "YES"
               Form2.T1.Enabled = True
                  While (YES = "YES")
                  DoEvents
                  Wend
               Else
     PCounter = AddHexPc(PCounter, 0)
     PushIT PCounter
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
    End If
   End If
End If
     
 Case "FC" 'CM
     PCounter = AddHexPc(PCounter, 1)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     
     If (Sign = 1) Then
           
     If ((byte2 + byte1) = Sub1Add) Then
       If (step2 <> 1) Then
               Form2.DISPLAY = Space(2) + displaydata + strDot + Space(2)
               Displayed = True
       End If
       
    If ((byte2 + byte1) = Sub2Add) Then
               Form2.T1.Interval = Int(0.00785 * tolong(D + E))
               If (Form2.T1.Interval <= 10) Then Form2.T1.Interval = 10
               YES = "YES"
               Form2.T1.Enabled = True
                  While (YES = "YES")
                  DoEvents
                  Wend
              
     Else
      PCounter = AddHexPc(PCounter, 0)
   PushIT PCounter
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
    End If
   End If
End If
     
 Case "2F" 'CMA
     A = CmpHex(A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = AddHexPc(PCounter, 0)
     
 Case "3F" 'CMC
     If (cy = 1) Then
     cy = 0
     Else
     cy = 1
     End If
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
 Case "BF" ' CMP A
     cy = 0
     SetFlags (A)
     Zero = 1
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
 Case "B8" 'CMP B
     Ti1 = toInt(A)
     Ti2 = toInt(B)
     If (Ti1 > Ti2) Then
        cy = 0
        Zero = 0
        End If
     If (Ti1 < Ti2) Then
       cy = 1
       Zero = 0
       End If
     If (Ti1 = Ti2) Then
       cy = 0
       Zero = 1
       End If
       SetFlags (A)
       PCounter = AddHexPc(PCounter, 1)
       PCounter = AddHexPc(PCounter, 0)
     
 Case "B9" 'CMP C
     Ti1 = toInt(A)
     Ti2 = toInt(c)
     If (Ti1 > Ti2) Then
        cy = 0
        Zero = 0
        End If
     If (Ti1 < Ti2) Then
       cy = 1
       Zero = 0
       End If
     If (Ti1 = Ti2) Then
       cy = 0
       Zero = 1
       End If
       SetFlags (A)
       PCounter = AddHexPc(PCounter, 1)
       PCounter = AddHexPc(PCounter, 0)
       
 Case "BA" 'CMP D
     Ti1 = toInt(A)
     Ti2 = toInt(D)
     If (Ti1 > Ti2) Then
        cy = 0
        Zero = 0
        End If
     If (Ti1 < Ti2) Then
       cy = 1
       Zero = 0
       End If
     If (Ti1 = Ti2) Then
       cy = 0
       Zero = 1
       End If
       SetFlags (A)
       PCounter = AddHexPc(PCounter, 1)
       PCounter = AddHexPc(PCounter, 0)
       
 Case "BB" 'CMP E
     Ti1 = toInt(A)
     Ti2 = toInt(E)
     If (Ti1 > Ti2) Then
        cy = 0
        Zero = 0
        End If
     If (Ti1 < Ti2) Then
       cy = 1
       Zero = 0
       End If
     If (Ti1 = Ti2) Then
       cy = 0
       Zero = 1
       End If
       SetFlags (A)
       PCounter = AddHexPc(PCounter, 1)
       PCounter = AddHexPc(PCounter, 0)
       
Case "BC" 'CMP H
     Ti1 = toInt(A)
     Ti2 = toInt(H)
     If (Ti1 > Ti2) Then
        cy = 0
        Zero = 0
        End If
     If (Ti1 < Ti2) Then
       cy = 1
       Zero = 0
       End If
     If (Ti1 = Ti2) Then
       cy = 0
       Zero = 1
       End If
       SetFlags (A)
       PCounter = AddHexPc(PCounter, 1)
       PCounter = AddHexPc(PCounter, 0)
       
Case "BD" 'CMP L
     Ti1 = toInt(A)
     Ti2 = toInt(l)
     If (Ti1 > Ti2) Then
        cy = 0
        Zero = 0
        End If
     If (Ti1 < Ti2) Then
       cy = 1
       Zero = 0
       End If
     If (Ti1 = Ti2) Then
       cy = 0
       Zero = 1
       End If
       SetFlags (A)
       PCounter = AddHexPc(PCounter, 1)
       PCounter = AddHexPc(PCounter, 0)
       
Case "BE" 'CMP M
     Ti1 = toInt(A)
     ts1 = H + l
     ts1 = AddHexPc(ts1)
     M = GetData(ts1)
     Ti2 = toInt(M)
     If (Ti1 > Ti2) Then
        cy = 0
        Zero = 0
        End If
     If (Ti1 < Ti2) Then
       cy = 1
       Zero = 0
       End If
     If (Ti1 = Ti2) Then
       cy = 0
       Zero = 1
       End If
        SetFlags (A)
      PCounter = AddHexPc(PCounter, 1)
      PCounter = AddHexPc(PCounter, 0)
      
Case "D4" 'CNC 16-BIT

     PCounter = AddHexPc(PCounter, 1)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     If (cy = 0) Then
     If ((byte2 + byte1) = Sub1Add) Then
       If (step2 <> 1) Then
               Form2.DISPLAY = Space(2) + displaydata + strDot + Space(2)
               Displayed = True
       End If
    Else
    If ((byte2 + byte1) = Sub2Add) Then
               Form2.T1.Interval = Int(0.00785 * tolong(D + E))
               If (Form2.T1.Interval <= 10) Then Form2.T1.Interval = 10
               YES = "YES"
               Form2.T1.Enabled = True
                  While (YES = "YES")
                  DoEvents
                  Wend
     Else
               PushIT PCounter
     PCounter = AddHexPc(PCounter, 0)
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
    End If
   End If

     End If

Case "C4" 'CNZ16-BIT
     PCounter = AddHexPc(PCounter, 1)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     If (Zero = 0) Then
     If ((byte2 + byte1) = Sub1Add) Then
       If (step2 <> 1) Then
               Form2.DISPLAY = Space(2) + displaydata + strDot + Space(2)
               Displayed = True
       End If
    Else
    If ((byte2 + byte1) = Sub2Add) Then
               Form2.T1.Interval = Int(0.00785 * tolong(D + E))
               If (Form2.T1.Interval <= 10) Then Form2.T1.Interval = 10
               YES = "YES"
               Form2.T1.Enabled = True
                  While (YES = "YES")
                  DoEvents
                  Wend
     Else
               PushIT PCounter
               PCounter = AddHexPc(PCounter, 0)
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
    End If
   End If

     End If
     
Case "F4" 'CP 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     If (Sign = 0) Then
     If ((byte2 + byte1) = Sub1Add) Then
       If (step2 <> 1) Then
               Form2.DISPLAY = Space(2) + displaydata + strDot + Space(2)
               Displayed = True
       End If
    Else
    If ((byte2 + byte1) = Sub2Add) Then
               Form2.T1.Interval = Int(0.00785 * tolong(D + E))
               If (Form2.T1.Interval <= 10) Then Form2.T1.Interval = 10
               YES = "YES"
               Form2.T1.Enabled = True
                  While (YES = "YES")
                  DoEvents
                  Wend
     Else
               PushIT PCounter
        PCounter = AddHexPc(PCounter, 0)
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
    End If
   End If

     End If
     
Case "EC" 'CPE 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     If (Parity = 1) Then
          PushIT PCounter
     If ((byte2 + byte1) = Sub1Add) Then
       If (step2 <> 1) Then
               Form2.DISPLAY = Space(2) + displaydata + strDot + Space(2)
               Displayed = True
       End If
    Else
    If ((byte2 + byte1) = Sub2Add) Then
               Form2.T1.Interval = Int(0.00785 * tolong(D + E))
               If (Form2.T1.Interval <= 10) Then Form2.T1.Interval = 10
               YES = "YES"
               Form2.T1.Enabled = True
                  While (YES = "YES")
                  DoEvents
                  Wend
     Else
   
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
    End If
   End If

     End If
     

Case "FE" 'CPI 8-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     Ti1 = toInt(A)
     Ti2 = toInt(byte1)
     If (Ti1 > Ti2) Then
        cy = 0
        Zero = 0
        End If
     If (Ti1 < Ti2) Then
       cy = 1
       Zero = 0
       End If
     If (Ti1 = Ti2) Then
       cy = 0
       Zero = 1
       End If
       SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)

Case "E4" 'CPO 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     If (Parity = 0) Then
     If ((byte2 + byte1) = Sub1Add) Then
       If (step2 <> 1) Then
               Form2.DISPLAY = Space(2) + displaydata + strDot + Space(2)
               Displayed = True
       End If
       
    If ((byte2 + byte1) = Sub2Add) Then
               Form2.T1.Interval = Int(0.00785 * tolong(D + E))
               If (Form2.T1.Interval <= 10) Then Form2.T1.Interval = 10
               YES = "YES"
               Form2.T1.Enabled = True
                  While (YES = "YES")
                  DoEvents
                  Wend
              
     Else
      PCounter = AddHexPc(PCounter, 0)
   PushIT PCounter
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
    End If
   End If

     End If
     
 Case "CC" 'CZ 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     byte2 = GetData(PCounter)
     
     PCounter = AddHexPc(PCounter, 1)
     If (Zero = 1) Then
If ((byte2 + byte1) = Sub1Add) Then
       If (step2 <> 1) Then
               Form2.DISPLAY = Space(2) + displaydata + strDot + Space(2)
               Displayed = True
       End If
       
    If ((byte2 + byte1) = Sub2Add) Then
               Form2.T1.Interval = Int(0.00785 * tolong(D + E))
               If (Form2.T1.Interval <= 10) Then Form2.T1.Interval = 10
               YES = "YES"
               Form2.T1.Enabled = True
                  While (YES = "YES")
                  DoEvents
                  Wend
              
     Else
      PCounter = AddHexPc(PCounter, 0)
   PushIT PCounter
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
    End If
   End If
     End If
     
     
     
Case "27" ' DAA
     A = DaaIt(A)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)

Case "09" 'DAD B
     ts1 = H + l
     ts2 = B + c
     ADD16BIT ts1, ts2
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "19" 'DAD D
     ts1 = H + l
     ts2 = D + E
     ADD16BIT ts1, ts2
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "29" 'DAD H
     ts1 = H + l
     ts2 = H + l
     ADD16BIT ts1, ts2
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "39" 'DAD SP
     ts1 = H + l
     ts2 = HByte("SP") + LByte("SP")
     ADD16BIT ts1, ts2
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "3D" 'DCR A
     A = SubInr(A, , 1)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     
Case "05" 'DCR B
     B = SubInr(B, , 1)
     SetFlags (B)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "0D" 'DCR C
     c = SubInr(c, , 1)
     SetFlags (c)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "15" 'DCR D
     D = SubInr(D, , 1)
     SetFlags (D)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "1D" 'DCR E
    E = SubInr(E, , 1)
     SetFlags (E)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "25" 'DCR H
    H = SubInr(H, , 1)
     SetFlags (H)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "2D" 'DCR L
     l = SubInr(l, , 1)
     SetFlags (l)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "35" 'DCR M
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     M = GetData(ts1)
     M = SubInr(M, , 1)
     SetFlags (M)
     SetData ts1, M
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "0B" 'DCX B
     ts1 = B + c
     ts1 = SubHexLong(ts1, , 1)
     B = Left(ts1, 2)
     c = Right(ts1, 2)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "1B" 'DCX D
     ts1 = D + E
     ts1 = SubHexLong(ts1, , 1)
     D = Left(ts1, 2)
     E = Right(ts1, 2)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "2B" 'DCX H
     ts1 = H + l
     ts1 = SubHexLong(ts1, , 1)
     H = Left(ts1, 2)
     l = Right(ts1, 2)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "3B" 'DCX SP
     ts1 = SP
     ts1 = SubHexLong(ts1, , 1)
     SP = ts1
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 1
   
Case "F3" 'DI
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "FB" 'EI
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)

Case "76" 'HLT
     ExecPgm = 1
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "DB" 'IN 8-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "3C" 'INR A
     A = AddInr(A, , 1)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "04" 'INR B
     B = AddInr(B, , 1)
     SetFlags (B)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "0C" 'INR C
     c = AddInr(c, , 1)
     SetFlags (c)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "14" 'INR D
     D = AddInr(D, , 1)
     SetFlags (D)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "1C" 'INR E
     E = AddInr(E, , 1)
     SetFlags (E)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "24" 'INR H
     H = AddInr(H, , 1)
     SetFlags (H)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "2C" 'INR L
     l = AddInr(l, , 1)
     SetFlags (l)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "34" 'INR M
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     M = GetData(ts1)
     M = AddInr(M, , 1)
     SetFlags (M)
     SetData ts1, M
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)

Case "03" 'INX B
     ts1 = B + c
     ts1 = AddHexLong(ts1, , 1)
     B = Left(ts1, 2)
     c = Right(ts1, 2)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)

Case "13" 'INX D
     ts1 = D + E
     ts1 = AddHexLong(ts1, , 1)
     D = Left(ts1, 2)
     E = Right(ts1, 2)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "23" 'INX H
     ts1 = H + l
     ts1 = AddHexLong(ts1, , 1)
     H = Left(ts1, 2)
     l = Right(ts1, 2)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
Case "33" 'INX SP
     ts1 = SP
     ts1 = AddHexLong(ts1, , 1)
     SP = ts1
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 2
     
  Case "DA" 'JC 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (cy = 1) Then
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
     End If
     

 Case "FA" 'JM 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (Sign = 1) Then
     
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
     End If

 Case "C3" 'JMP 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
     
 Case "D2" 'JNC 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (cy = 0) Then
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
     End If

 Case "C2" 'JNZ 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (Zero = 0) Then
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
     End If
     
 Case "F2" 'JP 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (Sign = 0) Then
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
     End If

 Case "EA" 'JPE 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (Parity = 1) Then
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
     End If
     
 Case "E2" 'JPO 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (Parity = 0) Then
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
     End If
     
 Case "CA" 'JZ 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (Zero = 1) Then
     PCounter = byte2 + byte1
     PCounter = AddHexPc(PCounter, 0)
     End If
     
  Case "3A" 'LDA 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     ts1 = byte2 + byte1
     ts1 = AddHexPc(ts1, 0)
     A = GetData(ts1)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
  Case "0A" 'LDAX B
     ts1 = B + c
     ts1 = AddHexPc(ts1, 0)
     A = GetData(ts1)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
  
  
   Case "1A" 'LDAX D
     ts1 = D + E
     ts1 = AddHexPc(ts1, 0)
     A = GetData(ts1)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
    Case "2A" 'LHLD 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     ts1 = byte2 + byte1
     ts1 = AddHexPc(ts1, 0)
     l = GetData(ts1)
     ts1 = AddHexLong(ts1, , 1)
     ts1 = AddHexPc(ts1, 0)
     H = GetData(ts1)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
    Case "01" 'LXI B,16-bit
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     B = byte2
     c = byte1
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "11" 'LXI D,16-bit
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     D = byte2
     E = byte1
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      Case "21" 'LXI H,16-bit
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     H = byte2
     l = byte1
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     M = GetData(ts1)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      Case "31" 'LXI SP,16-bit
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     
     SP = byte2 + byte1
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 3
     
     
     Case "7F" 'MOV A A
     A = A
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "78" 'MOV A B
     A = B
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "79" 'MOV A C
     A = c
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "7A" 'MOV A D
     A = D
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "7B" 'MOV A E
     A = E
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "7C" 'MOV A H
     A = H
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "7D" 'MOV A L
     A = l
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "7E" 'MOV A M
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     M = GetData(ts1)
     A = M
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "47" 'MOV B A
     B = A
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "40" 'MOV B B
     B = B
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "41" 'MOV B C
     B = c
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "42" 'MOV B D
     B = D
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     
      
     Case "43" 'MOV B E
     B = E
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "44" 'MOV B H
     B = H
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "45" 'MOV B L
     B = l
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "46" 'MOV B M
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     M = GetData(ts1)
     B = M
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "4F" 'MOV C A
     c = A
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "48" 'MOV C B
     c = B
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "49" 'MOV C C
     c = c
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "4A" 'MOV C D
     c = D
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "4B" 'MOV C E
     c = E
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "4C" 'MOV C H
     c = H
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "4D" 'MOV C L
     c = l
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "4E" 'MOV C M
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     M = GetData(ts1)
     c = M
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "57" 'MOV D A
     D = A
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "50" 'MOV D B
     D = B
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "51" 'MOV D C
     D = c
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "52" 'MOV D D
     D = D
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "53" 'MOV D E
     D = E
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "54" 'MOV D H
     D = H
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "55" 'MOV D L
     D = l
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "56" 'MOV D M
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     M = GetData(ts1)
     D = M
     
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "5F" 'MOV E A
     E = A
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "58" 'MOV E B
     E = B
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "59" 'MOV E C
     E = c
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "5A" 'MOV E D
     E = D
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "5B" 'MOV E E
     E = E
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "5C" 'MOV E H
     E = H
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "5D" 'MOV E L
     E = l
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "5E" 'MOV E M
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     M = GetData(ts1)
     E = M
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "67" 'MOV H A
     H = A
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "60" 'MOV H B
     H = B
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "61" 'MOV H C
     H = c
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "62" 'MOV H D
     H = D
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "63" 'MOV H E
     H = E
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "64" 'MOV H H
     H = H
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "65" 'MOV H L
     H = l
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "66" 'MOV H M
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     M = GetData(ts1)
     H = M
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "6F" 'MOV L A
     l = A
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "68" 'MOV L B
     l = B
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "69" 'MOV L C
     l = c
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     
     Case "6A" 'MOV L D
     l = D
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "6B" 'MOV L E
     l = E
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "6C" 'MOV L H
     l = H
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "6D" 'MOV L L
     l = l
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "6E" 'MOV L M
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     M = GetData(ts1)
     l = M
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "77" 'MOV M A
     M = A
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     SetData ts1, M
     BugAdd(1) = ts1
     BugData(1) = M
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "70" 'MOV M B
     M = B
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     SetData ts1, M
     BugAdd(1) = ts1
     BugData(1) = M
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "71" 'MOV M C
     M = c
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     SetData ts1, M
     BugAdd(1) = ts1
     BugData(1) = M
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "72" 'MOV M D
     M = D
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     SetData ts1, M
     BugAdd(1) = ts1
     BugData(1) = M

     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
      
     Case "73" 'MOV M E
     M = E
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     SetData ts1, M
     BugAdd(1) = ts1
     BugData(1) = M
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "74" 'MOV M H
     M = H
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     SetData ts1, M
     BugAdd(1) = ts1
     BugData(1) = M
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "75" 'MOV M L
     M = l
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     SetData ts1, M
     BugAdd(1) = ts1
     BugData(1) = M
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "3E" 'MVI A-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     A = byte1
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "06" 'MVI B-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     B = byte1
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "0E" 'MVI C-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     c = byte1
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)

     Case "16" 'MVI D-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     D = byte1
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)

     Case "1E" 'MVI E-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     E = byte1
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)

     Case "26" 'MVI H-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     H = byte1
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)

     Case "2E" 'MVI L-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     l = byte1
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)

     Case "36" 'MVI M-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     M = byte1
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     SetData ts1, M
     BugAdd(1) = ts1
     BugData(1) = M

     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)

     Case "00" 'NOP
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "B7"  'ORA A
     A = ORHex(A, A)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)

     Case "B0"  'ORA B
     A = ORHex(A, B)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "B1"  'ORA C
     A = ORHex(A, c)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)

     Case "B2"  'ORA D
     A = ORHex(A, D)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "B3"  'ORA E
     A = ORHex(A, E)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "B4"  'ORA H
     A = ORHex(A, H)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
          
     Case "B5"  'ORA L
     A = ORHex(A, l)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "B6"  'ORA M
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     M = GetData(ts1)
     A = ORHex(A, M)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
          
     Case "F6"  'ORI 8-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     A = ORHex(A, byte1)
     cy = 0
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
          
     Case "D3" 'OUT 8-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "E9" 'PCHL
     ts1 = H + l
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     PCounter = ts1
    
     
     Case "C1" 'POP B
     c = GetData(SP)
     SP = AddHexLong(SP, , 1)
     B = GetData(SP)
     SP = AddHexLong(SP, , 1)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 4
     
     Case "D1" 'POP D
     E = GetData(SP)
     SP = AddHexLong(SP, , 1)
     D = GetData(SP)
     SP = AddHexLong(SP, , 1)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 4
     
     Case "E1" 'POP H
     l = GetData(SP)
     SP = AddHexLong(SP, , 1)
     H = GetData(SP)
     SP = AddHexLong(SP, , 1)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 4
     
     Case "F1" 'POP PSW
     ts1 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     ts2 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     PSW = ts2 + ts1
     SetPSW (PSW)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 4
     
     Case "C5" 'PUSH B
     SP = SubHexLong(SP, , 1)
     SetData SP, B
     SP = SubHexLong(SP, , 1)
      SetData SP, c
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 5
     
     
 Case "D5" 'PUSH D
     SP = SubHexLong(SP, , 1)
     SetData SP, D
     SP = SubHexLong(SP, , 1)
      SetData SP, E
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 5
     
      Case "E5" 'PUSH H
     SP = SubHexLong(SP, , 1)
     SetData SP, H
     SP = SubHexLong(SP, , 1)
      SetData SP, l
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 5
     
      Case "F5" 'PUSH PSW
      SP = SubHexLong(SP, , 1)
      SetData SP, HByte("PSW")
     SP = SubHexLong(SP, , 1)
      SetData SP, LByte("PSW")
      
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 5
     
     
     Case "17" 'RAL
     A = RALCY(A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "1F" 'RAR
     A = RARCY(A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
         
     Case "D8" 'RC
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (cy = 1) Then
     ts1 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     ts2 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     PCounter = ts2 + ts1
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 4
     End If
     
     Case "C9" 'RET
     ts1 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     ts2 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     PCounter = ts2 + ts1
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 4
     
    Case "20" ' RIM
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     
     Case "07" 'RLC
     A = RLC(A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "F8" 'RM
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (Sign = 1) Then
     ts1 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     ts2 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     PCounter = ts2 + ts1
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 4
     End If
     
     Case "D0" 'RNC
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (cy = 0) Then
     ts1 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     ts2 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     PCounter = ts2 + ts1
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 4
     End If
     
     Case "C0" 'RNZ
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (Zero = 0) Then
     ts1 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     ts2 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     PCounter = ts2 + ts1
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 4
     End If
     
     Case "F0" 'RP
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (Sign = 0) Then
     ts1 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     ts2 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     PCounter = ts2 + ts1
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 4
     End If
     
     Case "E8" 'RPE
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (Parity = 1) Then
     ts1 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     ts2 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     PCounter = ts2 + ts1
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 4
     End If
     
     Case "E0" 'RPO
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (Parity = 0) Then
     ts1 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     ts2 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     PCounter = ts2 + ts1
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 4
     End If
     
     Case "0F" 'RRC
     A = RRC(A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "C7" 'RST 0
     PCounter = "0000"
     ExecPgm = 1
     Case "CF" 'RST 1
     PCounter = "0008"
     ExecPgm = 1
     Case "D7" 'RST 2
     PCounter = "0010"
     ExecPgm = 1
     Case "DF" 'RST 3
     PCounter = "0018"
     ExecPgm = 1
     Case "E7" 'RST 4
     PCounter = "0020"
     ExecPgm = 1
     Case "EF" 'RST 5
     PCounter = "0028"
     ExecPgm = 1
     Case "F7" 'RST 6
     PCounter = "0030"
     ExecPgm = 1
     Case "FF" 'RST 7
     PCounter = "0038"
     ExecPgm = 1
     
     Case "C8" 'RZ
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     If (Zero = 1) Then
     ts1 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     ts2 = GetData(SP)
     SP = AddHexLong(SP, , 1)
     PCounter = ts2 + ts1
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 4
     End If
  
     Case "9F" 'SBB A
     A = SubHex(A, A, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "98" 'SBB B
     A = SubHex(A, B, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "99" 'SBB C
     A = SubHex(A, c, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "9A" 'SBB D
     A = SubHex(A, D, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "9B" 'SBB E
     A = SubHex(A, E, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "9C" 'SBB H
     A = SubHex(A, H, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "9D" 'SBB L
     A = SubHex(A, l, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "9E" 'SBB M
     ts1 = H + l
     ts1 = AddHexPc(ts1, 0)
     M = GetData(ts1)
     A = SubHex(A, M, cy)
     SetFlags (A)
     SetData ts1, M
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "DE" 'SBI 8-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     A = SubHex(A, byte1, cy)
     SetFlags (A)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     Case "22" 'SHLD 16-BIT
     
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     
     ts1 = byte2 + byte1
     ts1 = AddHexPc(ts1, 0)
     SetData ts1, l
     BugAdd(1) = ts1
     BugData(1) = l
     ts1 = AddHexLong(ts1, , 1)
     ts1 = AddHexPc(ts1, 0)
     SetData ts1, H
     BugAdd(2) = ts1
     BugData(2) = H

     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
    Case "30" 'SIM
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     
     
     Case "F9" 'SPHL
     SP = H + l
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     StackCase = 3
    
     Case "32" 'STA 16-BIT
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte1 = GetData(PCounter)
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     byte2 = GetData(PCounter)
     ts1 = byte2 + byte1
     If (ts1 = displayadd) Then
     BugAdd(1) = ts1
     BugData(1) = A
     displaydata = A
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     Else
     ts1 = AddHexPc(ts1, 0)
     SetData ts1, A
     BugAdd(1) = ts1
     BugData(1) = A
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     End If
     
    Case "02" 'STAX B
    byte1 = B
    byte2 = c
    ts1 = byte1 + byte2
    If (ts1 = displayadd) Then
     BugAdd(1) = ts1
     BugData(1) = A
     displaydata = A
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     Else
     
    ts1 = AddHexPc(ts1, 0)
    SetData ts1, A
    BugAdd(1) = ts1
     BugData(1) = A
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
    End If
    
    Case "12" 'STAX D
    byte1 = D
    byte2 = E
    ts1 = byte1 + byte2
     If (ts1 = displayadd) Then
     BugAdd(1) = ts1
     BugData(1) = A
     displaydata = A
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
     Else
    ts1 = AddHexPc(ts1, 0)
    SetData ts1, A
    BugAdd(1) = ts1
     BugData(1) = A

    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    End If
    
    Case "37" 'STC
    cy = 1
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
    Case "97" 'SUB A
    A = SubHex(A, A, 0)
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
    Case "90" 'SUB B
    A = SubHex(A, B, 0)
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
    Case "91" 'SUB C
    A = SubHex(A, c, 0)
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
    Case "92" 'SUB D
    A = SubHex(A, D, 0)
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
    Case "93" 'SUB E
    A = SubHex(A, E, 0)
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
    Case "94" 'SUB H
    A = SubHex(A, H, 0)
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
    Case "95" 'SUB L
    A = SubHex(A, l, 0)
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
    Case "96" 'SUB M
    ts1 = H + l
    ts1 = AddHexPc(ts1, 0)
    M = GetData(ts1)
    A = SubHex(A, M, 0)
    SetFlags (A)
    SetData ts1, M
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
    Case "D6" 'SUI 8-BIT
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    byte1 = GetData(PCounter)
    A = SubHex(A, byte1, 0)
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
   Case "EB" 'XCHG
   ts1 = H
   ts2 = l
   H = D
   l = E
   D = ts1
   E = ts2
   PCounter = AddHexPc(PCounter, 1)
   PCounter = AddHexPc(PCounter, 0)
   
   Case "AF" 'XRA A
    A = XORHex(A, A)
    cy = 0
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
   Case "A8" 'XRA B
    A = XORHex(A, B)
    cy = 0
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
    Case "A9" 'XRA C
    A = XORHex(A, c)
    cy = 0
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
    Case "AA" 'XRA D
    A = XORHex(A, D)
    cy = 0
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
   Case "AB" 'XRA E
    A = XORHex(A, E)
    cy = 0
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
   Case "AC" 'XRA H
    A = XORHex(A, H)
    cy = 0
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
    Case "AD" 'XRA L
    A = XORHex(A, l)
    cy = 0
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
    Case "AE" 'XRA M
    ts1 = H + l
    ts1 = AddHexPc(ts1, 0)
    M = GetData(ts1)
    A = XORHex(A, M)
    cy = 0
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
    Case "EE" 'XRI 8-BIT
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    byte1 = GetData(PCounter)
    A = XORHex(A, byte1)
    cy = 0
    SetFlags (A)
    PCounter = AddHexPc(PCounter, 1)
    PCounter = AddHexPc(PCounter, 0)
    
    Case "E3" 'XTHL
     ts1 = l
     ts2 = H
     tempsp = SP
     l = GetData(SP)
     SetData SP, ts1
     SP = AddHexLong(SP, , 1)
     H = GetData(SP)
     SetData SP, ts2
     SP = tempsp
     PCounter = AddHexPc(PCounter, 1)
     PCounter = AddHexPc(PCounter, 0)
    
    End Select
    PSW = A + GetFlagReg

    If (step2 = 1) Then
    If (nonstop = 0) Then
    step1 = 0
     End If
     putdata stradd, strdisply
     End If
     DoEvents
    End If
  End If
 Loop Until (ExecPgm = 1 Or ErrPgm = 1 Or break = 1)
End Function

