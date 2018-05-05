Attribute VB_Name = "Print"
Option Explicit

Public strprint() As String
Public pstartadd As String

Public xmarg As Integer
Public ymarg As Integer
Sub getpstr()
Dim j As Integer
Dim I As Integer
Dim nextadd As String
Dim pstart As String
Dim pByte1 As String
Dim PByte2 As String
Dim Pdata As String
Dim Pnemo As String
Dim Pbyte As Integer
Dim SubCount As Integer
Dim Counter As Integer
Dim SubAdd(1 To 100) As String
Dim exitloop As Integer
Dim pcount As Integer
SubCount = 0
pstart = pstartadd
pcount = 1
exitloop = 0
 Do
               If (pcount <> 1) Then
               pstart = nextadd
               End If
               Pdata = GetData(pstart)
               For j = 1 To 246
               If (Pdata = InSet(j).OpCode) Then
               Pnemo = InSet(j).Nemo
               Pbyte = InSet(j).ByteCount
               Exit For
               End If
               Next j
               If (Pbyte = 1) Then
               pByte1 = ""
               nextadd = AddHexPc(pstart, 1)
               End If
            
               If (Pbyte = 2) Then
               pByte1 = GetData(AddHexPc(pstart, 1))
               nextadd = AddHexPc(AddHexPc(pstart, 1), 1)
               End If
               
              If (Pbyte = 3) Then
             pByte1 = GetData(AddHexPc(pstart, 1))
             PByte2 = GetData(AddHexPc(AddHexPc(pstart, 1), 1))
             nextadd = AddHexPc(AddHexPc(AddHexPc(pstart, 1), 1), 1)
             pByte1 = PByte2 + pByte1
            End If
            
             Select Case Pdata
             Case "CD" 'CALL
             SubCount = SubCount + 1
             SubAdd(SubCount) = pByte1
             pByte1 = "SUB" + STR(SubCount)
             
             Case "DC" 'CALLONCARY
             SubCount = SubCount + 1
             SubAdd(SubCount) = pByte1
             pByte1 = "SUB" + STR(SubCount)
             
             Case "FC" 'MC
             SubCount = SubCount + 1
             SubAdd(SubCount) = pByte1
             pByte1 = "SUB" + STR(SubCount)
             
             Case "D4" 'CNC
             SubCount = SubCount + 1
             SubAdd(SubCount) = pByte1
             pByte1 = "SUB" + STR(SubCount)
             
             Case "C4" 'CNZ
             SubCount = SubCount + 1
             SubAdd(SubCount) = pByte1
             pByte1 = "SUB" + STR(SubCount)
             
             Case "F4" 'CP
             SubCount = SubCount + 1
             SubAdd(SubCount) = pByte1
             pByte1 = "SUB" + STR(SubCount)
             
             Case "EC" 'CPE
             SubCount = SubCount + 1
             SubAdd(SubCount) = pByte1
             pByte1 = "SUB" + STR(SubCount)
             
             Case "E4" 'CPO
             SubCount = SubCount + 1
             SubAdd(SubCount) = pByte1
             pByte1 = "SUB" + STR(SubCount)
             
             Case "CC" 'CZ
             SubCount = SubCount + 1
             SubAdd(SubCount) = pByte1
             pByte1 = "SUB" + STR(SubCount)
             
             Case "76" 'halt
             Line1 = 1
             If (SubCount = 0) Then
             exitloop = 1
             Else
             Counter = 1
             nextadd = SubAdd(Counter)
              End If
             
              Case "C9"
               SubCount = SubCount - 1
               If (SubCount = 0) Then
               exitloop = 1
               Else
               Counter = Counter + 1
               nextadd = SubAdd(Counter)
               End If
               
            End Select
            
          ReDim Preserve strprint(pcount)
          strprint(pcount) = pstart + Space(20) + Pnemo + pByte1
          pcount = pcount + 1
          If (Line1 = 1) Then
          Line1 = 0
          ReDim Preserve strprint(pcount)
          strprint(pcount) = Space(200)
          pcount = pcount + 1
          End If
          
          If (pcount > 1500) Then exitloop = 1
          Loop Until (exitloop = 1)
         
End Sub



Public Sub preview(Optional np = 1)
Dim I As Long
Dim c As Long
Dim lx As Integer
Dim str1 As String
Dim ly As Integer
Dim j As Integer
lx = 50
ly = 50
j = 1
form1.Pic.Cls
form1.Pic.CurrentY = ymarg + ly
For I = np To UBound(strprint)
form1.Pic.CurrentX = xmarg + lx
form1.Pic.CurrentY = ymarg + (j - 1) * form1.Pic.TextHeight("SAMPLE") + ly
If (form1.Pic.CurrentY + form1.Pic.TextHeight("SAMPLE")) > form1.Pic.ScaleHeight Then
NextPage = I
Exit For
End If
str1 = Left(strprint(I), (form1.Pic.ScaleWidth - 2 * xmarg - lx) * 3 / form1.Pic.TextWidth("S  "))
form1.Pic.Print str1
j = j + 1
Next I
c = form1.Pic.ForeColor
form1.Pic.ForeColor = &H80&
form1.Pic.DrawStyle = 2
form1.Pic.Line (xmarg, 0)-(xmarg, form1.Pic.Height)
form1.Pic.Line (form1.Pic.Width - xmarg, 0)-(form1.Pic.Width - xmarg, form1.Pic.Height)
form1.Pic.ForeColor = c
If NextPage <> 0 Then form1.Npage.Enabled = True
End Sub

Public Sub printit()
On Error GoTo noprinter1
Dim I As Long
Dim c As Long
Dim lx As Integer
Dim ly As Integer
Printer.Font.Bold = form1.Pic.Font.FontBold
Printer.Font.Italic = form1.Pic.Font.Italic
Printer.Font.Name = form1.Pic.Font.Name
Printer.Font.Size = sx * form1.Pic.Font.Size

lx = 50
ly = 50
Printer.CurrentY = ymarg + ly
For I = 1 To UBound(strprint)
Printer.CurrentX = xmarg + lx
Printer.CurrentY = ymarg + (I - 1) * form1.Pic.TextHeight("SAMPLE") + ly
Printer.Print strprint(I)
Next I
Printer.EndDoc
noprinter1:
End Sub

