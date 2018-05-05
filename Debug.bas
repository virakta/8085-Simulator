Attribute VB_Name = "Debug"
Option Explicit
Public step1 As Integer
Public step2 As Integer
Public nonstop As Integer
Public break As Integer

Public strAdress As String * 4
Public strDot As String * 1
Public strdata As String * 2
Public TempAdress As String
Public TempData As String
Public ErrPgm As Integer
Public BAddCount As Integer
Public BAdd(1 To 3) As String
Public InsCount As String
Public InsData() As String
Public InsIndex As Integer
Public InsAdd As String
Public DelAdd(1 To 2) As String
Public delcount As Integer

Public ERRORU As Boolean
Public Regcount As Integer
Public ResetON As Boolean
Public AdressON As Boolean
Public AdressOK As Boolean
Public DataOK As Boolean
Public GoON As Boolean
Public ExaRegON As Boolean
Public ExaRegOK As Boolean
Public BMoveON As Boolean
Public InsON As Boolean
Public DelON As Boolean

Public BugAdd(1 To 2) As String
Public BugData(1 To 2) As String
Public BugCount As Integer
Public StackCase As Integer
Public StackCount As Integer



Sub DisplayError()
ERRORU = True
strAdress = "Err "
strdata = "  "
Form2.DISPLAY = strAdress + strDot + strdata
End Sub
Function InBetween(stradd As String) As Integer
Dim l As Long
InBetween = 0
l = tolong(stradd)
If ((49152 > l) Or (57343 < l)) Then InBetween = 1
End Function

Function AddHexPc(str1 As String, Optional c = 0) As String
Dim str2 As String
Dim Res As Integer
str2 = str1
If (c = 1) Then str2 = AddHexLong(str2, , c)
If (c = 0) Then
Res = InBetween(str2)
If (Res = 1) Then
ErrPgm = 1
DisplayError
End If
End If
AddHexPc = str2
End Function

Sub putdata(str1 As String, str2 As String)
Dim tstr1 As String
Dim tstr2 As String
 Dim itmX As ListItem
 Dim itmY As ListItem
      Set itmX = FrmBug.LView.ListItems.Add(, , str1)
      itmX.SubItems(1) = str2
      itmX.SubItems(2) = A
      itmX.SubItems(3) = B
      itmX.SubItems(4) = c
      itmX.SubItems(5) = D
      itmX.SubItems(6) = E
      itmX.SubItems(7) = H
      itmX.SubItems(8) = l
      itmX.SubItems(9) = M

          FrmBug.TxtS = STR(Sign)
          FrmBug.TxtP = STR(Parity)
          FrmBug.TxtZ = STR(Zero)
          FrmBug.TxtCY = STR(cy)
          If (BugAdd(1) <> "" And BugData(1) <> "") Then
          Set itmX = FrmBug.LView.ListItems.Item(BugCount)
          itmX.SubItems(10) = BugAdd(1)
          itmX.SubItems(11) = BugData(1)
          BugAdd(1) = ""
          BugData(1) = ""
          BugCount = BugCount + 1
          End If
          
          If (BugAdd(2) <> "" And BugData(2) <> "") Then
          Set itmX = FrmBug.LView.ListItems.Item(BugCount)
          itmX.SubItems(10) = BugAdd(2)
          itmX.SubItems(11) = BugData(2)
          BugAdd(2) = ""
          BugData(2) = ""
          BugCount = BugCount + 1
          End If
          Select Case StackCase
          Case 1:
               Set itmY = FrmBug.LV.ListItems.Add(1, , SP)
               itmY.SubItems(1) = GetData(SP)
               StackCount = StackCount + 1
          Case 2:
               On Error GoTo ER2
                FrmBug.LV.ListItems.Remove 1
                StackCount = StackCount - 1
ER2:
          Case 3:
               FrmBug.LV.ListItems.Clear
               Set itmY = FrmBug.LV.ListItems.Add(1, , SP)
               itmY.SubItems(1) = GetData(SP)
                StackCount = StackCount + 1

          Case 4:
               On Error GoTo ER1
               FrmBug.LV.ListItems.Remove 1
               FrmBug.LV.ListItems.Remove 1
               StackCount = StackCount - 2
ER1:
          Case 5:
               tstr1 = AddHexLong(SP, , 1)
               tstr2 = GetData(tstr1)
               Set itmY = FrmBug.LV.ListItems.Add(1, , tstr1)
               itmY.SubItems(1) = tstr2
               tstr1 = SP
               tstr2 = GetData(tstr1)
               Set itmY = FrmBug.LV.ListItems.Add(1, , tstr1)
               itmY.SubItems(1) = tstr2
                              

          End Select
          If (FrmBug.LV.Height < 225 * (3 + StackCount)) Then
          FrmBug.LV.Top = FrmBug.LV.Top - (225 * (4 + StackCount) - FrmBug.LV.Height)
          FrmBug.LV.Height = 225 * (4 + StackCount)
          End If
          If (FrmBug.LV.Height > 225 * (25 + StackCount)) Then
          FrmBug.LV.Top = FrmBug.LV.Top - (225 * (25 + StackCount) - FrmBug.LV.Height)
          FrmBug.LV.Height = Abs(225 * (25 + StackCount)) + 50
          End If
          StackCase = 0
       End Sub
Function BlockMove(str1 As String, str2 As String, str3 As String) As Integer
Dim I As Integer
Dim S1 As String
Dim S2 As String
Dim s3 As String
BlockMove = 0
If (InBetween(str1) = 1) Then
BlockMove = 1
Exit Function
End If
If (InBetween(str2) = 1) Then
BlockMove = 1
Exit Function
End If
If (InBetween(str2) = 1) Then
BlockMove = 1
Exit Function
End If
If (tolong(str2) - tolong(str1) < 0) Then
BlockMove = 1
Exit Function
End If
S1 = SubHexLong(str2, str1)
S2 = AddHexLong(str3, S1)
If (InBetween(S2) = 1) Then
BlockMove = 1
Exit Function
End If

S1 = str1
S2 = str3
While (S1 <> str2)
s3 = GetData(S1)
SetData S2, s3
S1 = AddHexLong(S1, , 1)
S2 = AddHexLong(S2, , 1)
Wend
End Function
Function Delete(str1 As String, str2 As String) As Integer
Dim S1 As String
Dim S2 As String
Delete = 0
If (InBetween(str1) = 1) Then
Delete = 1
Exit Function
End If
If (InBetween(str2) = 1) Then
Delete = 1
Exit Function
End If
If (tolong(str2) - tolong(str1) < 0) Then
Delete = 1
Exit Function
End If
S1 = AddHexLong(str2, , 1)
S2 = "DFFF"
Delete = BlockMove(S1, S2, str1)
End Function
