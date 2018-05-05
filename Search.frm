VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Move"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox TxtOPcode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox TxtNemo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Text            =   "       Use  Me"
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sn As Integer
Dim so As Integer
Function Getnemo(str As String) As String
Dim S1 As String
Dim l As Integer
Dim j As Integer
Dim k As Integer
Dim i As Integer
TxtOPcode = " "
l = Len(str)
k = 0
For i = 1 To 247
     S1 = str
     For j = 1 To l
          If (Mid(S1, j, 1) = Mid(InSet(i).Nemo, j, 1)) Then
          If (j > k) Then
          k = j
          Getnemo = InSet(i).Nemo
          TxtOPcode = InSet(i).OpCode
          End If
          Else
          Exit For
          End If
       Next j
          
  Next i
End Function

Private Sub Command1_Click()
TxtNemo = "MOV "
sn = Len(TxtNemo)
TxtNemo = Getnemo(TxtNemo.Text)
If (sn < 0) Then sn = 0
TxtNemo.SelStart = sn
If (Len(TxtNemo.Text) - sn > 0) Then
TxtNemo.SelLength = Len(TxtNemo.Text) - sn
Else
TxtNemo.SelLength = 0
End If
TxtNemo.SetFocus
End Sub

Private Sub Form_Load()
inicialise
End Sub

Private Sub TxtNemo_GotFocus()
TxtNemo.SelStart = sn
TxtNemo.SelLength = Len(TxtNemo)
End Sub



Private Sub TxtNemo_KeyUp(KeyCode As Integer, Shift As Integer)
SPressed = SPressed + Chr(KeyCode)
If (KeyCode = 8) Then Exit Sub
sn = Len(TxtNemo)
TxtNemo = Getnemo(TxtNemo.Text)
If (sn < 0) Then sn = 0
TxtNemo.SelStart = sn
If (Len(TxtNemo.Text) - sn > 0) Then
TxtNemo.SelLength = Len(TxtNemo.Text) - sn
Else
TxtNemo.SelLength = 0
End If

End Sub
Function GetOpcode(str As String) As String
Dim S1 As String
Dim l As Integer
Dim j As Integer
Dim k As Integer
Dim i As Integer
TxtNemo = " "
l = Len(str)
k = 0
For i = 1 To 247
     S1 = str
     For j = 1 To l
          If (Mid(S1, j, 1) = Mid(InSet(i).OpCode, j, 1)) Then
          If (j > k) Then
          k = j
          GetOpcode = InSet(i).OpCode
          TxtNemo = InSet(i).Nemo
          End If
          Else
          Exit For
          End If
       Next j
          
  Next i
End Function

Private Sub TxtOPcode_KeyUp(KeyCode As Integer, Shift As Integer)
If (KeyCode = 8) Then Exit Sub
so = Len(TxtOPcode)
TxtOPcode = GetOpcode(TxtOPcode.Text)
If (so < 0) Then so = 0
TxtOPcode.SelStart = so
If (Len(TxtOPcode.Text) - so > 0) Then
TxtOPcode.SelLength = Len(TxtOPcode.Text) - so
Else
TxtOPcode.SelLength = 0
End If

End Sub
