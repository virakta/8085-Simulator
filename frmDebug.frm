VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBug 
   Caption         =   "Debug  Window"
   ClientHeight    =   8580
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView LV 
      Height          =   855
      Left            =   5760
      TabIndex        =   18
      Top             =   6120
      Visible         =   0   'False
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   1508
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Stack "
         Object.Width           =   1766
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Contents"
         Object.Width           =   1766
      EndProperty
   End
   Begin VB.CommandButton CmdStack 
      Caption         =   "&STACK"
      Height          =   375
      Left            =   10560
      TabIndex        =   17
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton CmdClouse 
      Caption         =   "&CLOSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   720
      Width           =   825
   End
   Begin VB.CommandButton CmdBreak 
      Caption         =   "&BREAK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   720
      Width           =   825
   End
   Begin VB.CommandButton CmdReset 
      Caption         =   "RESE&T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   720
      Width           =   825
   End
   Begin VB.TextBox TxtCY 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10680
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   13
      Text            =   "CY"
      Top             =   720
      Width           =   500
   End
   Begin VB.TextBox TxtZ 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9480
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   12
      Text            =   "Z"
      Top             =   720
      Width           =   500
   End
   Begin VB.TextBox TxtP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8280
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   11
      Text            =   "P"
      Top             =   720
      Width           =   500
   End
   Begin VB.TextBox TxtS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7200
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   10
      Text            =   "S"
      Top             =   720
      Width           =   500
   End
   Begin VB.TextBox TxTAdress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "C000"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton CmdRun 
      Caption         =   "&RUN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   825
   End
   Begin VB.CommandButton CmdStep 
      Caption         =   "&STEP"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   825
   End
   Begin MSComctlLib.ListView LView 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9551
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label6 
      Caption         =   "CY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   9
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   8
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Flag registers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Starting Adress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FrmBug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ResDebug As Integer
Dim friststep As Integer
Dim tempcounter As String

Private Sub CmdReset_Click()
CmdStep.Enabled = True
CmdRun.Enabled = True
CmdBreak.Enabled = False
LView.ListItems.Clear
friststep = 1
StackCount = 0
PCounter = TxTAdress
LV.ListItems.Clear
End Sub

Private Sub CmdStack_Click()
LV.Top = CmdStack.Top - LV.Height
LV.Visible = Not LV.Visible
End Sub

Private Sub Cmdstep_Click()
Dim check As Integer
If (friststep = 1) Then
    check = InBetween(TxTAdress.Text)
    If (Len(TxTAdress) <> 4) Then check = 1
     If (check = 1) Then
     MsgBox "Specify  Starting Adress Between C000 - DFFFF", vbOKOnly + vbDefaultButton1 + vbCritical, "Starting Adress"
     TxTAdress.SetFocus
     Exit Sub
     End If
     CmdBreak.Enabled = True
     step1 = 1
     step2 = 1
     nonstop = 0
     friststep = 0
     break = 0
     PCounter = TxTAdress.Text
     ResDebug = 0
     ResDebug = ExecPgm()
     friststep = 0
     CmdBreak.Enabled = False
        If (ResDebug <> 1 And break = 0) Then
        ResDebug = 0
        MsgBox "Error in Writing to an Acessable Adress", vbDefaultButton1 + vbOKOnly + vbInformation, "ERROR"
        End If
  Else
     step1 = 1
     nonstop = 0
  End If
     End Sub

Private Sub Cmdrun_Click()
Dim check As Integer
If (friststep = 1) Then
  check = InBetween(TxTAdress)
  If (Len(TxTAdress) <> 4) Then check = 1
  If (check = 1) Then
     MsgBox "Specify  Starting Adress Between C000 - DFFFF", vbOKOnly + vbDefaultButton1 + vbCritical, "Starting Adress"
     TxTAdress.SetFocus
     Exit Sub
     End If
   CmdBreak.Enabled = True
   friststep = 0
   PCounter = TxTAdress
   break = 0
   nonstop = 1
   step1 = 1
   step2 = 1
   ResDebug = 0
   ResDebug = ExecPgm()
   CmdBreak.Enabled = False
   If (ResDebug <> 1 And break = 0) Then
     ResDebug = 0
     MsgBox "Error in Writing to an Acessable Adress", vbDefaultButton1 + vbOKOnly + vbInformation, "ERROR"
     End If
  Else
  step1 = 1
  nonstop = 1
  End If
End Sub

Private Sub CmdBreak_Click()
break = 1
CmdBreak.Enabled = False
End Sub

Private Sub CmdClouse_Click()
break = 1
Unload Me
End Sub



Private Sub Form_Load()
friststep = 1
StackCount = 0
LV.Width = 2025
LV.Height = 2415
LV.Left = CmdStack.Left - 1000
LV.Top = CmdStack.Top - LV.Height
tempcounter = PCounter
LView.Height = FrmBug.Height - 2500
LView.Width = FrmBug.Width - 300
TxtS = ""
TxtP = ""
TxtZ = ""
TxtCY = ""

   Dim clmX As ColumnHeader
   Set clmX = LView.ColumnHeaders. _
   Add(, , " Memory Adress", LView.Width * 3 / (10 * 3))
   Set clmX = LView.ColumnHeaders. _
   Add(, , "Instractions", LView.Width * 3 / (10 * 2))
   Set clmX = LView.ColumnHeaders. _
   Add(, , "A", LView.Width * (5 / (10 * 8)))
    Set clmX = LView.ColumnHeaders. _
   Add(, , "B", LView.Width * (5 / (10 * 8)))
    Set clmX = LView.ColumnHeaders. _
   Add(, , "C", LView.Width * (5 / (10 * 8)))
    Set clmX = LView.ColumnHeaders. _
   Add(, , "D", LView.Width * (5 / (10 * 8)))
   Set clmX = LView.ColumnHeaders. _
   Add(, , "E", LView.Width * (5 / (10 * 8)))
    Set clmX = LView.ColumnHeaders. _
   Add(, , "H", LView.Width * (5 / (10 * 8)))
    Set clmX = LView.ColumnHeaders. _
   Add(, , "L", LView.Width * (5 / (10 * 8)))
    Set clmX = LView.ColumnHeaders. _
   Add(, , "M", LView.Width * (5 / (10 * 8)))
   Set clmX = LView.ColumnHeaders.Add(, , "Mem Address", LView.Width * (2 / (10 * 2) + 1 / 40))
   Set clmX = LView.ColumnHeaders.Add(, , "Data ", LView.Width * (2 / (10 * 2) + 1 / 40))

   
   LView.BorderStyle = ccFixedSingle
   LView.View = lvwReport
End Sub

Private Sub Form_Unload(Cancel As Integer)
PCounter = tempcounter
LView.ListItems.Clear
End Sub
