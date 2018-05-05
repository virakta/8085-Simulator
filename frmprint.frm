VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form form1 
   Caption         =   "Print PreView"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Ppage 
      Caption         =   "Previous Page"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Npage 
      Caption         =   "NextPage"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdfont 
      Caption         =   "Set Font"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SetMargins"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   5640
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Printer Setup"
      FromPage        =   1
      ToPage          =   2
      Orientation     =   2
   End
   Begin VB.Frame Frame 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   -360
      TabIndex        =   0
      Top             =   840
      Width           =   11775
      Begin VB.PictureBox Pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   4440
         ScaleHeight     =   4095
         ScaleWidth      =   4815
         TabIndex        =   3
         Top             =   1680
         Width           =   4815
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub cmdfont_Click()
cdl1.Flags = cdlCFEffects + cdlCFForceFontExist + cdlCFWYSIWYG
On Error GoTo endit
cdl1.Flags = cdlCFScreenFonts
cdl1.ShowFont
Pic.Font.Name = cdl1.FontName
Pic.Font.Bold = cdl1.FontBold
Pic.Font.Italic = cdl1.FontItalic
Pic.Font.Size = cdl1.FontSize


NextPage = 0
getpstr
preview
endit:
End Sub

Private Sub CmdPrint_Click()
printit
End Sub

Private Sub Command1_Click()
Frmmarg.Show
End Sub

Private Sub Command2_Click()
preview
End Sub

Private Sub Form_Load()
On Error GoTo pin
Printer.FontBold = Pic.FontBold '
Printer.FontItalic = Pic.FontItalic
Printer.FontName = Pic.FontName
Printer.FontSize = Pic.FontSize
Printer.FontStrikethru = Pic.FontStrikethru
Printer.FontUnderline = Pic.FontUnderline
pin:
SModePH = Printer.ScaleHeight '15300
SModePW = Printer.ScaleWidth '11700
MaxX = 11000 '12120
MaxY = 7000 '8700
xmarg = 700
ymarg = 700

If (SModePH >= SModePW) Then
THeight = MaxY
If (SModePH > 0) Then
TWidth = Int(MaxY * SModePW / SModePH)
End If
sx = SModePH \ MaxY
Else
TWidth = MaxX
THeight = Int(MaxX * SModePH / SModePW)
sx = SModePW \ MaxX
End If
Pic.Top = form1.Top + 500
Pic.Left = form1.Left + 500
Pic.Height = THeight
Pic.Width = TWidth
Pic.Visible = True
NextPage = 0
getpstr
preview

End Sub


Private Sub Npage_Click()
Ppage.Enabled = True
preview NextPage
Npage.Enabled = False
End Sub

Private Sub Ppage_Click()
Ppage.Enabled = False
preview
End Sub
