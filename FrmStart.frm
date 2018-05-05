VERSION 5.00
Begin VB.Form frmstart 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   4320
      Top             =   600
   End
   Begin VB.Label Label1 
      Caption         =   "VI Sem E&&C "
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Caption         =   "MICROPROCESSOR  SIMULATOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3645
   End
   Begin VB.Label Label3 
      Caption         =   "Developer: SHIVAKUMAR VIRAKTAMATH"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "8085"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   885
   End
End
Attribute VB_Name = "frmstart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Dim ret As Long
 SetWindowPos Me.hwnd, conHwndNoTopmost, 250, 200, 313, 198, conSwpShowWindow
 ret = SetWindowLong(hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)

Me.Height = 2715
Me.Width = 4095

End Sub







Private Sub Label2_Click()

End Sub

Private Sub Timer2_Timer()

SubMain

End Sub

