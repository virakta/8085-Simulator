VERSION 5.00
Begin VB.Form Frmmarg 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Margins"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3915
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancle 
      Cancel          =   -1  'True
      Caption         =   "&Cancle"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "O&K"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Txtx 
      Height          =   405
      Left            =   1800
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Txty 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1200
      Width           =   615
   End
   Begin VB.VScrollBar VSc2 
      Height          =   375
      Left            =   2640
      Max             =   0
      Min             =   35
      TabIndex        =   4
      Top             =   600
      Width           =   255
   End
   Begin VB.VScrollBar VSc1 
      Height          =   375
      Left            =   2640
      Max             =   0
      Min             =   35
      TabIndex        =   3
      Top             =   1200
      Value           =   35
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Set Margins for print"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Y--Marin"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "X-- Margin"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Frmmarg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancle_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
xmarg = VSc2.Value
ymarg = VSc1.Value
xmarg = xmarg * 50
ymarg = ymarg * 50
NextPage = 0
preview
Unload Me
End Sub


Private Sub Form_Load()
VSc2.Value = xmarg \ 50
VSc1.Value = ymarg \ 50
Txtx = STR(VSc2.Value)
Txty = STR(VSc1.Value)
End Sub

Private Sub VSc1_Change()
Txty = STR(VSc1.Value)
End Sub

Private Sub VSc2_Change()
Txtx = STR(VSc2.Value)
End Sub
