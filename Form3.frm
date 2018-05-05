VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   810
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Text2 = DaaIt(Text1)
Text3 = cy
End Sub

Private Sub Command2_Click()
Text3 = AddInr(Text1, , 1)
Text6 = cy

End Sub

Private Sub Command3_Click()
Text3 = SubInr(Text1, , 1)
Text6 = cy

End Sub

Private Sub Form_Load()
Text4 = cy
End Sub

