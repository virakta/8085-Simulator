VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "CY"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k As Integer
Private Sub Command1_Click()
Text3 = DaaIt(Text1)

End Sub

Private Sub Command2_Click()
Text4 = str(valueof(Text3))
End Sub

Private Sub Form_Load()
Text4 = Val(CY)
k = 0
Initialise
Randamise
End Sub
