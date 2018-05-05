VERSION 5.00
Object = "{528EA9AE-5856-11D4-81F8-0080C8056F3D}#11.0#0"; "keycmd1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Microprocessor  Programing 8085 "
   ClientHeight    =   8220
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   15345
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "FormKey2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8220
   ScaleWidth      =   15345
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Tbar 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   714
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      ImageList       =   "Img"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "topen"
            Object.ToolTipText     =   " Open Files"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "tsave"
            Object.ToolTipText     =   "Save File"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "tprint"
            Object.ToolTipText     =   "Print Your Form"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "tpreview"
            Object.ToolTipText     =   "Preview File Print"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "tstep"
            Object.ToolTipText     =   "Step Run"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin VB.Timer T1 
         Enabled         =   0   'False
         Left            =   4320
         Top             =   500
      End
   End
   Begin VB.PictureBox keystate 
      Height          =   615
      Left            =   1200
      ScaleHeight     =   555
      ScaleWidth      =   795
      TabIndex        =   69
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "                             OPCODE   TABLE "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   30
      Top             =   2640
      Width           =   6015
      Begin VB.CommandButton CmdA 
         Caption         =   "PSW"
         Height          =   375
         Index           =   9
         Left            =   5040
         TabIndex        =   68
         TabStop         =   0   'False
         ToolTipText     =   "Select Any Register"
         Top             =   3840
         Width           =   495
      End
      Begin VB.CommandButton CmdA 
         Caption         =   "SP"
         Height          =   375
         Index           =   8
         Left            =   4200
         TabIndex        =   67
         TabStop         =   0   'False
         ToolTipText     =   "Select Any Register"
         Top             =   3840
         Width           =   495
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "JMP"
         Height          =   375
         Index           =   16
         Left            =   960
         TabIndex        =   66
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   3720
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "LXI"
         Height          =   375
         Index           =   9
         Left            =   1320
         TabIndex        =   65
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "INX"
         Height          =   375
         Index           =   8
         Left            =   600
         TabIndex        =   64
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "J%"
         Height          =   375
         Index           =   17
         Left            =   2760
         TabIndex        =   63
         TabStop         =   0   'False
         ToolTipText     =   "Click  J%  And  Then  Any Flag Conditions Ex: J% + PE= JPE"
         Top             =   3480
         Width           =   495
      End
      Begin VB.CommandButton CmdZero 
         Caption         =   "Z"
         Height          =   375
         Left            =   3360
         TabIndex        =   62
         TabStop         =   0   'False
         ToolTipText     =   "Click  J%  And  Then  Any Flag Conditions Ex: J% + PE= JPE"
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton CmdOparity 
         Caption         =   "PO"
         Height          =   375
         Left            =   2880
         TabIndex        =   61
         TabStop         =   0   'False
         ToolTipText     =   "Click  J%  And  Then  Any Flag Conditions Ex: J% + PE= JPE"
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton CmdParity 
         Caption         =   "P"
         Height          =   375
         Left            =   3360
         TabIndex        =   60
         TabStop         =   0   'False
         ToolTipText     =   "Click  J%  And  Then  Any Flag Conditions Ex: J% + PE= JPE"
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton CmdEParity 
         Caption         =   "PE"
         Height          =   375
         Left            =   2280
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "Click  J%  And  Then  Any Flag Conditions Ex: J% + PE= JPE"
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton CmdNozero 
         Caption         =   "NZ"
         Height          =   375
         Left            =   2280
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "Click  J%  And  Then  Any Flag Conditions Ex: J% + PE= JPE"
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton CmdNocarry 
         Caption         =   "NC"
         Height          =   375
         Left            =   3360
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "Click  J%  And  Then  Any Flag Conditions Ex: J% + PE= JPE"
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton CmdMinus 
         Caption         =   "M"
         Height          =   375
         Left            =   2880
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "Click  J%  And  Then  Any Flag Conditions Ex: J% + PE= JPE"
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton CmdCarry 
         Caption         =   "C"
         Height          =   375
         Left            =   2280
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Click  J%  And  Then  Any Flag Conditions Ex: J% + PE= JPE"
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "XRA"
         Height          =   375
         Index           =   15
         Left            =   1320
         TabIndex        =   54
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "DCX"
         Height          =   375
         Index           =   6
         Left            =   2040
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "MVI"
         Height          =   375
         Index           =   11
         Left            =   2760
         TabIndex        =   52
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "SUB"
         Height          =   375
         Index           =   14
         Left            =   600
         TabIndex        =   51
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "DAD"
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "MOV"
         Height          =   375
         Index           =   10
         Left            =   2040
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "SBB"
         Height          =   375
         Index           =   13
         Left            =   1320
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "ORA"
         Height          =   375
         Index           =   12
         Left            =   600
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "CMP"
         Height          =   375
         Index           =   3
         Left            =   2760
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "INR"
         Height          =   375
         Index           =   7
         Left            =   2760
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "DCR"
         Height          =   375
         Index           =   5
         Left            =   1320
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "ANA"
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "ADD"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "ADC"
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Click  Any Instruction For OPCode"
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton CmdA 
         Caption         =   "M"
         Height          =   375
         Index           =   7
         Left            =   4560
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Select Any Register"
         Top             =   3360
         Width           =   375
      End
      Begin VB.CommandButton CmdA 
         Caption         =   "L"
         Height          =   375
         Index           =   6
         Left            =   5040
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Select Any Register"
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton CmdA 
         Caption         =   "H"
         Height          =   375
         Index           =   5
         Left            =   4200
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Select Any Register"
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton CmdA 
         Caption         =   "E"
         Height          =   375
         Index           =   4
         Left            =   5040
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Select Any Register"
         Top             =   2400
         Width           =   375
      End
      Begin VB.CommandButton CmdA 
         Caption         =   "D"
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Select Any Register"
         Top             =   2400
         Width           =   375
      End
      Begin VB.CommandButton CmdA 
         Caption         =   "C"
         Height          =   375
         Index           =   2
         Left            =   5040
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Select Any Register"
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton CmdA 
         Caption         =   "B"
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Select Any Register"
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton CmdA 
         Caption         =   "A"
         Height          =   375
         Index           =   0
         Left            =   4680
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Select Any Register"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox TxtOPCode 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TxtNemo 
         Alignment       =   2  'Center
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
         Left            =   960
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "       Use  Me"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Line Line10 
         X1              =   120
         X2              =   1800
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line9 
         X1              =   5880
         X2              =   5880
         Y1              =   120
         Y2              =   4560
      End
      Begin VB.Line Line8 
         X1              =   120
         X2              =   5880
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line7 
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   4560
      End
      Begin VB.Line Line6 
         X1              =   3840
         X2              =   5880
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   5880
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line4 
         X1              =   2160
         X2              =   2160
         Y1              =   2760
         Y2              =   4560
      End
      Begin VB.Line Line3 
         X1              =   3840
         X2              =   3840
         Y1              =   1200
         Y2              =   4560
      End
      Begin VB.Line Line2 
         X1              =   2160
         X2              =   3840
         Y1              =   2760
         Y2              =   2760
      End
   End
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   3600
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.mcb"
   End
   Begin VB.TextBox DISPLAY 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   8040
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "SDA.85"
      Top             =   1320
      Width           =   7335
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   0
      Left            =   11160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   1
      Left            =   12720
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   2
      Left            =   14280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   3
      Left            =   15840
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   4
      Left            =   15840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   5
      Left            =   14280
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   6
      Left            =   12720
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   7
      Left            =   11160
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   8
      Left            =   11160
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   9
      Left            =   12720
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   10
      Left            =   14280
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   11
      Left            =   15840
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   12
      Left            =   11160
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   13
      Left            =   12720
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   14
      Left            =   14280
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   15
      Left            =   15840
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   16
      Left            =   6480
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   17
      Left            =   8040
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   18
      Left            =   9600
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   19
      Left            =   9600
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   20
      Left            =   8040
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   21
      Left            =   6480
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   22
      Left            =   6480
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   23
      Left            =   8040
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   24
      Left            =   9600
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   25
      Left            =   9600
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   26
      Left            =   8040
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin Project1.KeyCmdButton Key1 
      Height          =   615
      Index           =   27
      Left            =   6480
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
   End
   Begin VB.Line Line1 
      X1              =   8760
      X2              =   13200
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   7920
      Top             =   1200
      Width           =   7575
   End
   Begin ComctlLib.ImageList Img 
      Left            =   2880
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormKey2.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormKey2.frx":0984
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormKey2.frx":0EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormKey2.frx":1408
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormKey2.frx":194A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mfile 
      Caption         =   "&File"
      Begin VB.Menu mnew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu msave 
         Caption         =   "&Save"
         Shortcut        =   {F2}
      End
      Begin VB.Menu msaveas 
         Caption         =   "Save&as"
      End
      Begin VB.Menu ml1 
         Caption         =   "-"
      End
      Begin VB.Menu mprint 
         Caption         =   "&Print ..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mpview 
         Caption         =   "Print Pre&view"
      End
      Begin VB.Menu mpset 
         Caption         =   "Page Set&up.."
      End
      Begin VB.Menu L1 
         Caption         =   "-"
      End
      Begin VB.Menu nmuWind 
         Caption         =   "&Rescent Files"
         Begin VB.Menu nmufile 
            Caption         =   "File1"
            Enabled         =   0   'False
            Index           =   0
            Shortcut        =   +{F1}
         End
         Begin VB.Menu nmufile 
            Caption         =   "F2"
            Index           =   1
            Shortcut        =   +{F2}
            Visible         =   0   'False
         End
         Begin VB.Menu nmufile 
            Caption         =   "F3"
            Index           =   2
            Shortcut        =   +{F3}
            Visible         =   0   'False
         End
      End
      Begin VB.Menu NV 
         Caption         =   "-"
      End
      Begin VB.Menu mexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mdebug 
      Caption         =   "&Debug   "
      Begin VB.Menu mstep 
         Caption         =   "&Step"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mbreak 
         Caption         =   "&Break"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu mhelp 
      Caption         =   "&Help  "
      Begin VB.Menu mabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim sn As Integer
Dim so As Integer
Dim buttoncap(0 To 27) As String
Dim I As Integer




Private Sub CmdA_Click(Index As Integer)
Select Case Index
Case 0
     TxtNemo = Left(TxtNemo, sn) + "A "
Case 1
     TxtNemo = Left(TxtNemo, sn) + "B "
Case 2
     TxtNemo = Left(TxtNemo, sn) + "C "
Case 3
     TxtNemo = Left(TxtNemo, sn) + "D "
Case 4
     TxtNemo = Left(TxtNemo, sn) + "E "
Case 5
     TxtNemo = Left(TxtNemo, sn) + "H "
Case 6
     TxtNemo = Left(TxtNemo, sn) + "L "
Case 7
     TxtNemo = Left(TxtNemo, sn) + "M "
Case 8
     TxtNemo = Left(TxtNemo, sn) + "SP "
Case 9
     TxtNemo = Left(TxtNemo, sn) + "PSW "
End Select
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


Private Sub CmdA_GotFocus(Index As Integer)
Me.KeyPreview = False
'keystate = True
End Sub

Private Sub CmdCarry_Click()
TxtNemo = Left(TxtNemo, sn) + "C"
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

Private Sub CmdCarry_GotFocus()
Me.KeyPreview = False
'keystate = True
End Sub

Private Sub CmdEParity_Click()
TxtNemo = Left(TxtNemo, sn) + "PE"
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

Private Sub CmdEParity_GotFocus()
Me.KeyPreview = False
'keystate = True
End Sub

Private Sub CmdMinus_Click()
TxtNemo = Left(TxtNemo, sn) + "M"
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

Private Sub CmdMinus_GotFocus()
Me.KeyPreview = False
'keystate = True
End Sub

Private Sub CmdNocarry_Click()
TxtNemo = Left(TxtNemo, sn) + "NC"
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

Private Sub CmdNocarry_GotFocus()
Me.KeyPreview = False
'keystate = True
End Sub

Private Sub CmdNozero_Click()
TxtNemo = Left(TxtNemo, sn) + "NZ"
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

Private Sub CmdNozero_GotFocus()
Me.KeyPreview = False
'keystate = True
End Sub

Private Sub CmdOparity_Click()
TxtNemo = Left(TxtNemo, sn) + "PO"
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

Private Sub CmdOparity_GotFocus()
Me.KeyPreview = False
'keystate = True
End Sub

Private Sub CmdParity_Click()
TxtNemo = Left(TxtNemo, sn) + "P"
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

Private Sub CmdParity_GotFocus()
Me.KeyPreview = False
'keystate = True
End Sub

Private Sub CmdS_Click(Index As Integer)
Select Case Index
Case 0
     TxtNemo = "ADC "
Case 1
     TxtNemo = "ADD "
Case 2
     TxtNemo = "ANA "
Case 3
     TxtNemo = "CMP "
Case 4
     TxtNemo = "DAD "
Case 5
     TxtNemo = "DCR "
Case 6
     TxtNemo = "DCX "
Case 7
     TxtNemo = "INR "
Case 8
     TxtNemo = "INX "
Case 9
     TxtNemo = "LXI "
Case 10
     TxtNemo = "MOV "
Case 11
     TxtNemo = "MVI "
Case 12
     TxtNemo = "ORA "
Case 13
     TxtNemo = "SBB "
Case 14
     TxtNemo = "SUB "
Case 15
     TxtNemo = "XRA "
Case 16
     TxtNemo = "JMP"
Case 17
     TxtNemo = "J"
  End Select
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

Private Sub CmdS_GotFocus(Index As Integer)
Me.KeyPreview = False
'keystate = True
End Sub

Private Sub CmdZero_Click()

TxtNemo = Left(TxtNemo, sn) + "Z"
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

Private Sub CmdZero_GotFocus()
Me.KeyPreview = False
'keystate = True
End Sub

Private Sub DISPLAY_GotFocus()
Me.KeyPreview = True
'keystate = True
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
''keystate = True
If (KeyAscii = 27) Then Key1(27).KeyClick
If (KeyAscii = 13) Then Key1(21).KeyClick
If (KeyAscii = 47) Then Key1(16).KeyClick
If (KeyAscii = 46) Then Key1(20).KeyClick
If (KeyAscii = 9) Then Key1(19).KeyClick
If (KeyAscii = 32) Then Key1(23).KeyClick
If (KeyAscii = 8) Then Key1(25).KeyClick
If (KeyAscii = 48) Then Key1(0).KeyClick
If (KeyAscii = 49) Then Key1(1).KeyClick
If (KeyAscii = 50) Then Key1(2).KeyClick
If (KeyAscii = 51) Then Key1(3).KeyClick
If (KeyAscii = 52) Then Key1(4).KeyClick
If (KeyAscii = 53) Then Key1(5).KeyClick
If (KeyAscii = 54) Then Key1(6).KeyClick
If (KeyAscii = 55) Then Key1(7).KeyClick
If (KeyAscii = 56) Then Key1(8).KeyClick
If (KeyAscii = 57) Then Key1(9).KeyClick
If (KeyAscii = 65) Then Key1(10).KeyClick
If (KeyAscii = 66) Then Key1(11).KeyClick
If (KeyAscii = 67) Then Key1(12).KeyClick
If (KeyAscii = 68) Then Key1(13).KeyClick
If (KeyAscii = 69) Then Key1(14).KeyClick
If (KeyAscii = 70) Then Key1(15).KeyClick

End Sub

Private Sub Form_Load()

Me.WindowState = 2
frmcap = Me.Caption + "  ( "
InitialiseDataTable
Initialise
Randamise
If (file1 <> "") Then
nmufile(0).Caption = file1
nmufile(0).Enabled = True
nmufile(0).Visible = True
wincount = 1
End If
If (file2 <> "") Then
nmufile(1).Caption = file2
nmufile(1).Enabled = True
nmufile(1).Visible = True
wincount = 2
End If
If (file1 <> "") Then
nmufile(2).Caption = file3
nmufile(2).Enabled = True
nmufile(2).Visible = True
wincount = 3
End If
strDot = "."
savedfile = True
cdl1.CancelError = True
cdl1.Initdir = Initdir
ResetON = True
cdl1.Filter = "Microprocessor Machine Code (*.MPC)|*.MPC"
buttoncap(0) = "0"
buttoncap(1) = "1"
buttoncap(2) = "2"
buttoncap(3) = "3"
buttoncap(4) = "        4" + vbCrLf + "             SPH"
buttoncap(5) = "        5" + vbCrLf + "             SPL"
buttoncap(6) = "        6" + vbCrLf + "            PCH"
buttoncap(7) = "        7" + vbCrLf + "            PCL"
buttoncap(8) = "       8" + vbCrLf + "               H"
buttoncap(9) = "       9" + vbCrLf + "                L"
buttoncap(10) = "A"
buttoncap(11) = "B"
buttoncap(12) = "        C" + vbCrLf + "           CMPL"
buttoncap(13) = "        D" + vbCrLf + "              AD"
buttoncap(14) = "        E" + vbCrLf + "             CRT"
buttoncap(15) = "        F" + vbCrLf + "             FILL"
buttoncap(16) = "EXEC"
buttoncap(17) = "INS"
buttoncap(18) = "DEL"
buttoncap(19) = "    SUB" + vbCrLf + "          MEM"
buttoncap(20) = "GO"
buttoncap(21) = "NEXT"
buttoncap(22) = "  INTR" + vbCrLf + "         VECT"
buttoncap(23) = " EXAM" + vbCrLf + "          REG"
buttoncap(24) = "   SING" + vbCrLf + "           STEP "
buttoncap(25) = "PREV"
buttoncap(26) = "  BLOCK" + vbCrLf + "         MOVE"
buttoncap(27) = "RESET"

For I = 0 To 27
Key1(I).ByVel = 5
Key1(I).Font = Form2.Font
Key1(I).Caption = buttoncap(I)
Next I
SetWindowPos Me.hwnd, conHwndNoTopmost, 0, 0, 12000, 9000, conSwpShowWindow
Form2.Show
Unload frmstart
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.KeyPreview = True
'keystate = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Temps As String
Dim cb As Integer
Dim ret As Integer
Dim Inis1 As String
Dim Inis2 As String
Cancel = 0
If (savedfile = False) Then
ret = MsgBox("Do you want to save changes to " + pfilename, vbYesNoCancel + vbDefaultButton1 + vbQuestion, "SAVE")

If (ret = 2) Then
Cancel = 1
Exit Sub
End If
If (ret = 6) Then
On Error GoTo Errbuton
cdl1.ShowSave
On Error GoTo Errbuton
fileno = FreeFile
Open cdl1.FileName For Binary As #fileno
For cb = 0 To 8191
Temps = block(cb).Data
Put #fileno, , Temps
Next cb
Close #fileno
Errbuton:
End If
End If
regs = A + B + c + D + E + H + l + M
If (file1 <> "") Then WritePrivateProfileString "DATA", "FILE#1", file1, Inifile
If (file2 <> "") Then WritePrivateProfileString "DATA", "FILE#2", file2, Inifile
If (file3 <> "") Then WritePrivateProfileString "DATA", "FILE#3", file3, Inifile
If (path1 <> "") Then WritePrivateProfileString "DATA", "PATH#1", path1, Inifile
If (path2 <> "") Then WritePrivateProfileString "DATA", "PATH#2", path2, Inifile
If (path2 <> "") Then WritePrivateProfileString "DATA", "PATH#3", path3, Inifile
WritePrivateProfileString "DATA", "REGDATA", regs, Inifile
WritePrivateProfileString "DATA", "STACKPOINTER", SP, Inifile
End

End Sub



Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.KeyPreview = False
'keystate = True

End Sub

Private Sub Key1_Click(Index As Integer)
Dim Ti1 As Integer
Dim s As String
Dim Result As Integer
Dim Bret As Integer
Dim Dret As Integer
Select Case Index
     Case 27  'RESET
          ErrPgm = 0
          ResetON = True
          AdressON = False
          BMoveON = False
          AdressOK = False
          DataOK = False
          GoON = False
          ERRORU = False
          ExaRegON = False
          ExaRegOK = False
          Displayed = False
          strAdress = " SDA"
          strdata = "85"
          BAddCount = 0
          Bret = 0
          DISPLAY = strAdress + strDot + strdata
    Case 19 'SUBST MEM
    If (ERRORU = True) Then Exit Sub
          If (ResetON = True) Then
          ResetON = False
          AdressON = True
          strAdress = "    "
          strdata = "  "
          DISPLAY = strAdress + strDot + strdata
          Else
          DisplayError
          End If
   Case 20  'GO
   If (ERRORU = True) Then Exit Sub
          If (ResetON = True) Then
          ResetON = False
          AdressON = True
          GoON = True
          strAdress = "    "
          strdata = "  "
          DISPLAY = strAdress + strDot + strdata
          Else
          DisplayError
          End If
          
    Case 26 'BLOCKMOVE
               If (ResetON = True) Then
               BMoveON = True
               ResetON = False
               AdressON = True
               strAdress = "    "
               strdata = "  "
               DISPLAY = strAdress + strDot + strdata
               Else
               DisplayError
               End If
                              
    Case 21 'NEXT
         savedfile = False
         If (ERRORU = True) Then Exit Sub
         
         If (BMoveON = True) Then
               If (AdressOK = True) Then
               BAddCount = BAddCount + 1
               BAdd(BAddCount) = strAdress
               If (BAddCount = 3) Then
               BAddCount = 0
               Bret = BlockMove(BAdd(1), BAdd(2), BAdd(3))
               If (Bret <> 1) Then
               ResetON = True
                strAdress = " SDA"
                strdata = "85"
                DISPLAY = strAdress + strDot + strdata
                BMoveON = False
                Else
                DisplayError
                End If
               Exit Sub
               End If
             
               
               strAdress = "    "
               strdata = "  "
               DISPLAY = strAdress + strDot + strdata
               AdressOK = False
               DataOK = False
               Exit Sub
               Else
               DisplayError
               End If
           End If
           
           
           If (DelON = True) Then
               If (AdressOK = True) Then
                    Select Case delcount
                     Case 0: DelAdd(1) = strAdress
                     Case 1: DelAdd(2) = strAdress
                     
                         Dret = Delete(DelAdd(1), DelAdd(2))
                         If (Dret <> 1) Then
                         ResetON = True
                         strAdress = " SDA"
                         strdata = "85"
                         DISPLAY = strAdress + strDot + strdata
                         DelON = False
                         Else
                         DisplayError
                         Exit Sub
                         End If
                         Exit Sub
                     End Select
                    delcount = delcount + 1
                    strAdress = "    "
                    strdata = "  "
                    DISPLAY = strAdress + strDot + strdata
                    AdressON = True
                    AdressOK = False
                    Exit Sub
               Else
               DisplayError
               Exit Sub
               End If
            End If
           
                    If (ExaRegON = True) Then
           If (ExaRegOK = True) Then
             SetRegData Regcount, strdata
             Regcount = Regcount + 1
             If (Regcount > 15) Then Regcount = 4
             strAdress = GetRegCaption(Regcount)
             strdata = GetRegData(Regcount)
             DISPLAY = strAdress + strDot + strdata
            Else
         DisplayError
          End If
         Else
          If ((GoON = True) Or (AdressOK = False)) Then
          ERRORU = True
          strAdress = " Err"
          strdata = "  "
          DISPLAY = strAdress + strDot + strdata
          End If
          If (AdressOK = True And DataOK = False) Then
          Ti1 = InBetween(strAdress)
          If (Ti1 = 1) Then
          DisplayError
          Exit Sub
          End If
          strdata = GetData(strAdress)
          DISPLAY = strAdress + strDot + strdata
          DataOK = True
          AdressON = False
          Else
          If (DataOK = True) Then
             SetData strAdress, strdata
             If (strAdress = "DFFF") Then
             strAdress = "C000"
             Else
             strAdress = AddHexLong(strAdress, , 1)
             End If
             strdata = GetData(strAdress)
             DISPLAY = strAdress + strDot + strdata
          End If
          End If
          End If
    Case 23 'EXAM REGESTER
          If (ResetON = True) Then
          ResetON = False
          ExaRegON = True
          strAdress = "    "
          strdata = "  "
          DISPLAY = strAdress + strDot + strdata
          Else
         DisplayError

          End If
       
    Case 25   'PREVES
       If (ERRORU = True) Then Exit Sub
       If (ExaRegON = True) Then
          DisplayError
           Exit Sub
          End If
       
          If (DataOK = True) Then
          Ti1 = InBetween(strAdress)
          If (Ti1 = 1) Then
          DisplayError
          Exit Sub
          End If
          If (strAdress = "C000") Then
          strAdress = "DFFF"
          Else
          strAdress = SubHexLong(strAdress, , 1)
          End If
          strdata = GetData(strAdress)
          DISPLAY = strAdress + strDot + strdata
          Else
          DisplayError
          End If
    Case 0 To 3 'NUM PAD
      If (ERRORU = True) Then Exit Sub
      If (ExaRegON = True And ExaRegOK = False) Then
          DisplayError
          End If
          
          If (AdressON = True) Then
          TempAdress = strAdress
          s = GetKeyValue(Index)
          TempAdress = TempAdress + s
          strAdress = Right(TempAdress, 4)
          DISPLAY = strAdress + strDot + strdata
          If (Len(Trim(strAdress)) = 4) Then AdressOK = True
          End If
          
          If (DataOK = True) Then
          TempData = strdata
          s = GetKeyValue(Index)
          TempData = TempData + s
          strdata = Right(TempData, 2)
          DISPLAY = strAdress + strDot + strdata
          End If
          
     Case 4 To 15  'NUM PAD
      If (ERRORU = True) Then Exit Sub
      If (ExaRegON = True And ExaRegOK = False) Then
          Regcount = Index
          ExaRegOK = True
          DataOK = True
          strAdress = GetRegCaption(Regcount)
          strdata = GetRegData(Regcount)
          DISPLAY = strAdress + strDot + strdata
       Else
          If (AdressON = True) Then
          TempAdress = strAdress
          s = GetKeyValue(Index)
          TempAdress = TempAdress + s
          strAdress = Right(TempAdress, 4)
          DISPLAY = strAdress + strDot + strdata
          If (Len(Trim(strAdress)) = 4) Then AdressOK = True
          End If
          If (DataOK = True) Then
          TempData = strdata
          s = GetKeyValue(Index)
          TempData = TempData + s
          strdata = Right(TempData, 2)
          DISPLAY = strAdress + strDot + strdata
          End If
       End If
       Case 16 'EXEC
          If (GoON = True) Then
              If (AdressOK = True) Then
              Ti1 = InBetween(strAdress)
               If (Ti1 = 1) Then
               DisplayError
               Exit Sub
               End If
              PCounter = strAdress
              step1 = 1
              break = 0
              step2 = 0
              Result = ExecPgm()
              step1 = 0
              If (Result = 1) Then
              If (Displayed = False) Then
              strAdress = "E   "
              strdata = "  "
              DISPLAY = strAdress + strDot + strdata
              End If
              Else
              strAdress = "Err "
              strdata = "  "
              DISPLAY = strAdress + strDot + strdata
              End If
              AdressON = False
              DataOK = False
              Else
              DisplayError
            End If
          Else
          DisplayError
          End If
          Case 18: 'DELETE
                    
          If (ResetON = True) Then
          delcount = 0
          DelAdd(1) = ""
          DelAdd(2) = ""
          ResetON = False
          DelON = True
          strAdress = "    "
          strdata = "  "
          DISPLAY = strAdress + strDot + strdata
          AdressON = True
          Else
          DisplayError
          End If
          
          
        End Select
End Sub

Private Sub Key1_GotFocus(Index As Integer)
Me.KeyPreview = True
'keystate = True
End Sub

Private Sub mabout_Click()
frmAbout.Show
End Sub

Private Sub mbreak_Click()
Unload FrmBug
End Sub

Private Sub mexit_Click()
Unload Me
End Sub




Private Sub mnew_Click()
fname = ""
pfilename = "TEMP.MPU"
Form2.Caption = frmcap + pfilename + "  )"
End Sub

Private Sub mopen_Click()
Dim CD As Integer
Dim Temps As String

Temps = String(2, " ")
cdl1.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist + cdlOFNHideReadOnly
On Error GoTo erropen
cdl1.ShowOpen
pfilename = cdl1.FileTitle
wincount = wincount + 1
Select Case wincount
Case 1
file1 = cdl1.FileTitle
path1 = cdl1.FileName
Case 2
file2 = file1
path2 = path1
file1 = cdl1.FileTitle
path1 = cdl1.FileName
Case Else
     file3 = file2
     path3 = path2
     file2 = file1
     path2 = path1
     file1 = cdl1.FileTitle
     path1 = cdl1.FileName
End Select

If (file1 <> "") Then
nmufile(0).Caption = file1
nmufile(0).Enabled = True
nmufile(0).Visible = True
End If
If (file2 <> "") Then
nmufile(1).Caption = file2
nmufile(1).Enabled = True
nmufile(1).Visible = True
End If
If (file1 <> "") Then
nmufile(2).Caption = file3
nmufile(2).Enabled = True
nmufile(2).Visible = True
End If

fileno = FreeFile
Open cdl1.FileName For Binary As #fileno
For CD = 0 To 8191
Get #fileno, , Temps
block(CD).Data = Temps
Next CD
Form2.Caption = frmcap + pfilename + "  )"
erropen:
End Sub

Private Sub mprint_Click()
On Error GoTo noprinter
Dim I As Single
Dim c As Long
Dim lx As Integer
Dim ly As Integer
lx = 50
ly = 50
NextPage = 0
getpstr
preview
Printer.CurrentY = ymarg + ly
For I = 1 To UBound(strprint)
Printer.CurrentX = xmarg + lx
Printer.CurrentY = ymarg + (I - 1) * form1.Pic.TextHeight("SAMPLE") + ly
Printer.Print strprint(I)
Next I
Printer.EndDoc
noprinter:

End Sub

Private Sub mpset_Click()
On Error GoTo end1
form1.cdl1.Flags = cdlPDHidePrintToFile + cdlPDNoSelection
form1.cdl1.ShowPrinter
'pcopies = cdl1.Copies
end1:
End Sub

Private Sub mpview_Click()
If (pstartadd = "") Then
pstartadd = InputBox("Enter the Starting Adress", "Starting Adress", "C000")
End If
form1.Show
End Sub

Private Sub mrun_Click()
FrmBug.Show
End Sub

Private Sub msave_Click()
fileno = FreeFile
Dim cb As Integer
Dim Temps As String
If (fname = "") Then
cdl1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
On Error GoTo errop
cdl1.ShowSave
pfilename = cdl1.FileTitle
wincount = wincount + 1
Select Case wincount
Case 1
file1 = cdl1.FileTitle
path1 = cdl1.FileName
Case 2
file2 = file1
path2 = path1
file1 = cdl1.FileTitle
path1 = cdl1.FileName
Case Else
     file3 = file2
     path3 = path2
     file2 = file1
     path2 = path1
     file1 = cdl1.FileTitle
     path1 = cdl1.FileName
End Select

If (file1 <> "") Then
nmufile(0).Caption = file1
nmufile(0).Enabled = True
nmufile(0).Visible = True
End If
If (file2 <> "") Then
nmufile(1).Caption = file2
nmufile(1).Enabled = True
nmufile(1).Visible = True
End If
If (file1 <> "") Then
nmufile(2).Caption = file3
nmufile(2).Enabled = True
nmufile(2).Visible = True
End If
End If
Open fname For Binary As #fileno
savedfile = True
For cb = 0 To 8191
Temps = block(cb).Data
Put #fileno, , Temps
Next cb
Form2.Caption = frmcap + pfilename + "  )"
errop:
End Sub

Private Sub msaveas_Click()
Dim cb As Integer
Dim Temps As String * 2
cdl1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNHideReadOnly
On Error GoTo Errbuton
cdl1.ShowSave
pfilename = cdl1.FileTitle
wincount = wincount + 1
Select Case wincount
Case 1
file1 = cdl1.FileTitle
path1 = cdl1.FileName
Case 2
file2 = file1
path2 = path1
file1 = cdl1.FileTitle
path1 = cdl1.FileName
Case Else
     file3 = file2
     path3 = path2
     file2 = file1
     path2 = path1
     file1 = cdl1.FileTitle
     path1 = cdl1.FileName
End Select
If (file1 <> "") Then
nmufile(0).Caption = file1
nmufile(0).Enabled = True
nmufile(0).Visible = True
End If
If (file2 <> "") Then
nmufile(1).Caption = file2
nmufile(1).Enabled = True
nmufile(1).Visible = True
End If
If (file3 <> "") Then
nmufile(2).Caption = file3
nmufile(2).Enabled = True
nmufile(2).Visible = True
End If

fileno = FreeFile
Open cdl1.FileName For Binary As #fileno
savedfile = True
For cb = 0 To 8191
Temps = block(cb).Data
Put #fileno, , Temps
Next cb
Form2.Caption = frmcap + pfilename + "  )"
Errbuton:
End Sub

Private Sub mstep_Click()
FrmBug.Show
End Sub

Private Sub nmufile_Click(Index As Integer)

Dim CD As Integer
Dim Temps As String
Dim Tmpf As String
Dim Tmpp As String
Temps = String(2, " ")
Select Case Index
Case 0
     fname = path1
     Form2.Caption = frmcap + file1 + "  )"
     pfilename = file1
Case 1
     fname = path2
     Form2.Caption = frmcap + file2 + "  )"
     pfilename = file2
     Tmpf = file1
     Tmpp = path1
     file1 = file2
     path1 = path2
     file2 = Tmpf
     path2 = Tmpp
     
Case 2
     fname = path3
     Form2.Caption = frmcap + file3 + "  )"
     pfilename = file3
     Tmpf = file3
     Tmpp = path3
     file3 = file2
     path3 = path2
     file2 = file1
     path2 = path1
     file1 = Tmpf
     path1 = Tmpp
     End Select
     If (file1 <> "") Then
nmufile(0).Caption = file1
nmufile(0).Enabled = True
nmufile(0).Visible = True
End If
If (file2 <> "") Then
nmufile(1).Caption = file2
nmufile(1).Enabled = True
nmufile(1).Visible = True
End If
If (file1 <> "") Then
nmufile(2).Caption = file3
nmufile(2).Enabled = True
nmufile(2).Visible = True
End If

     Close (fileno)
   fileno = FreeFile
   On Error GoTo ED1
Open fname For Binary As #fileno
For CD = 0 To 8191
Get #fileno, , Temps
block(CD).Data = Temps
Next CD
  
ED1:
End Sub

Private Sub T1_Timer()
YES = "NO"
T1.Enabled = False
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key

   Case "topen"
          mopen_Click
   Case "tsave"
          msave_Click
   Case "tstep"
          mstep_Click
   Case "tprint"
     mprint_Click
   Case "tpreview"
     mpview_Click
      End Select
End Sub
Function Getnemo(STR As String) As String
Dim S1 As String
Dim l As Integer
Dim j As Integer
Dim k As Integer
Dim I As Integer
TxtOPCode = " "
l = Len(STR)
k = 0
For I = 1 To 247
     S1 = STR
     For j = 1 To l
          If (Mid(S1, j, 1) = Mid(InSet(I).Nemo, j, 1)) Then
          If (j > k) Then
          k = j
          Getnemo = InSet(I).Nemo
          TxtOPCode = InSet(I).OpCode
          End If
          Else
          Exit For
          End If
       Next j
          
  Next I
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


Private Sub TxtNemo_GotFocus()
Me.KeyPreview = False
'keystate = True
TxtNemo.SelStart = sn
TxtNemo.SelLength = Len(TxtNemo)
End Sub



Private Sub TxtNemo_KeyUp(KeyCode As Integer, Shift As Integer)
'keystate = True
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
Function GetOpcode(STR As String) As String
Dim S1 As String
Dim l As Integer
Dim j As Integer
Dim k As Integer
Dim I As Integer
TxtNemo = " "
l = Len(STR)
k = 0
For I = 1 To 247
     S1 = STR
     For j = 1 To l
          If (Mid(S1, j, 1) = Mid(InSet(I).OpCode, j, 1)) Then
          If (j > k) Then
          k = j
          GetOpcode = InSet(I).OpCode
          TxtNemo = InSet(I).Nemo
          End If
          Else
          Exit For
          End If
       Next j
          
  Next I
End Function

Private Sub TxtOPCode_GotFocus()
Me.KeyPreview = False
'keystate = True
TxtOPCode.SelStart = so
TxtOPCode.SelLength = Len(TxtOPCode)
End Sub

Private Sub TxtOPcode_KeyUp(KeyCode As Integer, Shift As Integer)
'keystate = True
If (KeyCode = 8) Then Exit Sub
so = Len(TxtOPCode)
TxtOPCode = GetOpcode(TxtOPCode.Text)
If (so < 0) Then so = 0
TxtOPCode.SelStart = so
If (Len(TxtOPCode.Text) - so > 0) Then
TxtOPCode.SelLength = Len(TxtOPCode.Text) - so
Else
TxtOPCode.SelLength = 0
End If

End Sub

