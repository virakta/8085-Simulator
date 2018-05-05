VERSION 5.00
Begin VB.UserControl KeyCmdButton 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3660
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2865
   ScaleWidth      =   3660
   ToolboxBitmap   =   "KeyCmdButton1.ctx":0000
End
Attribute VB_Name = "KeyCmdButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim mBackColor As OLE_COLOR
Dim mForeColor As OLE_COLOR

Dim mByVel As Integer
Dim bMouseDn As Boolean
'Default Property Values:
Const m_def_Caption = 0
Const m_def_BackColor = vbCyan
Const m_def_ForeColor = vbBlack
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_ByVel = 3
'Property Variables:
Dim m_Caption As Variant
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_Enabled As Boolean
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_ByVel As Integer
'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbcyan
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    mBackColor = m_BackColor
    PropertyChanged "BackColor"
    UserControl_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbblack
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    mForeColor = m_ForeColor
    PropertyChanged "ForeColor"
    UserControl_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    UserControl_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
    UserControl_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    UserControl_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,3
Public Property Get ByVel() As Integer
    ByVel = m_ByVel
End Property

Public Property Let ByVel(ByVal New_ByVel As Integer)
    m_ByVel = New_ByVel
    mByVel = m_ByVel
    PropertyChanged "ByVel"
    UserControl_Paint
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=13,0,0,Key Click

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub KeyClick()
 Dim i As Long
    Dim j As Integer
     bMouseDn = True
    UserControl_Paint
    For i = 0 To 100000
    j = j * 0 + j * 3 * 0 + j * (j * 6 * 0) + j ^ 3 * 0
    Next i
    bMouseDn = False
    UserControl_Paint
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
mBackColor = vbBlue
mForeColor = vbWhite
mByVel = 3
UserControl.BackColor = mBackColor
UserControl.ForeColor = mForeColor
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_ByVel = m_def_ByVel
    m_Caption = m_def_Caption
     Set UserControl.Font = Ambient.Font
     m_Caption = m_def_Caption
End Sub

Private Sub UserControl_Paint()
Dim inHeight As Integer
Dim inwidth As Integer
Dim i As Integer
Dim mb As Integer
Dim dw As Integer
dw = 1
With UserControl
mb = 25 * mByVel
inHeight = .Height - (dw * mb)
inwidth = .Width - (dw * mb)
.DrawWidth = 1
.FillStyle = 0
.FillColor = mBackColor
UserControl.Line (0, 0)-(inwidth, inHeight), , B
.DrawWidth = dw
If bMouseDn = False Then
.ForeColor = vbBlack
Else
.ForeColor = vbWhite
End If
For i = 1 To (mb * dw) Step dw
UserControl.Line ((dw * mb) - (i - 1), inHeight + i)-(inwidth + (mb * dw), inHeight + i)
UserControl.Line (inwidth + i, (dw * mb) - (i - 1))-(inwidth + i, inHeight + i)
Next i
If bMouseDn = False Then
.ForeColor = vbWhite
Else
.ForeColor = vbBlack
End If
For i = 1 To (mb * dw) Step dw
UserControl.Line (i, i)-(inwidth + (dw * mb) - (i - 1), i)
UserControl.Line (i, i)-(i, inHeight + (dw * mb) - (i - 1))
Next i
.ForeColor = mForeColor
.CurrentX = (.Width - .TextWidth(m_Caption1)) / 2
If .CurrentX < 5 Then .CurrentX = 5
.CurrentY = (.Height - .TextHeight(m_Caption1)) / 2
If .CurrentY < (mb * dw) Then .CurrentY = (mb * dw)
UserControl.Print m_Caption1
End With
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
bMouseDn = True
UserControl_Paint
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bMouseDn = False
UserControl_Paint
RaiseEvent Click
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_ByVel = PropBag.ReadProperty("ByVel", m_def_ByVel)

     Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
     m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
'    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("ByVel", m_ByVel, m_def_ByVel)
     Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
'     Call PropBag.WriteProperty("Caption2", m_Caption2, m_def_Caption2)
'     Call PropBag.WriteProperty("Font1", Text1.Font, Ambient.Font)
     Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
     Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
     Set UserControl.Font = New_Font
     PropertyChanged "Font"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=13,0,0,key
'Public Property Get Caption2() As String
'     Caption2 = m_Caption2
'End Property
'
'Public Property Let Caption2(ByVal New_Caption2 As String)
'     m_Caption2 = New_Caption2
'     PropertyChanged "Caption2"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Text1,Text1,-1,Font
'Public Property Get Font1() As Font
'     Set Font1 = Text1.Font
'End Property
'
'Public Property Set Font1(ByVal New_Font1 As Font)
'     Set Text1.Font = New_Font1
'     PropertyChanged "Font1"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Caption() As Variant
     Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As Variant)
     m_Caption = New_Caption
     PropertyChanged "Caption"
     UserControl_Paint
End Property

