Attribute VB_Name = "OnTop"
Option Explicit

Public Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const conHwndTopmost = -1
Public Const conHwndNoTopmost = -2
Public Const conSwpNoActivate = &H10
Public Const conSwpShowWindow = &H40
Public Const WS_EX_TRANSPARENT = &H20&
Public Const GWL_EXSTYLE = (-20)

