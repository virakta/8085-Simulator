Attribute VB_Name = "MainModule"
Option Explicit

Public startfile As String
Public Initdir As String
Public file1 As String
Public file2 As String
Public file3 As String
Public path1 As String
Public path2 As String
Public path3 As String
Public Inifile As String
Public regs As String
Public IniSP As String
Public wincount As Integer
Public LastFile As String

Public Sub Main()
frmstart.Show
End Sub

Sub SubMain()
Dim Inis1 As String
Dim Inis2 As String
Dim test As Long
Dim startup As Boolean
DoEvents
startfile = ""
wincount = 0
Initdir = CurDir
Inis1 = Space(20)
Inis2 = Space(200)
Inifile = "C:\WINDOWS\MP8085.INI"
GetProfileString "Software\Shivakumar\Microprocessor Programing 8085", "Startup", "YES", Inis1, Len(Inis1)
Inis1 = Trim(Inis1)
If (Left(Inis1, 1) = "Y") Then
startup = True
WriteProfileString "Software\Shivakumar\Microprocessor Programing 8085", "Startup", "NO"
End If
Inis2 = Space(20)
GetPrivateProfileString "STARTUP", "Starting", "YES", Inis2, Len(Inis2), Inifile
Inis2 = Trim(Inis2)
If (Left(Inis2, 1) = "Y") Then
WritePrivateProfileString "STARTUP", "Starting", "NO", Inifile
On Error GoTo errorpath
MkDir Initdir + "\MicroPgms"
Initdir = Initdir + "\MicroPgms"
ChDir Initdir
errorpath:
WritePrivateProfileString "STARTUP", "Initdir", Initdir, Inifile
Else
DoEvents
Inis2 = Space(500)
GetPrivateProfileString "STARTUP", "Initdir", Initdir, Inis2, Len(Inis2), Inifile
Initdir = Trim(Inis2)
Initdir = Left(Initdir, (Len(Initdir) - 1)) + "\MicoPgms"
On Error GoTo errdir
ChDir Initdir
errdir:
Inis2 = Space(20)
GetPrivateProfileString "DATA", "FILE#1", "", Inis2, Len(Inis2), Inifile
Inis2 = Trim(Inis2)
If (Len(Inis2) > 1) Then
file1 = Inis2
file1 = Left(file1, (Len(file1) - 1))
End If
DoEvents
Inis2 = Space(20)
GetPrivateProfileString "DATA", "FILE#2", "", Inis2, Len(Inis2), Inifile
Inis2 = Trim(Inis2)
If (Len(Inis2) > 1) Then
file2 = Inis2
file2 = Left(file2, (Len(file2) - 1))
End If
Inis2 = Space(20)
GetPrivateProfileString "DATA", "FILE#3", "", Inis2, Len(Inis2), Inifile
Inis2 = Trim(Inis2)
If (Len(Inis2) > 1) Then
file3 = Inis2
file3 = Left(file3, (Len(file3) - 1))
End If
Inis2 = Space(200)
GetPrivateProfileString "DATA", "PATH#1", "", Inis2, Len(Inis2), Inifile
Inis2 = Trim(Inis2)
If (Len(Inis2) > 1) Then
path1 = Inis2
startfile = path1
End If
Inis2 = Space(200)
DoEvents
GetPrivateProfileString "DATA", "PATH#2", "", Inis2, Len(Inis2), Inifile
Inis2 = Trim(Inis2)
If (Len(Inis2) > 1) Then path2 = Inis2

Inis2 = Space(200)
GetPrivateProfileString "DATA", "PATH#3", "", Inis2, Len(Inis2), Inifile
Inis2 = Trim(Inis2)
If (Len(Inis2) > 1) Then path3 = Inis2

Inis2 = Space(200)
GetPrivateProfileString "DATA", "REGDATA", "", Inis2, Len(Inis2), Inifile
Inis2 = Trim(Inis2)
If (Len(Inis2) > 1) Then regs = Inis2

Inis2 = Space(200)
GetPrivateProfileString "DATA", "STACKPOINTER", "", Inis2, Len(Inis2), Inifile
Inis2 = Trim(Inis2)
If (Len(Inis2) > 1) Then IniSP = Inis2
End If
DoEvents
Load Form2
 End Sub

