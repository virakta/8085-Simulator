Attribute VB_Name = "Data"
Option Explicit

Type reg
 Adress As String * 4
 Data As String * 2
End Type

Type Instraction
   Nemo As String
   OpCode As String
   ByteCount As Integer
End Type




Public savedfile As Boolean
Public LastFile As String

Public fileno As Integer
Public pfilename As String
Public frmcap As String
Public fname As String

Public InSet(1 To 247)   As Instraction




Public block(0 To 8191) As reg
Public StartAdress  As String * 4
Public PCounter As String * 4
Public SP As String * 4
Public PSW As String * 4

Public Sign As Integer
Public Parity As Integer
Public AC As Integer
Public Zero As Integer
Public cy As Integer
Public FLAGREG As String * 4


Public A As String * 2
Public B As String * 2
Public c As String * 2
Public D As String * 2
Public E As String * 2
Public F As String * 2
Public H As String * 2
Public l As String * 2
Public M As String * 2

Public displayadd As String
Public displaydata As String
Public delaytime As Long
Public Sub1Add As String
Public Sub2Add As String
Public YES As String
Public Displayed As Boolean
