Attribute VB_Name = "Module1"
Option Explicit

Public Type SHFILEOPSTRUCT
hwnd As Long
wFunc As Long
pFrom As String
pTo As String
fFlags As Integer
fAnyOperationsAborted As Long
hNameMappings As Long
lpszProgressTitle As Long
End Type

Public Declare Function SHFileOperation Lib "shell32.dll" Alias _
"SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Public Const FO_DELETE = &H3
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_SILENT = &H4
Public Type POINTAPI
x As Long
y As Long
End Type
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

