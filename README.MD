# windows-shortcut

用于在 windows 上读取和创建快捷方式(以.lnk 结尾)，关键的参数有 7 个。

`lnkPath`：表示快捷方式的位置，比如`D:/test.lnk`。

`TargetPath`：是目标位置，可以是一个程序，也可以是一个命令行。

`WindowStyle`：表示窗口样式，1 是默认，3 最大化，7 最小化。

`Hotkey`：是快捷方式，比如`'Ctrl+Shift+K'`。

`IconLocation`：是快捷方式图标，比如默认`',0'`或者`'path/to/icon.ico,0'`，或者`'path/to/app.exe,0'`。

`Description`：快捷方式的描述。

`WorkingDirectory`：工作目录。

## 环境需求

里面用到了`win32com`模块，需安装`pypiwin32`

```
pip install pypiwin32
```

## 用例

[example.py](example.py)

```python
import shortcut
shortcut.CreateShortCut("readme.lnk", TargetPath='README.MD')
print(shortcut.ReadShortCut("readme.lnk"))
```

<details>
<summary>python源代码</summary>

## python

[shortcut.py](shortcut.py)

```python
from win32com import client
import os

shell = client.Dispatch("WScript.Shell")

def ReadShortCut(lnkPath):
shortcut = shell.CreateShortCut(lnkPath)
result = {
"TargetPath": shortcut.TargetPath,
"WindowStyle": shortcut.WindowStyle,
"Hotkey": shortcut.Hotkey,
"IconLocation": shortcut.IconLocation,
"Description": shortcut.Description,
"WorkingDirectory": shortcut.WorkingDirectory
}
return result

def CreateShortCut(lnkPath, TargetPath, WindowStyle=1, Hotkey='', IconLocation=',0', Description='', WorkingDirectory=''):
lnkPath = os.path.abspath(lnkPath)
TargetPath = os.path.abspath(TargetPath)
WorkingDirectory = os.path.abspath(WorkingDirectory)

    shortcut = shell.CreateShortCut(lnkPath)
    shortcut.TargetPath = TargetPath
    shortcut.WindowStyle = WindowStyle
    shortcut.Hotkey = Hotkey
    shortcut.IconLocation = IconLocation
    shortcut.Description = Description
    shortcut.WorkingDirectory = WorkingDirectory
    shortcut.save()

```

</details>

<details>
<summary>可以参考VBS的代码</summary>

## vbs

[参考这里](https://www.cnblogs.com/sky20080101/articles/8309619.html)

```bash
Set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop") :'特殊文件夹“桌面”
'在桌面创建一个记事本快捷方式
set oShellLink = WshShell.CreateShortcut(strDesktop & "\记事本.lnk")
oShellLink.TargetPath = "C:\Windows\System32\notepad.exe" : '目标
oShellLink.WindowStyle = 3 :'参数 1 默认窗口激活，参数 3 最大化激活，参数 7 最小化
oShellLink.Hotkey = "Ctrl+Alt+e" : '快捷键
oShellLink.IconLocation = "C:\Windows\System32\notepad.exe, 0" : '图标
oShellLink.Description = "记事本快捷方式" : '备注
oShellLink.WorkingDirectory = strDesktop : '起始位置
oShellLink.Save : '创建保存快捷方式
```

</details>
