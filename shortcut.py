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


if __name__ == "__main__":
    CreateShortCut("readme.lnk", TargetPath='README.MD')
    print(ReadShortCut("readme.lnk"))
