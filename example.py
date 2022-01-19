import shortcut
shortcut.CreateShortCut("readme.lnk", TargetPath='README.MD')
print(shortcut.ReadShortCut("readme.lnk"))
