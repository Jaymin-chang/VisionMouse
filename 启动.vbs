' 视觉拟人鼠标 —— 隐藏窗口启动
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
Set sh = CreateObject("WScript.Shell")
sh.Run "D:\PYTHON\pythonw.exe """ & dir & "\app.py""", 0, False
