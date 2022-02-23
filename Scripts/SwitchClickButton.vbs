Option Explicit

Dim objShell

Set objShell = CreateObject("WScript.Shell")

objShell.Run "control main.cpl", 0, false

Wscript.Sleep(1000)

objShell.AppActivate("マウスのプロパティ")

objShell.SendKeys("s")

' objShell.SendKeys("{ENTER}")
