Option Explicit

Dim objShell

Set objShell = CreateObject("WScript.Shell")

objShell.Run "control main.cpl", 0, false

Wscript.Sleep(1000)

objShell.AppActivate("�}�E�X�̃v���p�e�B")

objShell.SendKeys("s")

' objShell.SendKeys("{ENTER}")
