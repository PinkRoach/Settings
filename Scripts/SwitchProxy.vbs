Option Explicit

Dim objShell
Dim SLEEP_TIME

' キー入力が受け付けられないことがあるので、スリープ
' 300（ギリギリ）400（時々失敗）500（ちょうど？）
SLEEP_TIME = 700

Set objShell = CreateObject("WScript.Shell")

objShell.Run "control inetcpl.cpl", 0, false

Wscript.Sleep(SLEEP_TIME)

' 接続タブへ移動
objShell.SendKeys("^{PGDN}")
objShell.SendKeys("^{PGDN}")
objShell.SendKeys("^{PGDN}")
objShell.SendKeys("^{PGDN}")

' LAN設定へ移動
objShell.SendKeys("l")

Wscript.Sleep(SLEEP_TIME)

' プロキシ設定切り替え
objShell.SendKeys("x")

Wscript.Sleep(SLEEP_TIME)

' 適用して終了
objShell.SendKeys("{ENTER}")

' インターネット設定終了
objShell.SendKeys("{ESC}")

