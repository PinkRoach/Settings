Option Explicit

Dim objShell
Dim SLEEP_TIME

' �L�[���͂��󂯕t�����Ȃ����Ƃ�����̂ŁA�X���[�v
' 300�i�M���M���j400�i���X���s�j500�i���傤�ǁH�j
SLEEP_TIME = 700

Set objShell = CreateObject("WScript.Shell")

objShell.Run "control inetcpl.cpl", 0, false

Wscript.Sleep(SLEEP_TIME)

' �ڑ��^�u�ֈړ�
objShell.SendKeys("^{PGDN}")
objShell.SendKeys("^{PGDN}")
objShell.SendKeys("^{PGDN}")
objShell.SendKeys("^{PGDN}")

' LAN�ݒ�ֈړ�
objShell.SendKeys("l")

Wscript.Sleep(SLEEP_TIME)

' �v���L�V�ݒ�؂�ւ�
objShell.SendKeys("x")

Wscript.Sleep(SLEEP_TIME)

' �K�p���ďI��
objShell.SendKeys("{ENTER}")

' �C���^�[�l�b�g�ݒ�I��
objShell.SendKeys("{ESC}")

