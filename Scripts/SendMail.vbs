Option Explicit
Dim outlook
Dim mail

Set outlook = CreateObject("Outlook.Application")
Set mail = outlook.CreateItem(0)

' �ݒ�
mail.To = "TODO"
mail.Cc = "TODO"
mail.Subject = "TODO(" & Month(Date) & "/" & Day(Date)& ")"
' Right("0" & Month(Date), 2) & "/" & Right("0" & Day(Date), 2)
mail.Body = "�{��" & vbCrLf & vbCrLf & "�ȏ�ł��B��낵�����肢���܂��B"

' �\��
mail.Display

' ���M����
' mail.Send

Set mail = Nothing
Set outlook = Nothing
