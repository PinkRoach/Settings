Option Explicit
Dim outlook
Dim mail

Set outlook = CreateObject("Outlook.Application")
Set mail = outlook.CreateItem(0)

' 設定
mail.To = "TODO"
mail.Cc = "TODO"
mail.Subject = "TODO(" & Month(Date) & "/" & Day(Date)& ")"
' Right("0" & Month(Date), 2) & "/" & Right("0" & Day(Date), 2)
mail.Body = "本文" & vbCrLf & vbCrLf & "以上です。よろしくお願いします。"

' 表示
mail.Display

' 送信処理
' mail.Send

Set mail = Nothing
Set outlook = Nothing
