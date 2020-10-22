Option Explicit

Sub Mail()

Dim OA1 As Outlook.Application
Dim OE1 As Outlook.MailItem

Set OA1 = New Outlook.Application
Set OE1 = OA1.CreateItem(olMailItem)

With OE1
    .BodyFormat = olFormatHTML
    .Display
    .HTMLBody = "Przesyłam raport danych.<br>" & "<br>" & .HTMLBody
    .Attachments.Add Environ("UserProfile") & "\Desktop\wyliczenia\Arkusz1.pdf"
    .To = "1@gmail.com; 2@gmail.com"
    .Subject = "Arkusz1"
    .Send
    
End With

End Sub