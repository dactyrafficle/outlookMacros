
Sub createHTMLMail()

 'declare the new mail item
 Dim objMail As Outlook.MailItem

 'create an outlook mail item
 Set objMail = Application.CreateItem(olMailItem)

 With objMail

  .BodyFormat = olFormatHTML
  .HTMLBody = "<HTML><BODY><H2 style='color: #ffaa00;'>The body of this mail item is in HTML</H2></BODY></HTML>"
  .Subject = "this one"
  .BCC = "sthing@email.com"
  .To = "sthing@email.com"
  .Display
  .Send
 
 End With
End Sub
