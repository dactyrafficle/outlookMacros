
Sub createHTMLMail()
 
 'declare the new mail item
 Dim objMail As Outlook.MailItem
 
 'create an outlook mail item
 Set objMail = Application.CreateItem(olMailItem)
 
 With objMail
 
 'set the body as HTML
 .BodyFormat = olFormatHTML
 
 'create the body - im styling with inline css
 .HTMLBody = "<HTML><BODY><H2 style='color: #ffaa00;'>The body of the outlook mail item is written in HTML</H2><p>And so on and so forth.</p></BODY></HTML>"
 
 'make the outlook mail item visible
 .Display
 
 End With
 
End Sub
