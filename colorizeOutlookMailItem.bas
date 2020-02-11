Option Explicit

' this subroutine uses the hsl2Rgb(h,s,l) function to only pick bright colors in hsl, then convert them to rgb
' atm, i need to convert all open mail items to html body:
' the html mail item has property .TextColor, while plaintext has the property .Color -> need to standardize so it works on all

Sub colorizeActiveMailItem()

  Dim NewMail As MailItem, oInspector As Inspector
    
  'an inspector object is just the window in outlook that has an item
  'its just the window - not the contents
  'this line returns the topmost window in outlook
  Set oInspector = Application.ActiveInspector

  If oInspector Is Nothing Then
    MsgBox "No active inspector."
    Exit Sub
  End If
    
  'an inspector object can contain mail items
  'the .BodyFormat shouldnt matter
  'Application.ActiveInspector.CurrentItem -> gets the current mail item inside the topmost window
  Set NewMail = oInspector.CurrentItem
    'NewMail.BodyFormat
    ' olFormatHTML
    ' olFormatPlain
    ' olFormatRichText
    ' olFormatUnspecified
 
  'interesting: try running the macro on an open window containing a sent mail, open mail, received mail
  'Application.ActiveInspector.CurrentItem.Sent (boolean)
  If NewMail.Sent Then
    MsgBox "This is not an editable email"
    Exit Sub
  End If
        
  'Application.ActiveInspector.IsWordMail (boolean)
  'if true, we can use Com, just like if it was a MS Word application
  If oInspector.IsWordMail Then
            
    Dim oDoc As Object, oWrdApp As Object, oSelection As Object
    Set oDoc = oInspector.WordEditor
    Set oSelection = oInspector.WordEditor.Application.Selection
                
    'OH WOW it works
    'oDoc.content.Font.Color = RGB(255, 0, 0)
                
    Dim n As Long
    n = oDoc.content.Words.Count
    Debug.Print "number of words: " & n
                
    Dim hue, sat, lum As Single
    hue = Int(Rnd() * 360)
    sat = 1
    lum = 0.6
                
    Dim minSize As Long, maxSize As Long, size As Long
    minSize = 12
    maxSize = 20
    size = minSize + Int(Rnd() * (maxSize - minSize))
                
    Dim names As Variant
    names = Array("Arial", "Times New Roman", "Ebrima", "Gadugi", "Kristen ITC", "Trebuchet MS", "Tahoma", "Algerian", "Verdana", "Calibri", "Cambria", "Comic Sans MS", "Impact", "Century Gothic", "Candara", "Consolas", "Georgia", "Papyrus")
                
    Dim w As Object
    For Each w In oDoc.content.Words
      With w.Font
        .size = minSize + Int(Rnd() * (maxSize - minSize))
        .Name = names(Int(Rnd() * UBound(names)) + 1)
        .Color = hslToRgb(Int(Rnd() * 360), sat, lum) 'textcolor to color textColor works for html but not plaintext, isnt that weird
      End With
    Next w
                            
    oSelection.Collapse 0
    Set oSelection = Nothing
    Set oWrdApp = Nothing
    Set oDoc = Nothing
  
  Else
  
    MsgBox "This isn't a word mail object so I can't work with it using COM"
    Exit Sub
  
  End If 'closing isWordMail if stmt

End Sub


