
' this subroutine uses the hsl2Rgb(h,s,l) function to only pick bright colors in hsl, then convert them to rgb
' atm, i need to convert all open mail items to html body:
'  the html mail item has property .TextColor, while plaintext has the property .Color -> need to standardize so it works on all

Sub colorizeActiveMailItem()

    Dim NewMail As MailItem, oInspector As Inspector
    
    'i think a new mail item gets thrown into the active inspector
    Set oInspector = Application.ActiveInspector
    
    
    If oInspector Is Nothing Then
        MsgBox "No active inspector"
    Else
    
        'is this only one item?
        Set NewMail = oInspector.CurrentItem
        
        'intersting thing
        If NewMail.Sent Then
            MsgBox "This is not an editable email"
        Else
        
            'isWordMail must be binary to make sure we can use COM
            If oInspector.IsWordMail Then
            
                ' i need to draw this thing out
                Dim oDoc As Object, oWrdApp As Object, oSelection As Object
                Set oDoc = oInspector.WordEditor
                Set oWrdApp = oDoc.Application
                Set oSelection = oWrdApp.Selection
                
                Dim n As Integer
                n = Len(oSelection.Document.content)
                MsgBox n
                
                
                Dim minSize, maxSize As Integer
                minSize = 12
                maxSize = 20
                
                
                  Dim size As Integer
                size = minSize + Int(Rnd() * (maxSize - minSize))
                
                Dim names As Variant
                names = Array("Arial", "Times New Roman", "Ebrima", "Gadugi", "Kristen ITC", "Trebuchet MS", "Tahoma", "Algerian", "Verdana", "Calibri", "Cambria", "Comic Sans MS", "Impact", "Century Gothic", "Candara", "Consolas", "Georgia", "Papyrus")
                
                Dim index As Integer
                index = Int(Rnd() * UBound(names))
                
                'this works too
                  Dim r, g, b As Single
                    r = Int(Rnd() * 256)
                    g = Int(Rnd() * 256)
                    b = Int(Rnd() * 256)
                    
                    Dim hue, sat, lum As Single
                    hue = Int(Rnd() * 360)
                    sat = 1
                    lum = 0.6
                

                
                  Dim i As Integer
                For i = 0 To n - 1
                
                  'if we encounter a space, then we need to change the font, size and color
                  
                  If oSelection.Document.Range(i, i + 1) = " " Then
                    
                    hue = 0 + Int(Rnd() * 360)
                    
                    'new size
                    size = minSize + Int(Rnd() * (maxSize - minSize))
                    
                    'new font
                    index = Int(Rnd() * UBound(names)) + 1
                    
                  End If
                  
                  ' i didn't know how to hook up a range - why it didn't work? but this did:
                  With oSelection.Document.Range(i, i + 1).Font
                    .size = size
                    .Name = names(index)
                    .TextColor = hslToRgb(hue, sat, lum)
                    
                  End With
                  
                Next i
    
                
                'oSelection.InsertAfter myText
                oSelection.Collapse 0
                Set oSelection = Nothing
                Set oWrdApp = Nothing
                Set oDoc = Nothing
            Else
                ' No object model to work with. Must manipulate raw text.
                Select Case NewMail.BodyFormat
                    Case olFormatPlain, olFormatRichText, olFormatUnspecified
                        NewMail.Body = NewMail.Body & myText
                    Case olFormatHTML
                        NewMail.HTMLBody = NewMail.HTMLBody & "<p>" & myText & " among other things </p>"
                End Select
            End If
        End If
    End If
End Sub
