

Function hueToRgb(t1, t2, hue) As Single

  'by GEORGE!
  Dim h As Single
  h = hue '<--- this i did it

  If h < 0 Then
    h = h + 6
  End If
  If h >= 6 Then
    h = h - 6
  End If
  If h < 1 Then
    hueToRgb = (t2 - t1) * h + t1
  ElseIf h < 3 Then
   hueToRgb = t2
  ElseIf h < 4 Then
   hueToRgb = (t2 - t1) * (4 - h) + t1
  Else
   hueToRgb = t1
  End If
End Function

' hue: [0, 360)
' sat: [0, 1)
' lit: [0, 1)
Function hslToRgb(hue, sat, light) As Long
  
  'these functions actually modify the argument
  'thats not good
  'thats why i need to make a copy
  'just like pjs and the PVector class
  Dim h As Single '<--and and and this!
  h = hue '<---and this!!!

  Dim t1, t2, r, g, b As Double
  h = h / 60
  
  If light <= 0.5 Then
    t2 = light * (sat + 1)
  Else
    t2 = light + sat - (light * sat)
  End If
  
  t1 = light * 2 - t2
  r = hueToRgb(t1, t2, h + 2) * 255
  g = hueToRgb(t1, t2, h) * 255
  b = hueToRgb(t1, t2, h - 2) * 255
  
  hslToRgb = RGB(r, g, b)
  
End Function
