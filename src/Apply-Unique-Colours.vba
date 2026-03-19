' ApplyUniqueColorsToBodies Macro
' Assigns a unique, evenly-spaced color to each body in the active SolidWorks part.

Dim swApp As SldWorks.SldWorks

Sub main()
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        MsgBox "Please open a part document.", vbCritical
        Exit Sub
    End If
    
    If swModel.GetType() <> swDocumentTypes_e.swDocPART Then
        MsgBox "This macro only works on part documents.", vbCritical
        Exit Sub
    End If
    
    Dim swPart As SldWorks.PartDoc
    Set swPart = swModel
    
    ' 0 represents swAllBodies enum value
    Dim vBodies As Variant
    vBodies = swPart.GetBodies2(0, False)
    
    If IsEmpty(vBodies) Then
        MsgBox "No bodies found in the active part.", vbInformation
        Exit Sub
    End If
    
    Dim totalBodies As Integer
    totalBodies = UBound(vBodies) + 1
    
    Dim i As Integer
    Dim hueStep As Double
    ' Distribute hues evenly across the 360 degree color wheel
    hueStep = 360 / totalBodies
    
    For i = 0 To UBound(vBodies)
        Dim swBody As SldWorks.Body2
        Set swBody = vBodies(i)
        
        Dim hue As Double
        hue = i * hueStep
        
        Dim rgbArr As Variant
        rgbArr = GetRGBFromHSV(hue, 1, 1) ' Full saturation and value
        
        ' SolidWorks MaterialPropertyValues2 expects an array of 9 doubles
        Dim dMatPrps(8) As Double
        dMatPrps(0) = rgbArr(0) ' R
        dMatPrps(1) = rgbArr(1) ' G
        dMatPrps(2) = rgbArr(2) ' B
        dMatPrps(3) = 1       ' Ambient
        dMatPrps(4) = 1       ' Diffuse
        dMatPrps(5) = 0.5     ' Specular
        dMatPrps(6) = 0.3125  ' Shininess
        dMatPrps(7) = 0       ' Transparency
        dMatPrps(8) = 0       ' Emission
        
        swBody.MaterialPropertyValues2 = dMatPrps
    Next
    
    ' Redraw the graphics area to display the new colors
    swModel.GraphicsRedraw2
    
    MsgBox "Successfully assigned unique colors to " & totalBodies & " bodies.", vbInformation
End Sub

' Function to convert HSV color space to an array of RGB doubles (0.0 to 1.0)
Function GetRGBFromHSV(H As Double, S As Double, V As Double) As Variant
    Dim C As Double, X As Double, m As Double
    Dim R As Double, G As Double, B As Double
    Dim H_prime As Double
    
    C = V * S
    H_prime = H / 60
    ' Calculate X with VBA robust modulo for floats
    X = C * (1 - Abs((H_prime - 2 * Int(H_prime / 2)) - 1))
    
    If H_prime >= 0 And H_prime < 1 Then
        R = C: G = X: B = 0
    ElseIf H_prime >= 1 And H_prime < 2 Then
        R = X: G = C: B = 0
    ElseIf H_prime >= 2 And H_prime < 3 Then
        R = 0: G = C: B = X
    ElseIf H_prime >= 3 And H_prime < 4 Then
        R = 0: G = X: B = C
    ElseIf H_prime >= 4 And H_prime < 5 Then
        R = X: G = 0: B = C
    ElseIf H_prime >= 5 And H_prime <= 6 Then
        R = C: G = 0: B = X
    Else
        R = 0: G = 0: B = 0
    End If
    
    m = V - C
    
    Dim rgbArray(2) As Double
    rgbArray(0) = R + m
    rgbArray(1) = G + m
    rgbArray(2) = B + m
    
    GetRGBFromHSV = rgbArray
End Function
