' ApplyUniqueColorsToBodies Macro - Version 4.1
' Assigns a unique, highly distinguishable color to each geometrically identical group of bodies or components.
' V4.1 Features:
' - Explicitly enforces per-configuration color assignments to prevent appearances from bleeding across multiple configurations in both Parts and Assemblies.
' V4 Features:
' - Uses Golden Angle algorithmic distribution for Hue to maximize contrast.
' - Cycles Saturation and Value (Brightness) through 7 distinct aesthetic tones to massively increase the number of distinguishable colors.
' V3.1 Features:
' - In Assemblies, correctly skips subassemblies so that individual leaf parts receive the color safely.
' V3 Features:
' - Avoids reusing colors already applied to individual faces.

Dim swApp As SldWorks.SldWorks

Sub main()
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        MsgBox "Please open a part or assembly document.", vbCritical
        Exit Sub
    End If
    
    Dim docType As Integer
    docType = swModel.GetType()
    
    If docType = swDocumentTypes_e.swDocPART Then
        ProcessPart swModel
    ElseIf docType = swDocumentTypes_e.swDocASSEMBLY Then
        ProcessAssembly swModel
    Else
        MsgBox "This macro only works on part or assembly documents.", vbCritical
        Exit Sub
    End If
    
    ' Redraw the graphics area to display the new colors
    swModel.GraphicsRedraw2
End Sub

' --- PART PROCESSING ---
Sub ProcessPart(swModel As SldWorks.ModelDoc2)
    Dim swPart As SldWorks.PartDoc
    Set swPart = swModel
    
    Dim vBodies As Variant
    vBodies = swPart.GetBodies2(0, False)
    
    If IsEmpty(vBodies) Then
        MsgBox "No bodies found in the active part.", vbInformation
        Exit Sub
    End If
    
    Dim totalBodies As Integer
    totalBodies = UBound(vBodies) + 1
    
    ' 1. Collect existing face colors to avoid
    Dim excludedHues() As Double
    Dim numExcludedHues As Integer
    numExcludedHues = 0
    ReDim excludedHues(1000)
    
    Dim i As Integer, j As Integer
    For i = 0 To totalBodies - 1
        Dim swBody As SldWorks.Body2
        Set swBody = vBodies(i)
        
        Dim vFaces As Variant
        vFaces = swBody.GetFaces
        If Not IsEmpty(vFaces) Then
            For j = 0 To UBound(vFaces)
                Dim swFace As SldWorks.Face2
                Set swFace = vFaces(j)
                
                Dim vMat As Variant
                vMat = swFace.MaterialPropertyValues
                If IsArray(vMat) Then
                    If UBound(vMat) >= 2 Then
                        Dim r As Double, g As Double, b As Double
                        r = vMat(0): g = vMat(1): b = vMat(2)
                        Dim faceHue As Double
                        faceHue = GetHSVHueFromRGB(r, g, b)
                        
                        If numExcludedHues >= UBound(excludedHues) Then
                            ReDim Preserve excludedHues(UBound(excludedHues) + 1000)
                        End If
                        excludedHues(numExcludedHues) = faceHue
                        numExcludedHues = numExcludedHues + 1
                    End If
                End If
            Next j
        End If
    Next i
    
    ' 2. Group bodies geometrically
    Dim numGroups As Integer
    numGroups = 0
    
    Dim groupVolume() As Double
    Dim groupArea() As Double
    Dim groupFaceCount() As Long
    
    ReDim groupVolume(totalBodies)
    ReDim groupArea(totalBodies)
    ReDim groupFaceCount(totalBodies)
    
    Dim bodyGroup() As Integer
    ReDim bodyGroup(totalBodies)
    
    For i = 0 To totalBodies - 1
        Set swBody = vBodies(i)
        
        Dim vProps As Variant
        vProps = swBody.GetMassProperties(1.0)
        
        Dim volume As Double, area As Double
        Dim faceCount As Long
        
        If IsArray(vProps) Then
            volume = vProps(3)            
            area = vProps(4)
        Else
            volume = 0
            area = 0
        End If
        
        faceCount = swBody.GetFaceCount
        
        Dim isMatch As Boolean
        isMatch = False
        Dim groupIndex As Integer
        
        For j = 0 To numGroups - 1
            Dim volLimit As Double, areaLimit As Double
            volLimit = Abs(groupVolume(j)) * 0.001 + 0.000000001
            areaLimit = Abs(groupArea(j)) * 0.001 + 0.000000001
            
            If Abs(groupVolume(j) - volume) <= volLimit And Abs(groupArea(j) - area) <= areaLimit And groupFaceCount(j) = faceCount Then
                isMatch = True
                groupIndex = j
                Exit For
            End If
        Next j
        
        If Not isMatch Then
            groupVolume(numGroups) = volume
            groupArea(numGroups) = area
            groupFaceCount(numGroups) = faceCount
            groupIndex = numGroups
            numGroups = numGroups + 1
        End If
        
        bodyGroup(i) = groupIndex
    Next i
    
    ' 3. Generate Colors and Avoid Excluded
    Dim groupHues() As Double
    If numGroups > 0 Then ReDim groupHues(numGroups - 1)
    
    For i = 0 To numGroups - 1
        Dim targetHue As Double
        targetHue = (i * 137.50776)
        targetHue = targetHue - 360 * Int(targetHue / 360)
        
        targetHue = FindSafeHue(targetHue, excludedHues, numExcludedHues)
        groupHues(i) = targetHue
    Next i
    
    Dim sVals(6) As Double
    Dim vVals(6) As Double
    sVals(0) = 1.00: vVals(0) = 1.00
    sVals(1) = 0.50: vVals(1) = 1.00
    sVals(2) = 1.00: vVals(2) = 0.50
    sVals(3) = 0.50: vVals(3) = 0.60
    sVals(4) = 0.75: vVals(4) = 0.75
    sVals(5) = 1.00: vVals(5) = 0.75
    sVals(6) = 0.75: vVals(6) = 1.00
    
    ' Explicitly retrieve the active configuration string name
    Dim configNames(0) As String
    Dim swConfig As SldWorks.Configuration
    Set swConfig = swModel.GetActiveConfiguration()
    If Not swConfig Is Nothing Then
        configNames(0) = swConfig.Name
    Else
        configNames(0) = "Default"
    End If
    
    ' 4. Apply Colors explicitly targeting the single configuration
    For i = 0 To totalBodies - 1
        Set swBody = vBodies(i)
        
        Dim grpIdx As Integer
        grpIdx = bodyGroup(i)
        
        Dim hue As Double
        hue = groupHues(grpIdx)
        
        Dim svIdx As Integer
        svIdx = grpIdx Mod 7
        
        Dim rgbArr As Variant
        rgbArr = GetRGBFromHSV(hue, sVals(svIdx), vVals(svIdx))
        
        Dim dMatPrps(8) As Double
        dMatPrps(0) = rgbArr(0)
        dMatPrps(1) = rgbArr(1)
        dMatPrps(2) = rgbArr(2)
        dMatPrps(3) = 1
        dMatPrps(4) = 1
        dMatPrps(5) = 0.5
        dMatPrps(6) = 0.3125
        dMatPrps(7) = 0
        dMatPrps(8) = 0
        
        Dim swEnt As SldWorks.Entity
        Set swEnt = swBody
        
        ' Select the body to apply material via extension
        swModel.ClearSelection2 True
        If Not swEnt Is Nothing Then
            If swEnt.Select4(False, Nothing) Then
                ' 1 = swInConfigurationOpts_e.swThisConfiguration
                swModel.Extension.SetMaterialPropertyValues dMatPrps, 1, configNames
            Else
                swBody.MaterialPropertyValues2 = dMatPrps
            End If
        Else
            swBody.MaterialPropertyValues2 = dMatPrps
        End If
    Next i
    
    MsgBox "Assigned vibrant unique colors to " & totalBodies & " bodies (" & numGroups & " geometric groups) in active configuration.", vbInformation
End Sub

' --- ASSEMBLY PROCESSING ---
Sub ProcessAssembly(swModel As SldWorks.ModelDoc2)
    Dim swAssy As SldWorks.AssemblyDoc
    Set swAssy = swModel
    
    Dim vComps As Variant
    vComps = swAssy.GetComponents(False) 
    
    If IsEmpty(vComps) Then
        MsgBox "No components found in the active assembly.", vbInformation
        Exit Sub
    End If
    
    Dim totalComps As Integer
    totalComps = UBound(vComps) + 1
    
    ' 1. Collect existing component-level colors to avoid
    Dim excludedHues() As Double
    Dim numExcludedHues As Integer
    numExcludedHues = 0
    ReDim excludedHues(1000)
    
    Dim i As Integer, j As Integer
    For i = 0 To totalComps - 1
        Dim swComp As SldWorks.Component2
        Set swComp = vComps(i)
        
        Dim path As String
        path = LCase(swComp.GetPathName())
        
        If Right(path, 7) = ".sldprt" Then
            Dim vMat As Variant
            vMat = swComp.MaterialPropertyValues
            If IsArray(vMat) Then
                If UBound(vMat) >= 2 Then
                    Dim r As Double, g As Double, b As Double
                    r = vMat(0): g = vMat(1): b = vMat(2)
                    Dim compHue As Double
                    compHue = GetHSVHueFromRGB(r, g, b)
                    
                    If numExcludedHues >= UBound(excludedHues) Then
                        ReDim Preserve excludedHues(UBound(excludedHues) + 1000)
                    End If
                    excludedHues(numExcludedHues) = compHue
                    numExcludedHues = numExcludedHues + 1
                End If
            End If
        End If
    Next i
    
    ' 2. Group components by Reference path (geometrically identical)
    Dim numGroups As Integer
    numGroups = 0
    
    Dim groupPaths() As String
    ReDim groupPaths(totalComps)
    
    Dim compGroup() As Integer
    ReDim compGroup(totalComps)
    
    Dim skippedComps As Integer
    skippedComps = 0
    
    For i = 0 To totalComps - 1
        Set swComp = vComps(i)
        
        path = LCase(swComp.GetPathName())
        
        If Right(path, 7) = ".sldprt" Then
            Dim isMatch As Boolean
            isMatch = False
            Dim groupIndex As Integer
            
            For j = 0 To numGroups - 1
                If groupPaths(j) = path Then
                    isMatch = True
                    groupIndex = j
                    Exit For
                End If
            Next j
            
            If Not isMatch Then
                groupPaths(numGroups) = path
                groupIndex = numGroups
                numGroups = numGroups + 1
            End If
            
            compGroup(i) = groupIndex
        Else
            compGroup(i) = -1
            skippedComps = skippedComps + 1
        End If
    Next i
    
    ' 3. Generate Component Colors
    Dim groupHues() As Double
    If numGroups > 0 Then ReDim groupHues(numGroups - 1)
    
    For i = 0 To numGroups - 1
        Dim targetHue As Double
        targetHue = (i * 137.50776)
        targetHue = targetHue - 360 * Int(targetHue / 360)
        
        targetHue = FindSafeHue(targetHue, excludedHues, numExcludedHues)
        groupHues(i) = targetHue
    Next i
    
    Dim sVals(6) As Double
    Dim vVals(6) As Double
    sVals(0) = 1.00: vVals(0) = 1.00
    sVals(1) = 0.50: vVals(1) = 1.00
    sVals(2) = 1.00: vVals(2) = 0.50
    sVals(3) = 0.50: vVals(3) = 0.60
    sVals(4) = 0.75: vVals(4) = 0.75
    sVals(5) = 1.00: vVals(5) = 0.75
    sVals(6) = 0.75: vVals(6) = 1.00
    
    ' Retrieve the active configuration string explicitly
    Dim configNames(0) As String
    Dim swConfig As SldWorks.Configuration
    Set swConfig = swModel.GetActiveConfiguration()
    If Not swConfig Is Nothing Then
        configNames(0) = swConfig.Name
    Else
        configNames(0) = "Default"
    End If
    
    ' 4. Apply Configuration-Specific Colors
    Dim coloredComps As Integer
    coloredComps = 0
    
    For i = 0 To totalComps - 1
        Set swComp = vComps(i)
        
        If compGroup(i) >= 0 Then
            Dim grpIdx As Integer
            grpIdx = compGroup(i)
            
            Dim hue As Double
            hue = groupHues(grpIdx)
            
            Dim svIdx As Integer
            svIdx = grpIdx Mod 7
            
            Dim rgbArr As Variant
            rgbArr = GetRGBFromHSV(hue, sVals(svIdx), vVals(svIdx))
            
            Dim dMatPrps(8) As Double
            dMatPrps(0) = rgbArr(0)
            dMatPrps(1) = rgbArr(1)
            dMatPrps(2) = rgbArr(2)
            dMatPrps(3) = 1
            dMatPrps(4) = 1
            dMatPrps(5) = 0.5
            dMatPrps(6) = 0.3125
            dMatPrps(7) = 0
            dMatPrps(8) = 0
            
            ' Pass exactly 1 for swThisConfiguration and the array of config names
            swComp.SetMaterialPropertyValues2 dMatPrps, 1, configNames
            coloredComps = coloredComps + 1
        End If
    Next i
    
    MsgBox "Assigned vibrant unique colors to " & coloredComps & " components (" & numGroups & " parts) explicitly in active configuration." & vbCrLf & _
           "Skipped " & skippedComps & " subassemblies.", vbInformation
End Sub

' Function to find a safe hue that avoids excluded hues
Function FindSafeHue(targetHue As Double, excludedHues() As Double, numExcludedHues As Integer) As Double
    Dim safeHue As Double
    safeHue = targetHue
    
    Dim isSafe As Boolean
    Dim loops As Integer
    loops = 0
    
    Do
        isSafe = True
        Dim i As Integer
        For i = 0 To numExcludedHues - 1
            Dim diff As Double
            diff = Abs(safeHue - excludedHues(i))
            ' Handle circular hue nature
            If diff > 180 Then diff = Abs(360 - diff)
            
            ' If hue is within 10 degrees of an excluded hue, shift it
            If diff < 10 Then
                isSafe = False
                safeHue = safeHue + 15
                If safeHue >= 360 Then safeHue = safeHue - 360
                Exit For
            End If
        Next i
        
        loops = loops + 1
        If loops > 36 Then Exit Do ' Prevent infinite shifting (max 36 steps)
    Loop Until isSafe
    
    FindSafeHue = safeHue
End Function

' Function to convert RGB to Hue (0-360)
Function GetHSVHueFromRGB(R As Double, G As Double, B As Double) As Double
    Dim maxVal As Double, minVal As Double, diff As Double
    
    maxVal = R
    If G > maxVal Then maxVal = G
    If B > maxVal Then maxVal = B
    
    minVal = R
    If G < minVal Then minVal = G
    If B < minVal Then minVal = B
    
    diff = maxVal - minVal
    
    If diff = 0 Then
        GetHSVHueFromRGB = 0
    ElseIf maxVal = R Then
        Dim tempHueR As Double
        tempHueR = (60 * ((G - B) / diff) + 360)
        GetHSVHueFromRGB = tempHueR - 360 * Int(tempHueR / 360)
    ElseIf maxVal = G Then
        GetHSVHueFromRGB = (60 * ((B - R) / diff) + 120)
    ElseIf maxVal = B Then
        GetHSVHueFromRGB = (60 * ((R - G) / diff) + 240)
    End If
End Function

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
