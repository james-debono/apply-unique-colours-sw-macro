' ApplyUniqueColorsToBodies Macro - Version 4.4
' Assigns a unique, highly distinguishable color to each geometrically identical group of bodies or components.
'
' --- FULL CHANGELOG ---
' V4.4 Features:
' - Eliminated standard Entity.Select4 COM Selection Errors.
' - String-Based Entity Selection using raw SelectByID2.
'
' V4.3 Features:
' - Granular Error Telemetry with detailed step tracing.
' - Preventative Null Reference (Nothing) COM Interface fixes.
'
' V4.2 Features:
' - Type Mismatch Fix for SolidWorks Variant matrices.
' - Strict SpecifyConfiguration targeting via API constants.
'
' V4.1 Features:
' - Hardened Per-Configuration Targeting array fixes to prevent bleeding.
'
' V4.0 Features:
' - Maximized Visual Contrast using Golden Angle distribution.
' - Complex 7-Stage Saturation & Brightness Variations for thousands of hues.
'
' V3.1 Features:
' - Exclusive Subassembly filtering for deepest-level parts.
'
' V3.0 Features:
' - Face Appearance override checks to prevent accidental body collisions.
' - Assembly Level Support grouping via path identifiers.
'
' V2.0 Features:
' - Geometric Pattern Grouping (Volume, Area, Face Count).
'
' V1.0 Features:
' - Base sequential HSV generation.

Dim swApp As SldWorks.SldWorks

Sub main()
    On Error GoTo mainError
    
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
    Exit Sub
    
mainError:
    MsgBox "An error occurred in main(): " & Err.Description & " (Error " & Err.Number & ")", vbCritical
End Sub

' --- PART PROCESSING ---
Sub ProcessPart(swModel As SldWorks.ModelDoc2)
    Dim currentStep As String
    On Error GoTo ErrorHandler
    
    currentStep = "Initializing Variables"
    Dim swPart As SldWorks.PartDoc
    Dim vBodies As Variant
    Dim totalBodies As Integer
    Dim excludedHues() As Double
    Dim numExcludedHues As Integer
    Dim i As Integer, j As Integer
    Dim numGroups As Integer
    Dim groupVolume() As Double
    Dim groupArea() As Double
    Dim groupFaceCount() As Long
    Dim bodyGroup() As Integer
    Dim groupHues() As Double
    Dim sVals(6) As Double
    Dim vVals(6) As Double
    
    currentStep = "Fetching Active Configuration"
    Dim vConfigNames As Variant
    Dim sNames(0) As String
    Dim swConfig As Object
    Set swConfig = swModel.GetActiveConfiguration()
    If Not swConfig Is Nothing Then
        sNames(0) = swConfig.Name
    Else
        sNames(0) = "Default"
    End If
    vConfigNames = sNames
    
    currentStep = "Loading Bodies"
    Set swPart = swModel
    vBodies = swPart.GetBodies2(0, False)
    
    If IsEmpty(vBodies) Then
        MsgBox "No bodies found in the active part.", vbInformation
        Exit Sub
    End If
    
    totalBodies = UBound(vBodies) + 1
    
    currentStep = "Collecting Excluded Face Colors"
    numExcludedHues = 0
    ReDim excludedHues(1000)
    
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
    
    currentStep = "Grouping Bodies"
    numGroups = 0
    
    ReDim groupVolume(totalBodies)
    ReDim groupArea(totalBodies)
    ReDim groupFaceCount(totalBodies)
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
    
    currentStep = "Generating Unique Colors"
    If numGroups > 0 Then ReDim groupHues(numGroups - 1)
    
    For i = 0 To numGroups - 1
        Dim targetHue As Double
        targetHue = (i * 137.50776)
        targetHue = targetHue - 360 * Int(targetHue / 360)
        
        targetHue = FindSafeHue(targetHue, excludedHues, numExcludedHues)
        groupHues(i) = targetHue
    Next i
    
    sVals(0) = 1.00: vVals(0) = 1.00
    sVals(1) = 0.50: vVals(1) = 1.00
    sVals(2) = 1.00: vVals(2) = 0.50
    sVals(3) = 0.50: vVals(3) = 0.60
    sVals(4) = 0.75: vVals(4) = 0.75
    sVals(5) = 1.00: vVals(5) = 0.75
    sVals(6) = 0.75: vVals(6) = 1.00
    
    ' 4. Apply Colors
    currentStep = "Applying Colors to Bodies"
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
        
        Dim vMatPrps As Variant
        vMatPrps = dMatPrps
        
        currentStep = "Selecting Body via SelectByID2"
        swModel.ClearSelection2 True
        Dim bRet As Boolean
        
        ' Select the exact body via its internal name, bypasses COM casting requirements
        bRet = swModel.Extension.SelectByID2(swBody.Name, "SOLIDBODY", 0, 0, 0, False, 0, Nothing, 0)
        If Not bRet Then
            bRet = swModel.Extension.SelectByID2(swBody.Name, "SURFACEBODY", 0, 0, 0, False, 0, Nothing, 0)
        End If
        
        currentStep = "Executing Explicit Extension Material Property"
        If bRet Then
            ' 3 = swSpecifyConfiguration
            swModel.Extension.SetMaterialPropertyValues vMatPrps, 3, vConfigNames
        Else
            ' Fallback if selection fails for some reason
            swBody.MaterialPropertyValues2 = vMatPrps
        End If
    Next i
    
    MsgBox "Assigned vibrant unique colors to " & totalBodies & " bodies (" & numGroups & " geometric groups) in active configuration.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "ERROR in ProcessPart at step [" & currentStep & "]" & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           "Please let the developer know where it stopped.", vbCritical
End Sub

' --- ASSEMBLY PROCESSING ---
Sub ProcessAssembly(swModel As SldWorks.ModelDoc2)
    Dim currentStep As String
    On Error GoTo ErrorHandler
    
    currentStep = "Initializing Variables"
    Dim swAssy As SldWorks.AssemblyDoc
    Dim vComps As Variant
    Dim totalComps As Integer
    Dim excludedHues() As Double
    Dim numExcludedHues As Integer
    Dim i As Integer, j As Integer
    Dim numGroups As Integer
    Dim groupPaths() As String
    Dim compGroup() As Integer
    Dim skippedComps As Integer
    Dim groupHues() As Double
    Dim sVals(6) As Double
    Dim vVals(6) As Double
    Dim coloredComps As Integer
    
    currentStep = "Fetching Active Configuration"
    Dim vConfigNames As Variant
    Dim sNames(0) As String
    Dim swConfig As Object
    Set swConfig = swModel.GetActiveConfiguration()
    If Not swConfig Is Nothing Then
        sNames(0) = swConfig.Name
    Else
        sNames(0) = "Default"
    End If
    vConfigNames = sNames
    
    currentStep = "Loading Components"
    Set swAssy = swModel
    vComps = swAssy.GetComponents(False) 
    
    If IsEmpty(vComps) Then
        MsgBox "No components found in the active assembly.", vbInformation
        Exit Sub
    End If
    
    totalComps = UBound(vComps) + 1
    
    currentStep = "Collecting Component Extents"
    numExcludedHues = 0
    ReDim excludedHues(1000)
    
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
    
    currentStep = "Grouping Components"
    numGroups = 0
    ReDim groupPaths(totalComps)
    ReDim compGroup(totalComps)
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
    
    currentStep = "Generating Colors"
    If numGroups > 0 Then ReDim groupHues(numGroups - 1)
    
    For i = 0 To numGroups - 1
        Dim targetHue As Double
        targetHue = (i * 137.50776)
        targetHue = targetHue - 360 * Int(targetHue / 360)
        
        targetHue = FindSafeHue(targetHue, excludedHues, numExcludedHues)
        groupHues(i) = targetHue
    Next i
    
    sVals(0) = 1.00: vVals(0) = 1.00
    sVals(1) = 0.50: vVals(1) = 1.00
    sVals(2) = 1.00: vVals(2) = 0.50
    sVals(3) = 0.50: vVals(3) = 0.60
    sVals(4) = 0.75: vVals(4) = 0.75
    sVals(5) = 1.00: vVals(5) = 0.75
    sVals(6) = 0.75: vVals(6) = 1.00
    
    currentStep = "Applying Component Material Properties"
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
            
            Dim vMatPrps As Variant
            vMatPrps = dMatPrps
            
            ' 3 = swSpecifyConfiguration
            currentStep = "Setting Component Property Values"
            swComp.SetMaterialPropertyValues2 vMatPrps, 3, vConfigNames
            coloredComps = coloredComps + 1
        End If
    Next i
    
    currentStep = "Routine Completed"
    MsgBox "Assigned vibrant unique colors to " & coloredComps & " components (" & numGroups & " parts) explicitly in active configuration." & vbCrLf & _
           "Skipped " & skippedComps & " subassemblies.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "ERROR in ProcessAssembly at step [" & currentStep & "]" & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           "Please let the developer know where it stopped.", vbCritical
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
            If diff > 180 Then diff = Abs(360 - diff)
            
            If diff < 10 Then
                isSafe = False
                safeHue = safeHue + 15
                If safeHue >= 360 Then safeHue = safeHue - 360
                Exit For
            End If
        Next i
        
        loops = loops + 1
        If loops > 36 Then Exit Do
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
