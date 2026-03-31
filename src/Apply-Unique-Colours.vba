' ApplyUniqueColorsToBodies Macro - Version 5.2
' Assigns a unique, highly distinguishable color to each geometrically identical group of bodies or components.
'
' --- MAJOR CHANGELOG ---
' V5 Features:
' - Improved part differentiation using pure Mathematical Physics Solver (Moments of Inertia array) with nano-tolerances and edge mapping.
' - Adaptive Equidistant Color Generation: Dynamically distributes interwoven Hues across 7 SV shading strata.
'
' V4 Features:
' - Display State Injection: Natively targets the active Display State and bypasses generic COM selections.
' - Initial implementation of Golden Angle contrast mapping.
'
' V3 Features:
' - Assembly Level Processing: Filters subassemblies strictly for deepest-level parts.
' - Face Appearance override checks to prevent accidental body collisions.
'
' V2 Features:
' - Recognise like bodies using Geometric Pattern Grouping (Volume, Area, Face Count).
'
' V1 Features:
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
    Dim groupEdgeCount() As Long
    
    Dim groupM1() As Double
    Dim groupM2() As Double
    Dim groupM3() As Double
    Dim bodyGroup() As Integer
    Dim groupHues() As Double
    
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
                        Dim faceHue As Double
                        faceHue = GetHSVHueFromRGB(CDbl(vMat(0)), CDbl(vMat(1)), CDbl(vMat(2)))
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
    ReDim groupEdgeCount(totalBodies)
    ReDim groupM1(totalBodies)
    ReDim groupM2(totalBodies)
    ReDim groupM3(totalBodies)
    ReDim bodyGroup(totalBodies)
    
    For i = 0 To totalBodies - 1
        Set swBody = vBodies(i)
        
        Dim volume As Double, area As Double, faceCount As Long, edgeCount As Long
        Dim m1 As Double, m2 As Double, m3 As Double
        m1 = 0: m2 = 0: m3 = 0
        
        currentStep = "Extracting Raw Origin Mass Tensor"
        Dim vProps As Variant
        vProps = swBody.GetMassProperties(1.0)
        
        If IsArray(vProps) Then
            Dim cx As Double, cy As Double, cz As Double
            Dim mass As Double
            Dim Ixx_o As Double, Iyy_o As Double, Izz_o As Double
            Dim Ixy_o As Double, Ixz_o As Double, Iyz_o As Double
            
            cx = vProps(0): cy = vProps(1): cz = vProps(2)
            volume = vProps(3)            
            area = vProps(4)
            Ixx_o = vProps(5): Iyy_o = vProps(6): Izz_o = vProps(7)
            Ixy_o = vProps(8): Ixz_o = vProps(9): Iyz_o = vProps(10)
            mass = vProps(11)
            
            ' Parallel Axis Theorem: Shift to Center of Mass
            Dim Ixx As Double, Iyy As Double, Izz As Double
            Dim Ixy As Double, Ixz As Double, Iyz As Double
            Ixx = Ixx_o - mass * (cy * cy + cz * cz)
            Iyy = Iyy_o - mass * (cx * cx + cz * cz)
            Izz = Izz_o - mass * (cx * cx + cy * cy)
            Ixy = Ixy_o - mass * (cx * cy)
            Ixz = Ixz_o - mass * (cx * cz)
            Iyz = Iyz_o - mass * (cy * cz)
            
            ' Inertia Tensor format for Eigenvalues
            Dim Txx As Double, Tyy As Double, Tzz As Double
            Dim Txy As Double, Txz As Double, Tyz As Double
            Txx = Ixx: Tyy = Iyy: Tzz = Izz
            Txy = -Ixy: Txz = -Ixz: Tyz = -Iyz
            
            ' Cubic Characteristic Equation Coefficients: x^3 + b2 x^2 + b1 x + b0 = 0
            Dim c2 As Double, c1 As Double, c0 As Double
            c2 = Txx + Tyy + Tzz
            c1 = (Txx * Tyy - Txy * Txy) + (Txx * Tzz - Txz * Txz) + (Tyy * Tzz - Tyz * Tyz)
            c0 = Txx * (Tyy * Tzz - Tyz * Tyz) - Txy * (Txy * Tzz - Txz * Tyz) + Txz * (Txy * Tyz - Tyy * Txz)
            
            Dim b2 As Double, b1 As Double, b0 As Double
            b2 = -c2: b1 = c1: b0 = -c0
            
            Dim Q As Double, R As Double, D As Double
            Q = (3 * b1 - b2 * b2) / 9
            R = (9 * b2 * b1 - 27 * b0 - 2 * b2 * b2 * b2) / 54
            D = Q * Q * Q + R * R
            
            If D <= 0 Then
                Dim Q3 As Double
                Q3 = -Q * Q * Q
                If Q3 <= 0 Then Q3 = 0.0000000001
                
                Dim ratio As Double
                ratio = R / Sqr(Q3)
                If ratio > 1 Then ratio = 1
                If ratio < -1 Then ratio = -1
                
                Dim theta As Double
                If ratio = 1 Then
                    theta = 0
                ElseIf ratio = -1 Then
                    theta = 3.14159265358979
                Else
                    theta = Atn(-ratio / Sqr(-ratio * ratio + 1)) + 2 * Atn(1)
                End If
                
                Dim sq_ngQ As Double
                If Q < 0 Then
                    sq_ngQ = Sqr(-Q)
                Else
                    sq_ngQ = 0
                End If
                
                m1 = 2 * sq_ngQ * Cos(theta / 3) - b2 / 3
                m2 = 2 * sq_ngQ * Cos((theta + 2 * 3.14159265358979) / 3) - b2 / 3
                m3 = 2 * sq_ngQ * Cos((theta + 4 * 3.14159265358979) / 3) - b2 / 3
            End If
        Else
            volume = 0: area = 0
        End If
        
        faceCount = swBody.GetFaceCount
        edgeCount = swBody.GetEdgeCount
        
        ' Sequence the Principal Moments
        Dim temp As Double
        If m1 > m2 Then temp = m1: m1 = m2: m2 = temp
        If m2 > m3 Then temp = m2: m2 = m3: m3 = temp
        If m1 > m2 Then temp = m1: m1 = m2: m2 = temp
        
        currentStep = "Comparing Geometric Tolerance Limits"
        Dim isMatch As Boolean
        isMatch = False
        Dim groupIndex As Integer
        
        For j = 0 To numGroups - 1
            Dim volLimit As Double, areaLimit As Double
            Dim m1L As Double, m2L As Double, m3L As Double
            
            volLimit = Abs(groupVolume(j)) * 0.00000001 + 0.000000000001
            areaLimit = Abs(groupArea(j)) * 0.00000001 + 0.000000000001
            m1L = Abs(groupM1(j)) * 0.000000001 + 0.000000000001
            m2L = Abs(groupM2(j)) * 0.000000001 + 0.000000000001
            m3L = Abs(groupM3(j)) * 0.000000001 + 0.000000000001
            
            If Abs(groupVolume(j) - volume) <= volLimit And Abs(groupArea(j) - area) <= areaLimit And groupFaceCount(j) = faceCount And groupEdgeCount(j) = edgeCount Then
                If Abs(groupM1(j) - m1) <= m1L And Abs(groupM2(j) - m2) <= m2L And Abs(groupM3(j) - m3) <= m3L Then
                    isMatch = True
                    groupIndex = j
                    Exit For
                End If
            End If
        Next j
        
        If Not isMatch Then
            groupVolume(numGroups) = volume
            groupArea(numGroups) = area
            groupFaceCount(numGroups) = faceCount
            groupEdgeCount(numGroups) = edgeCount
            groupM1(numGroups) = m1
            groupM2(numGroups) = m2
            groupM3(numGroups) = m3
            groupIndex = numGroups
            numGroups = numGroups + 1
        End If
        
        bodyGroup(i) = groupIndex
    Next i
    
    currentStep = "Generating Unique Colors (Adaptive Equidistant)"
    
    Dim totalLayers As Integer
    Dim itemsPerLayer As Integer
    If numGroups <= 8 Then
        totalLayers = 1
        itemsPerLayer = numGroups
    Else
        totalLayers = Int((numGroups - 1) / 8) + 1
        If totalLayers > 7 Then totalLayers = 7
        itemsPerLayer = Int((numGroups - 1) / totalLayers) + 1
    End If
    
    If numGroups > 0 Then ReDim groupHues(numGroups - 1)
    
    For i = 0 To numGroups - 1
        Dim layerIdx As Integer
        layerIdx = Int(i / itemsPerLayer)
        Dim indexInLayer As Integer
        indexInLayer = i Mod itemsPerLayer
        
        Dim currentLayerItems As Integer
        If layerIdx = totalLayers - 1 Then
            currentLayerItems = numGroups - (layerIdx * itemsPerLayer)
        Else
            currentLayerItems = itemsPerLayer
        End If
        
        Dim targetHue As Double
        If currentLayerItems > 1 Then
            targetHue = (indexInLayer * (360# / currentLayerItems))
            Dim offset As Double
            offset = (360# / currentLayerItems) / totalLayers
            targetHue = targetHue + (layerIdx * offset)
        Else
            targetHue = layerIdx * 45
        End If
        
        If targetHue >= 360 Then targetHue = targetHue - 360 * Int(targetHue / 360)
        targetHue = FindSafeHue(targetHue, excludedHues, numExcludedHues)
        groupHues(i) = targetHue
    Next i
    
    Dim sVals(6) As Double
    Dim vVals(6) As Double
    sVals(0) = 1.00: vVals(0) = 1.00 ' Bright
    sVals(1) = 0.45: vVals(1) = 1.00 ' Pastel
    sVals(2) = 1.00: vVals(2) = 0.50 ' Dark
    sVals(3) = 0.20: vVals(3) = 0.90 ' Pale
    sVals(4) = 0.60: vVals(4) = 1.00 ' Soft
    sVals(5) = 0.85: vVals(5) = 0.70 ' Muted
    sVals(6) = 1.00: vVals(6) = 0.35 ' Very Dark
    
    currentStep = "Configuring Display State Targets"
    Dim swDispStateSetts As SldWorks.DisplayStateSetting
    Set swDispStateSetts = swModel.Extension.GetDisplayStateSetting(swDisplayStateOpts_e.swThisDisplayState)
    
    If swDispStateSetts Is Nothing Then
        MsgBox "SolidWorks failed to generate a Display State Setting for this Body.", vbInformation
        Exit Sub
    End If
    
    swDispStateSetts.Option = swDisplayStateOpts_e.swThisDisplayState
    
    currentStep = "Applying Colors to Bodies via Display States"
    For i = 0 To totalBodies - 1
        Set swBody = vBodies(i)
        
        Dim grpIdx As Integer
        grpIdx = bodyGroup(i)
        Dim hue As Double
        hue = groupHues(grpIdx)
        
        Dim lIdx As Integer
        lIdx = Int(grpIdx / itemsPerLayer)
        If lIdx > 6 Then lIdx = 6
        
        Dim rgbArr As Variant
        rgbArr = GetRGBFromHSV(hue, sVals(lIdx), vVals(lIdx))
        
        Dim entities(0) As Object 
        Set entities(0) = swBody
        swDispStateSetts.Entities = entities
        
        Dim vAppearances As Variant
        vAppearances = swModel.Extension.DisplayStateSpecMaterialPropertyValues(swDispStateSetts)
        
        If Not IsEmpty(vAppearances) Then
            Dim appearSet As SldWorks.AppearanceSetting
            Set appearSet = vAppearances(0)
            
            appearSet.Color = RGB(Int(rgbArr(0) * 255), Int(rgbArr(1) * 255), Int(rgbArr(2) * 255))
            appearSet.Diffuse = 1#
            appearSet.Specular = 0.5
            appearSet.Luminous = 0#
            
            Dim newAppearances(0) As SldWorks.AppearanceSetting
            Set newAppearances(0) = appearSet
            
            swModel.Extension.DisplayStateSpecMaterialPropertyValues(swDispStateSetts) = newAppearances
        Else
            Dim dMatPrps(8) As Double
            dMatPrps(0) = rgbArr(0): dMatPrps(1) = rgbArr(1): dMatPrps(2) = rgbArr(2)
            dMatPrps(3) = 1: dMatPrps(4) = 1: dMatPrps(5) = 0.5: dMatPrps(6) = 0.3125: dMatPrps(7) = 0: dMatPrps(8) = 0
            swBody.MaterialPropertyValues2 = dMatPrps
        End If
    Next i
    
    MsgBox "Applied unique colours to bodies in active display state" & vbCrLf & _
           "Total Bodies: " & totalBodies & vbCrLf & _
           "Unique Bodies: " & numGroups & vbCrLf & vbCrLf & _
           "Macro Version: 5.2", vbInformation
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
    Dim coloredComps As Integer
    
    currentStep = "Loading Components"
    Set swAssy = swModel
    vComps = swAssy.GetComponents(False) 
    
    If IsEmpty(vComps) Then
        MsgBox "No components found in the active assembly.", vbInformation
        Exit Sub
    End If
    
    totalComps = UBound(vComps) + 1
    
    currentStep = "Collecting Excluded Component Colors"
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
                    Dim faceHue As Double
                    faceHue = GetHSVHueFromRGB(CDbl(vMat(0)), CDbl(vMat(1)), CDbl(vMat(2)))
                    If numExcludedHues >= UBound(excludedHues) Then
                        ReDim Preserve excludedHues(UBound(excludedHues) + 1000)
                    End If
                    excludedHues(numExcludedHues) = faceHue
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
        Dim pathStr As String
        pathStr = LCase(swComp.GetPathName())
        
        If Right(pathStr, 7) = ".sldprt" Then
            Dim isMatch As Boolean
            isMatch = False
            Dim groupIndex As Integer
            
            For j = 0 To numGroups - 1
                If groupPaths(j) = pathStr Then
                    isMatch = True
                    groupIndex = j
                    Exit For
                End If
            Next j
            
            If Not isMatch Then
                groupPaths(numGroups) = pathStr
                groupIndex = numGroups
                numGroups = numGroups + 1
            End If
            compGroup(i) = groupIndex
        Else
            compGroup(i) = -1
            skippedComps = skippedComps + 1
        End If
    Next i
    
    currentStep = "Generating Colors (Adaptive Equidistant)"
    
    Dim totalLayers As Integer
    Dim itemsPerLayer As Integer
    If numGroups <= 8 Then
        totalLayers = 1
        itemsPerLayer = numGroups
    Else
        totalLayers = Int((numGroups - 1) / 8) + 1
        If totalLayers > 7 Then totalLayers = 7
        itemsPerLayer = Int((numGroups - 1) / totalLayers) + 1
    End If
    
    If numGroups > 0 Then ReDim groupHues(numGroups - 1)
    
    For i = 0 To numGroups - 1
        Dim layerIdx As Integer
        layerIdx = Int(i / itemsPerLayer)
        
        Dim indexInLayer As Integer
        indexInLayer = i Mod itemsPerLayer
        
        Dim currentLayerItems As Integer
        If layerIdx = totalLayers - 1 Then
            currentLayerItems = numGroups - (layerIdx * itemsPerLayer)
        Else
            currentLayerItems = itemsPerLayer
        End If
        
        Dim targetHue As Double
        If currentLayerItems > 1 Then
            targetHue = (indexInLayer * (360# / currentLayerItems))
            Dim offset As Double
            offset = (360# / currentLayerItems) / totalLayers
            targetHue = targetHue + (layerIdx * offset)
        Else
            targetHue = layerIdx * 45
        End If
        
        If targetHue >= 360 Then targetHue = targetHue - 360 * Int(targetHue / 360)
        targetHue = FindSafeHue(targetHue, excludedHues, numExcludedHues)
        groupHues(i) = targetHue
    Next i
    
    Dim sVals(6) As Double
    Dim vVals(6) As Double
    sVals(0) = 1.00: vVals(0) = 1.00 ' Bright
    sVals(1) = 0.45: vVals(1) = 1.00 ' Pastel
    sVals(2) = 1.00: vVals(2) = 0.50 ' Dark
    sVals(3) = 0.20: vVals(3) = 0.90 ' Pale
    sVals(4) = 0.60: vVals(4) = 1.00 ' Soft
    sVals(5) = 0.85: vVals(5) = 0.70 ' Muted
    sVals(6) = 1.00: vVals(6) = 0.35 ' Very Dark
    
    currentStep = "Configuring Display State Targets"
    Dim swDispStateSetts As SldWorks.DisplayStateSetting
    Set swDispStateSetts = swModel.Extension.GetDisplayStateSetting(swDisplayStateOpts_e.swThisDisplayState)
    swDispStateSetts.Option = swDisplayStateOpts_e.swThisDisplayState
    swDispStateSetts.PartLevel = False
    
    currentStep = "Applying Component Material Properties via Display States"
    coloredComps = 0
    
    For i = 0 To totalComps - 1
        Set swComp = vComps(i)
        
        If compGroup(i) >= 0 Then
            Dim grpIdx As Integer
            grpIdx = compGroup(i)
            Dim hue As Double
            hue = groupHues(grpIdx)
            
            Dim lIdx As Integer
            lIdx = Int(grpIdx / itemsPerLayer)
            If lIdx > 6 Then lIdx = 6
            
            Dim rgbArr As Variant
            rgbArr = GetRGBFromHSV(hue, sVals(lIdx), vVals(lIdx))
            
            Dim entities(0) As Object
            Set entities(0) = swComp
            swDispStateSetts.Entities = entities
            
            Dim vAppearances As Variant
            vAppearances = swModel.Extension.DisplayStateSpecMaterialPropertyValues(swDispStateSetts)
            
            If Not IsEmpty(vAppearances) Then
                Dim appearSet As SldWorks.AppearanceSetting
                Set appearSet = vAppearances(0)
                
                appearSet.Color = RGB(Int(rgbArr(0) * 255), Int(rgbArr(1) * 255), Int(rgbArr(2) * 255))
                appearSet.Diffuse = 1#
                appearSet.Specular = 0.5
                appearSet.Luminous = 0#
                
                Dim newAppearances(0) As SldWorks.AppearanceSetting
                Set newAppearances(0) = appearSet
                
                swModel.Extension.DisplayStateSpecMaterialPropertyValues(swDispStateSetts) = newAppearances
            Else
                Dim dMatPrps(8) As Double
                dMatPrps(0) = rgbArr(0): dMatPrps(1) = rgbArr(1): dMatPrps(2) = rgbArr(2)
                dMatPrps(3) = 1: dMatPrps(4) = 1: dMatPrps(5) = 0.5: dMatPrps(6) = 0.3125: dMatPrps(7) = 0: dMatPrps(8) = 0
                swComp.SetMaterialPropertyValues2 dMatPrps, 1, Empty
            End If
            
            coloredComps = coloredComps + 1
        End If
    Next i
    
    MsgBox "Applied unique colours to components in active display state" & vbCrLf & _
           "Total Bodies: " & coloredComps & vbCrLf & _
           "Unique Bodies: " & numGroups & vbCrLf & _
           "Skipped Subassemblies: " & skippedComps & vbCrLf & vbCrLf & _
           "Macro Version: 5.2", vbInformation
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
