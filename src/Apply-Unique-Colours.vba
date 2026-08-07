'==============================================================================
' Apply Unique Colours
'
' Assigns a distinct, high-contrast colour to every geometrically unique body
' in a part, or to every unique component in an assembly. Bodies or components
' that are identical to one another share a colour, so anything that differs
' stands out immediately on screen.
'
' In a part, two bodies are treated as the same if the SOLIDWORKS geometry kernel
' can move one onto the other, or failing that if every vertex sits the same
' distance from the body's centre of mass to within a micron. Position and
' orientation are therefore irrelevant, while a genuine difference - a hole moved
' a fraction of a millimetre - separates them.
'
' In an assembly, components are grouped by file path and referenced
' configuration. Subassemblies are skipped; only bottom-level parts are
' coloured.
'
' Colours are spaced evenly around the hue wheel and corrected so that every
' colour carries the same perceived brightness, which stops blues and violets
' reading as near-black next to yellows.
'
' Colours are written to the active display state only, so other display states
' and configurations are left untouched.
'
' To use, open a part or assembly document and run the macro.
'
'   Version   0.10.1
'   Date      2026-08-07
'   Author    James Debono
'
'------------------------------------------------------------------------------
' CHANGELOG (summary - see CHANGELOG.md for the full history)
'
'   0.6.x   Body matching performed by the geometry kernel rather than by
'           numeric shape invariants. Luminance-corrected palette.
'   0.5.x   Body grouping driven by the mass-normalised principal moments of
'           inertia, invariant to position, orientation and material density.
'   0.4.x   Adaptive equidistant colour engine with interleaved saturation and
'           brightness layers. Colours applied through display states.
'   0.3.x   Assembly support, restricted to bottom-level parts. Existing face
'           colours avoided.
'   0.2.x   Identical bodies grouped by volume, surface area and face count.
'   0.1.x   Initial release. Sequential HSV colour per body.
'==============================================================================

Option Explicit

'--- Settings -----------------------------------------------------------------

' Compare bodies with the geometry kernel (IBody2::GetCoincidenceTransform2).
' This is an exact test: two bodies match only if one can be moved onto the
' other. Set to False to force the numeric invariant method used up to 0.5.18.
Const USE_KERNEL_COMPARISON As Boolean = True

' Give mirrored copies of a body the same colour as the original.
'
' The geometry kernel treats a mirror image as coincident, so this is on by
' default in SOLIDWORKS itself. Left off here, because a mirrored part is
' usually a genuinely different part: a bracket and its opposite hand are not
' interchangeable on the shop floor even though their geometry matches. Turn it
' on if mirrored instances of the same hardware should read as one part.
Const MATCH_MIRRORED As Boolean = False

' Shift generated colours away from colours already applied to individual faces.
' Costs one API call per face in the model, so turning it off is the single
' biggest speed-up available on models with no manual face colours.
Const AVOID_EXISTING_FACE_COLOURS As Boolean = True

' Report timings and per-group detail when the macro finishes. Group detail is
' written to the VBA Immediate window (Ctrl+G in the editor).
Const SHOW_DIAGNOSTICS As Boolean = True

' How far apart two bodies' volume and surface area may be and still be handed to
' the kernel for a proper comparison. Its only job is to skip pairs that obviously
' cannot match, so it is deliberately generous.
'
' Two instances of the same imported part can differ in computed volume by ten
' parts per million or more, because each insert-part operation re-derives the
' geometry rather than reusing one B-rep. 0.6.2 gated at one part per million on
' the assumption that identical bodies agree to near machine precision; that was
' wrong, and it rejected genuine matches before the kernel ever saw them. At one
' part per thousand the gate still discards the overwhelming majority of pairs
' while leaving the match decision where it belongs - with the kernel.
Const SIZE_GATE_TOLERANCE As Double = 0.001

' When the kernel reports two bodies are not coincident, compare their shape
' directly as a second opinion: the sorted distances of every vertex from the
' body's centre of mass, matched term by term.
'
' This exists because the kernel's comparison carries a tolerance of its own, and
' hardware imported through several separate insert-part features falls just
' outside it. The same rivet nut derived twice was measured nine parts per
' million apart in volume - a few tens of nanometres of geometry.
'
' Aggregate measurements cannot tell that apart from a real difference, and in
' fact get it backwards: two tubes with a hole in different places have volumes
' agreeing to fifteen decimal places, because moving a hole removes exactly as
' much material as before. Vertex positions do not have that blind spot - a hole
' that moves takes its vertices with it.
Const SHAPE_FALLBACK As Boolean = True

' How far two corresponding vertices may sit apart, in metres. One micron is far
' above the tens of nanometres separating re-derived copies of one part, and far
' below any feature displacement worth colouring differently.
Const SHAPE_TOLERANCE As Double = 0.000001

Const MACRO_VERSION As String = "0.10.1"

' Perceived brightness of each colour layer, darkest usable to lightest.
' Values are relative luminance, 0 = black, 1 = white.
Const LAYER_COUNT As Long = 7
Const ITEMS_PER_LAYER As Long = 8

Dim swApp As SldWorks.SldWorks

'--- Entry point --------------------------------------------------------------

Sub main()
    On Error GoTo mainError

    Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Please open a part or assembly document.", vbCritical
        Exit Sub
    End If

    Dim docType As Long
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
    MsgBox "An error occurred in main(): " & Err.Description & _
           " (Error " & Err.Number & ")", vbCritical
End Sub

'--- Part processing ----------------------------------------------------------

Sub ProcessPart(swModel As SldWorks.ModelDoc2)

    Dim currentStep As String
    Dim swView As SldWorks.ModelView
    On Error GoTo ErrorHandler

    Dim tStart As Single, tMeasured As Single, tGrouped As Single, tDone As Single
    tStart = Timer

    currentStep = "Loading bodies"
    Dim swPart As SldWorks.PartDoc
    Set swPart = swModel

    Dim vBodies As Variant
    vBodies = swPart.GetBodies2(swBodyType_e.swSolidBody, False)

    If IsEmpty(vBodies) Then
        MsgBox "No solid bodies found in the active part.", vbInformation
        Exit Sub
    End If

    Dim totalBodies As Long
    totalBodies = UBound(vBodies) + 1

    Set swView = swModel.ActiveView
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = False

    Dim i As Long, j As Long, k As Long

    ' --- Existing face colours to avoid ---------------------------------------
    currentStep = "Collecting existing face colours"
    Dim excludedHues() As Double
    Dim numExcludedHues As Long
    ReDim excludedHues(1000)
    numExcludedHues = 0

    If AVOID_EXISTING_FACE_COLOURS Then
        Dim swBody As SldWorks.Body2
        For i = 0 To totalBodies - 1
            Set swBody = vBodies(i)
            Dim vFaces As Variant
            vFaces = swBody.GetFaces()
            If Not IsEmpty(vFaces) Then
                Dim swFace As SldWorks.Face2
                For j = 0 To UBound(vFaces)
                    Set swFace = vFaces(j)
                    Dim vMat As Variant
                    vMat = swFace.MaterialPropertyValues
                    If IsArray(vMat) Then
                        If UBound(vMat) >= 2 Then
                            If numExcludedHues >= UBound(excludedHues) Then
                                ReDim Preserve excludedHues(UBound(excludedHues) + 1000)
                            End If
                            excludedHues(numExcludedHues) = GetHSVHueFromRGB( _
                                CDbl(vMat(0)), CDbl(vMat(1)), CDbl(vMat(2)))
                            numExcludedHues = numExcludedHues + 1
                        End If
                    End If
                Next j
            End If
        Next i
    End If

    ' --- Measure every body once ----------------------------------------------
    currentStep = "Measuring bodies"
    Dim bodyVolume() As Double, bodyArea() As Double, bodyFaces() As Long
    Dim bodyCX() As Double, bodyCY() As Double, bodyCZ() As Double
    ReDim bodyVolume(totalBodies - 1)
    ReDim bodyArea(totalBodies - 1)
    ReDim bodyFaces(totalBodies - 1)
    ReDim bodyCX(totalBodies - 1)
    ReDim bodyCY(totalBodies - 1)
    ReDim bodyCZ(totalBodies - 1)

    Dim vProps As Variant
    For i = 0 To totalBodies - 1
        Set swBody = vBodies(i)
        vProps = swBody.GetMassProperties(1#)
        If IsArray(vProps) Then
            bodyCX(i) = vProps(0)
            bodyCY(i) = vProps(1)
            bodyCZ(i) = vProps(2)
            bodyVolume(i) = vProps(3)
            bodyArea(i) = vProps(4)
        End If
        bodyFaces(i) = swBody.GetFaceCount()
    Next i

    ' Vertex signatures are expensive, so they are built only for the bodies that
    ' actually reach the shape fallback, and kept once built.
    Dim sigCache() As Variant
    Dim sigBuilt() As Boolean
    ReDim sigCache(totalBodies - 1)
    ReDim sigBuilt(totalBodies - 1)

    ' Decide which comparison engine to use. The kernel test is exact; the
    ' invariant test is the 0.5.18 fallback for builds that do not expose it.
    currentStep = "Selecting comparison engine"
    Dim useKernel As Boolean
    useKernel = False
    If USE_KERNEL_COMPARISON Then useKernel = KernelComparisonWorks(vBodies(0))

    Dim bodyJ1() As Double, bodyJ2() As Double, bodyJ3() As Double
    ReDim bodyJ1(totalBodies - 1)
    ReDim bodyJ2(totalBodies - 1)
    ReDim bodyJ3(totalBodies - 1)
    Dim bodyEdges() As Long
    ReDim bodyEdges(totalBodies - 1)

    If Not useKernel Then
        currentStep = "Computing shape invariants (fallback engine)"
        For i = 0 To totalBodies - 1
            Set swBody = vBodies(i)
            bodyEdges(i) = swBody.GetEdgeCount()
            ComputeInvariants swModel, swBody, bodyJ1(i), bodyJ2(i), bodyJ3(i)
        Next i
    End If

    tMeasured = Timer

    ' --- Group ----------------------------------------------------------------
    currentStep = "Grouping bodies"
    Dim groupRep() As Long
    Dim bodyGroup() As Long
    ReDim groupRep(totalBodies - 1)
    ReDim bodyGroup(totalBodies - 1)

    Dim numGroups As Long
    numGroups = 0

    Dim comparisons As Long
    Dim gateRejects As Long
    Dim notCoincident As Long
    Dim shapeMerges As Long
    Dim mirrorSplits As Long
    Dim kernelMatched As Long
    comparisons = 0
    gateRejects = 0
    notCoincident = 0
    shapeMerges = 0
    mirrorSplits = 0
    kernelMatched = 0

    ' Bucket groups by face count. Bodies with different face counts can never
    ' match, so this keeps the expensive comparison off almost every pair.
    Dim buckets As Object
    Set buckets = CreateObject("Scripting.Dictionary")

    For i = 0 To totalBodies - 1
        Set swBody = vBodies(i)

        Dim bucketKey As String
        bucketKey = CStr(bodyFaces(i))

        Dim matched As Boolean
        Dim groupIndex As Long
        matched = False
        groupIndex = -1

        If buckets.Exists(bucketKey) Then
            Dim bucket As Collection
            Set bucket = buckets(bucketKey)
            For k = 1 To bucket.Count
                Dim g As Long
                Dim rep As Long
                g = bucket(k)
                rep = groupRep(g)

                ' Cheap size gate first. Identical bodies agree on volume and
                ' area to far better than this, so it never rejects a true match.
                If SimilarSize(bodyVolume(i), bodyVolume(rep), _
                               bodyArea(i), bodyArea(rep)) Then
                    comparisons = comparisons + 1

                    Dim swRepBody As SldWorks.Body2
                    Set swRepBody = vBodies(rep)

                    Dim isSame As Boolean
                    Dim verdict As Long
                    If useKernel Then
                        verdict = CompareBodies(swBody, swRepBody)
                        If verdict = 1 Then
                            isSame = True
                            kernelMatched = kernelMatched + 1
                        ElseIf verdict = -1 Then
                            ' Same shape, opposite hand.
                            isSame = MATCH_MIRRORED
                            If Not isSame Then
                                mirrorSplits = mirrorSplits + 1
                                If SHOW_DIAGNOSTICS Then
                                    Debug.Print "mirror image, kept apart: " & swBody.Name & _
                                                "  vs  " & swRepBody.Name
                                End If
                            End If
                        Else
                            isSame = False
                        End If
                    Else
                        verdict = 0
                        isSame = InvariantsMatch(bodyJ1(i), bodyJ2(i), bodyJ3(i), _
                                                 bodyJ1(rep), bodyJ2(rep), bodyJ3(rep), _
                                                 bodyEdges(i), bodyEdges(rep), _
                                                 bodyVolume(i))
                    End If

                    ' The shape fallback runs only when the kernel found no
                    ' transform at all. A mirror we have just chosen to keep
                    ' apart must not reach it: distances from the centre of mass
                    ' are identical either side of a mirror, so it would undo the
                    ' decision immediately.
                    If Not isSame And useKernel And verdict = 0 Then
                        ' The kernel says these are different shapes. Its answer
                        ' carries a tolerance, though, and copies of one imported
                        ' part re-derived by separate features land outside it.
                        ' Ask the geometry directly before accepting the verdict.
                        notCoincident = notCoincident + 1

                        If SHOW_DIAGNOSTICS Then
                            Debug.Print "not coincident: " & swBody.Name & _
                                        "  vs  " & swRepBody.Name & _
                                        "   dVolume=" & _
                                        Format(RelDiff(bodyVolume(i), bodyVolume(rep)), "0.00E+00") & _
                                        "   dArea=" & _
                                        Format(RelDiff(bodyArea(i), bodyArea(rep)), "0.00E+00")
                        End If

                        If SHAPE_FALLBACK Then
                            If Not sigBuilt(i) Then
                                sigCache(i) = VertexSignature(swBody, _
                                    bodyCX(i), bodyCY(i), bodyCZ(i))
                                sigBuilt(i) = True
                            End If
                            If Not sigBuilt(rep) Then
                                sigCache(rep) = VertexSignature(swRepBody, _
                                    bodyCX(rep), bodyCY(rep), bodyCZ(rep))
                                sigBuilt(rep) = True
                            End If

                            If SignaturesMatch(sigCache(i), sigCache(rep), SHAPE_TOLERANCE) Then
                                isSame = True
                                shapeMerges = shapeMerges + 1
                                If SHOW_DIAGNOSTICS Then
                                    Debug.Print "   -> merged by shape: every vertex within " & _
                                                Format(SHAPE_TOLERANCE * 1000000#, "0.0") & " micron"
                                End If
                            End If
                        End If
                    End If

                    If isSame Then
                        matched = True
                        groupIndex = g
                        Exit For
                    End If
                Else
                    gateRejects = gateRejects + 1
                End If
            Next k
        End If

        If Not matched Then
            groupRep(numGroups) = i
            groupIndex = numGroups
            numGroups = numGroups + 1
            If Not buckets.Exists(bucketKey) Then
                Dim newBucket As Collection
                Set newBucket = New Collection
                buckets.Add bucketKey, newBucket
            End If
            buckets(bucketKey).Add groupIndex
        End If

        bodyGroup(i) = groupIndex
    Next i

    tGrouped = Timer

    ' --- Per-body detail, for working out why two bodies did or did not match --
    If SHOW_DIAGNOSTICS Then
        Debug.Print "=== Apply Unique Colours " & MACRO_VERSION & " - body detail ==="
        Debug.Print "engine: " & IIf(useKernel, "geometry kernel", "shape invariants") & _
                    ", bodies: " & totalBodies & ", groups: " & numGroups
        Debug.Print "group" & vbTab & "faces" & vbTab & "edges" & vbTab & _
                    "volume_m3" & vbTab & vbTab & "area_m2" & vbTab & vbTab & "name"
        For i = 0 To totalBodies - 1
            Set swBody = vBodies(i)
            Debug.Print bodyGroup(i) & vbTab & _
                        bodyFaces(i) & vbTab & _
                        swBody.GetEdgeCount() & vbTab & _
                        Format(bodyVolume(i), "0.000000000000") & vbTab & _
                        Format(bodyArea(i), "0.000000000000") & vbTab & _
                        swBody.Name
        Next i
        Debug.Print "=== end body detail ==="
    End If

    ' --- Build the palette ----------------------------------------------------
    currentStep = "Generating colours"
    Dim groupHues() As Double
    Dim groupLayer() As Long
    BuildPalette numGroups, excludedHues, numExcludedHues, groupHues, groupLayer

    ' --- Apply, one call per group rather than per body -----------------------
    currentStep = "Configuring display state"
    Dim swDispStateSetts As SldWorks.DisplayStateSetting
    Set swDispStateSetts = swModel.Extension.GetDisplayStateSetting( _
        swDisplayStateOpts_e.swThisDisplayState)

    If swDispStateSetts Is Nothing Then
        If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
        MsgBox "SOLIDWORKS did not return a display state setting for this document.", _
               vbInformation
        Exit Sub
    End If

    swDispStateSetts.Option = swDisplayStateOpts_e.swThisDisplayState

    currentStep = "Applying colours"
    Dim groupMembers() As Collection
    ReDim groupMembers(numGroups - 1)
    For j = 0 To numGroups - 1
        Set groupMembers(j) = New Collection
    Next j
    For i = 0 To totalBodies - 1
        groupMembers(bodyGroup(i)).Add vBodies(i)
    Next i

    Dim colouredBodies As Long
    colouredBodies = 0

    Dim groupColour() As Long
    ReDim groupColour(numGroups - 1)

    For j = 0 To numGroups - 1
        Dim rgbArr As Variant
        rgbArr = ColourAtLuminance(groupHues(j), LayerLuminance(groupLayer(j)))
        groupColour(j) = RGB(ToByte(rgbArr(0)), ToByte(rgbArr(1)), ToByte(rgbArr(2)))

        For k = 1 To groupMembers(j).Count
            Set swBody = groupMembers(j)(k)

            Dim applied As Boolean
            applied = ApplyColourToEntity(swModel, swDispStateSetts, _
                                          swBody, groupColour(j))
            If Not applied Then
                ' Display states unavailable on this build: write the colour
                ' straight onto the body instead. It will not be scoped to a
                ' display state, but a colour in every display state beats none.
                applied = WriteBodyColourDirect(swBody, rgbArr)
            End If

            If applied Then colouredBodies = colouredBodies + 1
        Next k

        If SHOW_DIAGNOSTICS Then
            Debug.Print "Group " & j & _
                        "  bodies=" & groupMembers(j).Count & _
                        "  faces=" & bodyFaces(groupRep(j)) & _
                        "  volume=" & Format(bodyVolume(groupRep(j)), "0.000000000") & _
                        "  hue=" & Format(groupHues(j), "0.0") & _
                        "  layer=" & groupLayer(j)
        End If
    Next j

    tDone = Timer

    ' Check the colours actually landed rather than trusting the apply calls.
    ' Timed separately so it does not distort the apply figure.
    Dim verified As Long
    Dim tVerified As Single
    verified = -1
    If SHOW_DIAGNOSTICS Then
        currentStep = "Verifying colours"
        verified = 0
        For i = 0 To totalBodies - 1
            Set swBody = vBodies(i)
            If ColourTookEffect(swModel, swDispStateSetts, swBody, _
                                groupColour(bodyGroup(i))) Then
                verified = verified + 1
            End If
        Next i
    End If
    tVerified = Timer

    ' Redraw before the dialog, otherwise the colours only appear once it is
    ' dismissed and the run looks like it did nothing.
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
    swModel.GraphicsRedraw2

    Dim msg As String
    msg = "Applied unique colours to bodies in active display state" & vbCrLf & _
          "Total Bodies: " & totalBodies & vbCrLf & _
          "Unique Bodies: " & numGroups & vbCrLf & vbCrLf & _
          "Macro Version: " & MACRO_VERSION

    If SHOW_DIAGNOSTICS Then
        msg = msg & vbCrLf & vbCrLf & _
              "Engine: " & IIf(useKernel, "geometry kernel", "shape invariants") & vbCrLf & _
              "Kernel comparisons: " & comparisons & vbCrLf & _
              "  matched outright: " & kernelMatched & vbCrLf & _
              "  matched as mirror: " & mirrorSplits & vbCrLf & _
              "  no match: " & notCoincident & vbCrLf & _
              "  of those, rescued by shape: " & shapeMerges & vbCrLf & _
              "Skipped by size gate: " & gateRejects & vbCrLf & _
              "Coloured: " & colouredBodies & " of " & totalBodies & vbCrLf & _
              "Verified: " & verified & " of " & totalBodies & vbCrLf & _
              "Measure: " & Format(tMeasured - tStart, "0.00") & " s" & vbCrLf & _
              "Group:   " & Format(tGrouped - tMeasured, "0.00") & " s" & vbCrLf & _
              "Apply:   " & Format(tDone - tGrouped, "0.00") & " s" & vbCrLf & _
              "Verify:  " & Format(tVerified - tDone, "0.00") & " s" & vbCrLf & _
              "Total:   " & Format(tDone - tStart, "0.00") & " s (excluding verify)"
    End If

    MsgBox msg, vbInformation
    Exit Sub

ErrorHandler:
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
    MsgBox "ERROR in ProcessPart at step [" & currentStep & "]" & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & _
           "Error " & Err.Number, vbCritical
End Sub

'--- Assembly processing ------------------------------------------------------

Sub ProcessAssembly(swModel As SldWorks.ModelDoc2)

    Dim currentStep As String
    Dim swView As SldWorks.ModelView
    On Error GoTo ErrorHandler

    Dim tStart As Single, tDone As Single
    tStart = Timer

    currentStep = "Loading components"
    Dim swAssy As SldWorks.AssemblyDoc
    Set swAssy = swModel

    Dim vComps As Variant
    vComps = swAssy.GetComponents(False)

    If IsEmpty(vComps) Then
        MsgBox "No components found in the active assembly.", vbInformation
        Exit Sub
    End If

    Dim totalComps As Long
    totalComps = UBound(vComps) + 1

    Set swView = swModel.ActiveView
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = False

    Dim i As Long, j As Long, k As Long

    ' Existing component colours are deliberately NOT collected here. Component
    ' level colour is exactly what this macro writes, so reading it back would
    ' make every second run treat its own previous output as a colour to avoid.

    currentStep = "Grouping components"
    Dim groupKeys As Object
    Set groupKeys = CreateObject("Scripting.Dictionary")

    Dim compGroup() As Long
    ReDim compGroup(totalComps - 1)

    Dim numGroups As Long
    Dim skippedComps As Long
    numGroups = 0
    skippedComps = 0

    Dim swComp As SldWorks.Component2
    For i = 0 To totalComps - 1
        Set swComp = vComps(i)
        Dim path As String
        path = LCase(swComp.GetPathName())

        If Right(path, 7) = ".sldprt" Then
            Dim compKey As String
            compKey = path & "::" & LCase(swComp.ReferencedConfiguration)

            If Not groupKeys.Exists(compKey) Then
                groupKeys.Add compKey, numGroups
                numGroups = numGroups + 1
            End If
            compGroup(i) = groupKeys(compKey)
        Else
            compGroup(i) = -1
            skippedComps = skippedComps + 1
        End If
    Next i

    currentStep = "Generating colours"
    Dim excludedHues() As Double
    ReDim excludedHues(0)
    Dim groupHues() As Double
    Dim groupLayer() As Long
    BuildPalette numGroups, excludedHues, 0, groupHues, groupLayer

    currentStep = "Configuring display state"
    Dim swDispStateSetts As SldWorks.DisplayStateSetting
    Set swDispStateSetts = swModel.Extension.GetDisplayStateSetting( _
        swDisplayStateOpts_e.swThisDisplayState)

    If swDispStateSetts Is Nothing Then
        If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
        MsgBox "SOLIDWORKS did not return a display state setting for this document.", _
               vbInformation
        Exit Sub
    End If

    swDispStateSetts.Option = swDisplayStateOpts_e.swThisDisplayState
    swDispStateSetts.PartLevel = False

    If numGroups = 0 Then
        If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
        MsgBox "No bottom-level parts found to colour." & vbCrLf & _
               "Skipped Subassemblies: " & skippedComps, vbInformation
        Exit Sub
    End If

    currentStep = "Applying colours"
    Dim groupMembers() As Collection
    ReDim groupMembers(numGroups - 1)
    For j = 0 To numGroups - 1
        Set groupMembers(j) = New Collection
    Next j

    Dim coloredComps As Long
    coloredComps = 0
    For i = 0 To totalComps - 1
        If compGroup(i) >= 0 Then
            groupMembers(compGroup(i)).Add vComps(i)
            coloredComps = coloredComps + 1
        End If
    Next i

    Dim colouredNow As Long
    colouredNow = 0

    For j = 0 To numGroups - 1
        Dim rgbArr As Variant
        rgbArr = ColourAtLuminance(groupHues(j), LayerLuminance(groupLayer(j)))

        Dim colourValue As Long
        colourValue = RGB(ToByte(rgbArr(0)), ToByte(rgbArr(1)), ToByte(rgbArr(2)))

        For k = 1 To groupMembers(j).Count
            Set swComp = groupMembers(j)(k)

            Dim applied As Boolean
            applied = ApplyColourToEntity(swModel, swDispStateSetts, _
                                          swComp, colourValue)
            If Not applied Then
                applied = WriteComponentColourDirect(swComp, rgbArr)
            End If

            If applied Then colouredNow = colouredNow + 1
        Next k
    Next j

    ' Redraw before the dialog, otherwise the colours only appear once it is
    ' dismissed and the run looks like it did nothing.
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
    swModel.GraphicsRedraw2
    tDone = Timer

    Dim msg As String
    msg = "Applied unique colours to components in active display state" & vbCrLf & _
          "Total Components: " & coloredComps & vbCrLf & _
          "Unique Components: " & numGroups & vbCrLf & _
          "Skipped Subassemblies: " & skippedComps & vbCrLf & vbCrLf & _
          "Macro Version: " & MACRO_VERSION

    If SHOW_DIAGNOSTICS Then
        msg = msg & vbCrLf & vbCrLf & _
              "Coloured: " & colouredNow & " of " & coloredComps & vbCrLf & _
              "Total: " & Format(tDone - tStart, "0.00") & " s"
    End If

    MsgBox msg, vbInformation
    Exit Sub

ErrorHandler:
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
    MsgBox "ERROR in ProcessAssembly at step [" & currentStep & "]" & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & _
           "Error " & Err.Number, vbCritical
End Sub

'--- Body comparison ----------------------------------------------------------

' True if this SOLIDWORKS build exposes IBody2::GetCoincidenceTransform2.
Function KernelComparisonWorks(ByVal swBody As SldWorks.Body2) As Boolean
    On Error GoTo NotSupported
    Dim swXform As SldWorks.MathTransform
    Dim result As Boolean
    result = swBody.GetCoincidenceTransform2(swBody, swXform)
    KernelComparisonWorks = True
    Exit Function
NotSupported:
    KernelComparisonWorks = False
End Function

' Asks the kernel whether one body can be moved onto the other, and reports what
' kind of movement it took. This is the kernel's own answer, so the match itself
' is exact - no tolerance is involved and no feature size is too small.
'
' Returns:
'    1  the same shape, same handedness
'    0  no transform exists between them
'   -1  the same shape, but only as a mirror image
Function CompareBodies(ByVal bodyA As SldWorks.Body2, ByVal bodyB As SldWorks.Body2) As Long
    On Error GoTo Failed

    CompareBodies = 0

    Dim swXform As SldWorks.MathTransform
    Dim found As Boolean

    ' Take the result into a Boolean before testing it. Writing this as
    ' "If Not bodyA.GetCoincidenceTransform2(...) Then" - as 0.8.0 did, to get an
    ' early exit - made every comparison report no match and silently disabled
    ' the whole kernel engine. Applying Not straight to the call does not invert
    ' what comes back the way it appears to.
    found = bodyA.GetCoincidenceTransform2(bodyB, swXform)

    If found Then
        If TransformIsReflection(swXform) Then
            CompareBodies = -1
        Else
            CompareBodies = 1
        End If
    End If
    Exit Function

Failed:
    CompareBodies = 0
End Function

' True when the transform turns a body inside out rather than simply moving it.
'
' The kernel counts a mirror image as coincident and returns a transform that
' includes the reflection, so the two cases can be told apart exactly: the
' determinant of a rotation matrix is always +1, and a reflection makes it -1.
' No handedness heuristic is needed, which is what made this awkward in 0.5.13.
Function TransformIsReflection(ByVal swXform As SldWorks.MathTransform) As Boolean
    On Error GoTo Failed

    TransformIsReflection = False
    If swXform Is Nothing Then Exit Function

    Dim v As Variant
    v = swXform.ArrayData
    If Not IsArray(v) Then Exit Function
    If UBound(v) < 8 Then Exit Function

    ' ArrayData holds the 3x3 rotation in elements 0 to 8, translation in 9 to 11
    ' and a scale factor in 12.
    Dim det As Double
    det = CDbl(v(0)) * (CDbl(v(4)) * CDbl(v(8)) - CDbl(v(5)) * CDbl(v(7))) _
        - CDbl(v(1)) * (CDbl(v(3)) * CDbl(v(8)) - CDbl(v(5)) * CDbl(v(6))) _
        + CDbl(v(2)) * (CDbl(v(3)) * CDbl(v(7)) - CDbl(v(4)) * CDbl(v(6)))

    TransformIsReflection = (det < 0#)
    Exit Function

Failed:
    TransformIsReflection = False
End Function

' Cheap gate applied before the expensive comparison. Its only job is to skip
' pairs that obviously cannot match, so it is deliberately generous.
' See SIZE_GATE_TOLERANCE in the settings block for why it is set where it is.
Function SimilarSize(ByVal volA As Double, ByVal volB As Double, _
                     ByVal areaA As Double, ByVal areaB As Double) As Boolean
    Dim volTol As Double, areaTol As Double
    volTol = Abs(volB) * SIZE_GATE_TOLERANCE + 0.000000000001
    areaTol = Abs(areaB) * SIZE_GATE_TOLERANCE + 0.000000001
    SimilarSize = (Abs(volA - volB) <= volTol) And (Abs(areaA - areaB) <= areaTol)
End Function

' Relative difference between two measurements, for diagnostics.
Function RelDiff(ByVal a As Double, ByVal b As Double) As Double
    If b = 0 Then
        RelDiff = 0
    Else
        RelDiff = Abs(a - b) / Abs(b)
    End If
End Function

'--- Shape comparison ---------------------------------------------------------

' The sorted distances of every vertex from the body's centre of mass.
'
' Moving or rotating a body changes neither the centre of mass relative to the
' body nor the distances to it, so two instances of one part sitting anywhere in
' the model produce the same list. Moving a feature does change it, because the
' vertices around that feature move and the centre of mass shifts with them.
'
' The distances are held individually and compared term by term rather than
' summed. Earlier attempts at this idea (0.5.12) added the distances together
' with fourth-power weighting, which buried small differences in accumulated
' rounding and made the result depend on how many vertices a body happened to
' have.
Function VertexSignature(ByVal swBody As SldWorks.Body2, ByVal cx As Double, _
                         ByVal cy As Double, ByVal cz As Double) As Variant
    On Error GoTo Failed

    Dim vVerts As Variant
    vVerts = swBody.GetVertices()
    If IsEmpty(vVerts) Then GoTo Failed

    Dim n As Long
    n = UBound(vVerts) + 1
    If n < 1 Then GoTo Failed

    Dim d() As Double
    ReDim d(n - 1)

    Dim i As Long
    Dim swVertex As SldWorks.Vertex
    Dim vPt As Variant
    Dim dx As Double, dy As Double, dz As Double

    For i = 0 To n - 1
        Set swVertex = vVerts(i)
        vPt = swVertex.GetPoint()
        dx = vPt(0) - cx
        dy = vPt(1) - cy
        dz = vPt(2) - cz
        d(i) = Sqr(dx * dx + dy * dy + dz * dz)
    Next i

    SortDoubles d
    VertexSignature = d
    Exit Function

Failed:
    VertexSignature = Empty
End Function

' Shell sort. The distances must be in a canonical order before two bodies can be
' compared term by term, since the kernel gives no guarantee about vertex order.
Sub SortDoubles(ByRef d() As Double)
    Dim n As Long, gap As Long, i As Long, j As Long
    Dim tmp As Double

    n = UBound(d) + 1
    gap = n \ 2

    Do While gap > 0
        For i = gap To n - 1
            tmp = d(i)
            j = i
            Do While j >= gap
                If d(j - gap) > tmp Then
                    d(j) = d(j - gap)
                    j = j - gap
                Else
                    Exit Do
                End If
            Loop
            d(j) = tmp
        Next i
        gap = gap \ 2
    Loop
End Sub

' Two signatures match when they hold the same number of vertices and every
' corresponding distance agrees within the tolerance.
Function SignaturesMatch(ByVal sigA As Variant, ByVal sigB As Variant, _
                         ByVal tol As Double) As Boolean
    SignaturesMatch = False

    If IsEmpty(sigA) Or IsEmpty(sigB) Then Exit Function
    If Not IsArray(sigA) Or Not IsArray(sigB) Then Exit Function
    If UBound(sigA) <> UBound(sigB) Then Exit Function

    Dim i As Long
    For i = 0 To UBound(sigA)
        If Abs(sigA(i) - sigB(i)) > tol Then Exit Function
    Next i

    SignaturesMatch = True
End Function

'--- Fallback engine (0.5.18 shape invariants) --------------------------------

' Principal moments of inertia normalised by mass, reduced to the three tensor
' invariants. Used only when the kernel comparison is unavailable.
Sub ComputeInvariants(ByVal swModel As SldWorks.ModelDoc2, ByVal swBody As SldWorks.Body2, _
                      ByRef j1 As Double, ByRef j2 As Double, ByRef j3 As Double)
    On Error GoTo Failed

    j1 = 0: j2 = 0: j3 = 0

    Dim swMass As SldWorks.MassProperty
    Set swMass = swModel.Extension.CreateMassProperty

    ' Force system units so results do not shift with the document's unit scheme.
    ' Older builds do not expose this, which is not fatal - skip it and carry on.
    On Error Resume Next
    swMass.UseSystemUnits = True
    Err.Clear
    On Error GoTo Failed

    Dim bArray(0) As Object
    Set bArray(0) = swBody

    If Not swMass.AddBodies(bArray) Then Exit Sub

    Dim localMass As Double
    localMass = swMass.Mass

    Dim Prin As Variant
    Prin = swMass.PrincipleMomentsOfInertia
    If Not IsArray(Prin) Then Exit Sub

    Dim Px As Double, Py As Double, Pz As Double
    If localMass > 0.000000001 Then
        Px = Prin(0) / localMass
        Py = Prin(1) / localMass
        Pz = Prin(2) / localMass
    Else
        Px = Prin(0): Py = Prin(1): Pz = Prin(2)
    End If

    j1 = Px + Py + Pz
    j2 = (Px * Py) + (Px * Pz) + (Py * Pz)
    j3 = Px * Py * Pz
    Exit Sub

Failed:
    j1 = 0: j2 = 0: j3 = 0
End Sub

Function InvariantsMatch(ByVal j1A As Double, ByVal j2A As Double, ByVal j3A As Double, _
                         ByVal j1B As Double, ByVal j2B As Double, ByVal j3B As Double, _
                         ByVal edgesA As Long, ByVal edgesB As Long, _
                         ByVal volume As Double) As Boolean
    If volume > 0.00001 Then
        If edgesA <> edgesB Then
            InvariantsMatch = False
            Exit Function
        End If
    End If

    Dim t1 As Double, t2 As Double, t3 As Double
    t1 = Abs(j1B) * 0.0001 + 0.000000005
    t2 = Abs(j2B) * 0.0001 + 0.000000005
    t3 = Abs(j3B) * 0.0001 + 0.000000005

    InvariantsMatch = (Abs(j1A - j1B) <= t1) And _
                      (Abs(j2A - j2B) <= t2) And _
                      (Abs(j3A - j3B) <= t3)
End Function

'--- Palette ------------------------------------------------------------------

' Spreads groups evenly around the hue wheel, in layers of differing brightness
' so that neighbouring hues are further separated. Each group's final hue is
' checked against every hue already assigned, so two groups can never end up
' sharing a colour - which the pre-0.6.0 collision avoidance allowed.
Sub BuildPalette(ByVal numGroups As Long, excludedHues() As Double, ByVal numExcludedHues As Long, _
                 ByRef groupHues() As Double, ByRef groupLayer() As Long)

    If numGroups <= 0 Then
        ReDim groupHues(0)
        ReDim groupLayer(0)
        Exit Sub
    End If

    ReDim groupHues(numGroups - 1)
    ReDim groupLayer(numGroups - 1)

    Dim totalLayers As Long, itemsPerLayer As Long
    If numGroups <= ITEMS_PER_LAYER Then
        totalLayers = 1
        itemsPerLayer = numGroups
    Else
        totalLayers = Int((numGroups - 1) / ITEMS_PER_LAYER) + 1
        If totalLayers > LAYER_COUNT Then totalLayers = LAYER_COUNT
        itemsPerLayer = Int((numGroups - 1) / totalLayers) + 1
    End If

    Dim assignedHues() As Double
    ReDim assignedHues(numGroups - 1)

    Dim i As Long
    For i = 0 To numGroups - 1
        Dim layerIdx As Long
        layerIdx = Int(i / itemsPerLayer)
        If layerIdx > LAYER_COUNT - 1 Then layerIdx = LAYER_COUNT - 1

        Dim indexInLayer As Long
        indexInLayer = i Mod itemsPerLayer

        Dim currentLayerItems As Long
        If layerIdx = totalLayers - 1 Then
            currentLayerItems = numGroups - (layerIdx * itemsPerLayer)
        Else
            currentLayerItems = itemsPerLayer
        End If
        If currentLayerItems < 1 Then currentLayerItems = 1

        Dim targetHue As Double
        If currentLayerItems > 1 Then
            targetHue = indexInLayer * (360# / currentLayerItems)
            targetHue = targetHue + layerIdx * ((360# / currentLayerItems) / totalLayers)
        Else
            targetHue = layerIdx * 45#
        End If
        targetHue = WrapHue(targetHue)

        targetHue = FindSafeHue(targetHue, excludedHues, numExcludedHues, assignedHues, i)

        assignedHues(i) = targetHue
        groupHues(i) = targetHue
        groupLayer(i) = layerIdx
    Next i
End Sub

' Perceived brightness for each layer. Nothing is darker than 0.25 or lighter
' than 0.74, which keeps every colour readable against a shaded model.
Function LayerLuminance(ByVal layerIdx As Long) As Double
    Select Case layerIdx
        Case 0: LayerLuminance = 0.58
        Case 1: LayerLuminance = 0.34
        Case 2: LayerLuminance = 0.72
        Case 3: LayerLuminance = 0.46
        Case 4: LayerLuminance = 0.25
        Case 5: LayerLuminance = 0.65
        Case Else: LayerLuminance = 0.4
    End Select
End Function

' Steps a hue away from colours already in the model and from hues already given
' to other groups. If nothing is clear, the evenly spaced hue is kept: clashing
' with a manual face colour is far less confusing than two groups matching.
Function FindSafeHue(ByVal targetHue As Double, excludedHues() As Double, ByVal numExcludedHues As Long, _
                     assignedHues() As Double, ByVal numAssigned As Long) As Double
    Dim attempt As Long
    Dim candidate As Double

    For attempt = 0 To 23
        candidate = WrapHue(targetHue + attempt * 15#)
        If Not HueClashes(candidate, excludedHues, numExcludedHues, 10#) Then
            If Not HueClashes(candidate, assignedHues, numAssigned, 6#) Then
                FindSafeHue = candidate
                Exit Function
            End If
        End If
    Next attempt

    FindSafeHue = targetHue
End Function

Function HueClashes(ByVal hue As Double, hues() As Double, ByVal count As Long, _
                    ByVal minGap As Double) As Boolean
    Dim i As Long
    Dim d As Double
    For i = 0 To count - 1
        d = Abs(hue - hues(i))
        If d > 180# Then d = 360# - d
        If d < minGap Then
            HueClashes = True
            Exit Function
        End If
    Next i
    HueClashes = False
End Function

Function WrapHue(ByVal hue As Double) As Double
    Dim h As Double
    h = hue
    Do While h >= 360#
        h = h - 360#
    Loop
    Do While h < 0#
        h = h + 360#
    Loop
    WrapHue = h
End Function

' Builds a colour of a given hue at a specific perceived brightness.
'
' Red, green and blue contribute very unequally to how bright a colour looks
' (roughly 21%, 72% and 7%). Plain HSV ignores this, so a "full brightness" blue
' reads as almost black beside a full brightness yellow. Because every RGB
' channel scales linearly with V, the brightness of a hue at full saturation is
' fixed - so darkening is done with V, and lightening by washing in white via S.
Function ColourAtLuminance(ByVal hue As Double, ByVal targetLum As Double) As Variant
    Dim pureRGB As Variant
    pureRGB = GetRGBFromHSV(hue, 1#, 1#)

    Dim lumPure As Double
    lumPure = 0.2126 * pureRGB(0) + 0.7152 * pureRGB(1) + 0.0722 * pureRGB(2)

    Dim S As Double, V As Double
    If targetLum <= lumPure Then
        S = 1#
        V = targetLum / lumPure
    Else
        V = 1#
        If lumPure < 0.999999 Then
            S = (1# - targetLum) / (1# - lumPure)
        Else
            S = 1#
        End If
    End If

    If S < 0# Then S = 0#
    If S > 1# Then S = 1#
    If V < 0# Then V = 0#
    If V > 1# Then V = 1#

    ColourAtLuminance = GetRGBFromHSV(hue, S, V)
End Function

'--- Colour application -------------------------------------------------------

' Applies one colour to a single body or component.
'
' This is deliberately one entity at a time. Passing several entities at once in
' DisplayStateSetting::Entities reports success but silently applies nothing, so
' 0.6.0 left every multi-body group - all the repeated hardware - uncoloured.
' Returns False if the display state route is unavailable, so the caller can
' fall back to setting the entity's material properties directly.
'
' 0.9.0 reused one appearance object across a group to save reading it back per
' body. Measured, that saved nothing - 119 ms per body before, 117 ms after -
' because the read costs nothing to begin with. The complication was removed.
Function ApplyColourToEntity(ByVal swModel As SldWorks.ModelDoc2, _
                             ByVal swDispStateSetts As SldWorks.DisplayStateSetting, _
                             ByVal swEntity As Object, ByVal colourValue As Long) As Boolean
    On Error GoTo Failed

    Dim ents(0) As Object
    Set ents(0) = swEntity
    swDispStateSetts.Entities = ents

    Dim vAppearances As Variant
    vAppearances = swModel.Extension.DisplayStateSpecMaterialPropertyValues(swDispStateSetts)

    If IsEmpty(vAppearances) Then
        ApplyColourToEntity = False
        Exit Function
    End If

    Dim appearSet As SldWorks.AppearanceSetting
    Set appearSet = vAppearances(0)
    appearSet.Color = colourValue
    appearSet.Diffuse = 1#
    appearSet.Specular = 0.5
    appearSet.Luminous = 0#

    Dim newAppearances(0) As SldWorks.AppearanceSetting
    Set newAppearances(0) = appearSet

    swModel.Extension.DisplayStateSpecMaterialPropertyValues(swDispStateSetts) = newAppearances
    ApplyColourToEntity = True
    Exit Function

Failed:
    ApplyColourToEntity = False
End Function

' Writes the colour straight onto a body, bypassing the display state machinery.
Function WriteBodyColourDirect(ByVal swBody As SldWorks.Body2, _
                               ByVal rgbArr As Variant) As Boolean
    On Error GoTo Failed
    swBody.MaterialPropertyValues2 = BuildMaterialArray(rgbArr)
    WriteBodyColourDirect = True
    Exit Function
Failed:
    WriteBodyColourDirect = False
End Function

' The same for an assembly component, scoped to the active configuration.
Function WriteComponentColourDirect(ByVal swComp As SldWorks.Component2, _
                                    ByVal rgbArr As Variant) As Boolean
    On Error GoTo Failed
    swComp.SetMaterialPropertyValues2 BuildMaterialArray(rgbArr), _
        swInConfigurationOpts_e.swThisConfiguration, Empty
    WriteComponentColourDirect = True
    Exit Function
Failed:
    WriteComponentColourDirect = False
End Function

' Reads the colour back off an entity and reports whether it actually took.
'
' The apply call can report success without having changed anything - that is
' exactly how the 0.6.0 batching fault went unnoticed - so diagnostic runs check
' the result rather than trusting it. This is what makes it safe to experiment
' with faster ways of applying colour.
Function ColourTookEffect(ByVal swModel As SldWorks.ModelDoc2, _
                          ByVal swDispStateSetts As SldWorks.DisplayStateSetting, _
                          ByVal swEntity As Object, ByVal colourValue As Long) As Boolean
    On Error GoTo Failed

    ColourTookEffect = False

    Dim ents(0) As Object
    Set ents(0) = swEntity
    swDispStateSetts.Entities = ents

    Dim vAppearances As Variant
    vAppearances = swModel.Extension.DisplayStateSpecMaterialPropertyValues(swDispStateSetts)
    If IsEmpty(vAppearances) Then Exit Function

    Dim appearSet As SldWorks.AppearanceSetting
    Set appearSet = vAppearances(0)
    ColourTookEffect = (appearSet.Color = colourValue)
    Exit Function

Failed:
    ColourTookEffect = False
End Function

Function BuildMaterialArray(ByVal rgbArr As Variant) As Variant
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
    BuildMaterialArray = dMatPrps
End Function

Function ToByte(ByVal value As Double) As Long
    Dim v As Long
    v = Int(value * 255# + 0.5)
    If v < 0 Then v = 0
    If v > 255 Then v = 255
    ToByte = v
End Function

'--- Colour space -------------------------------------------------------------

Function GetHSVHueFromRGB(ByVal R As Double, ByVal G As Double, ByVal B As Double) As Double
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
        GetHSVHueFromRGB = WrapHue(60# * ((G - B) / diff) + 360#)
    ElseIf maxVal = G Then
        GetHSVHueFromRGB = 60# * ((B - R) / diff) + 120#
    Else
        GetHSVHueFromRGB = 60# * ((R - G) / diff) + 240#
    End If
End Function

Function GetRGBFromHSV(ByVal H As Double, ByVal S As Double, ByVal V As Double) As Variant
    Dim C As Double, X As Double, m As Double
    Dim R As Double, G As Double, B As Double
    Dim H_prime As Double

    C = V * S
    H_prime = H / 60#
    X = C * (1# - Abs((H_prime - 2# * Int(H_prime / 2#)) - 1#))

    If H_prime >= 0# And H_prime < 1# Then
        R = C: G = X: B = 0
    ElseIf H_prime >= 1# And H_prime < 2# Then
        R = X: G = C: B = 0
    ElseIf H_prime >= 2# And H_prime < 3# Then
        R = 0: G = C: B = X
    ElseIf H_prime >= 3# And H_prime < 4# Then
        R = 0: G = X: B = C
    ElseIf H_prime >= 4# And H_prime < 5# Then
        R = X: G = 0: B = C
    ElseIf H_prime >= 5# And H_prime <= 6# Then
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
