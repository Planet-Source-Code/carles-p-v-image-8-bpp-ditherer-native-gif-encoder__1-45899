VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPal8bpp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'================================================
' Class:         cPal8bpp.cls [8bpp]
' Author:        (*)
' Dependencies:  cDIB.cls
' Last revision: 2003.05.25
'================================================

' (*)
'
'   From cPalette.cls
'   Copyright © 1999 Steve McMahon
'
'   The Octree Colour Quantisation Code (CreateOptimal) was written by
'   Brian Schimpf Copyright © 1999 Brian Schimpf
'   Note: Added RGB channels weigth coeffs.

Option Explicit

'-- API:

Private Type RGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Type PALETTEENTRY
    peR     As Byte
    peG     As Byte
    peB     As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE256
    palVersion       As Integer
    palNumEntries    As Integer
    palPalEntry(255) As PALETTEENTRY
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As Any) As Long
Private Declare Function GetNearestPaletteIndex Lib "gdi32" (ByVal hPalette As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)

'//

'-- Private Types:

Private Type tNode    'OCT-TREE node struct.
    bIsLeaf            As Boolean ' Leaf flag
    bAddedReduce       As Boolean ' Linked list flag
    vR                 As Long    ' Red Value
    vG                 As Long    ' Green Value
    vB                 As Long    ' Blue Value
    cClrs              As Long    ' Pixel count
    iChildren(1, 1, 1) As Long    ' Child pointers
    iNext              As Long    ' Next reducable node
End Type

'-- Public Enums.:

Public Enum palGreyLevelsCts
    [002_pgLevels] = 0
    [016_pgLevels]
    [256_pgLevels]
End Enum

Public Enum palHalftoneLevelsCts
    [008_phLevels] = 0
    [027_phLevels]
    [064_phLevels]
    [125_phLevels]
    [216_phLevels]
End Enum

'-- Private Property Variables:
Private m_Entries  As Integer
Private m_hPal     As Long
Private m_tPal()   As RGBQUAD

'-- Private Variables (Optimal):
Private aNodes()   As tNode
Private cNodes     As Long
Private nDepth     As Byte
Private TopGarbage As Long
Private cClr       As Long
Private aReduce()  As Long
Private wR         As Single
Private wG         As Single
Private wB         As Single



'========================================================================================
' Class
'========================================================================================

Private Sub Class_Initialize()
    '-- Default
    m_hPal = 0
    m_Entries = 0
End Sub

Private Sub Class_Terminate()
    '-- Destroy palette
    Destroy
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Sub Initialize(Optional ByVal InitialEntries As Integer = 0)
    
    '-- Destroy previous
    Destroy
    
    '-- Redim. palette
    m_Entries = InitialEntries
    ReDim m_tPal(255) As RGBQUAD
End Sub

Public Sub BuildLogicalPalette()

    '-- Build logical palette (-> hPal)
    pvBuildLogicalPalette
End Sub

Public Sub Destroy()

    '-- Delete palette array
    m_Entries = 0
    Erase m_tPal
    
    '-- Destroy logical palette
    If (m_hPal <> 0) Then
        DeleteObject m_hPal
        m_hPal = 0
    End If
End Sub

'//

Public Sub SetPalette(Palette() As Byte)

  Dim nEntries As Integer
    
    '-- Set current number of entries from source array
    nEntries = (UBound(Palette) + 1) \ 4
    Initialize nEntries
    '-- Copy entries
    CopyMemory m_tPal(0), Palette(0), 4 * nEntries
End Sub

Public Sub GetPalette(Palette() As Byte)

    '-- Redim. array and pass palette
    ReDim Palette(4 * m_Entries - 1)
    CopyMemory Palette(0), m_tPal(0), 4 * m_Entries
End Sub

'//

Public Sub ClosestColor(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, cR As Byte, cG As Byte, cb As Byte)
    
    '-- Get closest color
    With m_tPal(GetNearestPaletteIndex(m_hPal, RGB(B, G, R)))
        cR = .R
        cG = .G
        cb = .B
    End With
End Sub

Public Sub ClosestIndex(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, Index As Byte)
    
    '-- Get nearest palette color index
    Index = GetNearestPaletteIndex(m_hPal, RGB(B, G, R))
End Sub

'//

Public Function IsGreyScale() As Boolean
    
  Dim lEnt As Long
    
    '-- R = G = B
    IsGreyScale = (m_Entries > 0)
    For lEnt = 0 To m_Entries - 1
        With m_tPal(lEnt)
            If (.R <> .G Or .R <> .B Or .G <> .B) Then
                IsGreyScale = 0
                Exit For
            End If
        End With
    Next lEnt
End Function

'//

Public Sub CreateGreyScale(ByVal GreyLevels As palGreyLevelsCts)

  Dim lEnt As Long
  Dim lStp As Long
  
    '-- 2, 16, 256 greys palette:
    Initialize
        
    Select Case GreyLevels
      Case [002_pgLevels]: lStp = 255: m_Entries = 2
      Case [016_pgLevels]: lStp = 17: m_Entries = 16
      Case [256_pgLevels]: lStp = 1: m_Entries = 256
    End Select
    
    For lEnt = 0 To 255 \ lStp
        With m_tPal(lEnt)
            .R = lEnt * lStp
            .G = lEnt * lStp
            .B = lEnt * lStp
        End With
    Next lEnt
    
    pvBuildLogicalPalette
End Sub

Public Sub CreateEGA()

    '-- Standard EGA style 16 colour palette:
    Initialize 16
    
    pv16ColourLow8 0
    pv16ColourHigh8 8
    
    pvBuildLogicalPalette
End Sub

Public Sub CreateHalftone(ByVal HalftoneLevels As palHalftoneLevelsCts)

  Dim lStp As Long
  Dim lEnt As Long
  Dim R As Long, G As Long, B As Long
  Dim rA As Long, gA As Long, bA As Long

    '-- Halftone 8, 27, 64, 125, 216 colour palette
    Select Case HalftoneLevels
      Case [008_phLevels]: lStp = &H100: Initialize 8
      Case [027_phLevels]: lStp = &H80: Initialize 27
      Case [064_phLevels]: lStp = &H55: Initialize 64
      Case [125_phLevels]: lStp = &H40: Initialize 125
      Case [216_phLevels]: lStp = &H33: Initialize 216
    End Select
    
    For B = 0 To &H100 Step lStp
        If (B = &H100) Then
            bA = B - 1
          Else
            bA = B
        End If
        For G = 0 To &H100 Step lStp
            If (G = &H100) Then
                gA = G - 1
              Else
                gA = G
            End If
            For R = 0 To &H100 Step lStp
                If (R = &H100) Then
                    rA = R - 1
                  Else
                    rA = R
                End If
                With m_tPal(lEnt)
                    .R = rA
                    .G = gA
                    .B = bA
                End With
                lEnt = lEnt + 1
            Next R
        Next G
    Next B
    
    pvBuildLogicalPalette
End Sub

Public Sub CreateWebsafe()

  Dim lEnt As Long
  Dim R As Long, G As Long, B As Long
  Dim L As Long

    '-- Websafe
    Initialize 256
    
    pv16ColourLow8 0
    pv16ColourHigh8 248
    lEnt = 8
    
    For B = 0 To &HFF Step &H33
        For G = 0 To &HFF Step &H33
            For R = 0 To &HFF Step &H33
            
                '-- Ignore if the output is any combination of 0 and FF
                L = R + G + B
                If (L = 0 Or L = &H2FD) Then
                    '-- ignore
                  ElseIf (L = &H1FE And (R = 0 Or G = 0 Or B = 0)) Then
                    '-- ignore
                  ElseIf (L = &HFF And ((R = 0 And G = 0) Or (R = 0 And B = 0) Or (G = 0 And B = 0))) Then
                    '-- ignore
                  Else
                    '-- add
                    With m_tPal(lEnt)
                        .R = R
                        .G = G
                        .B = B
                    End With
                    lEnt = lEnt + 1
                End If
            Next R
        Next G
    Next B
    '-- Fill the remain entries with grey shades:
    R = 8: G = 8: B = 8
    For lEnt = 217 To 247
        With m_tPal(lEnt)
            .R = R: R = R + 8
            .G = G: G = G + 8
            .B = B: B = B + 8
        End With
    Next lEnt
    
    pvBuildLogicalPalette
End Sub

Public Sub CreateOptimal(oDIB32 As cDIB, _
                         ByVal nColors As Long, _
                         ByVal nLevels As Long, _
                         Optional ByVal RWeight As Single = 1, _
                         Optional ByVal GWeight As Single = 1, _
                         Optional ByVal BWeight As Single = 1)
                         
' Creates an optimal palette with the specified number
' of colors using octree quantisation.
'
' Inputs:
'   - oDIB32 (32-bpp DIB section for which palette is made)
'   - nLevels (Number of levels of color to create)
'   - nColors (Max. number of colors)
'   - RWeight, GWeight, BWeight (Channel wheights [0,1])
    
  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
  Dim x As Long, xMax As Long
  Dim y As Long, yMax As Long
  Dim lEnt As Long
  
    If (oDIB32.BPP = [32_bpp]) Then
  
        '--  R-G-B weights
        wR = RWeight
        wG = GWeight
        wB = BWeight
        
        '-- Map DIB bits
        pvBuildSA tSA, oDIB32
        CopyMemory ByVal VarPtrArray(Bits()), VarPtr(tSA), 4
        
        '-- Allocates initial storage
        ReDim aNodes(1 To 50) As tNode
        ReDim aReduce(1 To 8) As Long
        nDepth = nLevels
        cNodes = 1
        TopGarbage = 0
        cClr = 0
        
        xMax = oDIB32.Width - 1
        yMax = oDIB32.Height - 1
        
        For y = 0 To yMax
            For x = 0 To xMax
                '-- Adds the current pixel to the color octree
                pvAddClr 1, 1, 0, 255, 0, 255, 0, 255, Bits(x, y).R, Bits(x, y).G, Bits(x, y).B
                '-- Combine the levels to get down to desired palette size
                Do While (cClr > nColors)
                    If (pvCombineNodes = 0) Then Exit Do
                Loop
            Next x
        Next y
        CopyMemory ByVal VarPtrArray(Bits()), 0&, 4
        
        '-- Initialize palette
        Initialize
        '-- Go through octree and extract colors
        lEnt = 0
        For y = 1 To UBound(aNodes)
            If (aNodes(y).bIsLeaf) Then
                With aNodes(y)
                    m_tPal(lEnt).R = .vR / .cClrs
                    m_tPal(lEnt).G = .vG / .cClrs
                    m_tPal(lEnt).B = .vB / .cClrs
                    lEnt = lEnt + 1
                End With
            End If
        Next y
        m_Entries = lEnt
            
        '-- Create logical palette
        pvBuildLogicalPalette
    End If
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get lpPalette() As Long
    lpPalette = VarPtr(m_tPal(0))
End Property

Public Property Get hPal() As Long
    hPal = m_hPal
End Property

Public Property Get Entries() As Integer
   Entries = m_Entries
End Property
Public Property Let Entries(ByVal New_Entries As Integer)
   m_Entries = New_Entries
End Property

Public Property Get rgbR(ByVal PalEntry As Byte) As Byte
   rgbR = m_tPal(PalEntry).R
End Property
Public Property Let rgbR(ByVal PalEntry As Byte, ByVal R As Byte)
   m_tPal(PalEntry).R = R
End Property

Public Property Get rgbG(ByVal PalEntry As Byte) As Byte
   rgbG = m_tPal(PalEntry).G
End Property
Public Property Let rgbG(ByVal PalEntry As Byte, ByVal G As Byte)
   m_tPal(PalEntry).G = G
End Property

Public Property Get rgbB(ByVal PalEntry As Byte) As Byte
   rgbB = m_tPal(PalEntry).B
End Property
Public Property Let rgbB(ByVal PalEntry As Byte, ByVal B As Byte)
   m_tPal(PalEntry).B = B
End Property

'========================================================================================
' Private
'========================================================================================

Private Sub pvAddClr(ByVal iBranch As Long, ByVal nLevel As Long, _
                     ByVal vMinR As Byte, ByVal vMaxR As Byte, _
                     ByVal vMinG As Byte, ByVal vMaxG As Byte, _
                     ByVal vMinB As Byte, ByVal vMaxB As Byte, _
                     ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)

' <Recursive>
' Adds a color to the OctTree palette.
' Will call itself if not in correct level.
'
' Inputs:
'  - iBranch (Branch to look down)
'  - nLevel (Current level (depth) in tree)
'  - vMin(R, G, B) (The minimum branch value)
'  - vMax(R, G, B) (The maximum branch value)
'  - R, G, B (The Red, Green, and Blue color components)
  
  Dim iR As Byte, iG As Byte, iB As Byte
  Dim vMid As Byte, iIndex As Long
    
    '-- Find mid values for colors and decide which path to take.
    '   Also update max and min values for later call to self.
    
    vMid = vMinR / 2 + vMaxR / 2
    If (wR * R > vMid) Then iR = 1: vMinR = vMid Else iR = 0: vMaxR = vMid

    vMid = vMinG / 2 + vMaxG / 2
    If (wG * G > vMid) Then iG = 1: vMinG = vMid Else iG = 0: vMaxG = vMid

    vMid = vMinB / 2 + vMaxB / 2
    If (wB * B > vMid) Then iB = 1: vMinB = vMid Else iB = 0: vMaxB = vMid
    
    '-- If no child here then...
    If (aNodes(iBranch).iChildren(iR, iG, iB) = 0) Then
        '-- Get a new node index
        iIndex = pvGetFreeNode
        aNodes(iBranch).iChildren(iR, iG, iB) = iIndex
        aNodes(iBranch).cClrs = aNodes(iBranch).cClrs + 1
        '-- Clear/set data
        With aNodes(iIndex)
            .bIsLeaf = (nLevel = nDepth)
            .iNext = 0
            .cClrs = 0
            .vR = 0
            .vG = 0
            .vB = 0
        End With
      Else
        '-- Has a child here
        iIndex = aNodes(iBranch).iChildren(iR, iG, iB)
    End If
 
    '-- If it is a leaf
    If (aNodes(iIndex).bIsLeaf) Then
        With aNodes(iIndex)
            If (.cClrs = 0) Then cClr = cClr + 1
            .cClrs = .cClrs + 1
            .vR = .vR + R
            .vG = .vG + G
            .vB = .vB + B
        End With
      Else
        With aNodes(iIndex)
            '-- If 2 or more colors, add to reducable aNodes list
            If (.bAddedReduce = 0) Then
                .iNext = aReduce(nLevel)
                 aReduce(nLevel) = iIndex
                .bAddedReduce = -1
            End If
        End With
        '-- Search a level deeper
        pvAddClr iIndex, nLevel + 1, vMinR, vMaxR, vMinG, vMaxG, vMinB, vMaxB, R, G, B
   End If
End Sub

Private Function pvCombineNodes() As Boolean

' Combines octree aNodes to reduce the count of colors.
' Combines all children of a leaf into itself.
  
  Dim i As Long, iIndex As Long
  Dim iR As Byte, iG As Byte, iB As Byte
  Dim nR As Long, nG As Long, nB As Long, nPixel As Long

    '-- Find deepest reducable level
    For i = nDepth To 1 Step -1
        If (aReduce(i) <> 0) Then Exit For
    Next i

    If (i = 0) Then Exit Function
    iIndex = aReduce(i)
    aReduce(i) = aNodes(iIndex).iNext

    For i = 0 To 7
        If (i And 1) = 1 Then iR = 1 Else iR = 0
        If (i And 2) = 2 Then iG = 1 Else iG = 0
        If (i And 4) = 4 Then iB = 1 Else iB = 0
        
        '-- If there is a child
        If (aNodes(iIndex).iChildren(iR, iG, iB) <> 0) Then
            With aNodes(aNodes(iIndex).iChildren(iR, iG, iB))
                '-- Add red, green, blue, and pixel count to running total
                nR = nR + .vR
                nG = nG + .vG
                nB = nB + .vB
                nPixel = nPixel + .cClrs
                '-- Free the node
                pvFreeNode aNodes(iIndex).iChildren(iR, iG, iB)
                cClr = cClr - 1
            End With
            '-- Clear the link
            aNodes(iIndex).iChildren(iR, iG, iB) = 0
        End If
    Next i
    cClr = cClr + 1

    '-- Set the new node data
    With aNodes(iIndex)
        .cClrs = nPixel
        .bIsLeaf = -1
        .vR = nR
        .vG = nG
        .vB = nB
    End With
    pvCombineNodes = -1
End Function

Private Sub pvFreeNode(ByVal iNode As Long)

' Puts a node on the top of the garbage list.
' Inputs:
'  - iNode
'  - Index of node to free
    
    aNodes(iNode).iNext = TopGarbage
    TopGarbage = iNode
    aNodes(iNode).bIsLeaf = 0 ' Necessary for final loop through
    aNodes(iNode).bAddedReduce = 0
    cNodes = cNodes - 1
End Sub

Private Function pvGetFreeNode() As Long

' pvGetFreeNode: Gets a new node index from the trash list or the
' end of the list. Clears child pointers.
' Outputs:
'  - Node index
  
  Dim i  As Long
  Dim iR As Byte
  Dim iG As Byte
  Dim iB As Byte
  
    cNodes = cNodes + 1
    If (TopGarbage = 0) Then
        If (cNodes > UBound(aNodes)) Then
            i = cNodes * 1.1
            ReDim Preserve aNodes(1 To i)
        End If
        pvGetFreeNode = cNodes
      Else
        pvGetFreeNode = TopGarbage
        TopGarbage = aNodes(TopGarbage).iNext
        For i = 0 To 7
            If (i And 1) = 1 Then iR = 1 Else iR = 0
            If (i And 2) = 2 Then iG = 1 Else iG = 0
            If (i And 4) = 4 Then iB = 1 Else iB = 0
            aNodes(pvGetFreeNode).iChildren(iR, iG, iB) = 0
        Next i
    End If
End Function

Private Sub pv16ColourLow8(ByVal lStartIndex As Long)

    With m_tPal(lStartIndex + 0)
        .R = 0: .G = 0: .B = 0
    End With
    With m_tPal(lStartIndex + 1)
        .R = &H80: .G = 0: .B = 0
    End With
    With m_tPal(lStartIndex + 2)
        .R = 0: .G = &H80: .B = 0
    End With
    With m_tPal(lStartIndex + 3)
        .R = &H80: .G = &H80: .B = 0
    End With
    With m_tPal(lStartIndex + 4)
        .R = 0: .G = 0: .B = &H80
    End With
    With m_tPal(lStartIndex + 5)
        .R = &H80: .G = 0: .B = &H80
    End With
    With m_tPal(lStartIndex + 6)
        .R = 0: .G = &H80: .B = &H80
    End With
    With m_tPal(lStartIndex + 7)
        .R = &HC0: .G = &HC0: .B = &HC0
    End With
End Sub

Private Sub pv16ColourHigh8(ByVal lStartIndex As Long)
    
    With m_tPal(lStartIndex + 0)
        .R = &H80: .G = &H80: .B = &H80
    End With
    With m_tPal(lStartIndex + 1)
        .R = &HFF: .G = 0: .B = 0
    End With
    With m_tPal(lStartIndex + 2)
        .R = 0: .G = &HFF: .B = 0
    End With
    With m_tPal(lStartIndex + 3)
        .R = &HFF: .G = &HFF: .B = 0
    End With
    With m_tPal(lStartIndex + 4)
        .R = 0: .G = 0: .B = &HFF
    End With
    With m_tPal(lStartIndex + 5)
        .R = &HFF: .G = 0: .B = &HFF
    End With
    With m_tPal(lStartIndex + 6)
        .R = 0: .G = &HFF: .B = &HFF
    End With
    With m_tPal(lStartIndex + 7)
        .R = &HFF: .G = &HFF: .B = &HFF
    End With
End Sub

Private Sub pvBuildLogicalPalette()
    
  Dim logPal256 As LOGPALETTE256
    
    '-- Destroy previous
    If (m_hPal <> 0) Then
        DeleteObject m_hPal
        m_hPal = 0
    End If
    
    '-- Build logical palette
    With logPal256
        .palNumEntries = m_Entries
        .palVersion = &H300
         CopyMemory .palPalEntry(0), m_tPal(0), 1024
    End With
    m_hPal = CreatePalette(logPal256)
End Sub

Private Sub pvBuildSA(tSA As SAFEARRAY2D, oDIB32 As cDIB)

    With tSA
        .cbElements = IIf(App.LogMode = 1, 1, 4)
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = oDIB32.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = oDIB32.Width
        .pvData = oDIB32.lpBits
    End With
End Sub


