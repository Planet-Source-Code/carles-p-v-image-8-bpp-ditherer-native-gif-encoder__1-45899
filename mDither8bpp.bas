Attribute VB_Name = "mDither8bpp"
'================================================
' Module:        mDither8bpp.bas
' Author:        Carles P.V.
' Dependencies:  cDIB.cls
'                cPal8bpp.cls
' Last revision: 2003.05.25
'================================================

Option Explicit

'-- API:

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

Private Type RGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

'//

'-- Public Enums.:

Public Enum impPaletteCts
    [ipBrowser] = 0
    [ipOptimal]
End Enum

Public Enum impDitherMethodCts
    [idmNone] = 0
    [idmOrdered]
    [idmFloydSteinberg]
End Enum

'-- Property variables:

Private m_Palette      As impPaletteCts
Private m_DitherMethod As impDitherMethodCts

'-- Private variables:

Private m_tPal(&HFF) As RGBQUAD              '  8-bpp current palette entries

Private m_tSA32      As SAFEARRAY2D          ' 32-bpp SA
Private m_Bits32()   As RGBQUAD              ' 32-bpp mapping bits
Private m_tSA08      As SAFEARRAY2D          '  8-bpp SA
Private m_Bits08()   As Byte                 '  8-bpp mapping bits

Private m_x As Long, m_xIn As Long
Private m_y As Long, m_yIn As Long
Private m_W As Long
Private m_H As Long

'//

Private m_OD_Thr(3, 3)               As Long ' Ordered dither matrix thresholds
Private m_OD_Oiv(3, 3)               As Long ' RGB4096 (optimal palette) inc. values
Private m_OD_Hiv(3, 3)               As Long ' Halftone inc. values

Private m_RGB4096_Inv(&HF, &HF, &HF) As Byte ' RGB4096 palette inverse index LUT
Private m_RGB4096_Trn(-8 To 262)     As Long ' RGB4096 translation LUT (includes offsets)

Private m_HT216_Inv(&H5, &H5, &H5)   As Byte ' Halftone palette inverse index LUT
Private m_HT216_Trn(-51 To 280)      As Long ' Halftone translation LUT (includes offsets)

Private m_FS_Err1(-&HFF To &HFF)     As Long ' Floyd-Steinberg error coefs. LUTs
Private m_FS_Err3(-&HFF To &HFF)     As Long
Private m_FS_Err5(-&HFF To &HFF)     As Long
Private m_FS_Err7(-&HFF To &HFF)     As Long



'========================================================================================
' Module initialization
'========================================================================================

Public Sub InitializeLUTs()

  Dim lIdx As Long
  Dim R As Long
  Dim G As Long
  Dim B As Long
      
    '-- Ordered dither matrix thresholds (Bayer) and incs.
    '   m_OD_Thr [16,256/16]
    '   m_OD_Oiv [ 1, 16/ 1]
    '   m_OD_Hiv [ 3, 51/~3]
    
    m_OD_Thr(0, 0) = 16:  m_OD_Oiv(0, 0) = 1:  m_OD_Hiv(0, 0) = 3
    m_OD_Thr(1, 0) = 144: m_OD_Oiv(1, 0) = 9:  m_OD_Hiv(1, 0) = 29
    m_OD_Thr(2, 0) = 48:  m_OD_Oiv(2, 0) = 3:  m_OD_Hiv(2, 0) = 10
    m_OD_Thr(3, 0) = 176: m_OD_Oiv(3, 0) = 11: m_OD_Hiv(3, 0) = 35
    
    m_OD_Thr(0, 1) = 208: m_OD_Oiv(0, 1) = 13: m_OD_Hiv(0, 1) = 41
    m_OD_Thr(1, 1) = 80:  m_OD_Oiv(1, 1) = 5:  m_OD_Hiv(1, 1) = 16
    m_OD_Thr(2, 1) = 240: m_OD_Oiv(2, 1) = 15: m_OD_Hiv(2, 1) = 48
    m_OD_Thr(3, 1) = 112: m_OD_Oiv(3, 1) = 7:  m_OD_Hiv(3, 1) = 22
    
    m_OD_Thr(0, 2) = 64:  m_OD_Oiv(0, 2) = 4:  m_OD_Hiv(0, 2) = 13
    m_OD_Thr(1, 2) = 192: m_OD_Oiv(1, 2) = 12: m_OD_Hiv(1, 2) = 38
    m_OD_Thr(2, 2) = 32:  m_OD_Oiv(2, 2) = 2:  m_OD_Hiv(2, 2) = 6
    m_OD_Thr(3, 2) = 160: m_OD_Oiv(3, 2) = 10: m_OD_Hiv(3, 2) = 32
    
    m_OD_Thr(0, 3) = 256: m_OD_Oiv(0, 3) = 16: m_OD_Hiv(0, 3) = 51
    m_OD_Thr(1, 3) = 128: m_OD_Oiv(1, 3) = 8:  m_OD_Hiv(1, 3) = 26
    m_OD_Thr(2, 3) = 224: m_OD_Oiv(2, 3) = 14: m_OD_Hiv(2, 3) = 45
    m_OD_Thr(3, 3) = 96:  m_OD_Oiv(3, 3) = 6:  m_OD_Hiv(3, 3) = 19
    
    '-- Halfote-216 inverse indexes LUT
    For B = 0 To &H100 Step &H33
        For G = 0 To &H100 Step &H33
            For R = 0 To &H100 Step &H33
                '-- Set palette inverse index
                m_HT216_Inv(R \ &H33, G \ &H33, B \ &H33) = lIdx
                lIdx = lIdx + 1
            Next R
        Next G
    Next B
    '-- Halftone-216 translation LUT
    For lIdx = -51 To 280
        m_HT216_Trn(lIdx) = lIdx / &H33
        If (m_HT216_Trn(lIdx) < 0) Then m_HT216_Trn(lIdx) = 0
        If (m_HT216_Trn(lIdx) > &H5) Then m_HT216_Trn(lIdx) = &H5
    Next lIdx
    
    '-- RGB-4096 translation LUT
    For lIdx = -7 To 262
        m_RGB4096_Trn(lIdx) = lIdx / &H11
        If (m_RGB4096_Trn(lIdx) < 0) Then m_RGB4096_Trn(lIdx) = 0
        If (m_RGB4096_Trn(lIdx) > &HF) Then m_RGB4096_Trn(lIdx) = &HF
    Next lIdx
    
    '-- Floyd-Steinberg diffusion errors LUTs
    For lIdx = -&HFF To &HFF
        m_FS_Err1(lIdx) = (1 * lIdx) / &H10
        m_FS_Err3(lIdx) = (3 * lIdx) / &H10
        m_FS_Err5(lIdx) = (5 * lIdx) / &H10
        m_FS_Err7(lIdx) = (7 * lIdx) / &H10
    Next lIdx
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get Palette() As impPaletteCts
    Palette = m_Palette
End Property
Public Property Let Palette(ByVal New_Palette As impPaletteCts)
    m_Palette = New_Palette
End Property

Public Property Get DitherMethod() As impDitherMethodCts
    DitherMethod = m_DitherMethod
End Property
Public Property Let DitherMethod(ByVal New_DitherMethod As impDitherMethodCts)
    m_DitherMethod = New_DitherMethod
End Property

'========================================================================================
' Methods
'========================================================================================

Public Sub Dither(oDIB32 As cDIB, oDIB08 As cDIB, oPal8bpp As cPal8bpp)
  
  Dim oTmpDIB32  As New cDIB
  Dim aPal(1023) As Byte
  
  Dim bfW As Long, bfH As Long
  Dim bfx As Long, bfy As Long
  
    '-- Set palette
    
    Select Case m_Palette
            
        Case [ipOptimal]
            
            '-- Get Optimal from reduced DIB (speed up)
            '   Adjust this as you feel necessary.
            oTmpDIB32.CreateFromStdPicture oDIB32.Image, -1
            oTmpDIB32.GetBestFitInfo 150, 150, bfx, bfy, bfW, bfH
            oTmpDIB32.Resize bfW, bfH
            oPal8bpp.CreateOptimal oTmpDIB32, 256, 8 ' 6 levels would be enough
            '-- Build 4096-colors palette inverse indexes LUT
            pvBuildRGB4096LUT oPal8bpp

        Case [ipBrowser]
        
            '-- Create Halftne-216 (6x6x6)
            oPal8bpp.CreateHalftone [216_phLevels]
    End Select
    
    '-- Fill temp. palette copy (speed up)
    CopyMemory m_tPal(0), ByVal oPal8bpp.lpPalette, 1024
    
    '-- Build a 32-bpp copy from source image [?]
    If (m_DitherMethod = [idmFloydSteinberg]) Then
        oDIB32.CloneTo oTmpDIB32
    End If
    
    '-- Rebuild 8-bpp target DIB (create and set current palette)
    CopyMemory aPal(0), m_tPal(0), 1024
    oDIB08.Create oDIB32.Width, oDIB32.Height, [08_bpp]
    oDIB08.SetPalette aPal()
    
    '-- Map source and target DIB bits (32-bpp/8-bpp)
    If (m_DitherMethod = [idmFloydSteinberg]) Then
        pvBuild_32bppSA m_tSA32, oTmpDIB32
      Else
        pvBuild_32bppSA m_tSA32, oDIB32
    End If
    pvBuild_08bppSA m_tSA08, oDIB08
    CopyMemory ByVal VarPtrArray(m_Bits32()), VarPtr(m_tSA32), 4
    CopyMemory ByVal VarPtrArray(m_Bits08()), VarPtr(m_tSA08), 4
   
    '-- Get dimensions
    m_W = oDIB32.Width - 1
    m_H = oDIB32.Height - 1
   
    '-- Dither...
    Select Case m_Palette
            
        Case [ipBrowser]
            Select Case m_DitherMethod
                Case [idmNone]:           pvDitherToHT216
                Case [idmOrdered]:        pvDitherToHT216_Ordered
                Case [idmFloydSteinberg]: pvDitherToHT216_FloydSteinberg
            End Select
        
        Case [ipOptimal]
            Select Case m_DitherMethod
                Case [idmNone]:           pvDitherToPalette
                Case [idmOrdered]:        pvDitherToPalette_Ordered
                Case [idmFloydSteinberg]: pvDitherToPalette_FloydSteinberg
            End Select
    End Select
    
    '-- Unmap DIB bits
    CopyMemory ByVal VarPtrArray(m_Bits32()), 0&, 4
    CopyMemory ByVal VarPtrArray(m_Bits08()), 0&, 4
End Sub

Public Function OptimizePalette(oDIB08 As cDIB, oPal8bpp As cPal8bpp) As Byte

  Dim nBefore As Integer
  Dim aPal()  As Byte
  Dim bUsed() As Boolean
  Dim bTrnE() As Byte
  Dim bInvE() As Byte
  Dim lIdx    As Long
  Dim lMax    As Long
  
    Screen.MousePointer = vbHourglass
  
    '-- Store current number of entries and initialize arrays
    nBefore = oPal8bpp.Entries
    ReDim bUsed(nBefore - 1)
    ReDim bTrnE(nBefore - 1)
    ReDim bInvE(nBefore - 1)

    '-- Map current 8-bpp DIB bits
    pvBuild_08bppSA m_tSA08, oDIB08
    CopyMemory ByVal VarPtrArray(m_Bits08()), VarPtr(m_tSA08), 4
   
    '-- Get dimensions
    m_W = oDIB08.Width - 1
    m_H = oDIB08.Height - 1
    
    '-- Check used entries...
    For m_y = 0 To m_H
        For m_x = 0 To m_W
            bUsed(m_Bits08(m_x, m_y)) = -1
        Next m_x
    Next m_y
    
    '-- 'Strecth' palette...
    For lIdx = 0 To oPal8bpp.Entries - 1
        If (bUsed(lIdx)) Then
            bTrnE(lMax) = lIdx ' New index
            bInvE(lIdx) = lMax ' Inverse index
            lMax = lMax + 1    ' Current count
        End If
    Next lIdx
    
    '-- Any entry removed [?]
    If (lMax < oPal8bpp.Entries) Then
    
        If (lMax > 1) Then lMax = lMax - 1
        '-- Build temp. palette with only used entries
        ReDim aPal(4 * (lMax + 1) - 1)
        For lIdx = 0 To lMax
            aPal(4 * lIdx + 0) = oPal8bpp.rgbB(bTrnE(lIdx))
            aPal(4 * lIdx + 1) = oPal8bpp.rgbG(bTrnE(lIdx))
            aPal(4 * lIdx + 2) = oPal8bpp.rgbR(bTrnE(lIdx))
        Next lIdx
        
        '-- Rebuild current palette
        oPal8bpp.Initialize (lMax + 1)
        CopyMemory ByVal oPal8bpp.lpPalette, aPal(0), 4 * (lMax + 1)
        '-- and set as XOR DIB palette
        oDIB08.SetPalette aPal()
        
        '-- Set new indexes...
        For m_y = 0 To m_H
            For m_x = 0 To m_W
                m_Bits08(m_x, m_y) = bInvE(m_Bits08(m_x, m_y))
            Next m_x
        Next m_y
    End If

    '-- Unmap DIB bits
    CopyMemory ByVal VarPtrArray(m_Bits08()), 0&, 4
    
    Screen.MousePointer = vbDefault
    
    '-- Return removed entries
    OptimizePalette = (nBefore - oPal8bpp.Entries)
End Function

Public Function PaletteIndex(oDIB08 As cDIB, ByVal x As Long, ByVal y As Long) As Byte
    
    '-- Map DIB bits
    pvBuild_08bppSA m_tSA08, oDIB08
    CopyMemory ByVal VarPtrArray(m_Bits08()), VarPtr(m_tSA08), 4
    
    '-- Get 8-bpp index
    PaletteIndex = m_Bits08(x, y)

    '-- Unmap DIB bits
    CopyMemory ByVal VarPtrArray(m_Bits08()), 0&, 4
End Function

'========================================================================================
' Private
'========================================================================================

Private Sub pvBuildRGB4096LUT(oPal8bpp As cPal8bpp)

  Dim R As Long
  Dim G As Long
  Dim B As Long

    '-- Build 4096-colors palette inverse indexes LUT
    For R = 0 To &HF
    For G = 0 To &HF
    For B = 0 To &HF
        oPal8bpp.ClosestIndex R * &H11, G * &H11, B * &H11, m_RGB4096_Inv(R, G, B)
    Next B, G, R
End Sub

'//

Private Sub pvDitherToHT216()
  
    For m_y = 0 To m_H
        For m_x = 0 To m_W
            '-- Set 8-bpp palette index
            m_Bits08(m_x, m_y) = m_HT216_Inv(m_HT216_Trn(m_Bits32(m_x, m_y).R), m_HT216_Trn(m_Bits32(m_x, m_y).G), m_HT216_Trn(m_Bits32(m_x, m_y).B))
        Next m_x
    Next m_y
End Sub

Private Sub pvDitherToHT216_Ordered()
  
  Dim lodX As Long, lodY As Long
  Dim lodT As Long, lodI As Long
  Dim newR As Long, newG As Long, newB As Long
  
    For m_y = 0 To m_H
        For m_x = 0 To m_W
        
            '-- Threshold/Inc.
            lodT = m_OD_Thr(lodX, lodY)
            lodI = m_OD_Hiv(lodX, lodY)
            
            '-- Inc ord. matrix column
            lodX = lodX + 1
            If (lodX = 4) Then lodX = 0
           
            '-- Ordered dither
            If (m_HT216_Trn(m_Bits32(m_x, m_y).B) < lodT) Then
                newB = m_HT216_Trn(m_Bits32(m_x, m_y).B + &H1A - lodI)
              Else
                newB = m_HT216_Trn(m_Bits32(m_x, m_y).B - lodI)
            End If
            If (m_HT216_Trn(m_Bits32(m_x, m_y).G) < lodT) Then
                newG = m_HT216_Trn(m_Bits32(m_x, m_y).G + &H1A - lodI)
              Else
                newG = m_HT216_Trn(m_Bits32(m_x, m_y).G - lodI)
            End If
            If (m_HT216_Trn(m_Bits32(m_x, m_y).R) < lodT) Then
                newR = m_HT216_Trn(m_Bits32(m_x, m_y).R + &H1A - lodI)
              Else
                newR = m_HT216_Trn(m_Bits32(m_x, m_y).R - lodI)
            End If
            
            '-- Set 8-bpp palette index
            m_Bits08(m_x, m_y) = m_HT216_Inv(newR, newG, newB)
        Next m_x
        
        '-- Inc ord. matrix row
        lodX = 0
        lodY = lodY + 1
        If lodY = 4 Then lodY = 0
    Next m_y
End Sub

Private Sub pvDitherToHT216_FloydSteinberg()
    
  Dim aIdx As Byte
  Dim errR As Long, errG As Long, errB As Long
  Dim newR As Long, newG As Long, newB As Long
  
    For m_y = 0 To m_H
        For m_x = 0 To m_W

            '-- Get pre-calculated palette index
            aIdx = m_HT216_Inv(m_HT216_Trn(m_Bits32(m_x, m_y).R), m_HT216_Trn(m_Bits32(m_x, m_y).G), m_HT216_Trn(m_Bits32(m_x, m_y).B))
            '-- Set 8-bpp palette index
            m_Bits08(m_x, m_y) = aIdx

            '-- Diffuse error (Floyd-Steinberg)...
            errB = CLng(m_Bits32(m_x, m_y).B) - m_tPal(aIdx).B
            errG = CLng(m_Bits32(m_x, m_y).G) - m_tPal(aIdx).G
            errR = CLng(m_Bits32(m_x, m_y).R) - m_tPal(aIdx).R

            If (Abs(errB) + Abs(errG) + Abs(errR) > 3) Then
            
                '-- Recursive...
                m_Bits32(m_x, m_y) = m_tPal(aIdx)
                
                If (m_x < m_W) Then
                    m_xIn = m_x + 1
                    newB = m_Bits32(m_xIn, m_y).B + m_FS_Err7(errB)
                    newG = m_Bits32(m_xIn, m_y).G + m_FS_Err7(errG)
                    newR = m_Bits32(m_xIn, m_y).R + m_FS_Err7(errR)
                    If (newB < 0) Then newB = 0 Else If (newB > &HFF) Then newB = &HFF
                    If (newG < 0) Then newG = 0 Else If (newG > &HFF) Then newG = &HFF
                    If (newR < 0) Then newR = 0 Else If (newR > &HFF) Then newR = &HFF
                    m_Bits32(m_xIn, m_y).B = newB
                    m_Bits32(m_xIn, m_y).G = newG
                    m_Bits32(m_xIn, m_y).R = newR
                End If
                If (m_y < m_H) Then
                    m_yIn = m_y + 1
                    newB = m_Bits32(m_x, m_yIn).B + m_FS_Err5(errB)
                    newG = m_Bits32(m_x, m_yIn).G + m_FS_Err5(errG)
                    newR = m_Bits32(m_x, m_yIn).R + m_FS_Err5(errR)
                    If (newB < 0) Then newB = 0 Else If (newB > &HFF) Then newB = &HFF
                    If (newG < 0) Then newG = 0 Else If (newG > &HFF) Then newG = &HFF
                    If (newR < 0) Then newR = 0 Else If (newR > &HFF) Then newR = &HFF
                    m_Bits32(m_x, m_yIn).B = newB
                    m_Bits32(m_x, m_yIn).G = newG
                    m_Bits32(m_x, m_yIn).R = newR
                    If (m_x < m_W) Then
                        m_xIn = m_x + 1
                        newB = m_Bits32(m_xIn, m_yIn).B + m_FS_Err1(errB)
                        newG = m_Bits32(m_xIn, m_yIn).G + m_FS_Err1(errR)
                        newR = m_Bits32(m_xIn, m_yIn).R + m_FS_Err1(errG)
                        If (newB < 0) Then newB = 0 Else If (newB > &HFF) Then newB = &HFF
                        If (newG < 0) Then newG = 0 Else If (newG > &HFF) Then newG = &HFF
                        If (newR < 0) Then newR = 0 Else If (newR > &HFF) Then newR = &HFF
                        m_Bits32(m_xIn, m_yIn).B = newB
                        m_Bits32(m_xIn, m_yIn).G = newG
                        m_Bits32(m_xIn, m_yIn).R = newR
                    End If
                    If (m_x > 0) Then
                        m_xIn = m_x - 1
                        newB = m_Bits32(m_xIn, m_yIn).B + m_FS_Err3(errB)
                        newG = m_Bits32(m_xIn, m_yIn).G + m_FS_Err3(errG)
                        newR = m_Bits32(m_xIn, m_yIn).R + m_FS_Err3(errR)
                        If (newB < 0) Then newB = 0 Else If (newB > &HFF) Then newB = &HFF
                        If (newG < 0) Then newG = 0 Else If (newG > &HFF) Then newG = &HFF
                        If (newR < 0) Then newR = 0 Else If (newR > &HFF) Then newR = &HFF
                        m_Bits32(m_xIn, m_yIn).B = newB
                        m_Bits32(m_xIn, m_yIn).G = newG
                        m_Bits32(m_xIn, m_yIn).R = newR
                    End If
                End If
            End If
        Next m_x
    Next m_y
End Sub

'//

Private Sub pvDitherToPalette()

    For m_y = 0 To m_H
        For m_x = 0 To m_W
            '-- Set 8-bpp palette index
            m_Bits08(m_x, m_y) = m_RGB4096_Inv(m_RGB4096_Trn(m_Bits32(m_x, m_y).R), m_RGB4096_Trn(m_Bits32(m_x, m_y).G), m_RGB4096_Trn(m_Bits32(m_x, m_y).B))
        Next m_x
    Next m_y
End Sub

Private Sub pvDitherToPalette_Ordered()

  Dim lodX As Long, lodY As Long
  Dim lodT As Long, lodI As Long
  Dim newR As Long, newG As Long, newB As Long
    
    For m_y = 0 To m_H
        For m_x = 0 To m_W
            
            '-- Threshold/Inc.
            lodT = m_OD_Thr(lodX, lodY)
            lodI = m_OD_Oiv(lodX, lodY)

            '-- Inc ord. matrix column
            lodX = lodX + 1
            If (lodX = 4) Then lodX = 0
            
            '-- Ordered dither
            If (m_RGB4096_Trn(m_Bits32(m_x, m_y).B) < lodT) Then
                newB = m_RGB4096_Trn(m_Bits32(m_x, m_y).B + &H8 - lodI)
              Else
                newB = m_RGB4096_Trn(m_Bits32(m_x, m_y).B - lodI)
            End If
            If (m_RGB4096_Trn(m_Bits32(m_x, m_y).G) < lodT) Then
                newG = m_RGB4096_Trn(m_Bits32(m_x, m_y).G + &H8 - lodI)
              Else
                newG = m_RGB4096_Trn(m_Bits32(m_x, m_y).G - lodI)
            End If
            If (m_RGB4096_Trn(m_Bits32(m_x, m_y).R) < lodT) Then
                newR = m_RGB4096_Trn(m_Bits32(m_x, m_y).R + &H8 - lodI)
              Else
                newR = m_RGB4096_Trn(m_Bits32(m_x, m_y).R - lodI)
            End If
            
            '-- Set 8-bpp palette index
            m_Bits08(m_x, m_y) = m_RGB4096_Inv(newR, newG, newB)
        Next m_x
        
        '-- Inc ord. matrix row
        lodX = 0
        lodY = lodY + 1
        If lodY = 4 Then lodY = 0
    Next m_y
End Sub

Private Sub pvDitherToPalette_FloydSteinberg()
  
  Dim aIdx As Byte
  Dim errR As Long, errG As Long, errB As Long
  Dim newR As Long, newG As Long, newB As Long
    
    For m_y = 0 To m_H
        For m_x = 0 To m_W

            '-- Get pre-calculated palette index
            aIdx = m_RGB4096_Inv(m_RGB4096_Trn(m_Bits32(m_x, m_y).R), m_RGB4096_Trn(m_Bits32(m_x, m_y).G), m_RGB4096_Trn(m_Bits32(m_x, m_y).B))
            '-- Set 8-bpp palette index
            m_Bits08(m_x, m_y) = aIdx

            '-- Diffuse error (Floyd-Steinberg)...
            errB = CLng(m_Bits32(m_x, m_y).B) - m_tPal(aIdx).B
            errG = CLng(m_Bits32(m_x, m_y).G) - m_tPal(aIdx).G
            errR = CLng(m_Bits32(m_x, m_y).R) - m_tPal(aIdx).R

            If (Abs(errB) + Abs(errG) + Abs(errR) > 3) Then
            
                '-- Recursive...
                m_Bits32(m_x, m_y) = m_tPal(aIdx)
                
                If (m_x < m_W) Then
                    m_xIn = m_x + 1
                    newB = m_Bits32(m_xIn, m_y).B + m_FS_Err7(errB)
                    newG = m_Bits32(m_xIn, m_y).G + m_FS_Err7(errG)
                    newR = m_Bits32(m_xIn, m_y).R + m_FS_Err7(errR)
                    If (newB < 0) Then newB = 0 Else If (newB > &HFF) Then newB = &HFF
                    If (newG < 0) Then newG = 0 Else If (newG > &HFF) Then newG = &HFF
                    If (newR < 0) Then newR = 0 Else If (newR > &HFF) Then newR = &HFF
                    m_Bits32(m_xIn, m_y).B = newB
                    m_Bits32(m_xIn, m_y).G = newG
                    m_Bits32(m_xIn, m_y).R = newR
                End If
                If (m_y < m_H) Then
                    m_yIn = m_y + 1
                    newB = m_Bits32(m_x, m_yIn).B + m_FS_Err5(errB)
                    newG = m_Bits32(m_x, m_yIn).G + m_FS_Err5(errG)
                    newR = m_Bits32(m_x, m_yIn).R + m_FS_Err5(errR)
                    If (newB < 0) Then newB = 0 Else If (newB > &HFF) Then newB = &HFF
                    If (newG < 0) Then newG = 0 Else If (newG > &HFF) Then newG = &HFF
                    If (newR < 0) Then newR = 0 Else If (newR > &HFF) Then newR = &HFF
                    m_Bits32(m_x, m_yIn).B = newB
                    m_Bits32(m_x, m_yIn).G = newG
                    m_Bits32(m_x, m_yIn).R = newR
                    If (m_x < m_W) Then
                        m_xIn = m_x + 1
                        newB = m_Bits32(m_xIn, m_yIn).B + m_FS_Err1(errB)
                        newG = m_Bits32(m_xIn, m_yIn).G + m_FS_Err1(errR)
                        newR = m_Bits32(m_xIn, m_yIn).R + m_FS_Err1(errG)
                        If (newB < 0) Then newB = 0 Else If (newB > &HFF) Then newB = &HFF
                        If (newG < 0) Then newG = 0 Else If (newG > &HFF) Then newG = &HFF
                        If (newR < 0) Then newR = 0 Else If (newR > &HFF) Then newR = &HFF
                        m_Bits32(m_xIn, m_yIn).B = newB
                        m_Bits32(m_xIn, m_yIn).G = newG
                        m_Bits32(m_xIn, m_yIn).R = newR
                    End If
                    If (m_x > 0) Then
                        m_xIn = m_x - 1
                        newB = m_Bits32(m_xIn, m_yIn).B + m_FS_Err3(errB)
                        newG = m_Bits32(m_xIn, m_yIn).G + m_FS_Err3(errG)
                        newR = m_Bits32(m_xIn, m_yIn).R + m_FS_Err3(errR)
                        If (newB < 0) Then newB = 0 Else If (newB > &HFF) Then newB = &HFF
                        If (newG < 0) Then newG = 0 Else If (newG > &HFF) Then newG = &HFF
                        If (newR < 0) Then newR = 0 Else If (newR > &HFF) Then newR = &HFF
                        m_Bits32(m_xIn, m_yIn).B = newB
                        m_Bits32(m_xIn, m_yIn).G = newG
                        m_Bits32(m_xIn, m_yIn).R = newR
                    End If
                End If
            End If
        Next m_x
    Next m_y
End Sub

'//

Private Sub pvBuild_08bppSA(tSA As SAFEARRAY2D, oDIB As cDIB)

    '-- 8-bpp DIB mapping
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = oDIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = oDIB.BytesPerScanline
        .pvData = oDIB.lpBits
    End With
End Sub

Private Sub pvBuild_32bppSA(tSA As SAFEARRAY2D, oDIB As cDIB)

    '-- 32-bpp DIB mapping
    With tSA
        .cbElements = IIf(App.LogMode = 1, 1, 4)
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = oDIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = oDIB.Width
        .pvData = oDIB.lpBits
    End With
End Sub
