Attribute VB_Name = "mGlobal"
'================================================
' Module:        mGlobal.bas
' Author:        Carles P.V.
' Last revision: 2003.05.25
'================================================

Option Explicit

'-- API:

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (lpDst As Any, ByVal Length As Long, ByVal Fill As Byte)

'//

Public g_oDIB32         As New cDIB     ' Source 32-bpp DIB to dither from
Public g_oDIBXor        As New cDIB     ' XOR DIB (8-bpp)
Public g_oDIBAnd        As New cDIB     ' AND DIB (1-bpp)
Public g_oPal8bpp       As New cPal8bpp ' Current palette
Public g_oBack          As New cTile    ' Transparent color pattern
Public g_Picking        As Boolean      ' Picking color on canvas (flag)

Public g_FileLoad       As String       ' Temp. filename
Public g_FileSave       As String       ' Temp. filename

Public g_Transparent    As Boolean      ' GIF transparent mode
Public g_TransparentIdx As Byte         ' GIF transparent color index
Public g_Interlaced     As Boolean      ' GIF interlaced mode
Public g_Comment        As String       ' GIF comment (255 chars max.)



Public Sub TranslateTo8bpp(oDIB As cDIB)

  Dim aPal() As Byte
   
    '-- Get source palette
    oDIB.GetPalette aPal()
    '-- Build XOR and AND DIBs
    g_oDIBXor.Create oDIB.Width, oDIB.Height, [08_bpp]
    g_oDIBAnd.Create oDIB.Width, oDIB.Height, [01_bpp]
    '-- Translate 1,4,8 bpp formats to 8 bpp format
    g_oDIBXor.SetPalette aPal()
    Screen.MousePointer = vbHourglass
    g_oDIBXor.LoadBlt oDIB.hDIBDC
    Screen.MousePointer = vbDefault
    '-- Store to global palette
    g_oPal8bpp.SetPalette aPal()

    '-- Opaque image
    g_Transparent = 0
    g_TransparentIdx = 0
    fGIFOptions.chkTransparent = 0
End Sub

Public Sub DitherTo8bpp()
    
    '-- Build XOR and AND DIBs
    g_oDIBXor.Create g_oDIB32.Width, g_oDIB32.Height, [08_bpp]
    g_oDIBAnd.Create g_oDIB32.Width, g_oDIB32.Height, [01_bpp]
    '-- Dither
    Screen.MousePointer = vbHourglass
    mDither8bpp.Dither g_oDIB32, g_oDIBXor, g_oPal8bpp
    Screen.MousePointer = vbDefault
    '-- Opaque image
    g_Transparent = 0
    g_TransparentIdx = 0
    fGIFOptions.chkTransparent = 0
End Sub

Public Sub UpdateCanvas()
    
    With fMain.ucCanvas
        
        '-- Re-build canvas DIB
        .DIB.Create g_oDIBXor.Width, g_oDIBXor.Height, [32_bpp]
        .Resize
        '-- Paint background pattern
        g_oBack.Tile .DIB.hDIBDC, 0, 0, .DIB.Width, .DIB.Height
        '-- Paint image
        g_oDIBAnd.Stretch .DIB.hDIBDC, 0, 0, g_oDIBAnd.Width, g_oDIBAnd.Height, , , , , vbSrcAnd
        g_oDIBXor.Stretch .DIB.hDIBDC, 0, 0, g_oDIBXor.Width, g_oDIBXor.Height, , , , , vbSrcPaint
        '-- Refresh
        .Repaint
    End With
    
    '-- Update Info
    UpdateInfo
End Sub

Public Sub UpdateInfo()
    With fMain
        '-- Update info
        .ucInfo.TextFile = IIf(Len(g_FileSave), g_FileSave, "[Unnamed]")
        .ucInfo.TextInfo = g_oDIBXor.Width & "x" & g_oDIBXor.Height & "x" & g_oPal8bpp.Entries & "c" & IIf(g_oDIB32.hDIB <> 0, " [Dithered]", "")
        .ucInfo.Refresh
    End With
End Sub

Public Sub MaskImage(oDIBXor As cDIB, oDIBAnd As cDIB, oPal As cPal8bpp, Optional ByVal Transparent As Boolean = 0, Optional ByVal TransparentColorIndex As Byte = 0)
  
  Dim tPalXor(1023) As Byte
  Dim tPalMsk(1023) As Byte
  Dim tPalAnd(7)    As Byte
    
    '-- Temp. palettes
    CopyMemory tPalXor(0), ByVal oPal.lpPalette, 1024
    CopyMemory tPalMsk(0), ByVal oPal.lpPalette, 1024
    FillMemory tPalAnd(4), 3, &HFF
    
    '-- Set And DIB palette
    oDIBAnd.SetPalette tPalAnd()
    
    '-- Transparent [?]
    If (Transparent) Then
        
        FillMemory tPalXor(TransparentColorIndex * 4), 4, &H0
        FillMemory tPalMsk(0), 1024, &H0
        FillMemory tPalMsk(TransparentColorIndex * 4), 4, &HFF

        '-- And DIB (Transparent)
        oDIBXor.SetPalette tPalMsk()
        oDIBAnd.LoadBlt oDIBXor.hDIBDC
        '-- Xor DIB
        oDIBXor.SetPalette tPalXor()
        
      Else
        '-- And DIB (Not transparent)
        oDIBAnd.Cls &H0
        '-- Xor DIB
        oDIBXor.SetPalette tPalXor()
    End If
End Sub

'//

Public Sub UpdateGIFOptionsControls()
    
  Dim bImageExists As Boolean
  Dim bImageIs8bpp As Boolean
   
    bImageExists = (g_oDIBXor.hDIB <> 0)
    bImageIs8bpp = (g_oDIB32.hDIB <> 0)
    
    With fGIFOptions
        '-- Image exits
        .chkTransparent.Enabled = bImageExists
        .cmdPickColor.Enabled = (bImageExists And .chkTransparent)
        .chkInterlaced.Enabled = bImageExists
        .lblComment.Enabled = bImageExists
        .txtComment.Enabled = bImageExists
        '-- Image is 8-bpp color depth
        .fraPaletteImport.Enabled = bImageIs8bpp
        .lblPalette.Enabled = bImageIs8bpp
        .lblDitherMethod.Enabled = bImageIs8bpp
        .optPalette(0).Enabled = bImageIs8bpp
        .optPalette(1).Enabled = bImageIs8bpp
        .optDitherMethod(0).Enabled = bImageIs8bpp
        .optDitherMethod(1).Enabled = bImageIs8bpp
        .optDitherMethod(2).Enabled = bImageIs8bpp
    End With
    
    With fMain
        .mnuFile(3).Enabled = bImageExists
        .mnuFile(6).Enabled = bImageExists
    End With
End Sub

Public Sub FreeAll()
    
    '-- Destroy objects
    Set fMain.ucCanvas.DIB = Nothing
    Set g_oDIB32 = Nothing
    Set g_oDIBXor = Nothing
    Set g_oDIBAnd = Nothing
    Set g_oPal8bpp = Nothing
    Set g_oBack = Nothing
    
    '-- Destroy forms
    Unload fGIFOptions
    Set fGIFOptions = Nothing
    Set fMain = Nothing
End Sub
