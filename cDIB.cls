VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Class:         cDIB.cls
' Author:        Carles P.V.
' Dependencies:  None
' Last revision: 2003.05.25
'================================================

Option Explicit

'-- API:

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type BITMAPINFO_001
    bmiHeader       As BITMAPINFOHEADER
    bmiColors(7)    As Byte
End Type

Private Type BITMAPINFO_004
    bmiHeader       As BITMAPINFOHEADER
    bmiColors(63)   As Byte
End Type

Private Type BITMAPINFO_008
    bmiHeader       As BITMAPINFOHEADER
    bmiColors(1023) As Byte
End Type

Private Type BITMAPINFO_RGB
    bmiHeader       As BITMAPINFOHEADER
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PICTDESC
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Const DIB_RGB_COLORS      As Long = 0
Private Const COLORONCOLOR        As Long = 3
Private Const OBJ_BITMAP          As Long = 7
Private Const LR_LOADFROMFILE     As Long = &H10
Private Const IMAGE_BITMAP        As Long = 0
Private Const LR_CREATEDIBSECTION As Long = &H2000

Private Declare Function CreateDIBSection_001 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, lpBitsInfo As BITMAPINFO_001, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection_004 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, lpBitsInfo As BITMAPINFO_004, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection_008 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, lpBitsInfo As BITMAPINFO_008, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection_RGB Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, lpBitsInfo As BITMAPINFO_RGB, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, lpRGBQuad As Any) As Long
Private Declare Function GetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, lpRGBQuad As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetStretchBltMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

'//

'-- Public Enums.:
Public Enum dibBPPCts
    [01_bpp] = 1
    [04_bpp] = 4
    [08_bpp] = 8
    [16_bpp] = 16
    [24_bpp] = 24
    [32_bpp] = 32
End Enum

'-- Property Variables:
Private m_hDC              As Long
Private m_hDIB             As Long
Private m_lpBits           As Long
Private m_Width            As Long
Private m_Height           As Long
Private m_BPP              As dibBPPCts
Private m_BytesPerScanline As Long
Private m_Size             As Long

'-- Private Variables:
Private m_tBIH             As BITMAPINFOHEADER
Private m_hOldDIB          As Long



'========================================================================================
' Class
'========================================================================================

Private Sub Class_Terminate()
    '-- Destroy current DIB
    Destroy
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function Create(ByVal NewWidth As Long, ByVal NewHeight As Long, ByVal Newbpp As dibBPPCts) As Boolean
    
  Dim BI_001 As BITMAPINFO_001
  Dim BI_004 As BITMAPINFO_004
  Dim BI_008 As BITMAPINFO_008
  Dim BI_RGB As BITMAPINFO_RGB
  
    '-- Destroy previous
    Destroy
    
    '-- Define DIB header
    With m_tBIH
        .biSize = Len(m_tBIH)
        .biPlanes = 1
        .biBitCount = Newbpp
        .biWidth = NewWidth
        .biHeight = -NewHeight
    End With
    Select Case Newbpp
        Case [01_bpp]: BI_001.bmiHeader = m_tBIH
        Case [04_bpp]: BI_004.bmiHeader = m_tBIH
        Case [08_bpp]: BI_008.bmiHeader = m_tBIH
        Case Else:     BI_RGB.bmiHeader = m_tBIH
    End Select
    
    '-- Set private props.
    m_Width = NewWidth
    m_Height = NewHeight
    m_BPP = Newbpp
    m_BytesPerScanline = ((m_Width * Newbpp + 31) \ 32) * 4
    m_Size = m_BytesPerScanline * m_Height
    
    '-- Create DIB and select into a DC
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
        Select Case Newbpp
            Case [01_bpp]: m_hDIB = CreateDIBSection_001(m_hDC, BI_001, DIB_RGB_COLORS, m_lpBits, 0, 0)
            Case [04_bpp]: m_hDIB = CreateDIBSection_004(m_hDC, BI_004, DIB_RGB_COLORS, m_lpBits, 0, 0)
            Case [08_bpp]: m_hDIB = CreateDIBSection_008(m_hDC, BI_008, DIB_RGB_COLORS, m_lpBits, 0, 0)
            Case Else:     m_hDIB = CreateDIBSection_RGB(m_hDC, BI_RGB, DIB_RGB_COLORS, m_lpBits, 0, 0)
        End Select
        If (m_hDIB <> 0) Then
            m_hOldDIB = SelectObject(m_hDC, m_hDIB)
          Else
            Destroy
        End If
    End If
    
    '-- Success
    Create = (m_hDIB <> 0)
End Function

Public Function CreateFromStdPicture(Image As StdPicture, Optional ByVal Force32bpp As Boolean = 0) As Byte

  Dim tBI      As BITMAP
  Dim lhDC     As Long
  Dim lhOldBmp As Long
  Dim aPal()   As Byte
    
    If (Not Image Is Nothing) Then
        
        '-- Check object type
        If (GetObjectType(Image) = OBJ_BITMAP) Then
        
            '-- Get object info
            GetObject Image, Len(tBI), tBI
            
            '-- Create DIB
            If (Create(tBI.bmWidth, tBI.bmHeight, IIf(Force32bpp, [32_bpp], tBI.bmBitsPixel))) Then
                lhDC = CreateCompatibleDC(m_hDC)
                If (lhDC <> 0) Then
                    lhOldBmp = SelectObject(lhDC, Image.handle)
                    
                    With tBI
                        '-- Load palette [?]
                        If (.bmBitsPixel <= 8 And Not Force32bpp) Then
                            ReDim aPal(4 * (2 ^ .bmBitsPixel) - 1)
                            GetDIBColorTable lhDC, 0, 2 ^ .bmBitsPixel, aPal(0)
                            SetDIBColorTable m_hDC, 0, 2 ^ .bmBitsPixel, aPal(0)
                        End If
                        '-- Load bits
                        BitBlt m_hDC, 0, 0, .bmWidth, .bmHeight, lhDC, 0, 0, vbSrcCopy
                    End With
                    
                    '-- Destroy temp. DC
                    SelectObject lhDC, lhOldBmp
                    DeleteDC lhDC
                    
                    '-- Success
                    CreateFromStdPicture = tBI.bmBitsPixel
                End If
            End If
        End If
    End If
End Function

Public Function CreateFromBitmapFile(ByVal Filename As String) As Byte
'-- Use this function instead of CreateFromStdPicture (getting Image from VB LoadPicture) if
'   you want to preserve original image format (color depth). LoadPicture 'dithers' to current
'   screen color depth.

  Dim tBI As BITMAP
    
    '-- File exists [?]
    If (FileLen(Filename)) Then
    
        '-- Destroy previous
        Destroy

        '-- Create DIB and select into a DC
        m_hDC = CreateCompatibleDC(0)
        If (m_hDC <> 0) Then
            m_hDIB = LoadImage(App.hInstance, Filename, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
            If (m_hDIB <> 0) Then
                m_hOldDIB = SelectObject(m_hDC, m_hDIB)
              Else
                Destroy
            End If
        End If
        
        '-- Get DIB props.:
        If (m_hDIB <> 0) Then
            '-- Get object
            GetObject m_hDIB, Len(tBI), tBI
            '-- Get props.:
            With tBI
                m_Width = .bmWidth
                m_Height = .bmHeight
                m_BPP = .bmBitsPixel
                m_BytesPerScanline = ((m_Width * m_BPP + 31) \ 32) * 4
                m_Size = m_BytesPerScanline * m_Height
                '-- Bits pointer
                m_lpBits = .bmBits
            End With
            
            '-- Success
            CreateFromBitmapFile = m_BPP
        End If
    End If
End Function

Public Sub CloneTo(oDIB As cDIB)
    
  Dim aPal() As Byte
    
    '-- Create dest. DIB
    oDIB.Create m_Width, m_Height, m_BPP
    '-- Palette [?]
    If (m_BPP <= 8) Then
        GetPalette aPal()
        oDIB.SetPalette aPal()
    End If
    '-- Copy bits
    CopyMemory ByVal oDIB.lpBits, ByVal m_lpBits, m_Size
End Sub

Public Sub Destroy()

    '-- Destroy DIB
    If (m_hDC <> 0) Then
        If (m_hDIB <> 0) Then
            SelectObject m_hDC, m_hOldDIB
            DeleteObject m_hDIB
        End If
        DeleteDC m_hDC
    End If
    
    '-- Reset DIB vars.
    m_hDC = 0
    m_hDIB = 0
    m_hOldDIB = 0
    m_lpBits = 0
    
    '-- Reset DIB props.
    m_Width = 0
    m_Height = 0
    m_BPP = 0
    m_BytesPerScanline = 0
    m_Size = 0
End Sub

'//

Public Function LoadBlt(ByVal hSrcDC As Long, Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, Optional ByVal nWidth As Long, Optional ByVal nHeight As Long) As Long
    
    If (m_hDIB <> 0) Then
        
        '-- Check bounds
        If (nWidth <= 0) Then nWidth = m_Width
        If (nHeight <= 0) Then nHeight = m_Height
        
        '-- Paint bits (*)
        LoadBlt = BitBlt(m_hDC, x, y, nWidth, nHeight, hSrcDC, x, y, vbSrcCopy)
    End If
    
' (*) 32bpp format: Alpha will be lost.
'     So, use CloneTo or CopyMemory from/to bits pointers
End Function
  
Public Function Stretch(ByVal hDstDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, Optional ByVal nSrcWidth As Long, Optional ByVal nSrcHeight As Long, Optional ByVal lROP As RasterOpConstants = vbSrcCopy) As Long
  
  Dim lMode As Long
    
    If (m_hDIB <> 0) Then
        
        '-- Check bounds
        If (nSrcWidth <= 0) Then nSrcWidth = m_Width
        If (nSrcHeight <= 0) Then nSrcHeight = m_Height
        
        '-- BitBlt/StretchBlt
        lMode = GetStretchBltMode(hDstDC)
        SetStretchBltMode hDstDC, COLORONCOLOR
        Stretch = StretchBlt(hDstDC, x, y, nWidth, nHeight, m_hDC, xSrc, ySrc, nSrcWidth, nSrcHeight, lROP)
        SetStretchBltMode hDstDC, lMode
    End If
End Function

Public Sub Cls(Optional ByVal BackColor As Long = &H0)
  
  Dim tRect  As RECT2
  Dim lColor As Long
  Dim hBrush As Long
   
    '-- Clear background (*)
    If (m_hDIB <> 0) Then
        SetRect tRect, 0, 0, m_Width, m_Height
        OleTranslateColor BackColor, 0, lColor
        hBrush = CreateSolidBrush(lColor)
        FillRect m_hDC, tRect, hBrush
        DeleteObject hBrush
    End If
    
' (*) Paletted formats: Windows will match given color to
'     the closest DIB palette color.
End Sub

'//

Public Sub SetPalette(Palette() As Byte)
    
    '-- Use source array bounds
    If (m_hDIB <> 0 And m_BPP <= [08_bpp]) Then
        SetDIBColorTable m_hDC, 0, (UBound(Palette) + 1) \ 4, Palette(0)
    End If
End Sub

Public Sub GetPalette(Palette() As Byte)
    
    '-- Use current bpp as target array bounds
    If (m_hDIB <> 0 And m_BPP <= [08_bpp]) Then
        ReDim Palette(4 * (2 ^ m_BPP) - 1)
        GetDIBColorTable m_hDC, 0, 2 ^ m_BPP, Palette(0)
    End If
End Sub

'//

Public Sub Resize(ByVal NewWidth As Long, ByVal NewHeight As Long, Optional ByVal StretchDIB As Boolean = -1)

  Dim oDIB   As New cDIB
  Dim aPal() As Byte
    
    '-- Create resized DIB (temp.)
    oDIB.Create NewWidth, NewHeight, m_BPP
    '-- Palette [?]
    If (m_BPP <= 8) Then
        GetPalette aPal()
        oDIB.SetPalette aPal()
    End If
    '-- Set bits
    If (StretchDIB) Then
        Stretch oDIB.hDIBDC, 0, 0, NewWidth, NewHeight
      Else
        Stretch oDIB.hDIBDC, 0, 0, m_Width, m_Height
    End If
    
    '-- Create new and load bits
    Create NewWidth, NewHeight, m_BPP
    If (m_BPP <= 8) Then
        SetPalette aPal()
    End If
    LoadBlt oDIB.hDIBDC
End Sub

Public Sub GetBestFitInfo(ByVal DstW As Long, ByVal DstH As Long, bfx As Long, bfy As Long, bfW As Long, bfH As Long, Optional ByVal StretchFit As Boolean = 0)
  
  Dim cW As Single
  Dim cH As Single
    
    If (m_hDIB <> 0) Then
        '-- Get best fit dimensions
        If ((m_Width > DstW Or m_Height > DstH) Or StretchFit) Then
            cW = DstW / m_Width
            cH = DstH / m_Height
            If (cW < cH) Then
                bfW = DstW
                bfH = m_Height * cW
              Else
                bfH = DstH
                bfW = m_Width * cH
            End If
          Else
            bfW = m_Width
            bfH = m_Height
        End If
        '-- Get best fit offsets
        bfx = (DstW - bfW) \ 2
        bfy = (DstH - bfH) \ 2
    End If
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get hDIBDC() As Long
    hDIBDC = m_hDC
End Property

Public Property Get hDIB() As Long
    hDIB = m_hDIB
End Property

Public Property Get lpBits() As Long
    lpBits = m_lpBits
End Property

Public Property Get Width() As Long
    Width = m_Width
End Property

Public Property Get Height() As Long
    Height = m_Height
End Property

Public Property Get BPP() As dibBPPCts
    BPP = m_BPP
End Property

Public Property Get BytesPerScanline() As Long
    BytesPerScanline = m_BytesPerScanline
End Property

Public Property Get Size() As Long
    Size = m_Size
End Property

Property Get Image() As StdPicture
' by Vlad Vissoultchev.
'
' Returned Image color depth = current screen color depth.
' Notes: - Does exists a way (under True-Color screen color depth) to
'          return Image object with DIB color depth ?
'        - Same issue at time to copy to Clipboard...
    
  Dim hSrcDC        As Long
  Dim hDC           As Long
  Dim hBmp          As Long
  Dim hOldBmp       As Long
  Dim lpPictDesc    As PICTDESC
  Dim IID_IDispatch As GUID

    '-- Create temp. bitmap
    hSrcDC = GetDC(0)
    hDC = CreateCompatibleDC(hSrcDC)
    hBmp = CreateCompatibleBitmap(hSrcDC, m_Width, m_Height)
    hOldBmp = SelectObject(hDC, hBmp)
    Call ReleaseDC(0, hSrcDC)
    '-- Paint from DIB
    Stretch hDC, 0, 0, m_Width, m_Height
    '-- Cleanup
    Call SelectObject(hDC, hOldBmp)
    Call DeleteDC(hDC)
    
    '-- Fill image info
    With lpPictDesc
        .Size = Len(lpPictDesc)
        .Type = vbPicTypeBitmap
        .hBmp = hBmp
        .hPal = 0
    End With
    '-- Fill GUID info
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    '-- Create picture from bitmap handle
    Call OleCreatePictureIndirect(lpPictDesc, IID_IDispatch, 1, Image)
End Property
