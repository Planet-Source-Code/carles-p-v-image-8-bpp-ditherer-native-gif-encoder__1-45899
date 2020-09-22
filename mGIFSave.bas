Attribute VB_Name = "mGIFSave"
'================================================
' Module:        mGIFSave.bas
' Author:        Ron van Tilburg (*)
' Dependencies:  cDIB.cls,
'                cPal8bpp.cls
' Last revision: 2003.05.25
'================================================

' (*)
'
'   ©2001 Ron van Tilburg - All rights reserved 1.01.2001
'   Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'   Commercial use not permitted - email author please
'
'   Algorithm: use open addressing double hashing (no chaining) on the prefix code / next character
'   combination. We do a variant of Knuth's algorithm D (vol. 3, sec. 6.4) along with G. Knott's
'   relatively-prime secondary probe.  Here, the modular division first probe is gives way to a faster
'   exclusive-or manipulation. Also do block compression with an adaptive reset, whereby the code table
'   is cleared when the compression ratio decreases, but after the table fills. The variable-length output
'   codes are re-sized at this point, and a special CLEAR code is generated for the decompressor. Late
'   addition: construct the table according to file size for noticeable speed improvement on small files.
'   Please direct questions about this implementation to ames!jaw.
'
'   mGIFSave.bas - master file for writing GIF files
'   from the C copyright ©1997 Ron van Tilburg 25.12.1997
'   VB copyright ©2000 Ron van Tilburg 24.12.2000 'what xmas holidays are good for <:-)
'   and copyrights of the original C code from which this is derived are given in the body
'   Documentation of GIF structures is from the GIF standard as attached as html documents
'   All copyrights applying there continue to apply.
'
'   Unisys Corp believes it has the Copyright on all LZW algorithms for GIF files. If it worries you then
'   dont use this code. Read the HTML standards for the owner of the copyright of GIFs and its usability
'
'   ****************************************************************************
'   * GIF Image compression - Modified 'compress'
'   *
'   * Based on: compress.c - File compression ala IEEE Computer, June 1984.
'   *
'   * By Authors:  Spencer W. Thomas       (decvax!harpo!utah-cs!utah-gr!thomas)
'   *              Jim McKie               (decvax!mcvax!jim)
'   *              Steve Davies            (decvax!vax135!petsd!peora!srd)
'   *              Ken Turkowski           (decvax!decwrl!turtlevax!ken)
'   *              James A. Woods          (decvax!ihnp4!ames!jaw)
'   *              Joe Orost               (decvax!vax135!petsd!joe)
'   * VB code by   Ron van Tilburg          rivit@f1.net.au
'   *
'   ****************************************************************************
'   * FROM GIFCOMPR.C - GIF Image compression routines
'   *
'   * Lempel-Ziv compression based on 'compress'. GIF modifications by
'   * David Rowley (mgardi@watdcsu.waterloo.edu)
'   *
'   ****************************************************************************

'//

'   Modifications/additions:
'
'   Nothing special:
'   · Some variable/function names.
'   · Removed (in pvGetPixel function) 1 and 4 bpp cases: I always work with a 8-bpp source.
'   · Some Integer var. declarations for Long declarations.
'   · Finaly, main function: SaveGIF. Now, passed: the 8-bpp DIB and the 8-bpp palette objects.
'     Also added optionaly transparent mode and the addition of a comment block.

'//

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

Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

'//

'-- GIF format. Structures and Constants.

Private Const GIF_SIGNATURE                 As String = "GIF"
Private Const GIF_VERSION_87a               As String = "87a"
Private Const GIF_VERSION_89a               As String = "89a"
Private Const GIF_TRAILER                   As Byte = &H3B

Private Const GIF_EXTENSION_INTRODUCER      As Byte = &H21
Private Const GIF_IMAGE_DESCRIPTOR          As Byte = &H2C
Private Const GIF_BLOCK_TERMINATOR          As Byte = &H0

Private Const GIF_GRAPHIC_CONTROL_EXTENSION As Byte = &HF9
Private Const GIF_COMMENT_EXTENSION         As Byte = &HFE

Private Type tGIFScreenDescriptor
    sdLogicalScreenWidth     As Integer
    sdLogicalScreenHeight    As Integer
    sdFlags                  As Byte
    sdBackgroundColorIndex   As Byte
    sdPixelAspectRatio       As Byte
End Type

Private Type tGIFImageDescriptor
    idImageLeftPosition      As Integer
    idImageTopPosition       As Integer
    idImageWidth             As Integer
    idImageHeight            As Integer
    idFlags                  As Byte
End Type

Private Type tGIFGraphicControlExtension
    gceBlockSize             As Byte ' = &H4
    gceFlags                 As Byte
    gceDelayTime             As Integer
    gceTransparentColorIndex As Byte
End Type

'//

'-- GLOBAL VARIABLES for the Encoding Routines

Private Const MAX_BITS                    As Long = 12           ' User settable max - bits/code
Private Const MAX_BITSHIFT                As Long = 2 ^ MAX_BITS
Private Const MAX_CODE                    As Long = 2 ^ MAX_BITS ' Should NEVER generate this code
Private Const EOF_CODE                    As Long = -1           ' END of input
Private Const TABLE_SIZE                  As Long = 5003         ' 80% occupancy
Private m_lBits                           As Long                ' Number of bits/code
Private m_lMaxCode                        As Long                ' Maximum code, given m_lBits
Private m_lHashTable(0 To TABLE_SIZE - 1) As Long
Private m_lCodeTable(0 To TABLE_SIZE - 1) As Long
Private m_lFreeEntry                      As Long                ' First unused entry

'-- Block compression parameters.
'   After all codes are used up, and compression rate changes, start over.
Private m_lClearFlag     As Long
Private m_lInitBits      As Long
Private m_lClearCode     As Long
Private m_lEOFCode       As Long

'-- Variables for positioning and control
Private m_lx             As Long    ' Image current x pos.
Private m_ly             As Long    ' Image current y pos.
Private m_lImageWidth    As Long    ' Image Width
Private m_lImageHeight   As Long    ' Image Height
Private m_lPixelCount    As Long    ' Pixels left to do
Private m_lPass          As Long    ' Which m_lPass in interlaced mode
Private m_bInterlaced    As Boolean ' Use interlaced mode
Private m_lOutputBytes   As Long    ' Bytes output so far

'-- Variables for the code accumulator (pvOutputCode)
Private m_lOutputBucket  As Long
Private m_lOutputBits    As Long
Private m_lMask(0 To 16) As Long    ' Powers of 2 -1

'-- Variables for the output byte accumulator
Private m_lCharCount     As Long    ' Number of characters so far in this 'packet'
Private m_aChar()        As Byte    ' Will be max 256 bytes long, first byte is length

'//

'-- Global file handler
Private m_hFile   As Long
'-- 8bpp-DIB mapped bytes
Private m_aBits() As Byte



'========================================================================================
' Module initialization
'========================================================================================

Public Sub InitMasks()

  Dim lIdx As Long
    
    '-- Init LUT for fast 2 ^ x - 1
    m_lMask(0) = 0
    For lIdx = 1 To 16
        m_lMask(lIdx) = 2 * (m_lMask(lIdx - 1) + 1) - 1
    Next lIdx
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function SaveGIF(ByVal Filename As String, _
                        oDIB08 As cDIB, _
                        oPal08 As cPal8bpp, _
                        Optional ByVal TransparentColorIndex As Integer = -1, _
                        Optional ByVal Interlaced As Boolean = 0, _
                        Optional ByVal Comment As String) As Boolean
                        
  Dim tScreenDescriptor        As tGIFScreenDescriptor
  Dim tImageDescriptor         As tGIFImageDescriptor
  Dim tGraphicControlExtension As tGIFGraphicControlExtension
  Dim aGlobalPalette()         As Byte
  Dim aBPP                     As Byte
  
  Dim lIdx As Long
  Dim tSA  As SAFEARRAY2D
  
    '-- Check inputs
    If (oDIB08.BPP = [08_bpp] And oPal08.Entries >= 2) Then
  
        '-- Kill previous
        On Error Resume Next
        Kill Filename
        On Error GoTo 0
        
        '-- Get a free file handle and open a new one
        m_hFile = FreeFile()
        Open Filename For Binary Access Write As #m_hFile
        On Error GoTo ErrSave
        
        '-- Write GIF header
        Put #m_hFile, , GIF_SIGNATURE
        If (TransparentColorIndex > -1 Or Len(Comment)) Then
            Put #m_hFile, , GIF_VERSION_89a
          Else
            Put #m_hFile, , GIF_VERSION_87a
        End If
        
        '-- Prepare palette
        With oPal08
            '-- Get palette color depth
            Do
                aBPP = aBPP + 1
            Loop Until 2 ^ aBPP >= .Entries
            '-- Fill GIF palette (RGB triples)
            ReDim aGlobalPalette(0 To 3 * (2 ^ aBPP) - 1)
            For lIdx = 0 To .Entries - 1
                aGlobalPalette(lIdx * 3 + 0) = .rgbR(lIdx)
                aGlobalPalette(lIdx * 3 + 1) = .rgbG(lIdx)
                aGlobalPalette(lIdx * 3 + 2) = .rgbB(lIdx)
            Next lIdx
        End With
        
        '-- Prepare screen descriptor
        With tScreenDescriptor
            .sdLogicalScreenWidth = oDIB08.Width
            .sdLogicalScreenHeight = oDIB08.Height
            .sdFlags = &HF0 Or (aBPP - 1)
            .sdBackgroundColorIndex = 0
            .sdPixelAspectRatio = 0
        End With
        
        '-- Prepare graphic control extension [?] (transparent)
        If (TransparentColorIndex > -1) Then
            With tGraphicControlExtension
                .gceBlockSize = &H4
                .gceFlags = -(TransparentColorIndex > -1)
                .gceDelayTime = 0
                .gceTransparentColorIndex = CByte(TransparentColorIndex)
            End With
        End If
    
        '-- Prepare image descriptor
        With tImageDescriptor
            .idImageLeftPosition = 0
            .idImageTopPosition = 0
            .idImageWidth = oDIB08.Width
            .idImageHeight = oDIB08.Height
            .idFlags = IIf(Interlaced, (aBPP - 1) Or &H40, aBPP - 1)
        End With
        If (aBPP = 1) Then aBPP = 2
        
        '-- Write screen descriptor and global palette
        Put #m_hFile, , tScreenDescriptor
        Put #m_hFile, , aGlobalPalette
    
        '-- Write comment [?]
        If (Len(Comment)) Then
            Put #m_hFile, , GIF_EXTENSION_INTRODUCER
            Put #m_hFile, , GIF_COMMENT_EXTENSION
            Call pvCheckAndWriteComment(Comment)
            Put #m_hFile, , GIF_BLOCK_TERMINATOR
        End If
        
        '-- Write graphic control extension [?]
        If (TransparentColorIndex > -1) Then
            Put #m_hFile, , GIF_EXTENSION_INTRODUCER
            Put #m_hFile, , GIF_GRAPHIC_CONTROL_EXTENSION
            Put #m_hFile, , tGraphicControlExtension
            Put #m_hFile, , GIF_BLOCK_TERMINATOR
        End If
        
        '-- Write GIF image descriptor
        Put #m_hFile, , GIF_IMAGE_DESCRIPTOR
        Put #m_hFile, , tImageDescriptor
        '-- Write GIF-LZW code size
        Put #m_hFile, , aBPP
    
        '-- Prepare some vars. for compress and write image data
        m_lImageWidth = oDIB08.Width
        m_lImageHeight = oDIB08.Height
        m_lPixelCount = oDIB08.Width * oDIB08.Height
        m_bInterlaced = Interlaced * -(oDIB08.Height > 4)
        
        '-- Map 8bpp-DIB bits
        pvBuildSA tSA, oDIB08
        CopyMemory ByVal VarPtrArray(m_aBits()), VarPtr(tSA), 4
        
        '-- Compress/Write image data
        Call pvCompressAndWriteBits(aBPP + 1)
        Put #m_hFile, , GIF_BLOCK_TERMINATOR
        
        '-- Unmap 8bpp-DIB bits
        CopyMemory ByVal VarPtrArray(m_aBits()), 0&, 4
    
        '-- Finaly, write trailer label
        Put #m_hFile, , GIF_TRAILER
    
        '-- Close file: success
        Close #m_hFile
        SaveGIF = -1
    End If
    
ErrSave:
    On Error GoTo 0
End Function

'========================================================================================
' Private
'========================================================================================

Private Sub pvCompressAndWriteBits(nInitBits As Integer)

  Dim lIdx     As Long
  Dim lFCode   As Long
  Dim lC       As Long
  Dim lEnt     As Long
  Dim lDisp    As Long
  Dim m_lShift As Long
  
    '-- Set up where we are starting
    lIdx = 0
    m_lOutputBytes = 0
    m_lPass = 0
    m_lx = 0
    m_ly = 0

    '-- Set up the code accumulator
    m_lOutputBucket = 0
    m_lOutputBits = 0

    '-- Set up initial number of bits
    m_lInitBits = nInitBits

    '-- Set up the necessary values
    m_lClearFlag = 0
    m_lBits = m_lInitBits
    m_lMaxCode = m_lMask(m_lBits)
    m_lClearCode = 2 ^ (nInitBits - 1)
    m_lEOFCode = m_lClearCode + 1
    m_lFreeEntry = m_lClearCode + 2

    '-- Set up output buffers
    Call pvCharInit

    m_lShift = 0
    lFCode = TABLE_SIZE
    Do While lFCode < 65536
        m_lShift = m_lShift + 1
        lFCode = lFCode + lFCode
    Loop
    
    '-- Set hash code range bound for shifting
    m_lShift = 1 + m_lMask(8 - m_lShift)

    Call pvClearTable
    Call pvOutputCode(m_lClearCode)
    
    '-- Start...
    lEnt = pvGetPixel: lC = pvGetPixel
    
    Do While lC <> EOF_CODE

        lFCode = lC * MAX_BITSHIFT + lEnt
        lIdx = (lC * m_lShift) Xor lEnt      ' XOR hashing

        If (m_lHashTable(lIdx) = lFCode) Then
            lEnt = m_lCodeTable(lIdx)
            GoTo NextPixel
        ElseIf (m_lHashTable(lIdx) < 0) Then ' Empty slot
            GoTo NoMatch
        End If

        lDisp = TABLE_SIZE - lIdx            ' Secondary hash (after G. Knott)
        If (lIdx = 0) Then lDisp = 1

Probe:
        lIdx = lIdx - lDisp
        If (lIdx < 0) Then lIdx = lIdx + TABLE_SIZE

        If (m_lHashTable(lIdx) = lFCode) Then
            lEnt = m_lCodeTable(lIdx)
            GoTo NextPixel
        End If

        If (m_lHashTable(lIdx) > 0) Then GoTo Probe

NoMatch:
        Call pvOutputCode(lEnt)
        lEnt = lC

        If (m_lFreeEntry < MAX_CODE) Then
            m_lCodeTable(lIdx) = m_lFreeEntry
            m_lFreeEntry = m_lFreeEntry + 1  ' Code -> Hash table
            m_lHashTable(lIdx) = lFCode
          Else
            Call pvClearBlock
        End If
        
NextPixel:
        lC = pvGetPixel
        
    Loop

    '--  Put out the final code
    Call pvOutputCode(lEnt)
    Call pvOutputCode(m_lEOFCode)
End Sub

Private Function pvGetPixel() As Integer

    If (m_lPixelCount = 0) Then
        '-- End of data
        pvGetPixel = EOF_CODE
        
      Else
        '-- Return the next pixel from the image and increment positions
        pvGetPixel = m_aBits(m_lx, m_ly)
        
        m_lx = m_lx + 1
        If (m_lx = m_lImageWidth) Then
            m_lx = 0
            If (m_bInterlaced = 0) Then
                m_ly = m_ly + 1
              Else
                Select Case m_lPass
                    Case 0:
                        m_ly = m_ly + 8
                        If (m_ly >= m_lImageHeight) Then
                            m_lPass = m_lPass + 1
                            m_ly = 4
                        End If
                    Case 1:
                        m_ly = m_ly + 8
                        If (m_ly >= m_lImageHeight) Then
                            m_lPass = m_lPass + 1
                            m_ly = 2
                        End If
                    Case 2:
                        m_ly = m_ly + 4
                        If (m_ly >= m_lImageHeight) Then
                            m_lPass = m_lPass + 1
                            m_ly = 1
                        End If
                    Case 3:
                        m_ly = m_ly + 2
                End Select
            End If
        End If
        m_lPixelCount = m_lPixelCount - 1
    End If
End Function

Private Sub pvOutputCode(ByVal lCode As Long)
'-- Output the given code.
'   Assumptions:
'     - Chars are 8 bits long.
'   Algorithm:
'     - Maintain a MAX_BITS character long buffer (so that 8 codes will fit in it exactly).
'     - When the buffer fills up empty it and start over.

    m_lOutputBucket = m_lOutputBucket And m_lMask(m_lOutputBits)
    
    If (m_lOutputBits > 0) Then
        m_lOutputBucket = m_lOutputBucket Or (lCode * (1 + m_lMask(m_lOutputBits)))
      Else
        m_lOutputBucket = lCode
    End If
    m_lOutputBits = m_lOutputBits + m_lBits

    Do While (m_lOutputBits >= 8)
        Call pvCharOut(m_lOutputBucket And &HFF&)
        m_lOutputBucket = m_lOutputBucket \ 256&
        m_lOutputBits = m_lOutputBits - 8
    Loop

    '-- If the next entry is going to be too big for the code size, then increase it, if possible.
    If (m_lFreeEntry > m_lMaxCode Or m_lClearFlag = -1) Then
        If (m_lClearFlag = -1) Then
            m_lBits = m_lInitBits
            m_lMaxCode = m_lMask(m_lBits)
            m_lClearFlag = 0
          Else
            m_lBits = m_lBits + 1
            If (m_lBits = MAX_BITS) Then
                m_lMaxCode = MAX_CODE
              Else
                m_lMaxCode = m_lMask(m_lBits)
            End If
        End If
    End If

    '-- At EOF, write the rest of the buffer.
    If (lCode = m_lEOFCode) Then
        Do While (m_lOutputBits > 0)
            Call pvCharOut(m_lOutputBucket And &HFF&)
            m_lOutputBucket = m_lOutputBucket \ 256&
            m_lOutputBits = m_lOutputBits - 8
        Loop
        Call pvFlushChar
    End If
End Sub

Private Sub pvClearBlock()
'-- Clear out the hash table for block compress

    Call pvClearTable
    m_lFreeEntry = m_lClearCode + 2
    m_lClearFlag = -1
    Call pvOutputCode(m_lClearCode)
End Sub

Private Sub pvClearTable()
'-- Reset code table

  Dim lIdx As Long

    For lIdx = 0 To TABLE_SIZE - 1
        m_lHashTable(lIdx) = -1
    Next lIdx
End Sub

Private Sub pvCharInit()
'-- Set up the 'byte output' routine and define the storage for the packet accumulator

    m_lCharCount = 0
    ReDim m_aChar(0 To 255) As Byte
End Sub

Private Sub pvCharOut(ByVal lChar As Long)
'-- Add a character to the end of the current packet, and if it is 254 characters,
'   flush the packet to disk

    m_aChar(m_lCharCount + 1) = lChar              ' 0,...,n mapped to 1,...,n+1
    m_lCharCount = m_lCharCount + 1
    If (m_lCharCount >= 254) Then Call pvFlushChar
End Sub

Private Sub pvFlushChar()
'-- Flush the current packet to disk, and reset the accumulator

    If (m_lCharCount > 0) Then
        m_aChar(0) = m_lCharCount                          ' Set block length,
        ReDim Preserve m_aChar(0 To m_lCharCount) As Byte  ' and redimension to this length
        Put #m_hFile, , m_aChar()                          ' Write it to disk
        m_lOutputBytes = m_lOutputBytes + m_lCharCount + 1 ' Track bytes written
        Call pvCharInit
    End If
End Sub

'//

Private Sub pvCheckAndWriteComment(sComment As String)
    
  Dim aBuff()   As Byte
  Dim aBuffSize As Byte
    
    '-- 255 chars max.
    aBuffSize = Len(sComment)
    If (aBuffSize > 255) Then
        aBuffSize = 255
    End If
    '-- Fill byte array buffer
    ReDim aBuff(1 To aBuffSize)
    CopyMemory aBuff(1), ByVal sComment, aBuffSize
    
    '-- Write
    Put #m_hFile, , aBuffSize ' Block size
    Put #m_hFile, , aBuff()   ' Block itself
End Sub

'//

Private Sub pvBuildSA(tSA As SAFEARRAY2D, oDIB08 As cDIB)

    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = oDIB08.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = oDIB08.BytesPerScanline
        .pvData = oDIB08.lpBits
    End With
End Sub
