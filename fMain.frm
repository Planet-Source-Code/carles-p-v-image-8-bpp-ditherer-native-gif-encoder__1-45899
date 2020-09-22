VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "SaveAsGIF"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7200
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   StartUpPosition =   2  'CenterScreen
   Begin SaveAsGIF.ucInfo ucInfo 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      Top             =   5130
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   476
   End
   Begin SaveAsGIF.ucCanvas ucCanvas 
      Height          =   4965
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   8758
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "Import from &bitmap..."
         Index           =   0
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Import from &clipboard"
         Index           =   1
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save..."
         Index           =   3
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&GIF options"
         Index           =   5
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Optimize palette"
         Index           =   6
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   8
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   0
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' App.:          SaveAsGIF v1.0
' Author:        Carles P.V. (*)
' Last revision: 2003.05.25
'================================================

' (*) Hard code (image encoder itself) by Ron van Tilburg
'     - See original post at:
'       http://www.pscode.com/vb/scripts/showcode.asp?lngWId=1&txtCodeId=14210
'     - See mGIFSave module for full credits

Option Explicit



Private Sub Form_Load()
      
    '-- Initialize mGIFSave module
    mGIFSave.InitMasks
    
    '-- Initialize mDither8bpp module
    mDither8bpp.InitializeLUTs
    mDither8bpp.Palette = [ipBrowser]
    mDither8bpp.DitherMethod = [idmOrdered]
    '-- and update related controls (GIF options)
    mGlobal.UpdateGIFOptionsControls
    
    '-- Load transparent-layer pattern
    g_oBack.SetPatternFromStdPicture LoadResPicture("BITMAP_PATTERN_8", vbResBitmap)
    '-- Load pick cursor
    Set ucCanvas.UserIcon = LoadResPicture("CURSOR_PICKCOLOR", vbResCursor)
    
    '-- Hook mouse wheel (canvas zoom)
    mWheel.HookWheel
End Sub

Private Sub Form_Resize()

    '-- Resize canvas control
    If (Me.WindowState <> vbMinimized) Then
        On Error Resume Next
        ucCanvas.Move 0, 0, ScaleWidth, ScaleHeight - (ucInfo.Height + 2)
        On Error GoTo 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '-- Destroy all objects
    mGlobal.FreeAll
End Sub

'//

Private Sub mnuFile_Click(Index As Integer)
    
  Dim sTmpDIB      As New cDIB
  Dim sTmpFilename As String
  Dim aRemoved     As Byte
    
    Select Case Index
    
        Case 0 '-- Load image...
        
            '-- Show open file dialog
            sTmpFilename = mDialogFile.GetFileName(g_FileLoad, "Bitmaps (*.bmp)|*.BMP", , "Import from bitmap", -1)
            
            If (Len(sTmpFilename)) Then
                g_FileLoad = sTmpFilename
                
                '-- Reset Save path
                g_FileSave = ""
                
                '-- Import from bitmap...
                DoEvents
                If (sTmpDIB.CreateFromBitmapFile(g_FileLoad)) Then
                    If (sTmpDIB.BPP <= 8) Then
                        '-- We have a paletted format. Preserve it.
                        mGlobal.g_oDIB32.Destroy
                        mGlobal.TranslateTo8bpp sTmpDIB
                        mGlobal.UpdateCanvas
                        mGlobal.UpdateGIFOptionsControls
                      Else
                        '-- We have a RGB format. Dither it.
                        mGlobal.g_oDIB32.CreateFromStdPicture sTmpDIB.Image, -1
                        mGlobal.DitherTo8bpp
                        mGlobal.UpdateCanvas
                        mGlobal.UpdateGIFOptionsControls
                    End If
                  Else
                    MsgBox "Unexpected error loading bitmap.", vbExclamation
                End If
            End If
            
        Case 1 '-- Import from clipboard
            
            If (Clipboard.GetFormat(vbCFBitmap)) Then
                
                '-- Reset Save path
                g_FileSave = ""
                
                '-- Get from Clipboard
                DoEvents
                If (mGlobal.g_oDIB32.CreateFromStdPicture(Clipboard.GetData(vbCFBitmap), -1)) Then
                    mGlobal.DitherTo8bpp
                    mGlobal.UpdateCanvas
                    mGlobal.UpdateGIFOptionsControls
                  Else
                    MsgBox "Unexpected error loading image from Clipboard.", vbExclamation
                End If
              Else
                MsgBox "Nothing to import from Clipboard.", vbInformation
            End If
            
        Case 3 '-- Save...
            
            '-- Show save file dialog
            sTmpFilename = mDialogFile.GetFileName(g_FileSave, "GIF (*.gif)|*.GIF", , "Save GIF", 0)
            
            If (Len(sTmpFilename)) Then
                g_FileSave = pvCheckGIFext(sTmpFilename)
                
                '-- Save GIF image...
                DoEvents
                Screen.MousePointer = vbHourglass
                If (Not mGIFSave.SaveGIF(g_FileSave, g_oDIBXor, g_oPal8bpp, IIf(g_Transparent, g_TransparentIdx, -1), g_Interlaced, g_Comment)) Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Unexpected error savig GIF image", vbExclamation
                    g_FileSave = ""
                  Else
                    Screen.MousePointer = vbDefault
                    mGlobal.UpdateInfo
                End If
            End If
                    
        Case 5 '-- GIF options
            
            If (Not fGIFOptions.Visible) Then
                fGIFOptions.Show vbModeless, Me
            End If
            
        Case 6 '-- Optimize palette
        
            If (MsgBox(vbCrLf & _
                "Unused palette entries will be removed." & vbCrLf & vbCrLf & _
                "Please, confirm:", vbInformation Or vbYesNo) = vbYes) Then
                
                '-- Try to optimize..
                aRemoved = mDither8bpp.OptimizePalette(g_oDIBXor, g_oPal8bpp)
                '-- Show removed entries
                MsgBox "Removed entries after palette optimization: " & aRemoved, vbInformation
                mGlobal.UpdateInfo
                
                '-- Opaque image
                g_Transparent = 0
                g_TransparentIdx = 0
                fGIFOptions.chkTransparent = 0
            End If
            
        Case 8 '-- Exit
            
            Unload Me
    End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    
    '-- A simple About
    MsgBox "SaveAsGIF v" & App.Major & "." & App.Minor & vbCrLf & _
           "Image 8-bpp ditherer + native GIF encoder" & vbCrLf & _
           "Carles P.V. - 2003" & vbCrLf & vbCrLf & _
           "VB GIF encoder by Ron van Tilburg - ©2001"
End Sub

'//

Private Sub ucCanvas_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    
    If (g_Picking And ucCanvas.DIB.hDIB <> 0) Then
        ucCanvas_MouseMove Button, Shift, x, y
    End If
End Sub

Private Sub ucCanvas_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
    
    If (Button = vbLeftButton) Then
        If (g_Picking And ucCanvas.DIB.hDIB <> 0) Then
            If (x >= 0 And y >= 0 And x < ucCanvas.DIB.Width And y < ucCanvas.DIB.Height) Then
                '-- Get palette index (NOT color)
                g_TransparentIdx = mDither8bpp.PaletteIndex(g_oDIBXor, x, y)
                '-- Re-mask image
                mGlobal.MaskImage g_oDIBXor, g_oDIBAnd, g_oPal8bpp, g_Transparent, g_TransparentIdx
                '-- Update canvas
                mGlobal.UpdateCanvas
            End If
        End If
    End If
End Sub

Private Sub ucCanvas_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    
    If (g_Picking) Then
        g_Picking = 0
        ucCanvas.WorkMode = [cnvScrollMode]
    End If
End Sub

'//

Private Function pvCheckGIFext(ByVal sFilename As String) As String
    
    '-- Simple ext. checker
    If (UCase$(Right$(sFilename, 4)) <> ".GIF") Then
        pvCheckGIFext = sFilename & ".gif"
      Else
        pvCheckGIFext = sFilename
    End If
End Function
