VERSION 5.00
Begin VB.Form fGIFOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GIF options"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3135
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
   Icon            =   "fGIFOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   209
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraGIFSettings 
      Caption         =   "GIF settings"
      Height          =   2415
      Left            =   120
      TabIndex        =   10
      Top             =   1740
      Width           =   2895
      Begin VB.TextBox txtComment 
         Height          =   750
         Left            =   165
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1485
         Width           =   2565
      End
      Begin VB.CheckBox chkTransparent 
         Caption         =   "&Transparent"
         Height          =   315
         Left            =   165
         TabIndex        =   11
         Top             =   360
         Width           =   1290
      End
      Begin VB.CommandButton cmdPickColor 
         Caption         =   "&Pick color..."
         Enabled         =   0   'False
         Height          =   420
         Left            =   1500
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   300
         Width           =   1215
      End
      Begin VB.CheckBox chkInterlaced 
         Caption         =   "&Interlaced"
         Height          =   315
         Left            =   165
         TabIndex        =   13
         Top             =   810
         Width           =   1290
      End
      Begin VB.Label lblComment 
         Caption         =   "Comment:"
         Height          =   240
         Left            =   165
         TabIndex        =   14
         Top             =   1245
         Width           =   1845
      End
   End
   Begin VB.Frame fraPaletteImport 
      Caption         =   "Palette import"
      Height          =   1545
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   2895
      Begin VB.Frame fraPalette 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   150
         TabIndex        =   2
         Top             =   615
         Width           =   1140
         Begin VB.OptionButton optPalette 
            Caption         =   "&Browser"
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   1275
         End
         Begin VB.OptionButton optPalette 
            Caption         =   "&Optimal"
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   4
            Top             =   270
            Width           =   1275
         End
      End
      Begin VB.Frame fraDitherMethod 
         BorderStyle     =   0  'None
         Height          =   810
         Left            =   1305
         TabIndex        =   6
         Top             =   615
         Width           =   1500
         Begin VB.OptionButton optDitherMethod 
            Caption         =   "Solid (&None)"
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   1230
         End
         Begin VB.OptionButton optDitherMethod 
            Caption         =   "O&rdered"
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   8
            Top             =   270
            Width           =   1230
         End
         Begin VB.OptionButton optDitherMethod 
            Caption         =   "&Floyd-Steinberg"
            Height          =   225
            Index           =   2
            Left            =   0
            TabIndex        =   9
            Top             =   540
            Width           =   1500
         End
      End
      Begin VB.Label lblPalette 
         Caption         =   "Palette:"
         Height          =   240
         Left            =   180
         TabIndex        =   1
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label lblDitherMethod 
         Caption         =   "Dither method:"
         Height          =   240
         Left            =   1320
         TabIndex        =   5
         Top             =   315
         Width           =   1530
      End
   End
End
Attribute VB_Name = "fGIFOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Form:          fGIFOptions.frm
' Last revision: 2003.05.25
'================================================

Option Explicit

'-- API:

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'//

Private m_bLoading As Boolean



Private Sub Form_Load()

    '// Get current settings:

    m_bLoading = -1

    '-- Palette:
    optPalette(mDither8bpp.Palette) = -1
    '-- Dither method:
    optDitherMethod(mDither8bpp.DitherMethod) = -1
    '-- Transparent:
    chkTransparent = Abs(g_Transparent)
    '-- Interlaced:
    chkInterlaced = Abs(g_Interlaced)
    '-- Comment:
    txtComment = g_Comment
    txtComment.SelStart = Len(txtComment)

    m_bLoading = 0
    
    '-- Update controls
    mGlobal.UpdateGIFOptionsControls

    '-- A little change
    SendMessage cmdPickColor.hWnd, &HF4&, &H0&, 0&
End Sub

'//

Private Sub optPalette_Click(Index As Integer)
    mDither8bpp.Palette = Index
    If (Not m_bLoading) Then
        mGlobal.DitherTo8bpp
        mGlobal.UpdateCanvas
    End If
End Sub

Private Sub optDitherMethod_Click(Index As Integer)
    mDither8bpp.DitherMethod = Index
    If (Not m_bLoading) Then
        mGlobal.DitherTo8bpp
        mGlobal.UpdateCanvas
    End If
End Sub

'//

Private Sub chkTransparent_Click()
    g_Transparent = -chkTransparent
    cmdPickColor.Enabled = g_Transparent
    mGlobal.MaskImage g_oDIBXor, g_oDIBAnd, g_oPal8bpp, g_Transparent, g_TransparentIdx
    mGlobal.UpdateCanvas
End Sub

Private Sub cmdPickColor_Click()
    If (g_Transparent) Then
        g_Picking = -1
        fMain.ucCanvas.WorkMode = [cnvUserMode]
    End If
End Sub

'//

Private Sub chkInterlaced_Click()
    g_Interlaced = -chkInterlaced
End Sub

'//

Private Sub txtComment_Change()
    g_Comment = txtComment
End Sub
