VERSION 5.00
Begin VB.Form fGIFOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dither settings"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6675
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
   Icon            =   "fDitherSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDither 
      Caption         =   "Re-&dither"
      Default         =   -1  'True
      Height          =   450
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3870
      Width           =   1725
   End
   Begin VB.Frame fraDitherMethod 
      Caption         =   "Dither mehod"
      Height          =   1200
      Left            =   1020
      TabIndex        =   3
      Top             =   2205
      Width           =   1725
      Begin VB.OptionButton optDitherMethod 
         Caption         =   "Solid (&None)"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   300
         Width           =   1230
      End
      Begin VB.OptionButton optDitherMethod 
         Caption         =   "O&rdered"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   5
         Top             =   570
         Width           =   1230
      End
      Begin VB.OptionButton optDitherMethod 
         Caption         =   "&Floyd-Steinberg"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Top             =   840
         Width           =   1500
      End
   End
   Begin VB.Frame fraPalette 
      Caption         =   "Palette"
      Height          =   930
      Left            =   1020
      TabIndex        =   0
      Top             =   1215
      Width           =   1725
      Begin VB.OptionButton optPalette 
         Caption         =   "&Browser"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   1275
      End
      Begin VB.OptionButton optPalette 
         Caption         =   "&Optimal"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   570
         Width           =   1275
      End
   End
End
Attribute VB_Name = "fGIFOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Form:          fPaletteImports.frm
' Last revision: -
'================================================

Option Explicit

'-- API:

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'//



Private Sub Form_Load()

    '// Get current settings:
    
    '-- Palette:
    optPalette(mDither8bpp.Palette) = -1
    '-- Dither method:
    optDitherMethod(mDither8bpp.DitherMethod) = -1

    '-- Update main menu
    fMain.mnuView(0).Checked = -1
    
    '-- A little effect
    SendMessage cmdDither.hWnd, &HF4&, &H0&, 0&
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '-- [Esc]
    If (KeyCode = vbKeyEscape) Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '-- Update main menu
    fMain.mnuView(0).Checked = 0
End Sub

'//

Private Sub optPalette_Click(Index As Integer)
    mDither8bpp.Palette = Index
End Sub

Private Sub optDitherMethod_Click(Index As Integer)
    mDither8bpp.DitherMethod = Index
End Sub

'//

Private Sub cmdDither_Click()
    mGlobal.Dither
    mGlobal.UpdateCanvas
End Sub

