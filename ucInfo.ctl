VERSION 5.00
Begin VB.UserControl ucInfo 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   ClipControls    =   0   'False
   ScaleHeight     =   14
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
End
Attribute VB_Name = "ucInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucInfo.ctl
' Author:        Carles P.V.
' Dependencies:  None
' Last revision: 2003.03.28
'================================================

Option Explicit

'-- API:

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Const BDR_RAISEDINNER     As Long = &H4
Private Const BDR_SUNKENOUTER     As Long = &H2
Private Const BF_RECT             As Long = &HF

Private Const COLOR_BTNFACE       As Long = 15

Private Const DFC_SCROLL          As Long = &H3
Private Const DFCS_SCROLLSIZEGRIP As Long = &H8

Private Const WM_NCLBUTTONDOWN    As Long = &HA1
Private Const HTBOTTOMRIGHT       As Long = &H11

Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT2, lpSourceRect As RECT2) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT2, ByVal wFormat As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT2, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" (ByVal hDC As Long, ByVal pszPath As String, ByVal dx As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
                         
'//

'-- Property Variables:
Private m_TextFile         As String
Private m_TextInfo         As String

'-- Private Variables:
Private m_BarRect          As RECT2
Private m_SizeGripRect     As RECT2
Private m_EdgeRect(1 To 2) As RECT2
Private m_TextRect(1 To 2) As RECT2



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Paint()

  Dim lEdg       As Long
  Dim sCmpctPath As String
    
    '-- Erase background
    FillRect hDC, m_BarRect, GetSysColorBrush(COLOR_BTNFACE)
    '-- Draw edges
    For lEdg = 1 To 2
        DrawEdge hDC, m_EdgeRect(lEdg), BDR_SUNKENOUTER, BF_RECT
    Next lEdg
    '-- Draw size grip
    DrawFrameControl hDC, m_SizeGripRect, DFC_SCROLL, DFCS_SCROLLSIZEGRIP
    '-- Draw text
    sCmpctPath = pvCompactPath(m_TextFile, pvTextFileWidth)
    DrawText hDC, sCmpctPath, Len(sCmpctPath), m_TextRect(1), &H0
    DrawText hDC, m_TextInfo, Len(m_TextInfo), m_TextRect(2), &H0
End Sub

Private Sub UserControl_Resize()
    
  Const INFO_WIDTH As Long = 150
    
  Dim W    As Long
  Dim H    As Long
  Dim SG   As Long
  Dim lPnl As Long
  
    W = ScaleWidth
    H = 18
    Height = 18 * Screen.TwipsPerPixelY
    
    On Error Resume Next
    
    '-- Check parent form window state:
    '   Size Grip (Show/Hide)
    If (Parent.WindowState = vbMaximized) Then
        SG = 0
      Else
        SG = H
    End If
    
    '-- Set main Rect. and size grip Rect.
    SetRect m_BarRect, 0, 0, W, H
    SetRect m_SizeGripRect, W - SG, 0, W, H
    
    '-- Set text Rects. (Edge and text)
    SetRect m_EdgeRect(1), 0, 0, W - INFO_WIDTH - SG, H
    SetRect m_EdgeRect(2), W - INFO_WIDTH - SG, 0, W - SG, H
    For lPnl = 1 To 2
        CopyRect m_TextRect(lPnl), m_EdgeRect(lPnl)
        With m_TextRect(lPnl)
            .x1 = .x1 + 4
            .y1 = .y1 + 2
            .x2 = .x2 - 4
        End With
    Next lPnl
    
    '-- Refresh
    Refresh
    
    On Error GoTo 0
End Sub

'//

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = vbLeftButton And x > m_SizeGripRect.x1) Then
        ReleaseCapture
        SendMessage Parent.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal 0
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (x > m_SizeGripRect.x1) Then
        MousePointer = vbSizeNWSE
      Else
        MousePointer = vbDefault
   End If
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Sub Refresh()
    UserControl_Paint
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Let TextFile(ByVal New_TextFile As String)
Attribute TextFile.VB_MemberFlags = "400"
    m_TextFile = New_TextFile
End Property

Public Property Let TextInfo(ByVal New_TextInfo As String)
    m_TextInfo = New_TextInfo
End Property

'//

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Font = Ambient.Font
End Sub

'========================================================================================
' Private
'========================================================================================

Public Function pvTextFileWidth() As Long
    pvTextFileWidth = m_TextRect(1).x2 - m_TextRect(1).x1
End Function

Private Function pvCompactPath(ByVal sFullPath As String, ByVal lWidth As Long) As String
'-- KPD-Team 2000
'   URL: http://www.allapi.net/
'   E-Mail: KPDTeam@Allapi.net

  Dim lZeroPos As Long
  
    '-- Check width min.
    If (lWidth <= 25) Then
        pvCompactPath = ""
    
      Else
        '-- Compact
        PathCompactPath hDC, sFullPath, lWidth
    
        '-- Remove all trailing Chr$(0)'s
        lZeroPos = InStr(1, sFullPath, Chr$(0))
        If (lZeroPos > 0) Then
            pvCompactPath = Left$(sFullPath, lZeroPos - 1)
          Else
            pvCompactPath = sFullPath
        End If
    End If
End Function
