Attribute VB_Name = "mWheel"
'================================================
' Module:        mWheel.bas
' Author:        -
' Dependencies:  None
' Last revision: 2003.05.25
'================================================

Option Explicit

'-- API:

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const GWL_WNDPROC   As Long = (-4)
Private Const WM_MOUSEWHEEL As Long = &H20A

'//

'-- Private Variables:
Private m_OldWindowProc     As Long



Public Sub HookWheel()
    
    '-- New Window proc.
    m_OldWindowProc = SetWindowLong(fMain.hWnd, GWL_WNDPROC, AddressOf pvWindowProc)
End Sub

Private Function pvWindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Select Case wMsg
    
        Case WM_MOUSEWHEEL
        
            With fMain.ucCanvas
                If (.DIB.hDIB) Then
                    Select Case wParam
                    Case Is > 0
                        If (.Zoom < 10) Then
                            .Zoom = .Zoom + 1: .Resize: .Repaint
                        End If
                    Case Else
                        If (.Zoom > 1) Then
                            .Zoom = .Zoom - 1: .Resize: .Repaint
                        End If
                    End Select
                End If
            End With
    End Select
    
    pvWindowProc = CallWindowProc(m_OldWindowProc, hWnd, wMsg, wParam, lParam)
End Function
