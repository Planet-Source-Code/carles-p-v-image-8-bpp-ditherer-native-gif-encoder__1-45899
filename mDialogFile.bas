Attribute VB_Name = "mDialogFile"
'================================================
' Module:        mDialogFile.bas
' Author:        Ark (*)
' Dependencies:  None
' Last revision: 2003.03.28
'================================================

' (*)
'
'   From original work:
'   http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=41320&lngWId=1
'   Digital image processing, by Ark

Option Explicit

'-- API:

Private Type OPENFILENAME
    lStructSize       As Long
    hwndOwner         As Long
    hInstance         As Long
    lpstrFilter       As String
    lpstrCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As String
    nMaxFile          As Long
    lpstrFileTitle    As String
    nMaxFileTitle     As Long
    lpstrInitialDir   As String
    lpstrTitle        As String
    Flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type

Private Const OFN_HELPBUTTON      As Long = &H10
Private Const OFN_HIDEREADONLY    As Long = &H4
Private Const OFN_ENABLEHOOK      As Long = &H20
Private Const OFN_ENABLETEMPLATE  As Long = &H40
Private Const OFN_EXPLORER        As Long = &H80000
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_PATHMUSTEXIST   As Long = &H800
Private Const OFN_FILEMUSTEXISTS  As Long = &H1000
Private Const OFN_ENABLESIZING    As Long = &H800000

Private Const OFN_OPENFLAGS       As Long = &H881024
Private Const OFN_SAVEFLAGS       As Long = &H880026

Private Const MAX_PATH            As Long = 260

Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)



'========================================================================================
' Methods
'========================================================================================

Public Function GetFileName(Optional Path As String, Optional Filter As String, Optional FilterIndex As Long = 1, Optional Title As String, Optional OpenDialog As Boolean = -1) As String
   
 Dim OFN  As OPENFILENAME
 Dim lRet As Long
 Dim lFlt As Long
 
    For lFlt = 1 To Len(Filter)
        If (Mid$(Filter, lFlt, 1) = "|") Then
            Mid$(Filter, lFlt, 1) = vbNullChar
        End If
    Next lFlt
    
    If (Len(Filter) < MAX_PATH) Then
        Filter = Filter & String$(MAX_PATH - Len(Filter), 0)
      Else
        Filter = Filter & Chr(0) & Chr(0)
    End If

    With OFN
        .hwndOwner = fMain.hWnd
        .lStructSize = Len(OFN)
        .lpstrFilter = Filter
        .nFilterIndex = FilterIndex
        .lpstrTitle = Title
        .hInstance = App.hInstance
        .lpstrFile = Path & String(MAX_PATH - Len(Path), 0)
        .nMaxFile = MAX_PATH
    End With
    
    If (OpenDialog) Then
        OFN.Flags = OFN.Flags Or OFN_OPENFLAGS
        lRet = GetOpenFileName(OFN)
      Else
        OFN.Flags = OFN.Flags Or OFN_SAVEFLAGS
        lRet = GetSaveFileName(OFN)
    End If
    
    If (lRet) Then
        GetFileName = pvTrimNull(OFN.lpstrFile)
    End If
End Function

'========================================================================================
' Private
'========================================================================================

Private Function pvTrimNull(StartString As String) As String
  
  Dim lPos As Long
  
    lPos = InStr(StartString, Chr$(0))
    If (lPos) Then
        pvTrimNull = Left$(StartString, lPos - 1)
      Else
        pvTrimNull = StartString
    End If
End Function
