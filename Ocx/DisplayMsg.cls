VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisplayMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Custom Colors
Private Const vbGrey = 8421504
Private Const vbOffWhite = 16448250

Private Const SW_SHOWNOACTIVATE = 4
Private Const HWND_TOPMOST        As Integer = -1
Private Const HWND_NOTOPMOST      As Integer = -2
Private Const SWP_NOMOVE          As Long = &H2
Private Const SWP_NOSIZE          As Long = &H1
Private Const TOPMOST_FLAGS       As Double = SWP_NOMOVE Or SWP_NOSIZE
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const RGN_DIFF = 4

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    y, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long

Public Sub DisplayMessage(strMsg As String, _
                          clr As Long, _
                          size As Integer, _
                          Optional bBold As Boolean = True, _
                          Optional bItalic As Boolean = False, _
                          Optional bUnderline As Boolean = False, _
                          Optional fName As String = "MS Sans Serif", _
                          Optional fsclr1 As Long = vbGrey, _
                          Optional fsclr2 As Long = vbOffWhite)
Unload frmMessage
With frmMessage
      .Msg = strMsg
      .fClr = clr
      .sze = size
      .bld = bBold
      .itlic = bItalic
      .uLine = bUnderline
      .fName = fName
      .fShadowClr1 = fsclr1
      .fShadowClr2 = fsclr2
       ShowWindow .hWnd, SW_SHOWNOACTIVATE
      .Top = Screen.Height / 3
      .Left = Screen.Height / 3
End With
End Sub
