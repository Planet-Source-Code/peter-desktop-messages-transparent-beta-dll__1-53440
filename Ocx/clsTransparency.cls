VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransparency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'********************************* _
  Copyright © 2004 Peter Hart _
 *********************************
' Email:  peterbhart@hotmail.com _
 ********************************** _
  Additional Terms of Agreement: _
  You MAY NOT Sell This Code _
  You MAY NOT Sell Any Program Containing This Code _
  You use this code knowing I hold no responsibilities for any results _
  occuring from the use and/or misuse of this code

'This was faster than my own routine so I exchanged it!!
'Transparency Class
'Credits to: Chris Yates,
'who wrote this module.

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Sub DoTransparency(Frm As Form, transColor)
Dim rgn     As Long
Dim rgn2    As Long
Dim rgn3    As Long
Dim rgn4    As Long
Dim X1      As Long
Dim Y1      As Long
Dim i       As Long
Dim j       As Long
Dim tj      As Long

rgn = CreateRectRgn(0, 0, 0, 0)
rgn2 = CreateRectRgn(0, 0, 0, 0)
rgn3 = CreateRectRgn(0, 0, 0, 0)
i = 1

With Frm
     X1 = .Width / Screen.TwipsPerPixelX
     Y1 = .Height / Screen.TwipsPerPixelY
'     Object must have these setting's
     .AutoRedraw = True
     .ScaleMode = 3
End With

Do While i < X1
    j = 1
    Do While j < Y1
        If GetPixel(Frm.hDC, i, j) <> transColor Then
            tj = j
            Do While GetPixel(Frm.hDC, i, j + 1) <> transColor
                j = j + 1
                If j = Y1 Then Exit Do
            Loop
            rgn4 = CreateRectRgn(i, tj, i + 1, j + 1)
        
            CombineRgn rgn3, rgn2, rgn2, 5
            CombineRgn rgn2, rgn4, rgn3, 2
        
            DeleteObject rgn4
        End If
    j = j + 1
    Loop
    CombineRgn rgn3, rgn, rgn, 5
    CombineRgn rgn, rgn2, rgn3, 2
    i = i + 1
Loop
SetWindowRgn Frm.hWnd, rgn, True
'Clean up
DeleteObject rgn
DeleteObject rgn2
DeleteObject rgn3

End Sub
