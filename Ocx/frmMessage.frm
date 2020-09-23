VERSION 5.00
Begin VB.Form frmMessage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   406
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HWND_TOPMOST        As Integer = -1
Private Const HWND_NOTOPMOST      As Integer = -2
Private Const SWP_NOMOVE          As Long = &H2
Private Const SWP_NOSIZE          As Long = &H1
Private Const TOPMOST_FLAGS       As Double = SWP_NOMOVE Or SWP_NOSIZE
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const RGN_DIFF = 4

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
Public Msg As String
Public fName As String
Public fClr As Long
Public fShadowClr1 As Long
Public fShadowClr2 As Long
Public sze As Long
Public bld As Boolean
Public itlic As Boolean
Public uLine As Boolean
Public clsTrans As clsTransparency

Private Sub Form_Load()
Set clsTrans = New clsTransparency
PrintMessage
clsTrans.DoTransparency Me, RGB(255, 255, 255)
SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

If Button = vbLeftButton Then
    ReleaseCapture
    SendMessage frmMessage.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End If
End Sub

Private Sub Form_DblClick()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set clsTrans = Nothing
Set frmMessage = Nothing 'good practice to free resources VB doesn't normally free when you unload a form!
End Sub

Public Sub PrintMessage()
Dim curX As Integer
Dim curY As Integer
Dim str() As String
Dim i As Integer

With Me
     .ForeColor = fClr
     .FontName = fName
     .FontSize = sze
     .FontBold = bld
     .FontItalic = itlic
     .FontUnderline = uLine
End With

curX = 10
CurrentY = 10
str = Split(Msg, "||")

If UBound(str()) > 0 Then

   For i = 0 To UBound(str())
       CurrentX = curX
       CurrentY = curY
       Me.ForeColor = fShadowClr2
       Print str(i)
       
       CurrentX = curX - 2
       CurrentY = curY - 2
       Me.ForeColor = fShadowClr1
       Print str(i)
       
       CurrentX = curX - 1
       CurrentY = curY - 1
       Me.ForeColor = fClr
       Print str(i)
       curY = curY + 30
   Next i
 Else
   Me.ForeColor = fShadowClr2
   CurrentX = curX
   CurrentY = curY
   Print str(0)
   
   CurrentX = curX - 2
   CurrentY = curY - 2
   Me.ForeColor = fShadowClr1
   Print str(i)
       
   CurrentX = curX - 1
   CurrentY = curY - 1
   Me.ForeColor = fClr
   Print str(i)
End If

End Sub
