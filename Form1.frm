VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Transparent Msg Generator"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form2"
   ScaleHeight     =   2955
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkUnderline 
      Caption         =   "Underline"
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.TextBox txtfont 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Text            =   "Bookman Old Style"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CheckBox chkItalic 
      Caption         =   "Italic"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   3360
      Max             =   32
      Min             =   8
      TabIndex        =   5
      Top             =   1560
      Value           =   8
      Width           =   255
   End
   Begin VB.TextBox txtSize 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Text            =   "18"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   -600
   End
   Begin VB.TextBox txtMsg 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Wow||There's text on my desktop.||I've never seen that before!"
      Top             =   960
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Message"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Font Size:"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   10
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Font Face:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Message Text:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Multiline msg's can be seperated with ||"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Custom Colors
Private Const vbGrey = 8421504
Private Const vbOffWhite = 16448250

Private Sub Command1_Click()
                         ' strMsg As String, _
                         ' clr As Long, _
                         ' size As Integer, _
                         ' Optional bBold As Boolean = True, _
                         ' Optional bItalic As Boolean = False, _
                         ' Optional bUnderline As Boolean = False, _
                         ' Optional fName As String = "MS Sans Serif", _
                         ' Optional fsclr1 As Long = vbGrey, _
                         ' Optional fsclr2 As Long = RGB(250, 250, 250))
Dim objPlugIn As Object       'use a variable to define the plugin
Dim strResponse As String     'Variable contains plugin's response
Dim Indentity As String

On Error GoTo ErrHandler
'The format for CreateObject is [Project name].[Class module name]
Set objPlugIn = CreateObject("drmTransMsg.DisplayMsg")
    
'Call the entry function
strResponse = objPlugIn.DisplayMessage(txtMsg, _
                                       vbRed, _
                                       txtSize, _
                                       CBool(chkBold.Value), _
                                       CBool(chkItalic.Value), _
                                       CBool(chkUnderline.Value), _
                                       txtfont, _
                                       vbGrey, _
                                       vbOffWhite)
'if the plugin contains an error, show us in a message box
If strResponse <> vbNullString Then
   MsgBox strResponse
End If
  
Me.SetFocus
Exit Sub

ErrHandler:
Select Case Err.Number
    Case 429 'can't create object
        'The ProgID can't be found. Either it is misspelled or the component hasn't been registered!
        MsgBox "You have selected an invalid plug-in ID. Please check that the name is correct and the component is registered."
        Exit Sub
    Case 5 'Invalid proceedure call or argument
        'The 'DisplayAlert' function cannot be found in the class module
        MsgBox "The plug-in you have selected does not have a valid entry point. Please verify the object module with specified guidelines."
        Exit Sub
    Case Else
        MsgBox Err.Number & "  " & Err.Description
End Select
End Sub
 
Private Sub VScroll1_Change()
txtSize = VScroll1.Value
End Sub
