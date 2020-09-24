VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Grabber"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Default         =   -1  'True
      Height          =   495
      Left            =   1740
      TabIndex        =   1
      Top             =   1305
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   300
      Left            =   270
      MousePointer    =   8  'Größenänderung NW SO
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2640
      Width           =   645
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Enum GrabberConstants
    SM_CYHSCROLL = 3
    GWL_STYLE = -16
    GrabberBit = &H10
End Enum

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Form_Load()

  Dim PrevScaleMode     As Integer
  Dim Grabbersize       As Long

    PrevScaleMode = ScaleMode
    ScaleMode = vbPixels
    Grabbersize = GetSystemMetrics(SM_CYHSCROLL)
    With HScroll1
        'to prevent a blinking thumb the form must have at least one other control
        'which can receive the focus
        .TabStop = False
        .MousePointer = vbSizeNWSE
        .Move 0, 0, Grabbersize, Grabbersize
        SetWindowLong .hWnd, GWL_STYLE, GetWindowLong(.hWnd, GWL_STYLE) Or GrabberBit
    End With 'HSCROLL1
    ScaleMode = PrevScaleMode

End Sub

Private Sub Form_Resize()

    With HScroll1
        .Move ScaleWidth - .Width, ScaleHeight - .Height
    End With 'HSCROLL1

End Sub

':) Ulli's VB Code Formatter V2.19.3 (2005-Apr-18 02:02)  Decl: 10  Code: 37  Total: 47 Lines
':) CommentOnly: 2 (4,3%)  Commented: 2 (4,3%)  Empty: 11 (23,4%)  Max Logic Depth: 2
