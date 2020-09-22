VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   ScaleHeight     =   990
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   3000
      Top             =   240
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "This box will show the Hwnd of the window that the mouse is over."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
  x As Long
  y As Long
End Type
Private Sub Form_Load()
Text1.Text = ""
End Sub

Private Sub Timer1_Timer()
Dim CursorPos As POINTAPI
GetCursorPos CursorPos
Text1.Text = (WindowFromPoint(CursorPos.x, CursorPos.y))
End Sub

