VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configure"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   2790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Find A Forms hWnd"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start Transparancy"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop Transparancy"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "This example uses me.hwnd, you can make any thing transparent just find its hwnd using the button below"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim strinput As String
strinput = InputBox("Please type in the hWnd of the window you want to make not transparent - ( The Default value in the box is this forms hWnd )", "Hwnd", Me.hWnd)
MakeNotTransparent strinput
End Sub


Private Sub Command2_Click()
Dim strinput As String
Dim strValue As String
strinput = InputBox("Please type in the hWnd of the window you want to make transparent - ( The Default value in the box is this forms hWnd )", "Hwnd", Me.hWnd)
strValue = InputBox("What Level of transparancy would you like to set?", "Transparacy Level", "100")
If strValue < "1" Then GoTo oops
If strValue > "255" Then GoTo oops
MakeTransparent strValue, strinput
Exit Sub
oops:
MsgBox "Please Enter A Number From 1 to 255", vbCritical, "Error"
End Sub

Private Sub Command3_Click()
Form1.Show 1, frmOptions
End Sub
