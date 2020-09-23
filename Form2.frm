VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Which Ip To Bind To"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   3555
   LinkTopic       =   "Form2"
   ScaleHeight     =   3045
   ScaleWidth      =   3555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Select The IP Address To Bind To And Hit OK"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ipAddyUDP = List1.List(List1.ListIndex)
Form1.Show
Unload Me
End Sub

Private Sub Form_Load()
Call SocketsInitialize
bleh = Split(QueryIpAddress, vbCrLf)
For i = 0 To UBound(bleh) - 1
List1.AddItem bleh(i)
Next i
End Sub
