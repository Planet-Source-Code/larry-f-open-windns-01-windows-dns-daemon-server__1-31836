VERSION 5.00
Begin VB.Form DelW 
   Caption         =   "Delete Whole Domain"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Domain Name:       (such as e-pva.com)   "
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "DelW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If recordDNS.State > 0 Then recordDNS.Close

Dim IDee As Integer

If List1.ListIndex = -1 Then MsgBox "Select a Domain": Exit Sub

recordDNS.Open "Drop Table [" & List2.List(List1.ListIndex) & "];", dataDNS, adOpenKeyset, adLockPessimistic, adCmdText
If recordDNS.State > 0 Then recordDNS.Close
recordDNS.Open "DELETE * FROM DomainList Where ID=" & List2.List(List1.ListIndex) & ";", dataDNS, adOpenKeyset, adLockPessimistic, adCmdText
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If recordDNS.State > 0 Then
recordDNS.Close
End If
recordDNS.Open "select * from DomainList;", dataDNS, adOpenKeyset, adLockPessimistic, adCmdText
Do While Not recordDNS.EOF And Not recordDNS.BOF
List1.AddItem recordDNS.Fields("Domain Name")
List2.AddItem recordDNS.Fields("ID")
recordDNS.MoveNext
Loop
End Sub

Private Sub List4_Click()
'ip.Text = Trim$(List3.List(List4.ListIndex))

End Sub
