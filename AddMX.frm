VERSION 5.00
Begin VB.Form AddMX 
   Caption         =   "Add MX Domain"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Pref 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   2160
      Width           =   2055
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   4080
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   1920
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox ip 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox dom 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Preference"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "IP Address of MX Domain: "
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Mail Exchange Domain: (mx1.e-pva.com)"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
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
Attribute VB_Name = "AddMX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If recordDNS.State > 0 Then recordDNS.Close
'On Error Resume Next
Dim IDee As Integer
If dom.Text = "" Then MsgBox "MX Domain Name is Blank": Exit Sub
If List1.ListIndex = -1 Then MsgBox "Select a Domain": Exit Sub
If ip.Text = "" Then MsgBox "IP of MX is Blank": Exit Sub
If Pref.Text = "" Then MsgBox "Preference of MX is Blank": Exit Sub
recordDNS.Open "select * from " & List2.List(List1.ListIndex) & ";", dataDNS, adOpenKeyset, adLockPessimistic, adCmdText
Call recordDNS.AddNew
recordDNS("type") = "15"
recordDNS("name") = dom.Text
recordDNS("ip") = ip.Text
recordDNS("additional") = Pref.Text
recordDNS.Update
recordDNS.Close
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
