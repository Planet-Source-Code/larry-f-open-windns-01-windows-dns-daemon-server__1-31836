VERSION 5.00
Begin VB.Form AddW 
   Caption         =   "Add Whole Domain"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox ip 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox ns 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox dom 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "IP Address of Name Server: "
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Name Server:       (such as ns1.e-pva.com)   "
      Height          =   615
      Left            =   120
      TabIndex        =   2
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
Attribute VB_Name = "AddW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If recordDNS.State > 0 Then recordDNS.Close
'On Error Resume Next
Dim IDee As Integer
If dom.Text = "" Then MsgBox "Domain Name is Blank": Exit Sub
If ns.Text = "" Then MsgBox "Name Server is Blank": Exit Sub
If ip.Text = "" Then MsgBox "IP of Name Server is Blank": Exit Sub

recordDNS.Open "select * from DomainList;", dataDNS, adOpenKeyset, adLockPessimistic, adCmdText
Call recordDNS.AddNew("Domain Name", dom.Text)
IDee = recordDNS("ID")
recordDNS.Update
recordDNS.Close
Set recordDNS = dataDNS.Execute("CREATE TABLE [" & IDee & "] (type int, name char(255), IP char(255), Additional char(255));")
'Set recordDNS = Nothing
recordDNS.Open "select * from " & IDee & ";", dataDNS, adOpenKeyset, adLockPessimistic, adCmdText
Call recordDNS.AddNew
recordDNS("type") = "1"
recordDNS("name") = ns.Text
recordDNS("ip") = ip.Text
recordDNS.Update
recordDNS.Close
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
