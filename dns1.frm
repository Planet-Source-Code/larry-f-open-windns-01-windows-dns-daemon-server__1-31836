VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "open WinDNS .01"
   ClientHeight    =   765
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   3645
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2640
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label DaemonPort 
      Caption         =   "53"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Daemon Port:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label ServerIP 
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Server IP:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Quit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu Domains 
      Caption         =   "&Domains"
      Begin VB.Menu add 
         Caption         =   "&Add New"
         Begin VB.Menu dom 
            Caption         =   "&Whole Domain"
         End
         Begin VB.Menu A 
            Caption         =   "&A-Record"
         End
         Begin VB.Menu MX1 
            Caption         =   "&MX-Record"
         End
      End
      Begin VB.Menu Edit 
         Caption         =   "&Edit"
         Begin VB.Menu A2 
            Caption         =   "&A-Record"
         End
         Begin VB.Menu MX2 
            Caption         =   "&MX-Record"
         End
      End
      Begin VB.Menu del 
         Caption         =   "&Delete"
         Begin VB.Menu whole 
            Caption         =   "&Whole Domain"
         End
         Begin VB.Menu A3 
            Caption         =   "&A-Record/MX"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Made by Mike G, Please give credit to me if you use my code. Its Lame to say its yours so do it!
' Email Comments and Info to metallica999@metallica.com
Dim remotename As String
Dim remoteip As String
Dim yesno As String
Dim clientip As String
Dim clientMN As String
Dim rxdata As String
Dim SR As String
Dim step As Integer
Dim Command As String
Public start1 As String
Public start2 As String
Dim ServerMN As String





Private Sub A_Click()
AddA.Show 1
End Sub

Private Sub A2_Click()
EditA.Show 1
End Sub

Private Sub A3_Click()
DelA.Show 1
End Sub

Private Sub Command1_Click()
'MsgBox lookUpMX("bong.e-pva.com")
End Sub

Private Sub dom_Click()
AddW.Show 1
End Sub

Private Sub Form_Load()
ServerIP.Caption = ipAddyUDP
Winsock1.Bind "53", ipAddyUDP
'Winsock1.Bind "53"
step = 1
Set dataDNS = New ADODB.Connection
Set recordDNS = New ADODB.Recordset
If Dir(App.Path & "/data.mdb") <> "data.mdb" Then MsgBox "Error, database not found, contact technical support at techsupport@effortlessemail.com": End
With dataDNS
.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "/data.mdb;Persist Security Info=False"
.ConnectionTimeout = 30

End With

dataDNS.Open
recordDNS.Open "Select * from [DomainList];", dataDNS, adOpenKeyset, adLockPessimistic, adCmdText
If recordDNS.EOF Or recordDNS.BOF Then MsgBox "You have no domains listed in this DNS server, please do so"


End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
End Sub



Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text) + 1
If Len(Text1.Text) >= 40000 Then
    Text1.Text = ""
End If
End Sub

Private Sub MX1_Click()
AddMX.Show 1
End Sub

Private Sub MX2_Click()
EditMX.Show 1
End Sub

Private Sub Quit_Click()
Winsock1.Close
Unload Me
End Sub

Private Sub whole_Click()
DelW.Show 1
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error GoTo exits
Dim incom As String

Winsock1.GetData incom, vbString
step = step + 1


Dim rog As String
Clipboard.Clear
For i = 1 To Len(incom)
rog = rog & " " & Asc(Mid(incom, i, 1))
If i = 1 Then start1 = Asc(Mid(incom, i, 1))
If i = 2 Then start2 = Asc(Mid(incom, 2, 1))
Next i

Dim bleh As String
bleh = outPuta(rog)
Winsock1.SendData (bleh)


exits:

End Sub




