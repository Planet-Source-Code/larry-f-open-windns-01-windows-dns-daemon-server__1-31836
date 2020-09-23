Attribute VB_Name = "Module2"

' *********************************************************
' * Code was not originally written by me (Master Yoda).  *
' * I didn't spend a lot of time to fixed up this code so *
' * feel free to do so if nessesary.                      *
' *********************************************************

Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1
Public Const WSADescription_Len = 256
Public Const WSASYS_Status_Len = 128

Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Public Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type
Public ipAddyUDP As String
Public dataDNS As ADODB.Connection
Public recordDNS As ADODB.Recordset
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal Src As Long, ByVal cb As Long)
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)

Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function


Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function


Sub SocketsInitialize()
    Dim WSAD As WSADATA
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String
    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)


    If iReturn <> 0 Then
        MsgBox "Winsock.dll Error."
        End
    End If
    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
        WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then
        sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
        sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
        sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
        'sMsg = sMsg & " winsock.dll tarafindan desteklenmiyor. "
        MsgBox sMsg
        End
    End If

End Sub

Public Function CurrentIP(ReturnExternalIP As Boolean)
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String
    Dim ip As String

    If gethostname(hostname, 255) = SOCKET_ERROR Then
        MsgBox "Windows Socket Error " & Str(WSAGetLastError())
        Exit Function
    Else
        hostname = Trim$(hostname)
    End If
    hostent_addr = gethostbyname(hostname)


    If hostent_addr = 0 Then
        MsgBox "Winsock.dll error."
        Exit Function
    End If
    RtlMoveMemory host, hostent_addr, LenB(host)
    RtlMoveMemory hostip_addr, host.hAddrList, 4
   
    Do
        ReDim temp_ip_address(1 To host.hLength)
        RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength


        For i = 1 To host.hLength
            ip_address = ip_address & temp_ip_address(i) & "."
        
        Next
frmMain.List1.AddItem ip_address
        ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)

    ' Return Both LAN and External IP Fix
    ' Master Yoda 30-05-2000
    ' ##########################################
    '   HERE'S THE PROBLEM!!!
        'TheIP = TheIP + ip_address
    ' ##########################################
    '   HERE'S THE FIX!!!
        Internal = theip        ' Send ONLY the External IP to the CurrentIP Function
        EXTERNAL = ip_address   ' Send the External IP to the  function parameter External
        theip = ip_address      ' Send LAN IP to the function para Internal
        
        ' You don't really need to return parameters,
        ' it just allows you to get both IPs :)
    ' ##########################################

        ip_address = ""
        host.hAddrList = host.hAddrList + LenB(host.hAddrList)
        RtlMoveMemory hostip_addr, host.hAddrList, 4
    Loop While (hostip_addr <> 0)
    
    
If ReturnExternalIP = True Then
    CurrentIP = EXTERNAL
Else
    CurrentIP = Internal
End If
End Function

Sub SocketsCleanup()
    Dim lReturn As Long
    lReturn = WSACleanup()


    If lReturn <> 0 Then
        MsgBox "Socket Error " & Trim$(Str$(lReturn)) & " occurred In Cleanup "
        End
    End If
End Sub

Public Function QueryIpAddress() As String
 Dim hostname As String * 256
 Dim hostent_addr As Long
 Dim host As HOSTENT
 Dim hostip_addr As Long
 Dim Ip_Addresses As String
 Dim Tmp_ip_address(1 To 4) As Byte
 Dim i As Integer
 
 If gethostname(hostname, 256) = -1 Then _
  QueryIpAddress = "": Exit Function
 hostname = Trim$(hostname)

 hostent_addr = gethostbyname(hostname)
 If hostent_addr = 0 Then _
  QueryIpAddress = "": Exit Function
  
 Call MemCopy(host, hostent_addr, 16)
 
 Call MemCopy(hostip_addr, host.hAddrList, 4)
 
 Do
  Call MemCopy(Tmp_ip_address(1), hostip_addr, 4)

  For i = 1 To 4
   Ip_Addresses = Ip_Addresses & Tmp_ip_address(i) & "."
  Next i
  Ip_Addresses = Left$(Ip_Addresses, Len(Ip_Addresses) - 1) + vbCrLf

  host.hAddrList = host.hAddrList + 4
  Call MemCopy(hostip_addr, host.hAddrList, 4)
 Loop While (hostip_addr <> 0)
 QueryIpAddress = Ip_Addresses
End Function



