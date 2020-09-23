Attribute VB_Name = "Module1"

Public Function outPuta(inputAsInt As String) As String
On Error GoTo ender
' the input from query as integers
Dim eachInputInt() As String
' split apart each integer to handle
Dim domainString As String
'domain requested for
Dim domainBin As String
'original domain query
Dim pointAt As Integer
'keep position of int
Dim requestType As String
'type of request, mx , A
Dim typeOfType As String
'type of request inet..
inputAsInt = Mid(inputAsInt, 2, Len(inputAsInt))
'slice down first space
eachInputInt = Split(inputAsInt, " ")
'split up the ints
startstring = Chr(eachInputInt(0)) & Chr(eachInputInt(1))
'when replying it must have the same start
pointAt = 12
'skip to beginning of domain
Do While pointAt < UBound(eachInputInt)
length = eachInputInt(pointAt)
'since DNS does not like using periods it instead uses length of section
If length = 0 Then Exit Do
'ok done with domain
If domainBin = "" Then domainBin = domainBin & Chr(eachInputInt(pointAt))
For i = 1 To length
domainString = domainString & Chr(eachInputInt(pointAt + i))
domainBin = domainBin & Chr(eachInputInt(pointAt + i))
Next i
'add all characters up to that length (next period or end)
pointAt = pointAt + length + 1
'move position holder to next period or end
If Not eachInputInt(pointAt) = 0 Then
'not end
domainBin = domainBin & Chr(eachInputInt(pointAt))
domainString = domainString & "."
End If
Loop
pointAt = pointAt + 2
'skip to the query type, mx or A, etc...
requestType = eachInputInt(pointAt)

pointAt = pointAt + 2
'skip to type such as inet
typeOfType = eachInputInt(pointAt)
''''' time to make reply
Dim reply As String

Dim domainAnswer As String
Dim domainSection As String
Dim prefEr As String
'preference/priority
reply = startstring
If Not isDomainHosted(domainString) Then reply = reply & Chr(0) & Chr(0) & Chr(0): GoTo ender
'set the start to the same
reply = reply & Chr(133) & Chr(128)
'i dont know what the hell this does
reply = reply & Chr(0) & Chr(1)
'says it is answering one query
reply = reply & Chr(0) & Chr(1)
'says it has found these many answers
reply = reply & Chr(0) & Chr(0)
'says it has found these ns servers
reply = reply & Chr(0) & Chr(0)
'says how many extra things there are
reply = reply & domainBin
'say the domain you are replying to
reply = reply & Chr(0) & Chr(0)
'blank space
reply = reply & Chr(requestType) & Chr(0)
'telling the reply type
reply = reply & Chr(1) & Chr(192) & Chr(12)
'the first part identifies how many resuls, the second and third part identifies the pointer char and where the domain starts
reply = reply & Chr(0) & Chr(requestType)
'tell it the request type mx, a, etc...
reply = reply & Chr(0) & Chr(typeOfType)
Select Case requestType
Case 1
reply = reply & Chr(0) & Chr(0)
'tell it the RR Class (whatever that shit is)
reply = reply & Chr(81) & Chr(129) & Chr(0)
'time to live ... dont know how it is formated
'now we have to format the answer... which is a biatch so lets get started
reply = reply & lookUpIP(domainString) & Chr(0)
'reply = reply & Chr(4) & Chr(24) & Chr(189) & Chr(121) & Chr(87)
'MsgBox lookUpIP(domainString)
Case 15
'tell it the stupid type inet
reply = reply & Chr(0) & Chr(1)
'tell it the RR Class (whatever that shit is)
reply = reply & Chr(81) & Chr(129)
'time to live ... dont know how it is formated
'now we have to format the answer... which is a biatch so lets get started
domainAnswer = domainCompressed(domainString, lookUpMX(domainString, prefEr))

'compressed mx answer
domainSection = Chr(Len(domainAnswer))
'number of octets taken up by answer
reply = reply & Chr(0) & domainSection
'tell the number of octets
reply = reply & Chr(0) & Chr(Int(prefEr))
'priority
reply = reply & domainAnswer
reply = reply & Chr(0)
End Select
ender:
outPuta = reply

End Function
Public Function domainCompressed(domainString As String, mxString As String) As String
If mxString = domainString Then
    domainCompressed = Chr(192) & Chr(12)
    Exit Function
End If
Dim mxCopy As String
mxCopy = mxString
Dim mxDomainArry() As String
Dim finalDomain As String
Dim mxHold As String
Dim dmHold As String
Dim findPeriodM As Integer
Dim findPeriodD As Integer
Dim periodCountM As Integer
Dim periodCountD As Integer
Dim periodLocM As String
Dim periodLocD As String
Dim periodArryM() As String
Dim periodArryD() As String

periodCountM = 0
periodCountD = 0
findPeriodM = 0
findPeriodD = 0

While Not InStr(findPeriodD + 1, domainString, ".") < 1
    findPeriodD = InStr(findPeriodD + 1, domainString, ".")
    If Not periodLocD = "" Then
        periodLocD = periodLocD & " " & findPeriodD
    Else
        periodLocD = findPeriodD
    End If
    periodCountD = periodCountD + 1
Wend
periodArryD = Split(periodLocD, " ")

While Not InStr(findPeriodM + 1, mxString, ".") < 1
    findPeriodM = InStr(findPeriodM + 1, mxString, ".")
    If Not periodLocM = "" Then
        periodLocM = periodLocM & " " & findPeriodM
    Else
        periodLocM = findPeriodM
    End If
    periodCountM = periodCountM + 1
Wend
periodArryM = Split(periodLocM, " ")


For m = -1 To UBound(periodArryD)
    For i = -1 To UBound(periodArryM)
    trueBubble = False
        For z = UBound(periodArryM) + 1 To i + 1 Step -1
        If m = -1 Then
            dmHold = domainString
        Else
            dmHold = Mid(domainString, periodArryD(m) + 1, Len(domainString) - periodArryD(m))
        End If
        If i = -1 And z = UBound(periodArryM) + 1 Then
            mxHold = mxString
            
        ElseIf i = -1 And Not z = UBound(periodArryM) + 1 Then
            mxHold = Mid(mxString, 1, periodArryM(z) - 1)
            
        ElseIf Not i = -1 And z = UBound(periodArryM) + 1 Then
            mxHold = Mid(mxString, periodArryM(i) + 1, Len(mxString) - periodArryM(i))
            
        Else
            mxHold = Mid(mxString, periodArryM(i) + 1, periodArryM(z) - (periodArryM(i) + 1))
            
        End If
        If mxHold = dmHold Then mxCopy = Replace(mxCopy, mxHold, Chr(192 + 11 + Int(InStr(1, domainString, mxHold))))
        Next z
    Next i

Next m
Dim good As String
Dim whoaPeriodLength() As String
whoaPeriodLength = Split(mxCopy, ".")
Dim whoaPeriods As Integer
good = Chr(Len(whoaPeriodLength(0)))
whoaPeriods = 1
For i = 1 To Len(mxCopy)
If Not Asc(Mid(mxCopy, i, 1)) > 192 And Not Mid(mxCopy, i, 1) = "." Then
good = good & Mid(mxCopy, i, 1)
'End If
ElseIf Mid(mxCopy, i, 1) = "." And Not i = Len(mxCopy) - 1 Then
If Not Asc(Mid(mxCopy, i + 1, 1)) > 192 Then
    good = good & Chr(Len(whoaPeriodLength(whoaPeriods)))
    whoaPeriods = whoaPeriods + 1
    Else
    whoaPeriods = whoaPeriods + 1
    End If
ElseIf Mid(mxCopy, i, 1) = "." And i = Len(mxCopy) - 1 Then
whoaPeriods = whoaPeriods + 1

'End If
ElseIf Asc(Mid(mxCopy, i, 1)) > 192 And Not i = Len(mxCopy) Then
good = good & Chr(192) & Chr(Asc(Mid(mxCopy, i, 1)) - 192)

ElseIf Asc(Mid(mxCopy, i, 1)) > 192 And i = Len(mxCopy) Then
good = good & Chr(192) & Chr(Asc(Mid(mxCopy, i, 1)) - 192) & Chr(192) & Chr(12)
End If
Next i
'For i = 1 To Len(good)
'MsgBox Mid(good, i, 1)
'Next i
domainCompressed = good
End Function
Function isDomainHosted(domainString As String) As Boolean
Dim endPart As String
Dim Broken() As String
If recordDNS.State > 0 Then
recordDNS.Close
End If
Broken = Split(domainString, ".")
endPart = Broken(UBound(Broken) - 1) & "." & Broken(UBound(Broken))

recordDNS.Open "select * from DomainList WHERE [Domain Name] = '" & endPart & "';", dataDNS, adOpenKeyset, adLockPessimistic, adCmdText
If recordDNS.EOF Or recordDNS.BOF Then
isDomainHosted = False
Else
isDomainHosted = True
End If
End Function
Function lookUpIP(domainString As String) As String
Dim endPart As String
Dim Broken() As String
Dim preIP As String
Dim ipArry() As String
Dim tablA As String
Dim Answer As String
If recordDNS.State > 0 Then
recordDNS.Close
End If
Broken = Split(domainString, ".")
endPart = Broken(UBound(Broken) - 1) & "." & Broken(UBound(Broken))
recordDNS.Open "select * from DomainList WHERE [Domain Name] = '" & endPart & "';", dataDNS, adOpenKeyset, adLockPessimistic, adCmdText
tablA = recordDNS.Fields("ID")
recordDNS.Close
recordDNS.Open "select * from [" & tablA & "] WHERE Name='" & domainString & "';", dataDNS, adOpenKeyset, adLockPessimistic, adCmdText
If Not recordDNS.EOF And Not recordDNS.BOF Then
happy:
preIP = recordDNS.Fields("IP")
ipArry = Split(preIP, ".")
Answer = Chr(4)
For Each spot In ipArry
Answer = Answer & Chr(spot)
Next
lookUpIP = Answer
Exit Function
End If
recordDNS.Close
recordDNS.Open "select * from [" & tablA & "] WHERE Name='www." & endPart & "';", dataDNS, adOpenKeyset, adLockPessimistic, adCmdText
If Not recordDNS.EOF And Not recordDNS.BOF Then GoTo happy
recordDNS.Close
recordDNS.Open "select * from [" & tablA & "];", dataDNS, adOpenKeyset, adLockPessimistic, adCmdText
If Not recordDNS.EOF And Not recordDNS.BOF Then GoTo happy
lookUpIP = Chr(4) & Chr(0) & Chr(0) & Chr(0) & Chr(0)
End Function
Function lookUpMX(domainString As String, ByRef preF As String) As String
Dim endPart As String
Dim Broken() As String
Dim tablA As String
Dim Answer As String
If recordDNS.State > 0 Then
recordDNS.Close
End If
Broken = Split(domainString, ".")
endPart = Broken(UBound(Broken) - 1) & "." & Broken(UBound(Broken))
recordDNS.Open "select * from DomainList WHERE [Domain Name] = '" & endPart & "';", dataDNS, adOpenKeyset, adLockPessimistic, adCmdText
tablA = recordDNS.Fields("ID")
recordDNS.Close
recordDNS.Open "select * from [" & tablA & "] WHERE type=15;", dataDNS, adOpenKeyset, adLockPessimistic, adCmdText
If recordDNS.EOF Or recordDNS.BOF Then
Answer = "mx1." & endPart
preF = 50
Else
Answer = Trim$(recordDNS.Fields("Name"))
preF = Trim$(recordDNS.Fields("Additional"))
End If
lookUpMX = Answer

End Function

