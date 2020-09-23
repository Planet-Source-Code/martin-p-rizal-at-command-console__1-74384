Attribute VB_Name = "modGSM"
Option Explicit

Public Type TextMessage
MsgNumber As String
MsgHeader() As String
MsgBreak() As String
MobileNumber As String
MsgContent As String
End Type




Public NewMessage As TextMessage

Function ParseInput(ByVal strData As String) As String
Dim strBuffer As String
Dim strResponse As String
Dim lngPos As Long
Dim boDone As Boolean
Dim intI As Integer

        '
        ' At least RThreshold bytes are available
        ' Read them and append them to our buffer
        '
       ' strData = DeviceSource.Input
        strBuffer = strBuffer & strData
        
        Do
            '
            ' See if we've got a Prompt response from the 'phone
            ' (indicated by a ">" character)
            '
            lngPos = InStr(strBuffer, ">")
            
            If lngPos = 0 Then
                '
                ' If we haven't then look for the terminating vbCrLf
                '
                lngPos = InStr(strBuffer, vbCrLf)
            End If
            
            If lngPos <> 0 Then
                
                ' Yes - got one of those
                ' Unblock the Response and display it
                ' removing extraneous vbCrLf characters
                ' The 'phone is quite chatty and echos everything
                ' we send it back to us
                '
                strResponse = Mid$(strBuffer, 1, lngPos - 1)
                strResponse = Replace(strResponse, vbCr, "")
                strResponse = Replace(strResponse, vbLf, "")
                strResponse = Replace(strResponse, Chr(26), "")
                '
                ' When we receive a response of +CMGS: we know
                ' that the message has been sent to the Message Centre
                ' and it's on the way to the receiver
                '
                If InStr(strResponse, "ERR") > 0 Then
                    If InStr(strResponse, "500") > 0 Then
                        ParseInput = "Check Operator Service"
                    Else
                        ParseInput = "Message Not Sent"
                    End If
                End If
                
                If InStr(strResponse, "+CMGS") > 0 Then
                    ParseInput = "Message Sent"
                End If
                
                If InStr(strResponse, "+CNMA") > 0 Then
                    ParseInput = "New Message Acknowledgement"
                End If
                
                If InStr(strResponse, "+CNMI") > 0 Then
                    ParseInput = "New Message Indication"
                End If
                
                If InStr(strResponse, "+CMTI") > 0 Then
                    ParseInput = "New Message Received"
                    NewMessage.MsgNumber = GetMsgNumber(strResponse)
                End If
                
                If InStr(strResponse, "+CMGR") > 0 Then
                    NewMessage = ReadMessage(strResponse)
                End If
                
                'ParseInput = ParseInput & strResponse & vbCrLf
                'txtResponse.SelStart = Len(ParseInput) - 2
                '
                ' Are we at the end of our buffer?
                '
                If lngPos + 2 >= Len(strBuffer) Then
                    '
                    ' Yes - there's no more to do
                    ' Flush the buffer and signal to exit the loop
                    '
                    boDone = True
                    strBuffer = ""
                Else
                    '
                    ' No - there's something left that we haven't processsed
                    ' yet. Shift it to the front of the buffer
                    ' and go round the loop
                    '
                    strBuffer = Mid$(strBuffer, lngPos + 2)
                End If
                
            Else
                '
                ' We haven't received a complete response
                ' yet so just exit the loop and wait for
                ' the next comEvReceive event
                '
                boDone = True
            End If
            
        Loop Until boDone = True


End Function

Sub Ping(ByVal DeviceSource As MSComm)
DeviceSource.Output = "AT" & vbCrLf
End Sub

Sub SendSMS(TextInput As String, CellphoneNumber As String, ByVal DeviceSource As MSComm)

With DeviceSource
   .Output = "AT+CMGF=1" & vbCrLf
   .Output = "AT+CMGS=" & Chr(34) & CellphoneNumber & Chr(34) & vbCrLf
   .Output = TextInput & Chr(26)
End With

End Sub


Sub ConnectModem(ByRef DeviceSource As MSComm, PortNumber As Long, Optional BaudRate As Long = 9600)

'===================================
'Settings
'MSComm1.Settings = "115200,N,8,1"      'Change this with the Baud rate of your modem (The one you use with Hyper Terminal)
'MSComm1.CommPort = 3                  ' Change this with the port your Modem is attached,(go to device manager )
'=================================

'SendStat "Connecting to GSM modem on COM" & PortNumber & "..."

 With DeviceSource
 On Error GoTo K
        .CommPort = PortNumber     '    3G UI Pc Interface
        .Settings = BaudRate & ",N,8,1"
        .Handshaking = comRTS
        .RTSEnable = True
        .DTREnable = True
        .RThreshold = 1
        .SThreshold = 1
        .InputMode = comInputModeText
   
        .InputLen = 0
        .PortOpen = True 'must be the last
   
    'SendStat "Checking Parameters..."
    
  '  frmLogin.Text1.Text = frmLogin.Text1.Text & vbCr & buffer
 
    ''SendStat "Connection Successful!", vbBlue
    
   End With
   
Exit Sub

K:

'SendStat Err.Description & " (" & Err.Number & ")"

End Sub

Sub GetSMS(MsgNumber As String, DeviceSource As MSComm)
    DeviceSource.Output = "AT+CMGR=" & MsgNumber & vbCrLf
End Sub

Function ReadMessage(ByVal msgbuf As String) As TextMessage

If ParseSMS(msgbuf) = True Then
Dim I As Long

   ReadMessage.MsgBreak = Split(msgbuf, vbCrLf, , vbTextCompare)
   ReadMessage.MsgHeader = Split(ReadMessage.MsgBreak(0), ",", , vbTextCompare)
   ReadMessage.MobileNumber = Mid(Right(ReadMessage.MsgHeader(1), 11), 1, 10)
   msgbuf = ""
   For I = 1 To UBound(ReadMessage.MsgBreak(), 1)
       msgbuf = msgbuf & ReadMessage.MsgBreak(I) & vbCrLf
   Next I
   ReadMessage.MsgContent = "From: " & ReadMessage.MobileNumber & vbCrLf & vbCrLf & msgbuf
End If

End Function

Public Function ParseSMS(ByVal msgbuf As String) As Boolean
Dim I As Long
Dim StartPoint As Long
Dim EndPoint As Long
Dim Buffer1 As String
Dim Buffer2 As String
Buffer1 = msgbuf
StartPoint = InStr(1, Buffer1, "+CMGR:", vbTextCompare)
EndPoint = InStr(1, Buffer1, vbCrLf & "OK", vbTextCompare)
If StartPoint <> 0 And EndPoint > StartPoint Then
   I = StartPoint
   While I < EndPoint
    Buffer2 = Buffer2 & Mid(Buffer1, I, 1)
    I = I + 1
   Wend
   ParseSMS = True
   msgbuf = Buffer2
   Exit Function
End If
ParseSMS = False
End Function


Function isComplete(ByRef DeviceSource As MSComm) As Boolean

Dim prebuf As String

    Do
    DoEvents
    prebuf = prebuf & DeviceSource.Input
    Loop Until InStr(prebuf, "OK") > 0 Or InStr(prebuf, "ERROR") > 0
    
    If InStr(prebuf, "OK") > 0 Then
            'SendStat "Connection Failed!", vbRed
            'Code here
    isComplete = True
    Exit Function
    ElseIf InStr(prebuf, "ERROR") > 0 Then
            'SendStat "Connection Failed!", vbRed
            'Code here
    isComplete = False
    Exit Function
    End If
    
End Function


Sub DeleteMessage(ByVal MsgNumber As String, DeviceSource As MSComm)
DeviceSource.Output = "AT+CMGD=" & MsgNumber & vbCrLf
End Sub

Function GetMsgNumber(msgbuf As String) As String
Dim temp() As Variant
    temp = Split(msgbuf, ",")
    GetMsgNumber = temp(1)
End Function

Sub SetMsgStorage(Storage1 As String, Storage2 As String, Storage3 As String, DevSource As MSComm)
With DevSource
        .Output = "AT+CNMI=2,1,2,0,0" + Chr(13)  'Set SMS Storage
        .Output = "AT+CPMS=" & Storage1 & "," & Storage1 & "," & Storage1 & "" & Chr(13)
        .Output = "AT+CMGF=1" & vbCrLf
End With
End Sub


Sub ExecuteModem(Command As String, ByRef DevSource As MSComm)
On Error Resume Next
DevSource.Output = Command & vbCrLf
End Sub
