VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServerSession 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server Session"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   3720
      Top             =   2760
   End
   Begin MSWinsockLib.Winsock ServerSession 
      Left            =   4200
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmServerSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nick As String

Private Function isFreenick(text As String) As String
    Dim parseData() As String, output As String
    Dim pass As Long, contor As Long, isBusy As Long
    
    pass = 1
    If (nick = "") Then
        If (users <> "") Then
            parseData = Split(users, ",")
            For contor = 0 To UBound(parseData)
                If (LCase(Trim(text)) = LCase(Trim(parseData(contor)))) Then
                    pass = 0
                End If
            Next contor
            If (pass <> 0) Then
                Me.nick = Trim(text)
                users = users & "," & Trim(text)
            End If
        Else
            users = Trim(text)
            Me.nick = Trim(text)
        End If
    Else
        parseData = Split(users, ",")
        pass = 2
        isBusy = 0
        For contor = 0 To UBound(parseData)
            If (LCase(Trim(parseData(contor))) = LCase(Trim(text))) Then
                isBusy = 1
                pass = 0
            End If
        Next contor
        If (isBusy <> 1) Then
            For contor = 0 To UBound(parseData)
                If (LCase(Trim(parseData(contor))) = nick) Then
                    parseData(contor) = Trim(text)
                    pass = 1
                End If
            Next contor
            Me.nick = Trim(text)
            users = Join(parseData, ",")
        End If
    End If
    
    If (pass = 0) Then
        output = "/inick your nickname is allready in use."
    ElseIf (pass = 1) Then
        output = "your nickname is now " & Trim(text)
        For contor = 1 To nrClients
        If (serverS(contor).ServerSession.State = 7) Then
            If (serverS(contor).nick <> Trim(nick)) Then
                serverS(contor).ServerSession.SendData "/nick " & users
            End If
        End If
    Next contor
    ElseIf (pass = 2) Then
        output = "/inick no such nick"
    End If
    isFreenick = output
End Function

Private Sub ServerSession_Close()
    Dim contor As Long, tempCount As Long
    Dim parseData() As String
    Dim tempArr() As String
    
    tempCount = -1
    parseData = Split(users, ",")
    For contor = 0 To UBound(parseData)
        If (LCase(Trim(parseData(contor))) <> LCase(Trim(nick))) Then
            tempCount = tempCount + 1
            ReDim Preserve tempArr(tempCount) As String
            tempArr(tempCount) = parseData(contor)
        End If
    Next contor
    users = Join(tempArr, ",")
    For contor = 1 To nrClients
        If (serverS(contor).ServerSession.State = 7) Then
            If (serverS(contor).nick <> Trim(nick)) Then
                serverS(contor).ServerSession.SendData "/nick " & users
            End If
        End If
    Next contor
    Unload Me
End Sub

Private Sub sendMsg(text As String)
    Dim parseData() As String, msg As String
    Dim contor As Long, pass As Long
    
    parseData = Split(text, " ")
    If (parseData(1) <> "") Then
        pass = 0
        For contor = 1 To Len(text)
            If (Mid(text, contor, 1) = " ") Then
                pass = pass + 1
            End If
            If (pass = 2) Then
                msg = Mid(text, contor + 1, Len(text) - contor)
                Exit For
            End If
        Next contor
        For contor = 1 To nrClients
            If (LCase(serverS(contor).nick) = LCase(Trim(parseData(1)))) Then
                If (serverS(contor).ServerSession.State = 7) Then
                    serverS(contor).ServerSession.SendData "/msg " & parseData(1) & " " & Trim(msg)
                End If
                DoEvents
                Exit For
            End If
        Next contor
    End If
End Sub

Private Sub sendSpecialMsg(text As String)
    Dim parseData() As String, msg As String
    Dim contor As Long, pass As Long
    For contor = 1 To nrClients
            If (serverS(contor).ServerSession.State = 7) Then
                serverS(contor).ServerSession.SendData text
            End If
            DoEvents
    Next contor
End Sub

Private Sub ServerSession_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim parseData() As String
    Dim outBuffer As String, inBuffer As String
    
    ServerSession.GetData strData
    parseData = Split(strData, " ")
    If (UBound(parseData) >= 1) Then
        inBuffer = Trim(parseData(1))
        Select Case LCase(parseData(0))
            Case "/nick"
                outBuffer = outBuffer & isFreenick(inBuffer)
            Case "msg"
                sendMsg (strData)
            Case "/msg*"
                sendSpecialMsg (strData)
        End Select
    End If
    ServerSession.SendData outBuffer
    outBuffer = ""
End Sub

Private Sub Timer1_Timer()
    If (Me.ServerSession.State = 7) Then
        Me.ServerSession.SendData "/nick " & users
    End If
End Sub
