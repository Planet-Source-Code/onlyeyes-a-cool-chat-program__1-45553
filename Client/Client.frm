VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm MDIForm1 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7560
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11040
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   10440
      Top             =   6960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9960
      Top             =   120
   End
   Begin MSWinsockLib.Winsock Client 
      Left            =   10440
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   10000
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuline 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "&Status"
      End
      Begin VB.Menu mnuline1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuCommands 
      Caption         =   "&Commands"
      Begin VB.Menu mnuChange 
         Caption         =   "Change nick"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public host As String
Public myNick As String
Private Sub getNicks(text As String)
    Dim parsedata() As String
    Dim contor As Long
    
    parsedata = Split(text, ",")
    Form1.List1.Clear
    If (UBound(parsedata) > -1) Then
        For contor = 0 To UBound(parsedata)
            Form1.List1.AddItem Trim(parsedata(contor))
        Next contor
    End If
End Sub

Public Sub init(host As String)
    Client.RemoteHost = Trim(host)
    Client.RemoteHost = 10000
    Client.Connect
End Sub

Private Sub getMsg(text As String)
    Dim contor As Long, pass As Long
    Dim output As String
    Dim parsedata() As String
    
    pass = 0
    output = ""
    parsedata = Split(text, " ")
    For contor = 1 To Len(text)
        If (Mid(text, contor, 1) = " ") Then
            pass = pass + 1
        End If
        If (pass = 3) Then
            output = Mid(text, contor + 1, Len(text) - contor)
            Exit For
        End If
    Next contor
    If (parsedata(1) <> "") Then
        pass = 0
        For contor = 1 To nrComunicatior
            If frmComunicatior(contor).otherNick = Trim(parsedata(2)) Then
                frmComunicatior(contor).txtAll.text = frmComunicatior(contor).txtAll.text & "[" & Format(Timer, "hh:mm:ss") & "] <" & Trim(parsedata(2)) & "> " & output & vbCrLf
                frmComunicatior(contor).txtAll.SelStart = Len(frmComunicatior(contor).txtAll.text)
                pass = 1
            End If
        Next contor
        If (pass = 0) Then
            nrComunicatior = nrComunicatior + 1
            ReDim Preserve frmComunicatior(nrComunicatior) As New Form3
            frmComunicatior(nrComunicatior).nick = Trim(parsedata(1))
            frmComunicatior(nrComunicatior).otherNick = Trim(parsedata(2))
            frmComunicatior(nrComunicatior).txtAll.text = frmComunicatior(nrComunicatior).txtAll.text & "[" & Format(Timer, "hh:mm:ss") & "] <" & Trim(parsedata(2)) & "> " & output & vbCrLf
            frmComunicatior(contor).txtAll.SelStart = Len(frmComunicatior(contor).txtAll.text)
            frmComunicatior(nrComunicatior).Caption = Trim(parsedata(2))
            frmComunicatior(nrComunicatior).Show
        End If
    End If
End Sub

Private Sub getSpecialMsg(text As String)
    Dim contor As Long, pass As Long
    Dim output As String
    Dim parsedata() As String
    
    pass = 0
    output = ""
    parsedata = Split(text, " ")
    For contor = 1 To Len(text)
        If (Mid(text, contor, 1) = " ") Then
            pass = pass + 1
        End If
        If (pass = 2) Then
            output = "<" & parsedata(1) & ">" & Mid(text, contor + 1, Len(text) - contor)
            Exit For
        End If
    Next contor
    
    Form3.txtAll.text = Form3.txtAll.text & output & vbCrLf
End Sub

Private Sub inick(text As String)
    MDIForm1.myNick = ""
    Form3.txtAll = Form3.txtAll.text & text & vbCrLf
    Form3.txtAll.SelStart = Len(Form3.txtAll.text)
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
    Dim parsedata() As String
    Dim strData As String
    
    Client.GetData strData
    parsedata = Split(strData)
    Select Case (LCase(parsedata(0)))
        Case "/nick"
            getNicks (parsedata(1))
        Case "/inick"
            inick (parsedata(1))
        Case "/msg"
            getMsg (strData)
        Case "/msg*"
            getSpecialMsg (strData)
        Case Else
            Form3.txtAll.text = Form3.txtAll.text & strData & vbCrLf
    End Select
End Sub

Private Sub MDIForm_Load()
    Form1.Height = Me.Height - 440
    Form1.Width = 2000
    Form1.Left = Me.Width - Form1.Width - 130
    Form1.Top = 0
    Form1.Show
    
    'Form2.Height = Me.Height - 320
    'Form2.Width = 1800
    'Form2.Top = -60
    'Form2.Left = -60
    'Form2.Show
    
    Form4.Left = Me.Width / 2 - Form4.Width / 2
    Form4.Top = Me.Height / 2 - Form4.Width / 2
    Form4.Show
    Form4.ZOrder (0)

'    Form3.Left = 1700
'    'Form1.Left Form1.Width
'    Form3.Top = 0
'    Form3.Height = 500
'    Form3.Width = 500
    nrComunicatior = 0
    myNick = ""
    Form3.Caption = "Status"
    Form3.Show
    Form3.ZOrder (1)
End Sub

Private Sub MDIForm_Resize()
    If (MDIForm1.WindowState <> 1) Then
        Form1.Height = Me.Height - 440
        Form1.Width = 2000
        Form1.Left = Me.Width - Form1.Width - 130
        Form1.Top = 10

        Form2.Height = Me.Height - 320
        Form2.Width = 1800
        Form2.Top = -60
        Form2.Left = -60
    End If
'    Form3.Left = 1700
'    'Form1.Left Form1.Width
'    Form3.Top = 0
'    Form3.Height = 500
'    Form3.Width = 500
    'Form3.Height = Me.Height
    'Form3.Width = Me.Width - (Form1.Width + Form2.Width)
End Sub

Private Sub mnuChange_Click()
    frmChangeNick.Left = MDIForm1.Width / 2 - frmChangeNick.Width / 2
    frmChangeNick.Top = MDIForm1.Height / 2 - frmChangeNick.Height / 2
    frmChangeNick.Show
End Sub

Private Sub mnuConnect_Click()
    Form4.Left = Me.Width / 2 - Form4.Width / 2
    Form4.Top = Me.Height / 2 - Form4.Height / 2
    Form4.Show
End Sub

Private Sub mnuExit_Click()
    End
End Sub


Private Sub mnuStatus_Click()
    Dim x As New Form3
    On Error GoTo here
    Form3.SetFocus
    Exit Sub
here:
    x.Show
End Sub

Private Sub Timer1_Timer()
    'On Error Resume Next
    'Form3.txtAll.text = Form3.txtAll.text & vbCrLf & Client.State
    If (Client.State <> 7) Then
        mnuChange.Enabled = False
        Me.Client.Close
        Me.myNick = ""
        Me.Client.RemoteHost = Trim(host)
        Me.Client.RemotePort = 10000
        Me.Client.Connect
    ElseIf (Client.State = 7) Then
        If (MDIForm1.myNick = "") Then
            frmChangeNick.Left = Me.Width / 2 - frmChangeNick.Width / 2
            frmChangeNick.Top = Me.Height / 2 - frmChangeNick.Height / 2
            frmChangeNick.Show
            mnuChange.Enabled = True
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    
    If (Client.State <> 7) Then
        Select Case Client.State
            Case 3
                Form3.txtAll.text = Form3.txtAll.text & "Connencting Pending" & vbCrLf
            Case 4
                Form3.txtAll.text = Form3.txtAll.text & "Resolving Host" & vbCrLf
            Case 5
                Form3.txtAll.text = Form3.txtAll.text & "Host Resolved" & vbCrLf
            Case 6
                Form3.txtAll.text = Form3.txtAll.text & "Connecting..." & vbCrLf
            Case 8
                Form3.txtAll.text = Form3.txtAll.text & "Peer is closing the connection" & vbCrLf
            Case 9
                Form3.txtAll.text = Form3.txtAll.text & "Socket Error" & vbCrLf
        End Select
    End If
    
End Sub
