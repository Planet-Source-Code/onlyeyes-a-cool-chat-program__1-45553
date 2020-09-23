VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Server 
      Left            =   4080
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "0"
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub init()
    Server.LocalPort = 10000
    Server.Listen
End Sub


Private Sub Form_Load()
    init
    nrClients = 1
End Sub

Private Sub Form_Terminate()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Server_ConnectionRequest(ByVal requestID As Long)
    'ReDim Preserve serverS(nrClients + 1) As New frmServerSession
    
    serverS(nrClients).ServerSession.Accept requestID
    If (users <> "") Then
        serverS(nrClients).ServerSession.SendData "/nick " & users
    End If
    'serverS(nrClients).Show
    nrClients = nrClients + 1
End Sub

