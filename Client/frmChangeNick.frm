VERSION 5.00
Begin VB.Form frmChangeNick 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change nick"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Change nick"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Command1 
         Caption         =   "Set nick name"
         Default         =   -1  'True
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter nick name"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmChangeNick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim output As String
    If (Text1.text <> "") Then
        output = "/nick " & Trim(Text1.text)
        MDIForm1.myNick = Trim(Text1.text)
        MDIForm1.Client.SendData output
        Unload Me
    End If
End Sub
