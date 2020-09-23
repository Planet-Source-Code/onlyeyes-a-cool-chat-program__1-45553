VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   1905
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   8640
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    List1.Height = Me.Height + 150
    List1.Width = Me.Width
End Sub

Private Sub Form_Resize()
    List1.Height = Me.Height + 150
    List1.Width = Me.Width
End Sub

Private Sub List1_DblClick()
    nrComunicatior = nrComunicatior + 1
    ReDim Preserve frmComunicatior(nrComunicatior) As New Form3
    frmComunicatior(nrComunicatior).nick = MDIForm1.myNick
    frmComunicatior(nrComunicatior).otherNick = List1.List(List1.ListIndex)
    frmComunicatior(nrComunicatior).Caption = List1.List(List1.ListIndex)
    frmComunicatior(nrComunicatior).Show
End Sub
