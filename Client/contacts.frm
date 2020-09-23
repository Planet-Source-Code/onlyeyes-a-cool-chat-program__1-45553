VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   1695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   1695
   Begin VB.CommandButton Command1 
      Caption         =   "STATUS"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   100
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim x As New Form3
    On Error GoTo here
    Form3.SetFocus
    Exit Sub
here:
    x.Show
End Sub
