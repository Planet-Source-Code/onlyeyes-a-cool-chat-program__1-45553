VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCmd 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   7095
   End
   Begin RichTextLib.RichTextBox txtAll 
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6588
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"commands.frx":0000
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nick As String
Public otherNick As String
Private Sub txtCmd_KeyPress(KeyAscii As Integer)
    Dim parsedata() As String, output As String, out As String
    Dim contor As Long
    
    If (KeyAscii = 13) Then
        out = txtCmd.text
        If (MDIForm1.Client.State = 7) Then
            parsedata = Split(txtCmd.text, " ")
            If (parsedata(0) = "/nick") Then
                MDIForm1.myNick = Trim(parsedata(1))
            ElseIf (parsedata(0) = "msg") Then
                pass = 0
                output = ""
                parsedata = Split(txtCmd.text, " ")
                For contor = 1 To Len(txtCmd.text)
                    If (Mid(txtCmd.text, contor, 1) = " ") Then
                        pass = pass + 1
                    End If
                    If (pass = 3) Then
                        output = Mid(txtCmd.text, contor + 1, Len(txtCmd.text) - contor)
                        Exit For
                    End If
                Next contor
                Me.txtAll.text = Me.txtAll.text & "[" & Format(Timer, "hh:mm:ss") & "] <" & Trim(MDIForm1.myNick) & "> " & output & vbCrLf
                Me.txtAll.SelStart = Len(Me.txtAll.text)
            Else
                If (Me.otherNick = "") And (Me.nick = "") Then
                    out = "/msg* " & MDIForm1.myNick & " " & txtCmd.text
                Else
                    out = "msg " & otherNick & " " & nick & " " & txtCmd.text
                    Me.txtAll.text = Me.txtAll.text & "[" & Format(Timer, "hh:mm:ss") & "] <" & Trim(MDIForm1.myNick) & "> " & txtCmd.text & vbCrLf
                    Me.txtAll.SelStart = Len(Me.txtAll.text)
                End If
            End If
            MDIForm1.Client.SendData Trim(out)
            txtCmd.text = ""
        End If
    End If
End Sub
