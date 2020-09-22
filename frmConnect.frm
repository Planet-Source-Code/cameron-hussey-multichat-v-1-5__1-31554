VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multi chat Connect"
   ClientHeight    =   1335
   ClientLeft      =   7575
   ClientTop       =   2820
   ClientWidth     =   3510
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Text            =   "Name"
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton txtCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton txtConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtIPAddress 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "IP Adress"
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Enter IP Address Of Server:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
End
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtConnect_Click
End If
End Sub

Private Sub txtCancel_Click()
On Error Resume Next
frmConnect.Visible = False
End Sub

Private Sub txtConnect_Click()
On Error Resume Next
frmMultiChat.user.Text = Text1.Text
frmMultiChat.Text1.Text = Text1.Text
frmMultiChat.wsClient.RemotePort = 789
frmMultiChat.wsClient.Connect txtIPAddress.Text
DoEvents
frmConnect.Visible = False
frmMultiChat.Caption = "Connected As Client"
frmMultiChat.mnuconnectasclient.Enabled = False
frmMultiChat.mnustartserver.Enabled = False
frmMultiChat.mnuEndConnection.Enabled = True
frmMultiChat.wsClient.SendData (frmMultiChat.user.Text & " Has Entered the room.")
End Sub

Private Sub txtIPAddress_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
txtConnect_Click
End If
End Sub
