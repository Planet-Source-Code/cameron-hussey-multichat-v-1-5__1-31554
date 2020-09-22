VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMultiChat 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Multi Chat"
   ClientHeight    =   6285
   ClientLeft      =   8025
   ClientTop       =   7320
   ClientWidth     =   9990
   Icon            =   "MultiChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "X"
      Height          =   195
      Left            =   9600
      TabIndex        =   12
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8160
      TabIndex        =   11
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock wsClient 
      Left            =   3960
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsServer 
      Index           =   0
      Left            =   3600
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
      Caption         =   "&Change name"
      Height          =   375
      Left            =   6600
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton mnuEndConnection 
      Caption         =   "&End connection"
      Height          =   375
      Left            =   3360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Send"
      Height          =   255
      Left            =   9000
      TabIndex        =   10
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Whats my IP Adress?"
      Height          =   375
      Left            =   4920
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8160
      TabIndex        =   8
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox user 
      Height          =   285
      Left            =   8160
      TabIndex        =   6
      Text            =   "unknown"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton mnuconnectasclient 
      Caption         =   "&Connect as client"
      Height          =   375
      Left            =   1800
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton mnustartserver 
      Caption         =   "&Start server"
      Height          =   375
      Left            =   240
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   5280
      Width           =   8775
   End
   Begin VB.TextBox txtChatWindow 
      Height          =   4695
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   9495
   End
   Begin VB.Label lblUsersConnected 
      BackColor       =   &H80000007&
      Caption         =   "Total users connected: 0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "frmMultiChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SocketCount As Integer
Dim TotalUsersConnected As Integer

Private Sub Command1_Click()
On Error Resume Next
Text2.Text = user.Text
user.Text = Text1.Text
Text1.Text = ""
wsClient.SendData (Text2.Text & " Has changed his/her name to " & user.Text)
End Sub

Private Sub Command2_Click()
On Error Resume Next
txtMessage.Text = wsClient.LocalIP
End Sub

Private Sub Command3_Click()
On Error Resume Next
If txtMessage.Text = "" Then Exit Sub
If txtMessage.Text = " " Then Exit Sub
If txtMessage.Text = "  " Then Exit Sub
If txtMessage.Text = "   " Then Exit Sub
If txtMessage.Text = "    " Then Exit Sub
    wsClient.SendData (user & ": " & txtMessage.Text)
    DoEvents
    txtMessage.Text = ""
End Sub

Private Sub Command4_Click()
On Error Resume Next
wsClient.SendData (user.Text & " Has left the room.")
DoEvents
End
End Sub

Private Sub Form_Terminate()
On Error Resume Next
Unload frmConnect
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
End
End Sub

Private Sub mnuConnectAsClient_Click()
On Error Resume Next
frmConnect.Visible = True

End Sub

Private Sub mnuConnection_Click()

End Sub

Private Sub mnuEndConnection_Click()
On Error Resume Next
wsClient.SendData (user.Text & " Has left the room.")
frmConnect.txtIPAddress.Text = ""
frmConnect.Text1.Text = ""
wsServer(0).Close
wsClient.Close
mnustartserver.Enabled = True
mnuconnectasclient.Enabled = True
mnuEndConnection.Enabled = False
lblUsersConnected.Visible = False
frmMultiChat.Caption = "Multi Chat"
TotalUsersConnected = 0
lblUsersConnected.Caption = "Total users connected: 0"
txtChatWindow.Text = ""
End Sub

Private Sub mnuEndConnection1_Click()

End Sub

Private Sub mnuExit_Click()
Unload frmConnect
End

End Sub

Private Sub mnuLine0_Click()

End Sub

Private Sub mnuStartServer_Click()
On Error Resume Next
wsServer(0).LocalPort = 789
wsServer(0).Listen
ConnectServerAsClient
frmMultiChat.Caption = "Connected As Server"
mnuconnectasclient.Enabled = False
mnustartserver.Enabled = False
mnuEndConnection.Enabled = True
lblUsersConnected.Visible = True
wsClient.SendData ("Welcome to the room!")
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Command1_Click
End If
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
If txtMessage.Text = "" Then Exit Sub
If txtMessage.Text = " " Then Exit Sub
If txtMessage.Text = "  " Then Exit Sub
If txtMessage.Text = "   " Then Exit Sub
If txtMessage.Text = "    " Then Exit Sub
    wsClient.SendData (user & ": " & txtMessage.Text)
    DoEvents
    txtMessage.Text = ""
End If

End Sub

Private Sub wsClient_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim strDataRecived As String
wsClient.GetData strDataRecived
DoEvents
txtChatWindow.Text = txtChatWindow.Text & strDataRecived & vbCrLf
End Sub

Private Sub wsServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
SocketCount = SocketCount + 1
Load wsServer(SocketCount)
wsServer(SocketCount).Accept requestID
TotalUsersConnected = TotalUsersConnected + 1
lblUsersConnected.Caption = "Total users connected: " & TotalUsersConnected - 1
End Sub


Private Sub wsServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim strRecivedData As String
Dim SocketCheck As Integer
wsServer(Index).GetData strRecivedData
For SocketCheck = 0 To SocketCount Step 1
        If wsServer(SocketCheck).State = sckConnected Then
                wsServer(SocketCheck).SendData strRecivedData
                DoEvents
        End If
Next SocketCheck

End Sub

Public Sub ConnectServerAsClient()
On Error Resume Next
frmMultiChat.wsClient.RemotePort = 789
frmMultiChat.wsClient.Connect wsClient.LocalIP

End Sub
