VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPeerB 
   Caption         =   "Peer B"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   6615
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox txtOutput 
      Height          =   1335
      Left            =   1200
      TabIndex        =   1
      Top             =   2520
      Width           =   4455
   End
   Begin VB.TextBox txtSend 
      Height          =   1215
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin MSWinsockLib.Winsock udpPeerB 
      Left            =   120
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "frmPeerB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ' The control's name is udpPeerB.
    With udpPeerB
        ' IMPORTANT: be sure to change the RemoteHost
        ' value to the name of your computer.
        .RemoteHost = "127.0.0.1"
        .RemotePort = 1002    ' Port to connect to.
        .Bind 1001                ' Bind to the local port.
    End With
End Sub

Private Sub txtSend_Change()
    ' Send text as soon as it's typed.
    udpPeerB.SendData txtSend.Text
End Sub

Private Sub udpPeerB_DataArrival _
(ByVal bytesTotal As Long)
    Dim strData As String
    udpPeerB.GetData strData
    txtOutput.Text = strData
End Sub

