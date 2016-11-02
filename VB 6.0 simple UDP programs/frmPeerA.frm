VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPeerA 
   Caption         =   "Peer A"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   8490
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox txtOutput 
      Height          =   1335
      Left            =   2280
      TabIndex        =   1
      Top             =   2400
      Width           =   3975
   End
   Begin VB.TextBox txtSend 
      Height          =   1215
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin MSWinsockLib.Winsock udpPeerA 
      Left            =   240
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "frmPeerA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ' The control's name is udpPeerA
    With udpPeerA
        ' IMPORTANT: be sure to change the RemoteHost
        ' value to the name of your computer.
        .RemoteHost = "127.0.0.1"
        .RemotePort = 1001   ' Port to connect to.
        .Bind 1002                ' Bind to the local port.
    End With
    frmPeerB.Show                 ' Show the second form.
End Sub

Private Sub txtSend_Change()
    ' Send text as soon as it's typed.
    udpPeerA.SendData txtSend.Text
End Sub

Private Sub udpPeerA_DataArrival _
(ByVal bytesTotal As Long)
    Dim strData As String
    udpPeerA.GetData strData
    txtOutput.Text = strData
End Sub


