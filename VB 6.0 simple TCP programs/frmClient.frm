VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   Caption         =   "TCP Client"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   10935
   LinkTopic       =   "Form2"
   ScaleHeight     =   4815
   ScaleWidth      =   10935
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdConnect 
      Caption         =   "连接服务器"
      Height          =   735
      Left            =   7440
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtOutput 
      Height          =   855
      Left            =   3120
      TabIndex        =   1
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox txtSend 
      Height          =   735
      Left            =   3120
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   360
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label receive 
      Caption         =   "接收的数据："
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Send 
      Caption         =   "发送的数据："
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ' The name of the Winsock control is tcpClient.
    ' Note: to specify a remote host, you can use
    ' either the IP address (ex: "121.111.1.1") or
    ' the computer's "friendly" name, as shown here.
    tcpClient.RemoteHost = "127.0.0.1"
    tcpClient.RemotePort = 1001
End Sub

Private Sub cmdConnect_Click()
    ' Invoke the Connect method to initiate a
    ' connection.
    tcpClient.Connect
End Sub

Private Sub txtSend_Change()
    tcpClient.SendData txtSend.Text
End Sub

Private Sub tcpClient_DataArrival _
(ByVal bytesTotal As Long)
    Dim strData As String
    tcpClient.GetData strData
    txtOutput.Text = strData
End Sub



