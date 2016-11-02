VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Caption         =   "TCP Server"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   8325
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtOutput 
      Height          =   975
      Left            =   4200
      TabIndex        =   1
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtSendData 
      Height          =   975
      Left            =   4200
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin MSWinsockLib.Winsock tcpServer 
      Left            =   240
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label receive 
      Caption         =   "接收的数据："
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Send 
      Caption         =   "发送的数据："
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ' Set the LocalPort property to an integer.
    ' Then invoke the Listen method.
    tcpServer.LocalPort = 1001
    tcpServer.Listen
    frmClient.Show ' Show the client form.
End Sub

Private Sub tcpServer_ConnectionRequest _
(ByVal requestID As Long)
    ' Check if the control's State is closed. If not,
    ' close the connection before accepting the new
    ' connection.
    If tcpServer.State <> sckClosed Then _
    tcpServer.Close
    ' Accept the request with the requestID
    ' parameter.
    tcpServer.Accept requestID
End Sub

Private Sub txtSendData_Change()
    ' The TextBox control named txtSendData
    ' contains the data to be sent. Whenever the user
    ' types into the  textbox, the  string is sent
    ' using the SendData method.
    tcpServer.SendData txtSendData.Text
End Sub

Private Sub tcpServer_DataArrival _
(ByVal bytesTotal As Long)
    ' Declare a variable for the incoming data.
    ' Invoke the GetData method and set the Text
    ' property of a TextBox named txtOutput to
    ' the data.
    Dim strData As String
    tcpServer.GetData strData
    txtOutput.Text = strData
End Sub

