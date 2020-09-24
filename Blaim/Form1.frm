VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   0
      Width           =   4575
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   7920
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Send"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Host"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7920
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim dTs As String


Private Sub Command1_Click()
Port = InputBox("What port do you want to host on?")
Winsock1.LocalPort = Port
Winsock1.Listen
Winsock2.LocalPort = 100
Winsock2.Listen
End Sub

Private Sub Command2_Click()
host = InputBox("Enter the host's computer name or ip address:")
Port = InputBox("Enter the host's port to connect to:")
Winsock1.Connect host, Port
End Sub

Private Sub Command3_Click()
Text = InputBox("Send what text?")
Winsock1.SendData Text
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then Winsock1.Close
Winsock1.Accept requestID
End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim data As String
Winsock1.GetData data
Text1.Text = Text1.Text & vbNewLine _
                   & data
                   
                    
End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
If Winsock2.State <> sckClosed Then Winsock2.Close
    Winsock2.Accept requestID
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim data As String
Winsock2.GetData data
MsgBox data
End Sub

