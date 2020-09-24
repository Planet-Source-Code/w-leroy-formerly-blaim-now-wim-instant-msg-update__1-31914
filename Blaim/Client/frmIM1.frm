VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIM1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blaim Client"
   ClientHeight    =   3000
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   2625
      Width           =   600
   End
   Begin VB.TextBox txtIM 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2625
      Width           =   4650
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtxtIM 
      Height          =   2625
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   4630
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmIM1.frx":0000
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuIM 
      Caption         =   "&IM"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuSend 
         Caption         =   "&Send"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmIM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                                    Option Explicit
' Declarations for winsock
    Dim pORT1 As String
        Dim dATA1 As String
            Dim Host As String

Private Sub cmdSend_Click()
SendCommand txtIM.Text

            
End Sub
Private Sub SendCommand(ByVal data As Variant)
 Winsock1.SendData data
 rtxtIM.Text = rtxtIM.Text & txtIM.Text & vbNewLine
 txtIM.Text = ""
 
End Sub
Private Sub mnuConnect_Click()
' Save the address of the server here
    Host = InputBox("Enter the IP or name of the Server?", "Server:")
        ' What port is the Server on?
            'pORT1 = InputBox("What port is this Server on?", "Enter the Port :")
                ' Now connect winsock to the previous entered info.
                    Winsock1.Connect Host, 9456
                    
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuSend_Click()
' Again this will send the msg also
    cmdSend_Click
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
' If someone is sending us a msg we need to handle it.
    Winsock1.GetData dATA1
        ' now lets add it to our txtbox for us to view
            rtxtIM.Text = rtxtIM.Text & dATA1 & vbNewLine
End Sub


