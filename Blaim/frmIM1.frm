VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmIM1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Instant Message"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5250
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtIM 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2625
      Width           =   4650
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   4650
      TabIndex        =   1
      Top             =   2625
      Width           =   600
   End
   Begin RichTextLib.RichTextBox rTxtIMs 
      Height          =   2625
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   4630
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmIM1.frx":0000
   End
End
Attribute VB_Name = "frmIM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                                    Option Explicit

Private Sub cmdSend_Click()


    ' The # for which winsock it is in the array is stored in the .tag prop of txtIM
    ' then we send the data

    frmMain.sck(Int(txtIM.Tag)).SendData txtIM.Text
    
     
End Sub

Private Sub txtIM_KeyPress(KeyAscii As Integer)

    ' Just a short-cut for the User.
    If KeyAscii = 13 Then cmdSend_Click
    ' 13 = Enter key on the Key board for those that don't know.

End Sub
