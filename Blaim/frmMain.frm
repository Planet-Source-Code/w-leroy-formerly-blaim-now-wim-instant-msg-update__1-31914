VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H0063C7ED&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WIM"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data db 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   11040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Width           =   1140
   End
   Begin VB.TextBox tWarn 
      DataSource      =   "db"
      Height          =   495
      Left            =   11280
      TabIndex        =   15
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox tFro 
      DataSource      =   "db"
      Height          =   495
      Left            =   11280
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox tDate 
      DataSource      =   "db"
      Height          =   495
      Left            =   11280
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox tC 
      DataSource      =   "db"
      Height          =   495
      Left            =   11280
      TabIndex        =   12
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox tPW 
      DataSource      =   "db"
      Height          =   495
      Left            =   11280
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox tID 
      DataSource      =   "db"
      Height          =   495
      Left            =   11280
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock server 
      Left            =   4320
      Top             =   5175
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9457
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   6960
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   7560
      Picture         =   "frmMain.frx":1594
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   4920
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Frame fMenu 
      BackColor       =   &H0063C7ED&
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9015
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H80000018&
         Height          =   495
         Index           =   2
         Left            =   8400
         Picture         =   "frmMain.frx":1E5E
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Violate User"
         Top             =   170
         Width           =   495
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H80000018&
         DownPicture     =   "frmMain.frx":2728
         Height          =   495
         Index           =   1
         Left            =   735
         Picture         =   "frmMain.frx":2FF2
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Show Client Connections"
         Top             =   170
         Width           =   495
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H80000018&
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":38BC
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Turn Server On"
         Top             =   170
         Width           =   495
      End
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4560
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
            MinWidth        =   3969
            Text            =   "Server State here"
            TextSave        =   "Server State here"
            Key             =   "Server"
            Object.Tag             =   "Server"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
            MinWidth        =   3969
            Text            =   "Port"
            TextSave        =   "Port"
            Key             =   "Users"
            Object.Tag             =   "Users"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7673
            MinWidth        =   7673
            Text            =   "Last msg Sent"
            TextSave        =   "Last msg Sent"
            Key             =   "Message"
            Object.Tag             =   "Message"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList il1 
      Left            =   8160
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4586
            Key             =   "Warn"
            Object.Tag             =   "Warn"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E62
            Key             =   "Blue"
            Object.Tag             =   "Blue"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":573E
            Key             =   "Red"
            Object.Tag             =   "Red"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":601A
            Key             =   "White"
            Object.Tag             =   "White"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":68F6
            Key             =   "Stop"
            Object.Tag             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71D2
            Key             =   "Go"
            Object.Tag             =   "Go"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7EAE
            Key             =   "Fire"
            Object.Tag             =   "Fire"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8B8A
            Key             =   "hGlass"
            Object.Tag             =   "hGlass"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8F82
            Key             =   "pSound"
            Object.Tag             =   "pSound"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":93BE
            Key             =   "oServer"
            Object.Tag             =   "oServer"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":995A
            Key             =   "oDoor"
            Object.Tag             =   "oDoor"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9E4E
            Key             =   "cDoor"
            Object.Tag             =   "cDoor"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Time"
         Object.Tag             =   "Time"
         Text            =   "Time Stamp"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Origin"
         Object.Tag             =   "Origin"
         Text            =   "Origin"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Message"
         Object.Tag             =   "Message"
         Text            =   "Message / Command"
         Object.Width           =   7673
      EndProperty
   End
   Begin MSWinsockLib.Winsock sck 
      Index           =   0
      Left            =   7920
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9456
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      ToolTipText     =   "Shut down BLAIM"
      Top             =   4680
      Width           =   600
   End
   Begin VB.CommandButton cmdSnd 
      Caption         =   "&Send"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Sends a Msg. To the client High-Lighted"
      Top             =   4680
      Width           =   600
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuSFile 
         Caption         =   "&SendFile"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server"
      Visible         =   0   'False
      Begin VB.Menu mnuStart 
         Caption         =   "&Start"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop"
      End
      Begin VB.Menu mnuIP 
         Caption         =   "&IPs"
      End
      Begin VB.Menu mnuSmsg 
         Caption         =   "&Smsg"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                                    Option Explicit
' Twinb.MyFtp.org

'© Coded by William LeRoy       William_j_leroy@yahoo.com  send me a msg or email me
'**************************************************************************************
' this is far from being finished
' so far i have only spent about 30 minutes on this

    Dim cUser As Integer

    ' These are the declarations for Winsock
    Dim Data1 As String
    
    ' Flag for server
    Dim IsServer As Boolean
    
    
    

Private Sub cmdButton_Click(Index As Integer)

    
    'This is my recreation of the ToolBar cntrl kind of
    Select Case Index 'The # of which command in the array pushed is stored here
        
        Case 0 'If it was cmdButton(0) then
        
            If IsServer = False Then ' If the server is off then
            
                sck(0).Listen 'Turen it on
                
                cmdButton(0).Picture = Picture1.Picture 'Change the button pic
                
                cmdButton(0).ToolTipText = "Turn Server Off" 'Set tool tip
                
                'Change the flag to the server is on
                IsServer = True
               
            Else 'else if the server was on
            
                sck(0).Close 'Turn it off
                
                cmdButton(0).Picture = Picture2.Picture 'Bring our picture back
                
                cmdButton(0).ToolTipText = "Turn Server On" 'Change tool tip back
                
                IsServer = False 'And set flag to false server is off
                
            End If
        
        Case 1 'if cmdButton(1) is pressed then
        
            frmClients.Show 'Show the Client Connections form
            
        Case 2 'cmdButton(2)
            
            
    
    End Select 'End our select
    

End Sub

Private Sub cmdExit_Click()

    Unload frmIM1
    
    Unload Me
    
End Sub

Private Sub cmdSnd_Click()

    ' Open the Form we need to send Msg.
    frmIM1.Show
    
End Sub





Private Sub Form_Load()

    ' Add to the listview some startup text
    AddtoList lv1
    
    ' Add to the Status Bar some startup text
    sBar sb, server.LocalIP
    
    ' Set the flag to FALSE for IsServer
    IsServer = False
    
    'Set the database name to point here
    db.DatabaseName = App.Path & "\env\server.mdb"
    
    'The source will be the Client field in the db
    db.RecordSource = "Client"
    
    'Associate the proper datafields with their respectable owners
    tID.DataField = "UserID" ' Sorry Gregg
    
    tPW.DataField = "uPassword"
    
    tC.DataField = "uConnected"
    
    tDate.DataField = "ccDate"
    
    tFro.DataField = "uFrozen"
    
    tWarn.DataField = "uWarnings"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me

End Sub

Private Sub mnuExit_Click()

    Unload frmIM1
    
    Unload Me
    
End Sub

Private Sub mnuSFile_Click()

    frmClients.Show

End Sub

Private Sub mnuStart_Click()

    ' Start the server
    sck(0).Listen
                        
                        
End Sub



Private Sub sb_PanelClick(ByVal Panel As MSComctlLib.Panel)
    
    'Just playing
    Panel = InputBox("Please enter the new information", _
                    "Status Bar Change", "Type Here")
                    

End Sub

Private Sub sck_ConnectionRequest(Index As Integer, _
                                  ByVal requestID As Long)

    ' first load new winsock into array
    cUser = cUser + 1
    
    Load sck(cUser)
    
    ' then set the port to 0 for random port
    sck(cUser).LocalPort = 0
    
    ' accept the connection
    sck(cUser).Accept requestID
    
    ' now store what #winsock user is on in .tag of txtIM
    frmIM1.txtIM.Tag = cUser

End Sub

Private Sub sck_DataArrival(Index As Integer, _
                            ByVal bytesTotal As Long)
                            
    ' Store the data as a String to Data1
    sck(cUser).GetData Data1, vbString
    
    'Check the message
    Check_Message Data1
    
    
End Sub



