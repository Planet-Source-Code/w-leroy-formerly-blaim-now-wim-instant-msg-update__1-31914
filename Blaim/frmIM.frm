VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form frmIM 
   BackColor       =   &H80000018&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Message From :"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   528
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H0063C7ED&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   4
      Text            =   "Message to send Here"
      Top             =   3195
      Width           =   5370
   End
   Begin MSComctlLib.ImageList il 
      Left            =   7080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM.frx":08DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM.frx":11B8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton cmdSend 
      Height          =   525
      Left            =   6840
      TabIndex        =   3
      ToolTipText     =   "Send Message"
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   16711680
      EMBOSSS         =   8421631
      MPTR            =   0
      MICON           =   "frmIM.frx":1A94
      ALIGN           =   1
      IMGLST          =   "il"
      IMGICON         =   "1"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   525
      Left            =   5760
      TabIndex        =   2
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "&Send a File"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIM.frx":1AB0
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4048
      _Version        =   393217
      BackColor       =   6539245
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmIM.frx":1ACC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H0063C7ED&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   0
      Text            =   "From:"
      Top             =   150
      Width           =   2850
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      FillColor       =   &H0063C7ED&
      FillStyle       =   0  'Solid
      Height          =   525
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   5520
   End
   Begin VB.Shape sFrom 
      BorderWidth     =   5
      FillColor       =   &H0063C7ED&
      FillStyle       =   0  'Solid
      Height          =   525
      Left            =   75
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   3000
   End
End
Attribute VB_Name = "frmIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
