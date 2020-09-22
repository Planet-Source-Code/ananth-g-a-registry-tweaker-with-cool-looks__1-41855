VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ..."
   ClientHeight    =   3465
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4110
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2391.604
   ScaleMode       =   0  'User
   ScaleWidth      =   3859.502
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   5220
      TabIndex        =   0
      Top             =   3540
      Width           =   1245
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "About.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   405
      Left            =   2730
      Picture         =   "About.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2910
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Ananthforu@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   570
      TabIndex        =   8
      Top             =   2580
      Width           =   2910
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "If you any bug please inform me mail to"
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   660
      TabIndex        =   7
      Top             =   2310
      Width           =   2730
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "By Ananth.G"
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   1470
      TabIndex        =   6
      Top             =   600
      Width           =   900
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00800000&
      Caption         =   $"About.frx":0C34
      ForeColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   120
      TabIndex        =   3
      Top             =   1095
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "RegTweak "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1200
      TabIndex        =   4
      Top             =   210
      Width           =   1590
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Version 0.99"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1470
      TabIndex        =   5
      Top             =   870
      Width           =   885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Show
    Dim ShapeCtrl As clsTransForm
    Set ShapeCtrl = New clsTransForm  'instantiate the object from the class
    ShapeCtrl.ShapeMe cmdOK, RGB(0, 0, 0), True, ""
    'ShapeCtrl.ShapeMe cmdBack, RGB(0, 0, 0), True, ""
    Set ShapeCtrl = Nothing
    
    
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

