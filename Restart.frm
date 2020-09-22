VERSION 5.00
Begin VB.Form frmRestart 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restart"
   ClientHeight    =   1875
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDummy 
      Caption         =   "Command1"
      Height          =   525
      Left            =   5280
      TabIndex        =   0
      Top             =   1830
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   2880
      Picture         =   "Restart.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1050
      Width           =   2025
   End
   Begin VB.CommandButton cmdRestart 
      Height          =   495
      Left            =   600
      Picture         =   "Restart.frx":1C62
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1050
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Are you sure you want to restart ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   630
      TabIndex        =   3
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "frmRestart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRestart_Click()
    ExitWindowsEx EWX_REBOOT, 0
End Sub

Private Sub Form_Load()
    
    Me.Show
    Dim ShapeCtrl As clsTransForm
    Set ShapeCtrl = New clsTransForm  'instantiate the object from the class
    ShapeCtrl.ShapeMe cmdRestart, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe cmdCancel, RGB(0, 0, 0), True, ""
    Set ShapeCtrl = Nothing
    cmdDummy.Visible = False
End Sub
