VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "RegTweak"
   ClientHeight    =   3405
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   2625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   2625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Height          =   495
      Left            =   2250
      Picture         =   "Main.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   405
   End
   Begin VB.CommandButton cmdDummy 
      Caption         =   "Command1"
      Height          =   285
      Left            =   2010
      TabIndex        =   0
      Top             =   3330
      Width           =   1245
   End
   Begin VB.CommandButton cmdAbout 
      Height          =   495
      Left            =   300
      Picture         =   "Main.frx":05E2
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Height          =   495
      Left            =   300
      Picture         =   "Main.frx":2244
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2730
      Width           =   1935
   End
   Begin VB.CommandButton cmdSecurity 
      Height          =   495
      Left            =   300
      Picture         =   "Main.frx":3EA6
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdWindows 
      Height          =   495
      Left            =   300
      Picture         =   "Main.frx":59A0
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdIe 
      Height          =   495
      Left            =   300
      Picture         =   "Main.frx":7602
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdNetwork 
      Height          =   495
      Left            =   300
      Picture         =   "Main.frx":9264
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cmdAbout_Click()
    Load frmAbout
End Sub

Private Sub cmdClose_Click()
    cmdExit_Click
End Sub

Private Sub cmdExit_Click()
    If MsgBox("Do you realy want to exit?", vbQuestion + vbDefaultButton2 + vbYesNo) = vbYes Then
        End
    End If
End Sub

Private Sub cmdIe_Click()
    Load frmIE
End Sub

Private Sub cmdNetwork_Click()
    Load frmNetwork
End Sub

Private Sub cmdSecurity_Click()
    Load frmSecurity
End Sub

Private Sub cmdWindows_Click()
    Load frmWindows
End Sub

Private Sub Form_Load()
    Me.Show
    Dim ShapeCtrl As clsTransForm
    Set ShapeCtrl = New clsTransForm  'instantiate the object from the class
    ShapeCtrl.ShapeMe cmdNetwork, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe cmdSecurity, RGB(0, 0, 0), False, ""
    ShapeCtrl.ShapeMe cmdIe, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe cmdWindows, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe cmdAbout, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe cmdExit, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe cmdClose, RGB(0, 0, 0), True, ""
    
    Set ShapeCtrl = Nothing
    
    cmdDummy.Visible = False
End Sub
