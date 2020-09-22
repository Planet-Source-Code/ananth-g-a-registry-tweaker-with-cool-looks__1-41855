VERSION 5.00
Begin VB.Form frmWindows 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Windows Folders"
   ClientHeight    =   2700
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   2085
   End
   Begin VB.CommandButton Command1 
      Height          =   405
      Left            =   2220
      Picture         =   "Windows.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   270
      Width           =   765
   End
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   60
      TabIndex        =   13
      Top             =   930
      Width           =   2085
   End
   Begin VB.CommandButton Command2 
      Height          =   405
      Left            =   2220
      Picture         =   "Windows.frx":08FA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   900
      Width           =   765
   End
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   60
      TabIndex        =   11
      Top             =   1590
      Width           =   2085
   End
   Begin VB.CommandButton Command3 
      Height          =   405
      Left            =   2220
      Picture         =   "Windows.frx":11F4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1590
      Width           =   765
   End
   Begin VB.TextBox Text4 
      Height          =   345
      Left            =   60
      TabIndex        =   9
      Top             =   2250
      Width           =   2085
   End
   Begin VB.CommandButton Command4 
      Height          =   405
      Left            =   2220
      Picture         =   "Windows.frx":1AEE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2250
      Width           =   765
   End
   Begin VB.TextBox Text5 
      Height          =   345
      Left            =   3240
      TabIndex        =   7
      Top             =   270
      Width           =   2085
   End
   Begin VB.CommandButton Command5 
      Height          =   405
      Left            =   5400
      Picture         =   "Windows.frx":23E8
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   345
      Left            =   3240
      TabIndex        =   5
      Top             =   930
      Width           =   2085
   End
   Begin VB.CommandButton Command6 
      Height          =   405
      Left            =   5400
      Picture         =   "Windows.frx":2CE2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   900
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Height          =   345
      Left            =   3240
      TabIndex        =   3
      Top             =   1590
      Width           =   2085
   End
   Begin VB.CommandButton Command7 
      Height          =   405
      Left            =   5400
      Picture         =   "Windows.frx":35DC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      Height          =   405
      Left            =   5250
      Picture         =   "Windows.frx":3ED6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2220
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Start Menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   22
      Top             =   90
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "History"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   870
      TabIndex        =   21
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "StartUp"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   20
      Top             =   1380
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Personal"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   810
      TabIndex        =   19
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Favorites"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3945
      TabIndex        =   18
      Top             =   60
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "SendTo"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3975
      TabIndex        =   17
      Top             =   720
      Width           =   570
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Programs"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3930
      TabIndex        =   16
      Top             =   1380
      Width           =   660
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   6180
      TabIndex        =   15
      ToolTipText     =   "Close"
      Top             =   -240
      Width           =   165
   End
End
Attribute VB_Name = "frmWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Res As Long, keyHwnd As Long, keyHwnd1&, KeyHwnd2&
Dim Data As Long, dType As Long, lenData
Dim szData As String * 256

Private Sub GetStart()
    RegQueryValueEx keyHwnd, "Start Menu", 0, REG_SZ, ByVal szData, Len(szData)
    Text1.Text = szData
End Sub
Private Sub GetHistory()
    RegQueryValueEx keyHwnd, "History", 0, REG_SZ, ByVal szData, Len(szData)
    Text2.Text = szData
End Sub
Private Sub GetStartup()
    RegQueryValueEx keyHwnd, "StartUP", 0, REG_SZ, ByVal szData, Len(szData)
    Text3.Text = szData
End Sub
Private Sub GetPersonal()
    RegQueryValueEx keyHwnd, "Personal", 0, REG_SZ, ByVal szData, Len(szData)
    Text4.Text = szData
End Sub
Private Sub GetFavorites()
    RegQueryValueEx keyHwnd, "Favorites", 0, REG_SZ, ByVal szData, Len(szData)
    Text5.Text = szData
End Sub
Private Sub GetSendto()
    RegQueryValueEx keyHwnd, "SendTo", 0, REG_SZ, ByVal szData, Len(szData)
    Text6.Text = szData
End Sub
Private Sub GetPrograms()
    RegQueryValueEx keyHwnd, "Programs", 0, REG_SZ, ByVal szData, Len(szData)
    Text7.Text = szData
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    RegSetValueEx keyHwnd, "Start Menu", 0, REG_SZ, ByVal Text1.Text, Len(Text1.Text)
    Load frmRestart
End Sub

Private Sub Command2_Click()
    RegSetValueEx keyHwnd, "History", 0, REG_SZ, ByVal Text2.Text, Len(Text2.Text)
    Load frmRestart
End Sub

Private Sub Command3_Click()
    RegSetValueEx keyHwnd, "StartUp", 0, REG_SZ, ByVal Text3.Text, Len(Text3.Text)
    Load frmRestart
End Sub

Private Sub Command4_Click()
    RegSetValueEx keyHwnd, "Personal", 0, REG_SZ, ByVal Text4.Text, Len(Text4.Text)
    Load frmRestart
End Sub

Private Sub Command5_Click()
    RegSetValueEx keyHwnd, "Favorites", 0, REG_SZ, ByVal Text5.Text, Len(Text5.Text)
    Load frmRestart
End Sub

Private Sub Command6_Click()
    RegSetValueEx keyHwnd, "SendTo", 0, REG_SZ, ByVal Text6.Text, Len(Text6.Text)
    Load frmRestart
End Sub

Private Sub Command7_Click()
    RegSetValueEx keyHwnd, "Programs", 0, REG_SZ, ByVal Text7.Text, Len(Text7.Text)
    Load frmRestart
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static moved As Boolean, StartX As Long, StartY As Long

    If Button = 1 Then
        If moved = True Then
            Me.Left = Me.Left - (StartX - x)
            Me.Top = Me.Top - (StartY - y)
            x = StartX
            y = StartY
        Else
            StartX = x
            StartY = y
            moved = True
        End If
    Else
        moved = False
    End If

End Sub

Private Sub Form_Load()
    Me.Show
    
    Dim ShapeCtrl As clsTransForm
    Set ShapeCtrl = New clsTransForm  'instantiate the object from the class
    ShapeCtrl.ShapeMe Command1, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe Command2, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe Command3, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe Command4, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe Command5, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe Command6, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe Command7, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe cmdBack, RGB(0, 0, 0), True, ""
    Set ShapeCtrl = Nothing
    

    
    MsgBox "Changing the System Folders may cause troubles", vbCritical, "Warning"
    Res = RegOpenKey(HKEY_CURRENT_USER, "software\microsoft\windows\currentversion\explorer\shell folders", keyHwnd)
    
    GetStart
    GetHistory
    GetStartup
    GetPrograms
    GetPersonal
    GetFavorites
    GetSendto
    
    
End Sub

Private Sub Label8_Click()
    Unload Me
End Sub

