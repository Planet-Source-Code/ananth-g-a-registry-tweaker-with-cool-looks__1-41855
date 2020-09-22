VERSION 5.00
Begin VB.Form frmSecurity 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Security & Privacy"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5970
   Icon            =   "frmSec.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Control Panel"
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   3840
      TabIndex        =   21
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton cmdRunMru 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox chkCmd 
         BackColor       =   &H00000000&
         Caption         =   "Disable Command      prompt"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chkReal 
         BackColor       =   &H00000000&
         Caption         =   "Disable Real Mode   Ms-Dos Applications"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkFileMenu 
         BackColor       =   &H00000000&
         Caption         =   "Remove File Menu     in Explorer"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "  Remove Run    command list"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   26
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton cmdDeny 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   19
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmdStartMenuHelp 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   360
      Width           =   255
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Display Properties"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   4575
      Begin VB.CheckBox chkSett 
         BackColor       =   &H00000000&
         Caption         =   "Disable Settings Page"
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   2040
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox chkScr 
         BackColor       =   &H00000000&
         Caption         =   "Disable Screen Saver Page"
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   2040
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkWall 
         BackColor       =   &H00000000&
         Caption         =   "Disable Background Page"
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkApp 
         BackColor       =   &H00000000&
         Caption         =   "Disable Appearance Page"
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Control Panel"
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   1680
      TabIndex        =   8
      Top             =   120
      Width           =   2055
      Begin VB.CheckBox chkDisp 
         BackColor       =   &H00000000&
         Caption         =   "Deny Access to Display Properties"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chkNet 
         BackColor       =   &H00000000&
         Caption         =   "Deny Access to Network"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkPrint 
         BackColor       =   &H00000000&
         Caption         =   "Deny Access to Printers"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox chkPwd 
         BackColor       =   &H00000000&
         Caption         =   "Deny Access to Passwords"
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdBack 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   4770
      Picture         =   "frmSec.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Start Menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      Begin VB.CheckBox chkRun 
         BackColor       =   &H00000000&
         Caption         =   "&Run"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkFind 
         BackColor       =   &H00000000&
         Caption         =   "&Find"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   690
         Width           =   855
      End
      Begin VB.CheckBox chkFav 
         BackColor       =   &H00000000&
         Caption         =   "Favorites"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1020
         Width           =   975
      End
      Begin VB.CheckBox chkLogoff 
         BackColor       =   &H00000000&
         Caption         =   "&Logoff"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1350
         Width           =   855
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00000000&
         Caption         =   "&Documents"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdApply 
      Height          =   495
      Left            =   4620
      Picture         =   "frmSec.frx":1C30
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1425
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Res As Long, keyHwnd As Long, keyHwnd1&, KeyHwnd2&
Dim Data As Long, dType As Long
Dim szData As String * 256

Dim RgnHwnd As Long

Private Sub Set_Size()

End Sub
Private Sub cmdIcon_Click()
    RegOpenKey HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", keyHwnd1
    RegSetValueEx keyHwnd1, "Shell Icon BPP", 0, REG_SZ, "16", Len("16")
    RegCloseKey keyHwnd1
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub

'********************************
'SET SIZE
Private Sub Change_Size()
    RgnHwnd = CreateEllipticRgn(3, 3, cmdBack.Width / Screen.TwipsPerPixelX - 3, cmdBack.Height / Screen.TwipsPerPixelY - 3)
    SetWindowRgn cmdBack.hWnd, RgnHwnd, 1
End Sub

Private Sub cmdDeny_Click()
    MsgBox "You can select the following options to deny access to particular applet in the Control Panel"
End Sub

Private Sub cmdMoreHelp_Click()

End Sub

Private Sub cmdRunMru_Click()
    RegOpenKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer", keyHwnd
    RegDeleteKey keyHwnd, "RunMru"
    RegCreateKey keyHwnd, "RunMru", Res
    RegCloseKey keyHwnd
End Sub

Private Sub cmdStartMenuHelp_Click()
    MsgBox "The items given below appear in the Start Menu. Select the item to make it visible or don't if you want to hide it"
End Sub

Private Sub cmdApply_Click()

    Res = RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", keyHwnd)
    If Res <> ERROR_SUCCESS Then
        RegOpenKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies", keyHwnd
        RegCreateKey keyHwnd, "Explorer", keyHwnd1
        keyHwnd = keyHwnd1
    End If

'***************************
'run
    If chkRun.Value = 1 Then
        Data = 0
        RegSetValueEx keyHwnd, "NoRun", 0, REG_DWORD, Data, Len(Data)
    Else
        Data = 1
        RegSetValueEx keyHwnd, "NoRun", 0, REG_DWORD, Data, Len(Data)
    End If
'***************************
'favorites
If chkFav.Value = 1 Then
    Data = 0
    RegSetValueEx keyHwnd, "NoFavoritesMenu", 0, REG_DWORD, Data, Len(Data)
Else
    Data = 1
    RegSetValueEx keyHwnd, "NoFavoritesMenu", 0, REG_DWORD, Data, Len(Data)
End If
'***************************
'documents menu
If chkDoc.Value = 1 Then
    Data = 0
    RegSetValueEx keyHwnd, "NoRecentDocsMenu", 0, REG_DWORD, Data, Len(Data)
Else
    Data = 1
    RegSetValueEx keyHwnd, "NoRecentDocsMenu", 0, REG_DWORD, Data, Len(Data)
End If
'***************************
'find
If chkFind.Value = 1 Then
    Data = 0
    RegSetValueEx keyHwnd, "NoFind", 0, REG_DWORD, Data, Len(Data)
Else
    Data = 1
    RegSetValueEx keyHwnd, "NoFind", 0, REG_DWORD, Data, Len(Data)
End If
'***************************
'logoff
If chkLogoff.Value = 1 Then
    Data = 0
    RegSetValueEx keyHwnd, "NoLogOff", 0, REG_DWORD, Data, Len(Data)
Else
    Data = 1
    RegSetValueEx keyHwnd, "NoLogOff", 0, REG_DWORD, Data, Len(Data)
End If
'*************************************
'file menu in explorer
If chkFileMenu.Value = 1 Then
        RegSetValueEx keyHwnd, "NoFileMenu", 0, REG_DWORD, 1, Len(Data)
    Else
        RegSetValueEx keyHwnd, "NoFileMenu", 0, REG_DWORD, 0, Len(Data)
    End If
'*************************************
'printers in control panel
If chkPrint.Value = 1 Then
        RegSetValueEx keyHwnd, "NoPrinters", 0, REG_DWORD, 1, Len(Data)
    Else
        RegSetValueEx keyHwnd, "NoPrinters", 0, REG_DWORD, 0, Len(Data)
    End If
RegCloseKey keyHwnd
'*************************************
'ms dos mode
Res = RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", keyHwnd1)
If Res <> ERROR_SUCCESS Then
    RegOpenKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies", keyHwnd
    RegCreateKey keyHwnd, "Network", keyHwnd1
    keyHwnd = keyHwnd1
End If

If chkCmd.Value = 1 Then
    RegSetValueEx keyHwnd1, "Disabled", 0, REG_DWORD, 1, Len(Data)
Else
    RegSetValueEx keyHwnd1, "Disabled", 0, REG_DWORD, 0, Len(Data)
End If

RegCloseKey keyHwnd1

'*************************************
'real ms dos mode
Res = RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", keyHwnd1)
If Res <> ERROR_SUCCESS Then
    RegOpenKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies", keyHwnd
    RegCreateKey keyHwnd, "WinOldApp", keyHwnd1
    keyHwnd = keyHwnd1
End If

If chkReal.Value = 1 Then
    RegSetValueEx keyHwnd1, "NoRealMode", 0, REG_DWORD, 1, Len(Data)
Else
    RegSetValueEx keyHwnd1, "NoRealMode", 0, REG_DWORD, 0, Len(Data)
End If

RegCloseKey keyHwnd1

'*************************************
'deny access to display properties
Res = RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", keyHwnd)
If Res <> ERROR_SUCCESS Then
    RegOpenKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies", keyHwnd
    RegCreateKey keyHwnd, "System", keyHwnd1
    keyHwnd = keyHwnd1
End If

    If chkDisp.Value = 1 Then
        RegSetValueEx keyHwnd, "NoDispCPL", 0, REG_DWORD, 1, Len(Data)
    Else
        RegSetValueEx keyHwnd, "NoDispCPL", 0, REG_DWORD, 0, Len(Data)
    End If
'**************************************************
' deny access to password and user in cpl
    If chkPwd.Value = 1 Then
        RegSetValueEx keyHwnd, "NoSecCPL", 0, REG_DWORD, 1, Len(Data)
    Else
        RegSetValueEx keyHwnd, "NoSecCPL", 0, REG_DWORD, 0, Len(Data)
    End If
'**************************************************
' deny access appearance page in display properties
    If chkApp.Value = 1 Then
        RegSetValueEx keyHwnd, "NoDispAppearancePage", 0, REG_DWORD, 1, Len(Data)
    Else
        RegSetValueEx keyHwnd, "NoDispAppearancePage", 0, REG_DWORD, 0, Len(Data)
    End If
'**************************************************
' deny access background page in display properties
    If chkWall.Value = 1 Then
        RegSetValueEx keyHwnd, "NoDispBackgroundPage", 0, REG_DWORD, 1, Len(Data)
    Else
        RegSetValueEx keyHwnd, "NoDispBackgroundPage", 0, REG_DWORD, 0, Len(Data)
    End If
'**************************************************
' deny access screen saver page in display properties
    If chkScr.Value = 1 Then
        RegSetValueEx keyHwnd, "NoDispScrSavPage", 0, REG_DWORD, 1, Len(Data)
    Else
        RegSetValueEx keyHwnd, "NoDispScrSavPage", 0, REG_DWORD, 0, Len(Data)
    End If
'**************************************************
' deny access settings page in display properties
    If chkSett.Value = 1 Then
        RegSetValueEx keyHwnd, "NoDispSettingsPage", 0, REG_DWORD, 1, Len(Data)
    Else
        RegSetValueEx keyHwnd, "NoDispSettingsPage", 0, REG_DWORD, 0, Len(Data)
    End If



RegCloseKey keyHwnd

'*************************************
'deny access to network in cpl
Res = RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", keyHwnd)
If Res <> ERROR_SUCCESS Then
    RegOpenKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies", keyHwnd
    RegCreateKey keyHwnd, "Network", keyHwnd1
    keyHwnd = keyHwnd1
End If

    If chkNet.Value = 1 Then
        RegSetValueEx keyHwnd, "NoNetSetup", 0, REG_DWORD, 1, Len(Data)
    Else
        RegSetValueEx keyHwnd, "NoNetSetup", 0, REG_DWORD, 0, Len(Data)
    End If
    
RegCloseKey keyHwnd

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command5_Click()
MsgBox "Use these options to hide or show the pages given below which appear in the Display Properties"
End Sub

Private Sub Form_Load()
    Me.Show
    Dim ShapeCtrl As clsTransForm
    Set ShapeCtrl = New clsTransForm  'instantiate the object from the class
    ShapeCtrl.ShapeMe cmdApply, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe cmdBack, RGB(0, 0, 0), True, ""
    Set ShapeCtrl = Nothing



Set_Size


RegOpenKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", keyHwnd
'*********************************************
' FOR RUN MENU
Res = RegQueryValueEx(keyHwnd, "NoRun", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkRun.Value = 0
    Else
        chkRun.Value = 1
    End If
Else
    chkRun.Value = 1
End If
'*********************************************
'FOR find MENU
Res = RegQueryValueEx(keyHwnd, "NoFind", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkFind.Value = 0
    Else
        chkFind.Value = 1
    End If
Else
    chkFind.Value = 1
End If
'*********************************************
'FOR FAVORITES MENU
Res = RegQueryValueEx(keyHwnd, "NoFavoritesMenu", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkFav.Value = 0
    Else
        chkFav.Value = 1
    End If
Else
    chkFav.Value = 1
End If
'*********************************************
'FOR LOGOFF MENU
Res = RegQueryValueEx(keyHwnd, "NoLogoff", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkLogoff.Value = 0
    Else
        chkLogoff.Value = 1
    End If
Else
    chkLogoff.Value = 1
End If
'*********************************************
'FOR DOCUMENTS MENU

Res = RegQueryValueEx(keyHwnd, "NoRecentDocsMenu", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkDoc.Value = 0
    Else
        chkDoc.Value = 1
    End If
Else
    chkDoc.Value = 1
End If

RegCloseKey keyHwnd

'***********************************

'ms dos mode

RegOpenKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", keyHwnd1
Res = RegQueryValueEx(keyHwnd1, "Disabled", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkCmd.Value = 1
    Else
        chkCmd.Value = 0
    End If
Else
        chkCmd.Value = 0
End If

RegCloseKey keyHwnd1
'**************************************
'real ms dos mode

RegOpenKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", keyHwnd1
Res = RegQueryValueEx(keyHwnd1, "NoRealMode", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkReal.Value = 1
    Else
        chkReal.Value = 0
    End If
Else
        chkReal.Value = 0
End If

RegCloseKey keyHwnd1
'**************************************
'file menu in explorer
RegOpenKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", keyHwnd
Res = RegQueryValueEx(keyHwnd, "NoFileMenu", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkFileMenu.Value = 1
    Else
        chkFileMenu.Value = 0
    End If
Else
    chkFileMenu.Value = 0
End If
RegCloseKey keyHwnd
'**************************************
'deny access to display properties
RegOpenKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", keyHwnd
Res = RegQueryValueEx(keyHwnd, "NoDispCPL", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkDisp.Value = 1
    Else
        chkDisp.Value = 0
    End If
Else
    chkDisp.Value = 0
End If
RegCloseKey keyHwnd
'**************************************
'deny access to network in control panel
RegOpenKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", keyHwnd
Res = RegQueryValueEx(keyHwnd, "NoNetSetup", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkNet.Value = 1
    Else
        chkNet.Value = 0
    End If
Else
    chkNet.Value = 0
End If
RegCloseKey keyHwnd
'**************************************
'deny access to printers in control panel
RegOpenKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", keyHwnd
Res = RegQueryValueEx(keyHwnd, "NoPrinters", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkPrint.Value = 1
    Else
        chkPrint.Value = 0
    End If
Else
    chkPrint.Value = 0
End If
RegCloseKey keyHwnd
'**************************************
'deny access to printers in control panel
RegOpenKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", keyHwnd
Res = RegQueryValueEx(keyHwnd, "NoSecCPL", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkPwd.Value = 1
    Else
        chkPwd.Value = 0
    End If
Else
    chkPwd.Value = 0
End If
'**************************************
'deny access appearance page in display properties
Res = RegQueryValueEx(keyHwnd, "NoDispAppearancePage", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkApp.Value = 1
    Else
        chkApp.Value = 0
    End If
Else
    chkApp.Value = 0
End If
'**************************************
'deny access background page in display properties
Res = RegQueryValueEx(keyHwnd, "NoDispBackgroundPage", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkWall.Value = 1
    Else
        chkWall.Value = 0
    End If
Else
    chkWall.Value = 0
End If
'**************************************
'deny access screensaver page in display properties
Res = RegQueryValueEx(keyHwnd, "NoDispScrSavPage", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkScr.Value = 1
    Else
        chkScr.Value = 0
    End If
Else
    chkScr.Value = 0
End If
'**************************************
'deny access to display settings page in display properties
Res = RegQueryValueEx(keyHwnd, "NoDispSettingsPage", 0, dType, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    If Data = 1 Then
        chkSett.Value = 1
    Else
        chkSett.Value = 0
    End If
Else
    chkSett.Value = 0
End If

RegCloseKey keyHwnd
End Sub


