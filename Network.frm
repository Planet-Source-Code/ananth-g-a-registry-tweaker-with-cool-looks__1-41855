VERSION 5.00
Begin VB.Form frmNetwork 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Network"
   ClientHeight    =   2790
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4380
   Icon            =   "Network.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBack 
      Height          =   405
      Left            =   3210
      Picture         =   "Network.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2250
      Width           =   1125
   End
   Begin VB.CommandButton cmdApply 
      Height          =   405
      Left            =   1860
      Picture         =   "Network.frx":1C30
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2250
      Width           =   1335
   End
   Begin VB.Frame fraTweak1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   2775
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   4335
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Control Panel"
         ForeColor       =   &H00FFFFFF&
         Height          =   2595
         Left            =   90
         TabIndex        =   7
         Top             =   60
         Width           =   2055
         Begin VB.CheckBox chkDialPass 
            BackColor       =   &H00000000&
            Caption         =   "Disable Save Password in DUN"
            ForeColor       =   &H00FFFFFF&
            Height          =   555
            Left            =   240
            TabIndex        =   11
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CheckBox chkDisHidden 
            BackColor       =   &H00000000&
            Caption         =   "Disable Hidden Shares"
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   240
            TabIndex        =   10
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CheckBox chkDisPrinter 
            BackColor       =   &H00000000&
            Caption         =   "Deny Printer Sharing"
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkDisFile 
            BackColor       =   &H00000000&
            Caption         =   "Disable File  Sharing"
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   270
            TabIndex        =   0
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox chkDisRAS 
            BackColor       =   &H00000000&
            Caption         =   "Disconnect RAS Callers"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   270
            TabIndex        =   8
            Top             =   2130
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Control Panel"
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   2190
         TabIndex        =   2
         Top             =   60
         Width           =   2055
         Begin VB.CheckBox chkMapDrive 
            BackColor       =   &H00000000&
            Caption         =   "Remove Map Disconnect Net Drive"
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   1845
         End
         Begin VB.CheckBox chkHideShare 
            BackColor       =   &H00000000&
            Caption         =   "Disable Real Mode   Ms-Dos Applications"
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   120
            TabIndex        =   5
            Top             =   570
            Width           =   1815
         End
         Begin VB.CheckBox chkHideUser 
            BackColor       =   &H00000000&
            Caption         =   "Don't Show Last User Name"
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   120
            TabIndex        =   4
            Top             =   90
            Width           =   1815
         End
         Begin VB.CheckBox chkDialIn 
            BackColor       =   &H00000000&
            Caption         =   "Disable Dialin Access"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   1590
            Width           =   1275
         End
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4200
      TabIndex        =   12
      ToolTipText     =   "Close"
      Top             =   -270
      Width           =   165
   End
End
Attribute VB_Name = "frmNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Res As Long, keyHwnd As Long, keyHwnd1&, KeyHwnd2&
Dim Data As Long, dType As Long
Dim szData As String * 256

Dim RgnHwnd As Long

Private Sub CheckNetwork()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", keyHwnd)
If Res <> ERROR_SUCCESS Then
    RegOpenKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\", keyHwnd
    RegCreateKey keyHwnd, "Network", keyHwnd1
End If
RegCloseKey keyHwnd
RegCloseKey keyHwnd1
End Sub

Private Sub GetDenyFileShare()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", keyHwnd)
Res = RegQueryValueEx(keyHwnd, "NoFileSharing", 0, REG_DWORD, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    chkDisFile.Value = Data
Else
    Res = RegSetValueEx(keyHwnd, "NoFileSharing", 0, REG_DWORD, Data, Len(Data))
    chkDisFile.Value = 0
    
End If
RegCloseKey keyHwnd
End Sub

Private Sub GetDenyPrintShare()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", keyHwnd)
Res = RegQueryValueEx(keyHwnd, "NoPrintSharing", 0, REG_DWORD, Data, Len(Data))
If Res = ERROR_SUCCESS Then
        chkDisPrinter.Value = Data
Else
    Res = RegSetValueEx(keyHwnd, "NoPrintSharing", 0, REG_DWORD, Data, Len(Data))
    chkDisPrinter.Value = 0
End If
RegCloseKey keyHwnd
End Sub

Private Sub SetDenyFileShare()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", keyHwnd)
dType = REG_DWORD
Data = chkDisFile.Value
Res = RegSetValueEx(keyHwnd, "NoFileSharing", 0, REG_DWORD, Data, Len(Data))
RegCloseKey keyHwnd
End Sub


Private Sub SetDenyPrintShare()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", keyHwnd)
dType = REG_DWORD
Data = chkDisPrinter.Value
Res = RegSetValueEx(keyHwnd, "NoPrintSharing", 0, REG_DWORD, Data, Len(Data))
RegCloseKey keyHwnd
End Sub

Private Sub SetDenyHidden()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\LanmanServer\Parameters", keyHwnd)
dType = REG_DWORD
Data = chkDisHidden.Value
Res = RegSetValueEx(keyHwnd, "AutoShareWks", 0, REG_DWORD, Data, Len(Data))
RegCloseKey keyHwnd
End Sub

Private Sub GetDenyHidden()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\LanmanServer\Parameters", keyHwnd)
If Res <> ERROR_SUCCESS Then
    Res = RegOpenKey(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services", keyHwnd)
    RegCreateKey keyHwnd, "LanManServer", keyHwnd1
    RegCreateKey keyHwnd1, "Parameters", KeyHwnd2
    keyHwnd = KeyHwnd2
End If

Res = RegQueryValueEx(keyHwnd, "AutoShareWks", 0, REG_DWORD, Data, Len(Data))

If Res = ERROR_SUCCESS Then
    chkDisHidden.Value = Data
Else
    Res = RegSetValueEx(keyHwnd, "AutoShareWks", 0, REG_DWORD, Data, Len(Data))
    chkDisHidden.Value = 0
    
End If
RegCloseKey keyHwnd

End Sub

Private Sub SetDialPass()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\RasMan\Parameters", keyHwnd)
dType = REG_DWORD
Data = chkDialPass.Value
Res = RegSetValueEx(keyHwnd, "DisableSavePassword", 0, REG_DWORD, Data, Len(Data))
RegCloseKey keyHwnd

End Sub

Private Sub GetDialPass()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\RasMan\Parameters", keyHwnd)
keyHwnd1 = keyHwnd
If Res <> ERROR_SUCCESS Then
    Res = RegOpenKey(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services", keyHwnd)
    RegCreateKey keyHwnd, "RsMan", keyHwnd1
    RegCreateKey keyHwnd1, "Parameters", KeyHwnd2
    keyHwnd = KeyHwnd2
End If
Res = RegQueryValueEx(keyHwnd1, "DisableSavePassword", 0, REG_DWORD, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    chkDialPass.Value = Data
Else
    Res = RegSetValueEx(keyHwnd1, "DisableSavePassword", 0, REG_DWORD, Data, Len(Data))
    chkDialPass.Value = 0
End If
RegCloseKey keyHwnd
End Sub

Private Sub SetHideUser()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\RasMan\Parameters", keyHwnd)
dType = REG_DWORD
Data = chkHideUser.Value
Res = RegSetValueEx(keyHwnd, "DontDisplayLastUserName", 0, REG_DWORD, Data, Len(Data))
RegCloseKey keyHwnd
End Sub

Private Sub GetHideUser()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", keyHwnd)
If Res <> ERROR_SUCCESS Then
    Res = RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", keyHwnd)
    RegCreateKey keyHwnd, "WinLogon", keyHwnd1
    keyHwnd = keyHwnd1
End If
Res = RegQueryValueEx(keyHwnd, "DontDisplayLastUserName", 0, REG_DWORD, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    chkHideUser.Value = Data
Else
    Res = RegSetValueEx(keyHwnd, "DontDisplayLastUserName", 0, REG_DWORD, Data, Len(Data))
    chkHideUser.Value = 0
End If
RegCloseKey keyHwnd
End Sub


Private Sub SetSharePass()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", keyHwnd)
dType = REG_DWORD
Data = chkHideShare.Value
Res = RegSetValueEx(keyHwnd, "HideSharePwds", 0, REG_DWORD, Data, Len(Data))
RegCloseKey keyHwnd
End Sub

Private Sub GetSharePass()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", keyHwnd)
If Res <> ERROR_SUCCESS Then
    Res = RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies", keyHwnd)
    RegCreateKey keyHwnd, "Network", keyHwnd1
    keyHwnd = keyHwnd1
End If
Res = RegQueryValueEx(keyHwnd, "HideSharePwds", 0, REG_DWORD, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    chkHideShare.Value = Data
Else
    Res = RegSetValueEx(keyHwnd, "HideSharePwds", 0, REG_DWORD, Data, Len(Data))
    chkHideShare.Value = 0
End If
RegCloseKey keyHwnd
End Sub

Private Sub SetMapDrive()
Res = RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", keyHwnd)
dType = REG_DWORD
Data = chkMapDrive.Value
Res = RegSetValueEx(keyHwnd, "NoNetConnectDisconnect", 0, REG_DWORD, Data, Len(Data))
RegCloseKey keyHwnd
End Sub

Private Sub GetMapDrive()
Res = RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", keyHwnd)
Res = RegQueryValueEx(keyHwnd, "NoNetConnectDisconnect", 0, REG_DWORD, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    chkMapDrive.Value = Data
Else
    Res = RegSetValueEx(keyHwnd, "NoNetConnectDisconnect", 0, REG_DWORD, Data, Len(Data))
    chkMapDrive.Value = 0
End If
RegCloseKey keyHwnd
End Sub

Private Sub GetDialIn()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", keyHwnd)
Res = RegQueryValueEx(keyHwnd, "NoDialIn", 0, REG_DWORD, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    chkDialIn.Value = Data
Else
    Res = RegSetValueEx(keyHwnd, "NoDialIn", 0, REG_DWORD, Data, Len(Data))
    chkDialIn.Value = 0
End If
RegCloseKey keyHwnd
End Sub

Private Sub SetDialIn()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", keyHwnd)
dType = REG_DWORD
Data = chkDialIn.Value
Res = RegSetValueEx(keyHwnd, "NoDialIn", 0, REG_DWORD, Data, Len(Data))
RegCloseKey keyHwnd
End Sub

Private Sub SetRASDis()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\RemoteAccess\Parameters", keyHwnd)
dType = REG_DWORD
Data = chkDisRAS.Value
Res = RegSetValueEx(keyHwnd, "AutoDisconnect", 0, REG_DWORD, Data, Len(Data))
RegCloseKey keyHwnd
End Sub

Private Sub GetRASDis()
Res = RegOpenKey(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\RemoteAccess\Parameters", keyHwnd)
If Res <> ERROR_SUCCESS Then
    Res = RegOpenKey(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\RemoteAccess", keyHwnd)
    RegCreateKey keyHwnd, "Parameters", keyHwnd1
    keyHwnd = keyHwnd1
End If
Res = RegQueryValueEx(keyHwnd, "AutoDisconnect", 0, REG_DWORD, Data, Len(Data))
If Res = ERROR_SUCCESS Then
    chkDisRAS.Value = Data
Else
    Res = RegSetValueEx(keyHwnd, "AutoDisconnect", 0, REG_DWORD, Data, Len(Data))
    chkDisRAS.Value = 0
End If
RegCloseKey keyHwnd
End Sub

Private Sub cmdApply_Click()
SetDenyFileShare
SetDenyPrintShare
SetDenyHidden
SetDialPass
SetHideUser
SetSharePass
SetMapDrive
SetDialIn
SetRASDis
frmRestart.Show 1
Unload Me
End Sub

Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub Form_Load()

    Me.Show
    Dim ShapeCtrl As clsTransForm
    Set ShapeCtrl = New clsTransForm  'instantiate the object from the class
    ShapeCtrl.ShapeMe cmdApply, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe cmdBack, RGB(0, 0, 0), True, ""
    Set ShapeCtrl = Nothing


CheckNetwork
GetDenyPrintShare
GetDenyFileShare
GetDenyHidden
GetDialPass
GetHideUser
GetSharePass
GetMapDrive
GetDialIn
GetRASDis
End Sub


