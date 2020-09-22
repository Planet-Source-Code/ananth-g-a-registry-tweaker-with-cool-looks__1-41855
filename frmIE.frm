VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIE 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internet Explorer"
   ClientHeight    =   3180
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   1170
      Top             =   2670
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtBack 
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   330
      Width           =   1995
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   405
      Left            =   2040
      Picture         =   "frmIE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   330
      Width           =   1335
   End
   Begin VB.TextBox txtSearch 
      Height          =   345
      Left            =   90
      TabIndex        =   6
      Top             =   1020
      Width           =   1995
   End
   Begin VB.CommandButton cmdSet 
      Height          =   345
      Left            =   2280
      Picture         =   "frmIE.frx":1076
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1020
      Width           =   945
   End
   Begin VB.TextBox txtTitle 
      Height          =   345
      Left            =   90
      TabIndex        =   4
      Top             =   1710
      Width           =   1995
   End
   Begin VB.CommandButton cmdSetTitle 
      Height          =   345
      Left            =   2280
      Picture         =   "frmIE.frx":1970
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1710
      Width           =   945
   End
   Begin VB.CheckBox chkHideIE 
      BackColor       =   &H00800000&
      Caption         =   "Hide Internet Explorer Icon from Desktop"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2220
      Width           =   1845
   End
   Begin VB.CommandButton cmdBack 
      Height          =   405
      Left            =   2040
      Picture         =   "frmIE.frx":226A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Explorer Background Image"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   90
      Width           =   2925
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Explorer Search Engine (with Query)"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   780
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Explorer Title"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1470
      Width           =   1545
   End
End
Attribute VB_Name = "frmIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Res As Long, keyHwnd As Long, keyHwnd1&, KeyHwnd2&
Dim Data As Long, dType As Long, lenData
Dim szData As String * 256

Private Sub GetIEback()
Res = RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar", keyHwnd)
Res = RegQueryValueEx(keyHwnd, "BackBitmap", 0, REG_SZ, ByVal szData, Len(szData))
txtBack.Text = szData
RegCloseKey keyHwnd
End Sub
Private Sub GetHideIE()
On Error Resume Next
Res = RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", keyHwnd)
RegQueryValueEx keyHwnd, "NoInternetIcon", 0, REG_DWORD, Data, Len(Data)
chkHideIE.Value = Data
RegCloseKey keyHwnd
End Sub

Private Sub GetIESearch()
Res = RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\SearchUrl", keyHwnd)
Res = RegQueryValue(keyHwnd, vbNullString, ByVal szData, Len(szData))
If Res = ERROR_SUCCESS Then
Else
RegSetValueEx keyHwnd, "(default)", 0, REG_SZ, ByVal "", 0
txtSearch.Text = ""
End If
txtSearch.Text = szData
RegCloseKey keyHwnd
End Sub
Private Sub GetIETitle()
Res = RegOpenKey(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Internet Explorer\Main", keyHwnd)
Res = RegQueryValueEx(keyHwnd, "Window Title", 0, REG_SZ, ByVal szData, Len(szData))
If Res = ERROR_SUCCESS Then
txtTitle.Text = szData
Else
RegSetValueEx keyHwnd, "Window Title", 0, REG_SZ, ByVal "", 0
txtTitle.Text = ""
End If
RegCloseKey keyHwnd
End Sub


Private Sub SetIEback()
Res = RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar", keyHwnd)
Res = RegSetValueEx(keyHwnd, "BackBitmap", 0, REG_SZ, ByVal txtBack.Text, Len(txtBack.Text))
RegCloseKey keyHwnd
End Sub

Private Sub SetIESearch()
Res = RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\SearchUrl", keyHwnd)
RegSetValueEx keyHwnd, "(Default)", 0, REG_SZ, ByVal txtSearch.Text, Len(txtSearch.Text)
RegCloseKey keyHwnd
End Sub

Private Sub chkHideIE_Click()
Res = RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", keyHwnd)
Data = chkHideIE.Value
RegSetValueEx keyHwnd, "NoInternetIcon", 0, REG_DWORD, Data, Len(Data)
If Data = 1 Then
    chkHideIE.Value = 1
Else
    chkHideIE.Value = 0
End If
End Sub

Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub cmdBrowse_Click()
Dialog.DialogTitle = "BackGround Image"
Dialog.Filter = "Pictures (*.bmp;*.jpg)|*.bmp;*.jpg"
Dialog.ShowOpen
txtBack.Text = Dialog.FileName
SetIEback
End Sub

Private Sub cmdSet_Click()
SetIESearch
End Sub

Private Sub cmdSetTitle_Click()
Res = RegOpenKey(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Internet Explorer\Main", keyHwnd)
RegSetValueEx keyHwnd, "Window Title", 0, REG_SZ, ByVal txtTitle.Text, Len(txtTitle.Text)
RegCloseKey keyHwnd
End Sub



Private Sub Form_Load()
    
    Me.Show
    Dim ShapeCtrl As clsTransForm
    Set ShapeCtrl = New clsTransForm  'instantiate the object from the class
    ShapeCtrl.ShapeMe cmdBrowse, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe cmdBack, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe cmdSet, RGB(0, 0, 0), True, ""
    ShapeCtrl.ShapeMe cmdSetTitle, RGB(0, 0, 0), True, ""
    Set ShapeCtrl = Nothing


GetIEback
GetHideIE
GetIESearch
GetIETitle
End Sub
