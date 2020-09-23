VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmSetKoneksi 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00993300&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konfigurasi Koneksi"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "FrmSetKoneksi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   6840
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   6840
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   60
      ScaleHeight     =   2415
      ScaleWidth      =   6375
      TabIndex        =   6
      Top             =   120
      Width           =   6375
      Begin VB.TextBox TxtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   255
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1560
         Width           =   4155
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtUser 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   2
         Top             =   1200
         Width           =   4155
      End
      Begin VB.TextBox Txtmaster 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   4
         Top             =   1920
         Width           =   4155
      End
      Begin VB.TextBox TxtDB 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   1
         Top             =   840
         Width           =   4155
      End
      Begin VB.TextBox TxtServer 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   0
         Top             =   480
         Width           =   4155
      End
      Begin VB.PictureBox Shape2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -360
         ScaleHeight     =   255
         ScaleWidth      =   6975
         TabIndex        =   13
         Top             =   0
         Width           =   6975
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   ", Jika Bingung HUB. STAF IT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   210
            Left            =   480
            TabIndex        =   14
            Top             =   0
            Width           =   6135
         End
      End
      Begin VB.CommandButton CmdBrowse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1920
         Width           =   495
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DB Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DB User Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Master Exe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   930
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Konfigurasi Koneksi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   150
         Width           =   1650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FrmSetKoneksi.frx":0442
      Top             =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   7680
      TabIndex        =   19
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   4895
      View            =   2
      Arrange         =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Started"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Stoped"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Paused"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FrmSetKoneksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
    End Type


Private Type ENUM_SERVICE_STATUS
    lpServiceName As Long
    lpDisplayName As Long
    ServiceStatus As SERVICE_STATUS
    End Type
    '***************************************
    '     ****************
    'Constants
    '***************************************
    '     ****************
    Private Const ERROR_MORE_DATA = 234
    Private Const SERVICE_ACTIVE = &H1
    Private Const SERVICE_INACTIVE = &H2
    Private Const SERVICE_COMPLETE_LIST = &H3
    Private Const SC_MANAGER_ENUMERATE_SERVICE = &H4
    Private Const SERVICE_WIN32_OWN_PROCESS As Long = &H10
    Private Const SERVICE_WIN32_SHARE_PROCESS As Long = &H20
    Private Const SERVICE_WIN32 As Long = SERVICE_WIN32_OWN_PROCESS _
    + SERVICE_WIN32_SHARE_PROCESS
    Private Const SERVICE_CONTROL_PAUSE = &H2
    Private Const SERVICE_CONTROL_STOP = &H1
    Private Const SERVICE_CONTROL_CONTINUE = &H3
    Private Const GENERIC_EXECUTE = &H20000000
    Private Const GENERIC_READ = &H80000000
    Private Const GENERIC_WRITE = &H40000000
    Private Const GENERIC_ALL = GENERIC_EXECUTE + GENERIC_READ + GENERIC_WRITE
    Private Const DELETE = &H10000
    Private Const SERVICE_ERROR_IGNORE = &H0
    Private Const SERVICE_AUTO_START = &H2


Private Declare Function OpenSCManager Lib "advapi32.dll" _
    Alias "OpenSCManagerA" (ByVal lpMachineName As String, _
    ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long


Private Declare Function EnumServicesStatus Lib "advapi32.dll" _
    Alias "EnumServicesStatusA" (ByVal hSCManager As Long, _
    ByVal dwServiceType As Long, ByVal dwServiceState As Long, _
    lpServices As Any, ByVal cbBufSize As Long, _
    pcbBytesNeeded As Long, lpServicesReturned As Long, _
    lpResumeHandle As Long) As Long


Private Declare Function CloseServiceHandle Lib "advapi32.dll" _
    (ByVal hSCObject As Long) As Long


Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" _
    (szDest As String, szcSource As Long) As Long


Private Declare Function ControlService Lib "advapi32.dll" _
    (ByVal hService As Long, ByVal dwControl As Long, _
    lpServiceStatus As SERVICE_STATUS) As Long


Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" _
    (ByVal hService As Long, _
    ByVal dwNumServiceArgs As Long, _
    ByVal lpServiceArgVectors As Long) As Long


Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" _
    (ByVal hSCManager As Long, _
    ByVal lpServiceName As String, _
    ByVal dwDesiredAccess As Long) As Long
    


Private Declare Function ChangeServiceConfig Lib "advapi32.dll" Alias _
    "ChangeServiceConfigA" (ByVal hService As Long, _
    ByVal dwServiceType As Long, _
    ByVal dwStartType As Long, _
    ByVal dwErrorControl As Long, _
    ByVal lpBinaryPatHargaBelime As String, _
    ByVal lpLoadOrderGroup As String, _
    lpdwTagId As Long, _
    ByVal lpDependencies As String, _
    ByVal lpServiceStartName As String, _
    ByVal lpPassword As String, ByVal lpDisplayName As String) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim myNamaMenu As String
Dim LenFile As Long
Dim LenName As Long
'Dim fso As New Scripting.FileSystemObject

Private Sub CmdBrowse_Click()
On Error Resume Next
With CD
   .InitDir = "\\" & StrServer
   .Filter = "File Exe (*.EXE)|*.EXE"
   .FileName = App.EXEName & ".Exe"
   .DialogTitle = "Folder File Master " & App.EXEName
   .ShowOpen
   LenFile = Len(.FileTitle)
   LenName = Len(.FileName)
   LenFile = Abs(LenName - LenFile) - 1
   MasterUpdate = Trim(.FileName)
   MasterFile = Trim(.FileTitle)
   Txtmaster = MasterUpdate
End With
End Sub

Private Sub cmdOk_Click()
Dim Retval As String
Dim LokasiFile As String
Dim Str2 As String
If TxtServer = "" Then
   MsgBox "Silahkan Masukan Nama Servernya", vbInformation
   TxtServer.SetFocus
   Exit Sub
ElseIf TxtDB = "" Then
   MsgBox "Silahkan Masukan Nama Databasenya", vbInformation
   TxtDB.SetFocus
   Exit Sub
ElseIf TxtUser = "" Then
   MsgBox "Silahkan Masukan Database User", vbInformation
   TxtUser.SetFocus
   Exit Sub
ElseIf Txtmaster = "" Then
   MsgBox "Master File Belum Diisi", vbInformation
   CmdBrowse.SetFocus
   Exit Sub
Else
Tanya = MsgBox("Save Configuration?", vbYesNo + vbInformation, "Configuration Message")
    If Tanya = vbYes Then
        Open App.Path & "\Mod.NFO" For Output As #1
        Print #1, """" & TxtServer & """;""" & TxtDB & """;""" & MasterUpdate & """;""" & MasterFile & """;""" & TxtUser & """;""" & TxtPassword & """"
        Close #1
    End If
  
   Unload Me
End If
End Sub

Private Sub CmdClose_Click()
'    Tanya = MsgBox("Yakin Anda Ingin Keluar Dari Program?", vbYesNo + vbInformation, "Incentive Message")
'    If Tanya = vbYes Then
       Unload Me
'    End If
'    Unload Me
'    FrmSplash.show
    MsgBox "Untuk MeRefresh Settingan, Silahkan Logout Aplikasi", vbInformation
End Sub

Private Sub Command1_Click()
On Error Resume Next
With CD
'If StrRestore = "" Then
   .InitDir = StrRestore
   .Filter = "File Database (my.ini)|*.ini"
   .DialogTitle = "Lokasi Database MySQL"
   .FileName = "my.ini"
   .ShowOpen
   LenFile = Len(.FileTitle)
   LenName = Len(.FileName)
'   FileZip = .FileTitle
   LenFile = Abs(LenName - LenFile) - 1
   StrRestore = Mid(.FileName, 1, LenFile)
   TxtSource = .FileName
'   GetList
End With
End Sub

Private Sub Form_Load()
If skinsFileName = "" Then
    Call SaveSetting(App.EXEName, App.EXEName, "Skins", "DE.skn")
    skinsFileName = GetSetting(App.EXEName, App.EXEName, "Skins")
End If
If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path & "\Skins\" & skinsFileName
      Skin1.ApplySkin hwnd
    End If

ReadKoneksi
TxtServer = StrServer
TxtDB = StrDatabase
TxtUser = StrUserDB
TxtPassword = strPasswordDB
Txtmaster = MasterUpdate
Label3 = StrServer & " Mati , Silahkan Hubungi, STAF IT"
If strGroup = "IT" Then Picture1.Enabled = True
End Sub

Private Sub TxtDB_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtServer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtSource_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
