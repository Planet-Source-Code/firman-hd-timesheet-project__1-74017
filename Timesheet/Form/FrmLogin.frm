VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Begin VB.Form FrmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Timesheet"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "FrmLogin.frx":27A2
      Top             =   2760
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      Picture         =   "FrmLogin.frx":29D6
      ScaleHeight     =   1575
      ScaleWidth      =   5175
      TabIndex        =   2
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox TxtPass 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1515
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   3300
      End
      Begin VB.TextBox TxtUser 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1515
         TabIndex        =   0
         Top             =   360
         Width           =   3300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NIP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   345
         TabIndex        =   3
         Top             =   878
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   5175
      TabIndex        =   5
      Top             =   0
      Width           =   5175
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Masukan User name && password untuk dapat masuk kedalam aplikasi Timesheet ...!!!"
         ForeColor       =   &H00000040&
         Height          =   420
         Left            =   1035
         TabIndex        =   9
         Top             =   540
         Width           =   3930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log In User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   1035
         TabIndex        =   6
         Top             =   90
         Width           =   1995
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   135
         Picture         =   "FrmLogin.frx":14492
         Stretch         =   -1  'True
         Top             =   45
         Width           =   1005
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSProUmum 
      Height          =   1695
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6735
      _cx             =   11880
      _cy             =   2990
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   4194304
      BackColorFixed  =   15648682
      ForeColorFixed  =   0
      BackColorSel    =   12648447
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   0
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmLogin.frx":17CFF
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancel_Click()
 End
End Sub

Private Sub cmdOk_Click()
Dim RsShow As New ADODB.Recordset
Dim InfoLogin As Boolean
Dim NetR As NETRESOURCE
Dim ErrInfo As Long
Dim MyPass As String, MyUSer As String
 

If TxtUser.Text = "" Then
   MsgBox "User Name Belum Diisi", vbInformation
   TxtUser.SetFocus
   Exit Sub
ElseIf TxtPass.Text = "" Then
   MsgBox "Password Belum Diisi", vbInformation
   TxtPass.SetFocus
   Exit Sub
End If
StrUser = TxtUser
'StrPassword = Enc.DecryptString(TxtPass)
StrPassword = TxtPass
On Error GoTo Adaerror
ReadKoneksi
If CN.State = adStateOpen Then CN.Close
CN = "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog= '" & StrDatabase & "';Data Source='" & StrServer & "';User ID='" & StrUserDB & "';Password = '" & strPasswordDB & "';"

If CN.State = adStateClosed Then CN.Open
If RsShow.State = adStateOpen Then RsShow.Close
'    StrSQL = "Select * From login Where NIP ='" & StrUser & "' "
    StrSQL = "SELECT tbldata_user.ID,tbldata_user.NIP, tbldata_user.Password, tbldata_user.Type, Karyawan.Nama,"
    StrSQL = StrSQL & " Divisi.NM_DIV,Divisi.KD_DIV"
    StrSQL = StrSQL & " FROM tbldata_user INNER JOIN"
    StrSQL = StrSQL & " Karyawan ON tbldata_user.NIP = Karyawan.NIP INNER JOIN Divisi ON Karyawan.Kd_Divisi = Divisi.KD_DIV Where Karyawan.NIP ='" & StrUser & "'"
    RsShow.Open StrSQL, CN, adOpenStatic
     
If Not RsShow.EOF Then
    If StrUser <> RsShow!NIP Then
       MsgBox "NIP Tidak Terdaftar", vbInformation
       TxtUser = ""
       TxtUser.SetFocus
       Exit Sub
    ElseIf StrUser = RsShow!NIP And StrPassword <> Enc.DecryptString(RsShow!Password) Then
       MsgBox "Password Anda Salah, Isikan password anda dengan benar", vbInformation
       TxtPass = ""
       TxtPass.SetFocus
       Exit Sub
   ElseIf StrUser = RsShow!NIP And StrPassword = Enc.DecryptString(RsShow!Password) Then
        InfoLogin = True
        StrUser = TxtUser
        StrNamaUser = RsShow!Nama
        StrPassword = TxtPass
        StrIDUser = RsShow!ID
        StrNIPUser = RsShow!NIP
        strGroup = RsShow!Type
        NamaDivisi = RsShow!NM_DIV
        KodeDivisi = RsShow!kd_div
        
        If StrUser = 242 Then KodeDivisi = 70   ' Hrjanto/MK
        If StrUser = 112 Then KodeDivisi = 45   ' fauziah
        If StrUser = 286 Then KodeDivisi = 45   'budihartono
        If StrUser = 2865 Then KodeDivisi = 16   'intiadi
        GetTheIP
        StrSQL = "Update tbldata_user Set ComputerName = '" & ComputerName & "',IpAddress = '" & IPPc & "',UserComp = '" & UserName & "',Statuslogin = 1, last_Login = '" & Now & "',NamaUser = '" & StrNamaUser & "',NamaDivisi = '" & NamaDivisi & "',Versiapp ='" & App.Revision & "' Where NIP = '" & TxtUser & "'"
        CN.Execute StrSQL
        SocketsCleanup
        Getumum (KodeDivisi)
        Unload FrmLogin
        Set FrmLogin = Nothing
        
        If StrUser = "3578" Then
           MDIMENU.Mnukrepair.Visible = True

        Else
            CekMenu
           MDIMENU.Mnukrepair.Visible = False
           StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(ServerTime, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Login Timesheet, " & ComputerName & " &  " & IPPc & "','Login')"
            CN.Execute StrSQL
        End If
        
       
       
        MDIMENU.MnuKepUtility.Enabled = True
        MDIMENU.mnuUC.Enabled = True
        MDIMENU.mnuUN.Enabled = True
        MDIMENU.mnuUWE.Enabled = True
        MDIMENU.Mnusetkoneksi.Enabled = True
        MDIMENU.Label1 = "NIP : " & StrUser
        MDIMENU.Label2 = "NAMA : " & StrNamaUser
        MDIMENU.Label3 = "DIVISI : " & NamaDivisi
        MDIMENU.Label4 = "Server : " & StrServer
        MDIMENU.StyleButton2.Enabled = True
        MDIMENU.show
        
         If StrUser <> 3578 Then MDIMENU.show_menu (True)
        frmDateChecker.show vbModal
         MDIMENU.show_menu (False)
        If KodeDivisi = 46 Then
            LoadForm FrmTimesheet
        ElseIf UCase(strGroup) = "USER" Then
            LoadForm FrmTimesheet2
        End If
        If RsShow.State = adStateOpen Then RsShow.Close
        Set RsShow.ActiveConnection = Nothing
   End If
Else
    MsgBox "Isikan User Name dan Password Anda Dengan Benar", vbCritical, "Login"
    TxtPass.Text = ""
    TxtPass.SetFocus
End If
'End With
Exit Sub:
Adaerror:
Debug.Print err.Number
If err.Number = "-2147467259" Then
   MsgBox "Koneksi Ke Server Gagal", vbCritical
   
Else
    MsgBox err.Description
End If
 
End Sub
 
Sub Getumum(Divisi As String)
Dim x, Selisih As Integer

Dim RsUmum As New ADODB.Recordset

If RsUmum.State = adStateOpen Then RsUmum.Close
RsUmum.Open "Select kd_divisi,Kode,Nama From Project Where Nama Like '%Umum%' And kd_divisi = " & Divisi & " Order By ID", CN, adOpenStatic
Set VSProUmum.DataSource = RsUmum

With VSProUmum
        For x = 1 To .Rows - 1
            If .TextMatrix(x, 1) = Divisi Then
               ProjectUmum = .TextMatrix(x, 2)
               NMProjectUmum = .TextMatrix(x, 3)
                Exit For
            End If
        Next
    End With
End Sub
Private Sub Form_Activate()
    If Trim(TxtUser) <> "" Then
        TxtPass.SetFocus
    Else
        TxtUser.SetFocus
    End If
End Sub

Private Sub Form_GotFocus()
    TxtUser = GetSetting(App.EXEName, App.EXEName, "LastUser")

End Sub

Private Sub Form_Load()
    If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path & "\Skins\" & skinsFileName
      Skin1.ApplySkin hwnd
    End If
'    MapDriveServer
    TxtUser = GetSetting(App.EXEName, App.EXEName, "LastUser")
    
End Sub

Private Sub TxtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOk.SetFocus
End Sub

Private Sub TxtUser_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then TxtPass.SetFocus
End Sub

Private Sub TxtUser_LostFocus()
   Call SaveSetting(App.EXEName, App.EXEName, "LastUser", TxtUser)
End Sub
