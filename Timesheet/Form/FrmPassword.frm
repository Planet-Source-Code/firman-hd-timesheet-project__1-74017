VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Begin VB.Form FrmPassword 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4875
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EEC7AA&
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   135
      ScaleHeight     =   1185
      ScaleWidth      =   4560
      TabIndex        =   4
      Top             =   180
      Width           =   4560
      Begin VB.TextBox Txt2 
         Appearance      =   0  'Flat
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   2685
      End
      Begin VB.TextBox Txt1 
         Appearance      =   0  'Flat
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   90
         Width           =   2685
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verify Password :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1470
      End
   End
   Begin VB.CommandButton CmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      Picture         =   "FrmPassword.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3915
      Picture         =   "FrmPassword.frx":5266
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FrmPassword.frx":A5C6
      Top             =   0
   End
End
Attribute VB_Name = "FrmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim MyUSer As String
Dim MyPassword As String
Dim MyIDGroup As String
Dim MyRow As Long
Dim MyIDUser As String
Dim MyOldPassword As String
Dim MyNamaGroup As String
Public Add As Boolean
Private Sub CmdClose_Click()
Unload Me
Set FrmPassword = Nothing
End Sub

Private Sub CmdExit_Click()
Set FrmPassword = Nothing
Unload Me
End Sub

Private Sub cmdOk_Click()
    If Txt1 = Txt2 And Txt1 <> "" And Txt2 <> "" Then
        UserPassword = Txt1
        SimpanData
    ElseIf Txt1 <> Txt2 And Txt1 = "" And Txt2 = "" Then
        If MsgBox("Password tidak boleh dikosongkan ! " & Chr(13) & "Password standard adalah '12345'. Setuju ?", vbCritical + vbYesNo) = vbYes Then
            Exit Sub
        Else
            Txt1 = ""
            Txt2 = ""
            Txt1.SetFocus
            Exit Sub
        End If
'    Else
'        MsgBox "Password harus sama !" & Chr(13) & "Harap masukkan kembali password !", vbCritical
'        Txt1 = ""
'        Txt2 = ""
'        Txt1.SetFocus
'        Cancel = True
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
      If cmdOk.Enabled = True Then cmdOk_Click
ElseIf KeyCode = vbKeyF9 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
    Me.Caption = "Password " & Me.NamaUser
     If Len(skinsFileName) <> 0 Then
        Skin1.LoadSkin App.Path & "\Skins\" & skinsFileName
        Skin1.ApplySkin hwnd
      End If
      Txt1.Text = UserPassword
      Txt2.Text = UserPassword
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmPassword = Nothing
End Sub

Private Sub Txt1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Txt2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Public Property Get NamaUser() As String
    NamaUser = MyUSer
End Property

Public Property Let NamaUser(ByVal vNewValue As String)
    MyUSer = vNewValue
End Property

Public Property Get UserPassword() As String
    UserPassword = MyPassword
End Property

Public Property Let UserPassword(ByVal vNewValue As String)
    MyPassword = vNewValue
End Property

Public Property Get IDGroup() As String
    IDGroup = MyIDGroup
End Property

Public Property Let IDGroup(ByVal vNewValue As String)
    MyIDGroup = vNewValue
End Property

Public Property Get RowFlex() As Long
    RowFlex = MyRow
End Property

Public Property Let RowFlex(ByVal vNewValue As Long)
    MyRow = vNewValue
End Property

Public Property Get IDUser() As String
    IDUser = MyIDUser
End Property

Public Property Let IDUser(ByVal vNewValue As String)
    MyIDUser = vNewValue
End Property

Public Property Get OldPassword() As String
    OldPassword = MyOldPassword
End Property

Public Property Let OldPassword(ByVal vNewValue As String)
    MyOldPassword = vNewValue
End Property

Public Property Get NamaGroup() As String
    NamaGroup = MyNamaGroup
End Property

Public Property Let NamaGroup(ByVal vNewValue As String)
    MyNamaGroup = vNewValue
End Property

Private Sub SimpanData()
On Error GoTo Adaerror
If IDUser = "" Then
    
    StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(ServerTime, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Tambah User, User " & NamaUser & "','Password User')"
    PerintahExecute (StrSQL)
     
    StrSQL = "Insert into tbldata_user (NIP,Password,Type,StatusLogin) VALUES ('" & NamaUser & "','" & Enc.EncryptString(UserPassword) & "','" & IDGroup & "',0)"
    CN.Execute StrSQL

    Rscek.Open "Select * from tbldata_user Where type='" & IDGroup & "' and NIP='" & NamaUser & "'", CN, adOpenStatic
    If Not Rscek.EOF Then
        IDUser = Rscek!ID
        FrmDataUser.VsFlex.TextMatrix(RowFlex, 2) = IDUser
    End If
    
    MsgBox "Data berhasil disimpan !", vbInformation
    If Add = True Then FrmDataUser.VsFlex.TextMatrix(RowFlex, 6) = UserPassword
    Unload Me
Else
   
    StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(ServerTime, "yyyy-MM-dd HH:mm:ss") & "','" & StrUser & "','Rubah Password, User " & NamaUser & "','Password User')"
    PerintahExecute (StrSQL)
    
    StrSQL = "Update tbldata_user set NIP = '" & NamaUser & "', Password='" & Enc.EncryptString(UserPassword) & "' Where ID='" & IDUser & "'"
    PerintahExecute (StrSQL)
    
    MsgBox "Data berhasil disimpan !", vbInformation
     
     If Add = True Then FrmDataUser.VsFlex.TextMatrix(RowFlex, 6) = UserPassword
    
    Unload Me
End If
    
    
Exit Sub
Adaerror:
MsgBox err.Description & " Silahkan Ulangi Lagi", vbInformation
Cancel = True
End Sub



