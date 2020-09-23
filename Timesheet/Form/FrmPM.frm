VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Begin VB.Form FrmPM 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PM"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5700
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Send Message"
      Height          =   495
      Left            =   1080
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   5775
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   3840
         OleObjectBlob   =   "FrmPM.frx":0000
         Top             =   2160
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   5
         Top             =   120
         Width           =   4095
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   3
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   2
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   1
         Top             =   1560
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Project"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Project"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Divisi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "NIP PM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama PM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5640
         Y1              =   2040
         Y2              =   2040
      End
   End
End
Attribute VB_Name = "FrmPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NIP As String
Public Function Showdata(ByVal Kode As String, KodeDivisi As String)
Dim Rscek As New ADODB.Recordset
If Rscek.State = adStateOpen Then Rscek.Close
Kode = Replace(Kode, "*", "")
Kode = Trim(Kode)
StrSQL = "SELECT Project.Kd_Divisi, Project.NIP_PM,Project.Kode, Project.Nama,Karyawan.Nama AS NamaPM, Divisi.NM_DIV FROM Karyawan INNER JOIN Project ON Karyawan.NIP = Project.Nip_PM INNER JOIN Divisi ON Project.Kd_Divisi = Divisi.KD_DIV"
If UCase(strGroup) = "PTW" Or UCase(strGroup) = "IT" Then
    StrSQL = StrSQL & " WHERE PROJECT.KODE = '" & Kode & "'"
Else
    StrSQL = StrSQL & " WHERE PROJECT.KODE = '" & Kode & "' AND PROJECT.kd_divisi = '" & KodeDivisi & "'"
End If
Rscek.Open StrSQL, CN, adOpenStatic
If Not Rscek.EOF Then
    Text(0) = Rscek!Kode
    Text(1) = Rscek!NM_DIV
    Text(1).Tag = Rscek!kd_divisi
    Text(2) = Rscek!Nama
    Text(3) = Rscek!NIP_PM
    Text(4) = Rscek!NamaPM
End If

End Function
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
With FrmMessage
    .NIPKry = NIP
    .NIPPM = Text(3)
    .Project = Text(0)
    .NamaPM = Trim(Text(4))
    .kdDivisi = Text(1).Tag
    .TxtFrom = StrNamaUser
    .TxtFrom.Tag = StrUser
    .show vbModal
End With
End Sub

Private Sub Form_Activate()
 If Text(3).Text = "" Then Command2.Enabled = False
End Sub

Private Sub Form_Load()
 If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
      Skin1.ApplySkin hwnd
    End If
   
End Sub
