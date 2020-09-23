VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmProject 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Project"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6675
      TabIndex        =   14
      Top             =   3360
      Width           =   6735
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Clos&e"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   17
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3375
      ScaleWidth      =   6615
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   2160
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   6120
         OleObjectBlob   =   "FrmProject.frx":0000
         Top             =   2160
      End
      Begin VB.TextBox txtketerangan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2520
         Width           =   4935
      End
      Begin VB.ComboBox combostatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   3
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtnamaproject 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   4935
      End
      Begin VB.TextBox txtkode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
      Begin VSFlex8Ctl.VSFlexGrid TxtNIPPM 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   2160
         Width           =   2385
         _cx             =   4207
         _cy             =   556
         Appearance      =   0
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
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   16777215
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   0
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmProject.frx":0234
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   60
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
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
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   52232195
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   52232195
         CurrentDate     =   39931
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Dari "
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
         Left            =   840
         TabIndex        =   21
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "S.D"
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
         Left            =   840
         TabIndex        =   20
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
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
         Left            =   240
         TabIndex        =   13
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama PM"
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
         Left            =   360
         TabIndex        =   12
         Top             =   2160
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   720
         TabIndex        =   11
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Project"
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
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Project"
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
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1200
      End
   End
End
Attribute VB_Name = "FrmProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancel_Click()
Aktif
bersih
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub CmdEdit_Click()
DeAktif
'TxtNIPPM.Text = Text1
Text1.Visible = False
End Sub

Private Sub CmdNew_Click()
DeAktif
bersih
txtkode.SetFocus
End Sub
Sub bersih()
txtketerangan.Text = ""
txtnamaproject.Text = ""
txtkode.Tag = ""
txtkode.Text = ""
txtnamaproject.Text = ""
TxtNIPPM.Text = ""
combostatus.Text = ""
Text1.Text = ""
DTPicker1.CustomFormat = "dd/MM/yyyy"
DTPicker2.CustomFormat = "dd/MM/yyyy"
DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub

Private Sub cmdSave_Click()
Dim strkddiv As String
If Rscek.State = adStateOpen Then Rscek.Close
Rscek.Open "SELECT * FROM KARYAWAN WHERE Nama = '" & TxtNIPPM.Text & "'", CN, adOpenStatic
If Not Rscek.EOF Then
   strkddiv = KodeDivisi  'Rscek!kd_divisi
   TxtNIPPM.Text = Rscek!NIP
Else
    MsgBox "Divisi NIP PM Tidak Diketahui", vbCritical
    Exit Sub
End If
 
If txtkode.Tag = "" Then
    If Rscek.State = adStateOpen Then Rscek.Close
       Rscek.Open "Select * From Project Where Kode = '" & txtkode & "'"
       If Not Rscek.EOF Then
          MsgBox "Kode Project Sudah Ada", vbCritical
          Exit Sub
       End If
    StrSQL = "insert into project (kd_divisi,kode,nama,status,nip_pm,keterangan, " & _
            "tgl_update,Tgl_akhir,UserInput) values ('" & strkddiv & "','" & txtkode.Text & "' " & _
            ",'" & txtnamaproject.Text & "','" & combostatus.Text & "' " & _
            ",'" & TxtNIPPM.Text & "','" & txtketerangan.Text & "','" & Format(DTPicker1, "yyyy/MM/dd") & "','" & Format(DTPicker2, "yyyy/MM/dd") & "','" & StrUser & "' )"
     CN.Execute StrSQL
     StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Tambah Data Project, " & txtkode & "','Project')"
     PerintahExecute (StrSQL)
Else
    If Text1.Tag <> strkddiv Then
        MsgBox "Maaf Anda Tidak berhak Mengubah Project Ini", vbCritical
        Exit Sub
    End If
     StrSQL = "update project set kode = '" & txtkode.Text & "', nama = '" & txtnamaproject.Text & "' " & _
            ", status = '" & combostatus.Text & "', nip_pm = '" & TxtNIPPM.Text & "' " & _
            ", keterangan = '" & txtketerangan.Text & "', tgl_update = '" & Format(DTPicker1, "yyyy/MM/dd") & "',tgl_Akhir = '" & Format(DTPicker2, "yyyy/MM/dd") & "'" & _
            ", UserInput='" & StrUser & "' where id = '" & txtkode.Tag & "'"
    CN.Execute StrSQL
    StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Rubah Data Project, " & txtkode & "','Project')"
    PerintahExecute (StrSQL)
    Text1.Tag = ""
End If
bersih
Aktif
End Sub

Private Sub combostatus_GotFocus()
combostatus.BackColor = &HC0FFFF
End Sub

Private Sub combostatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub combostatus_LostFocus()
combostatus.BackColor = vbWhite
End Sub

Private Sub Form_Load()
combostatus.AddItem "Terpakai"
combostatus.AddItem "Tidak Terpakai"
Aktif
AddKaryawan
bersih
 If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
      Skin1.ApplySkin hwnd
    End If

End Sub
Sub AddKaryawan()

    Dim Cboid     As String
    Dim cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    cboid1 = vbNullString
    StrSQL = "SELECT karyawan.NIP, karyawan.Nama, Divisi.NM_DIV FROM karyawan INNER JOIN Divisi ON karyawan.Kd_Divisi = Divisi.KD_DIV Where Status <> '14' And Len(NIP) < 5 order by Nama,NM_DIV"
    Rscek.Open StrSQL, CN, adOpenStatic
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("Nama") & vbTab & Rscek("NIP") & vbTab & Rscek("NM_DIV")
      cboid1 = cboid1 + Cboid
       Rscek.MoveNext
    Loop
    TxtNIPPM.ColComboList(0) = cboid1
End Sub

Public Sub Showdata(ID As String)
On Error Resume Next
Dim RsProject As New ADODB.Recordset
If RsProject.State = adStateOpen Then RsProject.Close
RsProject.Open "Select * From Project Where ID = '" & ID & "'", CN, adOpenStatic
If Not RsProject.EOF Then
    txtkode = RsProject!Kode
    Text1 = RsProject!NIP_PM
    txtketerangan = RsProject!Keterangan
    txtnamaproject.Text = RsProject!Nama
    Text1.Tag = RsProject!kd_divisi
    combostatus.Text = RsProject!status
    txtkode.Tag = RsProject!ID
    DTPicker1 = Format(RsProject!Tgl_Update, "dd/MM/yyyy")
    If Trim(RsProject!Tgl_akhir) <> "" Then DTPicker2 = Format(RsProject!Tgl_akhir, "dd/MM/yyyy")
    CmdEdit.Enabled = True
    Text1.Visible = False
End If
If RsProject.State = adStateOpen Then RsProject.Close
RsProject.Open "SELECT * FROM KARYAWAN WHERE NIP = '" & Text1 & "'", CN, adOpenStatic
If Not RsProject.EOF Then
   TxtNIPPM.Text = RsProject!Nama
'   Text1.Text = Rscek!Nama
End If
End Sub
Sub Aktif()
CmdNew.Enabled = True
CmdEdit.Enabled = False
cmdSave.Enabled = False
CmdCancel.Enabled = False
Picture1.Enabled = False
End Sub
Sub DeAktif()
CmdNew.Enabled = False
CmdEdit.Enabled = False
cmdSave.Enabled = True
CmdCancel.Enabled = True
Picture1.Enabled = True
Text1.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmProject = Nothing
End Sub

Private Sub txtketerangan_GotFocus()
txtketerangan.BackColor = &HC0FFFF
End Sub

Private Sub txtketerangan_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtketerangan_LostFocus()
txtketerangan.BackColor = vbWhite
End Sub

Private Sub txtkode_GotFocus()
txtkode.BackColor = &HC0FFFF
End Sub

Private Sub txtkode_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96
If KeyAscii = 13 Then
    If CekGanda(txtkode) = False Then
        SendKeys "{TAB}"
    Else
        MsgBox "Kode Project Sudah Ada", vbCritical
        txtkode.SetFocus
        Exit Sub
    End If
End If
End Sub
Function CekGanda(Nama As String) As Boolean
If Rscek.State = adStateOpen Then Rscek.Close
CekGanda = False
Rscek.Open "Select * From Project Where Kode = '" & Nama & "' And kd_divisi = '" & KodeDivisi & "'", CN, adOpenStatic
If Not Rscek.EOF Then
    CekGanda = True
End If
End Function
Private Sub txtkode_LostFocus()
txtkode.BackColor = vbWhite
End Sub

Private Sub txtnamaproject_GotFocus()
txtnamaproject.BackColor = &HC0FFFF
End Sub

Private Sub txtnamaproject_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtnamaproject_LostFocus()
txtnamaproject.BackColor = vbWhite
End Sub

Private Sub txtnippm_GotFocus()
TxtNIPPM.BackColor = &HC0FFFF
End Sub

Private Sub txtnippm_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtnippm_LostFocus()
TxtNIPPM.BackColor = vbWhite

End Sub
