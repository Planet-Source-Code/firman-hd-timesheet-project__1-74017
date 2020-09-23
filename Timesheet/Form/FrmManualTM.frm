VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmManualTM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual Timesheet"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2295
      ScaleWidth      =   5535
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.TextBox TxtNama 
         Appearance      =   0  'Flat
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
         Left            =   2160
         TabIndex        =   11
         Top             =   120
         Width           =   3255
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   5040
         OleObjectBlob   =   "FrmManualTM.frx":0000
         Top             =   1680
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Keluar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Simpan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox TxtNIP 
         Appearance      =   0  'Flat
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
         TabIndex        =   1
         Top             =   120
         Width           =   615
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20643843
         CurrentDate     =   39833
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20643843
         CurrentDate     =   39833
      End
      Begin VSFlex8Ctl.VSFlexGrid cboFlex 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   480
         Width           =   1665
         _cx             =   2937
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
         BackColorSel    =   12648447
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   16777215
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   3
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
         FormatString    =   $"FrmManualTM.frx":0234
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
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5640
         Y1              =   1440
         Y2              =   1440
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
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3240
         TabIndex        =   7
         Top             =   840
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Awal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "NIP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   330
      End
   End
End
Attribute VB_Name = "FrmManualTM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim Temptgl As Date
Dim JmlTgl As Integer
Dim Hari As String
Dim RsTS As New ADODB.Recordset
If cboFlex.Text = "" Then MsgBox "Project Belum Diisi", vbCritical: cboFlex.SetFocus: Exit Sub
 
 If MsgBox("Apakah Anda yakin ingin Menyimpan  Data ?", vbQuestion + vbYesNo, "Konfirmasi Simpan Data") = vbNo Then
    Exit Sub
 End If
JmlTgl = DateDiff("d", DTPicker1, DTPicker2)
For i = 0 To JmlTgl
    If i = 0 Then
      Temptgl = Format(DTPicker1, "MM/dd/yyyy")
    Else
      Temptgl = DateAdd("d", i, DTPicker1)
    End If
      Hari = Format(Temptgl, "ddd")
      If Hari <> "Sat" And Hari <> "Sun" And Hari <> "Sabtu" And Hari <> "Minggu" Then
        StrSQL = "select tanggallibur from kalender " & _
            "where tanggallibur = '" & Format(Temptgl, "MM/dd/yyyy") & "'"
        If Rscek.State = adStateOpen Then Rscek.Close
        Rscek.Open StrSQL, CN, adOpenStatic
        If Rscek.EOF Then
'        For Lrow = 1 To VsFlex.Rows - 2
             If RsTS.State = adStateOpen Then RsTS.Close
                RsTS.Open "Select *  from tbltimesheet Where Tanggal = '" & Format(Temptgl, "yyyy/MM/dd") & "' And Nip = '" & TxtNIP & "' And Status ='Actual'", CN, adOpenStatic
            If RsTS.EOF Then
'            SkinLabel1.Caption = Format(Temptgl, "dd/MM/yyyy")
            
                '-----Ambil ID -------------------
                    GetNomorID ("tbltimesheet")
                    StrKodeID = NewID
                '---------------------------------
                StrSQL = "Insert Into tbltimesheet(IDtimesheet,NIP,kd_divisi,Status,NoProject,Tanggal,JamAwal,JamAkhir,keterangan,last_update,last_user,Hari,Total,StatusPM,StatusDivisi) Values "
                StrSQL = StrSQL & "('" & StrKodeID & "','" & TxtNIP & "','" & KodeDivisi & "','Actual','" & cboFlex & "','" & Format(Temptgl, "MM/dd/yyyy") & "','08:00','17:00','Timesheet','" & Now & "','" & StrUser & "','Kerja','480',1,1)"
                CN.Execute StrSQL
                
                StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Tambah Manual Timesheet, " & StrNIPUser & " &  " & Temptgl & "','Manual Timesheet')"
                CN.Execute StrSQL
                           
            End If
'         Next
        End If
    End If
Next
MsgBox "Data Berhasil Disimpan", vbInformation
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
AddProject
If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
      Skin1.ApplySkin hwnd
    End If
 DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker1.Value = DateSerial(Year(Now), Month(Now), 1)
DTPicker1.Value = DateAdd("M", 0, DTPicker1.Value)
DTPicker2.Value = DateSerial(Year(Now), Month(Now), 1)
DTPicker2.Value = DateAdd("M", 1, DTPicker2.Value) - 1
DTPicker1.CustomFormat = "dd/MMM/yyyy"
DTPicker2.CustomFormat = "dd/MMM/yyyy"
End Sub
Private Sub AddProject()

    Dim Cboid     As String
    Dim Cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    Cboid1 = vbNullString
    StrSQL = "select Kode,Nama from project " & _
        "where kd_divisi='" & KodeDivisi & "' group by kode,Nama " & _
        "order by kode"
    Rscek.Open StrSQL, CN, adOpenStatic
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("Kode") & vbTab & Rscek("Nama")
      Cboid1 = Cboid1 + Cboid
      Rscek.MoveNext
    Loop
    cboFlex.ColComboList(0) = Cboid1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmManualTM = Nothing
End Sub

Private Sub TxtNIP_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
