VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTimesheetPlan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planning Timesheet"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5400
   Begin VB.PictureBox Picture2 
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5355
      TabIndex        =   10
      Top             =   3120
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "List Plan"
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
         Left            =   3960
         TabIndex        =   14
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
         Left            =   2640
         TabIndex        =   13
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
         Left            =   240
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
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
         Left            =   1440
         TabIndex        =   11
         Top             =   120
         Width           =   1095
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5760
      OleObjectBlob   =   "FrmTimesheetPlan.frx":0000
      Top             =   360
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3135
      ScaleWidth      =   5520
      TabIndex        =   0
      Top             =   0
      Width           =   5520
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   45088771
         CurrentDate     =   39828
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   45088771
         CurrentDate     =   39828
      End
      Begin VSFlex8Ctl.VSFlexGrid VsFlex 
         Height          =   2055
         Left            =   1560
         TabIndex        =   6
         ToolTipText     =   "Tekan Enter Untuk Menambah Baris"
         Top             =   960
         Width           =   1935
         _cx             =   3413
         _cy             =   3625
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
         BackColorAlternate=   16707036
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
         Rows            =   50
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmTimesheetPlan.frx":0234
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
      Begin VSFlex8Ctl.VSFlexGrid cboFlex 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   120
         Width           =   1545
         _cx             =   2725
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
         ForeColorSel    =   255
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
         FormatString    =   $"FrmTimesheetPlan.frx":0279
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
      Begin VB.Label SkinLabel1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   8
         Top             =   1560
         Width           =   1815
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "NIP Karyawan"
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
         TabIndex        =   7
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
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
         Left            =   3240
         TabIndex        =   5
         Top             =   480
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         TabIndex        =   4
         Top             =   480
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "No Project"
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
         TabIndex        =   3
         Top             =   120
         Width           =   885
      End
   End
End
Attribute VB_Name = "FrmTimesheetPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancel_Click()
bersih
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()

Dim Temptgl As Date
Dim JmlTgl As Integer
Dim RsTS As New ADODB.Recordset
On Error GoTo AdaError
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
        For Lrow = 1 To VSFlex.Rows - 2
             If RsTS.State = adStateOpen Then RsTS.Close
                RsTS.Open "Select *  from tbltimesheet Where Tanggal = '" & Format(Temptgl, "yyyy/MM/dd") & "' And Nip = '" & VSFlex.TextMatrix(Lrow, 1) & "' And Status ='Plan'", CN, adOpenStatic
            If RsTS.EOF Then
            SkinLabel1.Caption = VSFlex.TextMatrix(Lrow, 1) & vbCrLf & Format(Temptgl, "dd/MM/yyyy")
            
                '-----Ambil ID -------------------
                    GetNomorID ("tbltimesheet")
                    StrKodeID = NewID
                '---------------------------------
                StrSQL = "Insert Into tbltimesheet(IDtimesheet,NIP,kd_divisi,Status,NoProject,Tanggal,JamAwal,JamAkhir,keterangan,last_update,last_user,Hari,Total) Values "
                StrSQL = StrSQL & "('" & StrKodeID & "','" & VSFlex.TextMatrix(Lrow, 1) & "','" & KodeDivisi & "','Plan','" & cboFlex & "','" & Format(Temptgl, "MM/dd/yyyy") & "','08:00','17:00','Timesheet','" & Now & "','" & StrUser & "','Kerja','480')"
                PerintahExecute (StrSQL)
                
                StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Tambah Plan Timesheet, " & StrNIPUser & " &  " & Temptgl & "','Plan Timesheet')"
                PerintahExecute (StrSQL)
                           
            End If
         Next
        End If
    End If
Next
MsgBox "Data Berhasil Disimpan", vbInformation
bersih
Exit Sub
AdaError:
MsgBox err.Description
End Sub

Private Sub Command1_Click()
LoadForm FrmListPlan
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
'    DeAktif
    bersih
    AddKaryawan
    AddProject
    If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
      Skin1.ApplySkin hwnd
    End If
   
     
End Sub
Private Sub AddKaryawan()

    Dim Cboid     As String
    Dim Cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    Cboid1 = vbNullString
    StrSQL = "select * from Karyawan Where kd_divisi = '" & KodeDivisi & "' And Status <> '14' Order By Nama"
    Rscek.Open StrSQL, CN, adOpenStatic
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("NIP") & vbTab & Rscek("Nama")
      Cboid1 = Cboid1 + Cboid
      Rscek.MoveNext
    Loop
    VSFlex.ColComboList(1) = Cboid1
     cboFlex.CellAlignment = flexAlignLeftCenter
End Sub
Private Sub AddProject()

    Dim Cboid     As String
    Dim Cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    Cboid1 = vbNullString
    StrSQL = "SELECT Project.Kode, Project.Nama FROM Project Where kd_Divisi = '" & KodeDivisi & "' And Project.Status ='Terpakai' Order By Kode"

    Rscek.Open StrSQL, CN, adOpenStatic
    Cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("Kode") & vbTab & Rscek("Nama")
      Cboid1 = Cboid1 + Cboid
      Rscek.MoveNext
    Loop
    cboFlex.ColComboList(0) = Cboid1
End Sub
Sub bersih()
cboFlex.Text = ""
SkinLabel1.Caption = ""
VSFlex.Rows = 1
VSFlex.Rows = 2
 DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker1.Value = DateSerial(Year(Now), Month(Now), 1)
    DTPicker1.Value = DateAdd("M", 0, DTPicker1.Value)
    DTPicker2.Value = DateSerial(Year(Now), Month(Now), 1)
    DTPicker2.Value = DateAdd("M", 1, DTPicker2.Value) - 1
    DTPicker1.CustomFormat = "dd/MMM/yyyy"
    DTPicker2.CustomFormat = "dd/MMM/yyyy"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmTimesheetPlan = Nothing
End Sub
Sub Aktif()
'CmdNew.Enabled = True
'CmdEdit.Enabled = True
cmdSave.Enabled = False
CmdCancel.Enabled = False
Picture1.Enabled = False
End Sub
Sub DeAktif()
'CmdNew.Enabled = False
'CmdEdit.Enabled = False
cmdSave.Enabled = True
CmdCancel.Enabled = True


End Sub

Private Sub VSFlex_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With VSFlex
    .TextMatrix(.Row, 0) = .Row
    If .TextMatrix(.Rows - 1, 1) <> "" Then
       .Rows = .Rows + 1
    End If
End With
End Sub

