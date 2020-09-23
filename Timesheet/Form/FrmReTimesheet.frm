VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmReTimesheet 
   Caption         =   "Repair Timesheet"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   10830
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   10770
      TabIndex        =   0
      Top             =   0
      Width           =   10830
      Begin VB.CommandButton Command6 
         Caption         =   "Update Hari"
         Height          =   495
         Left            =   8520
         TabIndex        =   12
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Load Hari"
         Height          =   495
         Left            =   9480
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Load Timesheet"
         Height          =   495
         Left            =   10320
         TabIndex        =   9
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Fix Timesheet"
         Height          =   495
         Left            =   12240
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Fix"
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load lembur"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   120
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
         Format          =   55836675
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   6240
         TabIndex        =   4
         Top             =   120
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
         Format          =   55836675
         CurrentDate     =   39931
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   9975
      Left            =   -360
      TabIndex        =   5
      Top             =   840
      Width           =   8175
      _cx             =   14420
      _cy             =   17595
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483634
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
      GridLines       =   1
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
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   1
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
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
      Editable        =   0
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
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   55836674
         CurrentDate     =   39940
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   55836674
         CurrentDate     =   39940
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlex 
      Height          =   9975
      Left            =   7920
      TabIndex        =   10
      Top             =   840
      Width           =   7215
      _cx             =   12726
      _cy             =   17595
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483634
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
      GridLines       =   1
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
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   1
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
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
      Editable        =   0
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
End
Attribute VB_Name = "FrmReTimesheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim Lrow, JmlTgl As Long
Dim Str As Integer
Dim RsLembur As New ADODB.Recordset
Dim Temptgl As Date
Dim FixedName As Boolean
Dim SelRecord As Long
'StrSQL = " SELECT Timesheet2.IDTimesheet, Timesheet2.ID, Timesheet2.NIP,"
'StrSQL = StrSQL & "Timesheet2.Kd_Divisi, Timesheet2.Tanggal,"
'StrSQL = StrSQL & "Timesheet2.Status, Timesheet2.Slot, Timesheet2.JamAwal,"
'StrSQL = StrSQL & "Timesheet2.JamAkhir, Absensi.Masuk,"
'StrSQL = StrSQL & "Timesheet2.Keterangan, Timesheet2.StatusPM,"
'StrSQL = StrSQL & "Timesheet2.StatusDivisi, Timesheet2.last_update,Timesheet2.last_user"
'StrSQL = StrSQL & " FROM Timesheet2 INNER JOIN Absensi ON Timesheet2.NIP = Absensi.NIP AND Timesheet2.Tanggal = Absensi.Tgl Order by Timesheet2.Tanggal, Timesheet2.ID,Timesheet2.NIP"
'Set fg.DataSource = Rscek
'Set Rscek = Nothing
'With Fg
'    FixedName = False
'    JmlTgl = DateDiff("d", DTPicker1, DTPicker2)
'    For I = 0 To JmlTgl
'        If I = 0 Then
'          TempTgl = Format(DTPicker1, "MM/dd/yyyy")
'        Else
'          TempTgl = DateAdd("d", I, DTPicker1)
'        End If
'            If Rscek.State = adStateOpen Then Rscek.Close
'
'            StrSQL = "Select * From Timesheet2 Where Tanggal = '" & Format(TempTgl, "yyyy/MM/dd") & "' Order by Tanggal,ID"
'            Rscek.Open StrSQL, CN, adOpenStatic
'        If Not Rscek.EOF Then
'            If FixedName = False Then
'                    lCols = Rscek.Fields.Count
''                   .Redraw = flexRDNone
'                    .Rows = 1
'                    .Rows = 2
'                    .Cols = lCols
'                    .Row = 0
'                    For lCol = 1 To lCols - 1
'                        .Col = lCo1
'                        .TextMatrix(0, lCol) = Rscek(lCol).Name
''                        .ColWidth(lCol) = 1500
'                        .FixedAlignment(lCol) = flexAlignCenterCenter
'                    Next
'                    FixedName = True
'            End If
''            Exit Sub
''    Do Until Rscek.EOF
'           If Rscek.RecordCount > 0 Then
'                Rscek.AbsolutePosition = 1
'                    Lrow = .Rows - 1
'             For SelRecord = 1 To Rscek.RecordCount
'                        For lCol = 1 To lCols - 1
'                            .Col = lCol
'                            .Row = Lrow
'                            .Text = Rscek.Fields(lCol).Value
'                        Next
'                         .TextMatrix(Lrow, 0) = SelRecord
'                         .TextMatrix(Lrow, 1) = ""
'                         Rscek.MoveNext
'                         Lrow = Lrow + 1
'                        .Rows = Lrow + 1
'            Next
'            End If
'StrSQL = Shell(App.HelpFile, vbNormalFocus
Dim Jam1, Jam2 As String
With fg
If Rscek.State = adStateOpen Then Rscek.Close
StrSQL = "Select * From Lembur1 Where Tanggal Between '" & Format(DTPicker1, "yyyy/MM/dd") & "' And '" & Format(DTPicker2, "yyyy/MM/dd") & "' And Keterangan <>'OK' Order By Tanggal ASC"
Rscek.Open StrSQL, CN, adOpenStatic
Set fg.DataSource = Rscek

Sleep 1000
Set Rscek = Nothing

.Cols = .Cols + 5
.TextMatrix(0, 13) = "Kode Divisi"
.TextMatrix(0, 14) = "Masuk"
.TextMatrix(0, 15) = "Keluar"
.TextMatrix(0, 16) = "Status"
.TextMatrix(0, 17) = "Keterangan"
For Lrow = 1 To .Rows - 1
If Rscek.State = adStateOpen Then Rscek.Close
Rscek.Open "SELECT * From Absensi Where NIP = '" & .TextMatrix(Lrow, 2) & "' And Tgl = '" & Format(.TextMatrix(Lrow, 5), "MM/dd/yyyy") & "'", CN, adOpenStatic
     If Not Rscek.EOF Then
'        .TextMatrix(Lrow, 13) = Rscek!kd_divisi
        .TextMatrix(Lrow, 14) = Format(Rscek!masuk, "HH:mm")
        .TextMatrix(Lrow, 15) = Format(Rscek!keluar, "HH:mm")
       
    End If
    .TextMatrix(Lrow, 16) = "Actual"
     .TextMatrix(Lrow, 17) = "Lembur"
    If Rscek.State = adStateOpen Then Rscek.Close
        Rscek.Open "SELECT * From Karyawan Where NIP = '" & .TextMatrix(Lrow, 2) & "'", CN, adOpenStatic
     If Not Rscek.EOF Then .TextMatrix(Lrow, 13) = Rscek!kd_divisi
      Str = InStr(.TextMatrix(Lrow, 4), "*")
    If Str = 0 Then
       .TextMatrix(Lrow, 9) = 1
       .TextMatrix(Lrow, 10) = 1
    Else
         .TextMatrix(Lrow, 9) = 0
       .TextMatrix(Lrow, 10) = 0
    End If
    If .TextMatrix(Lrow, 6) = "" Then .TextMatrix(Lrow, 6) = "00:00"
    Jam1 = Replace(.TextMatrix(Lrow, 6), ".", ":")
    Jam2 = Replace(.TextMatrix(Lrow, 7), ".", ":")
    If Jam2 = "24:00" Then Jam2 = "00:00"
    If Jam1 = "24:00" Then Jam1 = "00:00"
    DTPicker3 = CDate(Jam1)
    DTPicker4 = CDate(Jam2)
    .TextMatrix(Lrow, 8) = Abs(DateDiff("n", Format(DTPicker3, "hh:mm"), Format(DTPicker4, "hh:mm")))

Next
End With
End Sub


Private Sub Command2_Click()
Dim Lrow As Long
Dim Hari, StatusHari As String
With fg
 For Lrow = 1 To .Rows - 1
    .TextMatrix(Lrow, 4) = Replace(.TextMatrix(Lrow, 4), "*", "")
    .TextMatrix(Lrow, 6) = Replace(.TextMatrix(Lrow, 6), ".", ":")
    .TextMatrix(Lrow, 7) = Replace(.TextMatrix(Lrow, 7), ".", ":")
    Command2.Caption = Lrow
        Hari = Format(.TextMatrix(Lrow, 5), "ddd")
        If Hari <> "Sat" And Hari <> "Sun" And Hari <> "Sabtu" And Hari <> "Minggu" Then
            StatusHari = "Kerja"
             StrSQL = "select tanggallibur from kalender " & _
                 "where tanggallibur = '" & Format(DTPicker1, "MM/dd/yyyy") & "'"
             If Rscek.State = adStateOpen Then Rscek.Close
             Rscek.Open StrSQL, CN, adOpenStatic
             If Not Rscek.EOF Then
                 StatusHari = "Libur"
             End If
         Else
             StatusHari = "Libur"
         End If
    If Rscek.State = adStateOpen Then Rscek.Close
    Set Rscek = Nothing
    Rscek.Open "Select * From TblTimesheet Where NIP ='" & .TextMatrix(Lrow, 2) & "' And Tanggal = '" & Format(.TextMatrix(Lrow, 5), "MM/dd/yyyy") & "'", CN, adOpenStatic
    If Rscek.EOF Then
        StrSQL = "Update Lembur1 Set Keterangan = 'Ok' Where NIp ='" & .TextMatrix(Lrow, 2) & "' And Tanggal='" & Format(.TextMatrix(Lrow, 5), "MM/dd/yyyy") & "' And ID='" & .TextMatrix(Lrow, 1) & "'"
        CN.Execute StrSQL
         '-----Ambil ID -------------------
            GetNomorID ("tbltimesheet")
            StrKodeID = NewID
        '---------------------------------
       
        StrSQL = "Insert Into tbltimesheet (IDtimesheet,NIP,Hari,NoProject,Tanggal,JamAwal,JamAkhir,Total,StatusPM,StatusDivisi,Last_update ,last_user,kd_divisi,Masuk,Keluar,Status,Keterangan)"
        StrSQL = StrSQL & "Values('" & StrKodeID & "','" & .TextMatrix(Lrow, 2) & "','" & StatusHari & "','" & .TextMatrix(Lrow, 4) & "','" & Format(.TextMatrix(Lrow, 5), "yyyy/MM/dd") & "','" & Format(.TextMatrix(Lrow, 6), "hh:mm") & "','" & Format(.TextMatrix(Lrow, 7), "hh:mm") & "','" & .TextMatrix(Lrow, 8) & "','" & .TextMatrix(Lrow, 9) & "','" & .TextMatrix(Lrow, 10) & "','" & Now & "','" & StrUser & "','" & .TextMatrix(Lrow, 13) & "','" & .TextMatrix(Lrow, 14) & "','" & .TextMatrix(Lrow, 15) & "','" & .TextMatrix(Lrow, 16) & "','Lembur')"
        CN.Execute StrSQL
    Else
        StrSQL = "Update Lembur1 Set Keterangan = 'Ok' Where NIp ='" & .TextMatrix(Lrow, 2) & "' And Tanggal='" & Format(.TextMatrix(Lrow, 5), "MM/dd/yyyy") & "' And ID='" & .TextMatrix(Lrow, 1) & "'"
        CN.Execute StrSQL
        
       StrSQL = "Update tbltimesheet Set noProject = '" & .TextMatrix(Lrow, 4) & "',JamAwal= '" & Format(.TextMatrix(Lrow, 6), "hh:mm") & "', JamAkhir = '" & Format(.TextMatrix(Lrow, 7), "hh:mm") & "',Hari = '" & StatusHari & "' Where NIP ='" & .TextMatrix(Lrow, 2) & "' And Tanggal = '" & Format(.TextMatrix(Lrow, 5), "MM/dd/yyyy") & "'"
       CN.Execute StrSQL
    End If
Next
End With
fg.Rows = 1
MsgBox "Finish"
End Sub

Private Sub Command3_Click()
Dim Lrow As Long
Dim Hari, StatusHari As String
With VsFlex
 For Lrow = 1 To .Rows - 1
      Hari = Format(.TextMatrix(Lrow, 5), "ddd")
        If Hari <> "Sat" And Hari <> "Sun" And Hari <> "Sabtu" And Hari <> "Minggu" Then
            StatusHari = "Kerja"
             StrSQL = "select tanggallibur from kalender " & _
                 "where tanggallibur = '" & Format(DTPicker1, "MM/dd/yyyy") & "'"
             If Rscek.State = adStateOpen Then Rscek.Close
             Rscek.Open StrSQL, CN, adOpenStatic
             If Not Rscek.EOF Then
                 StatusHari = "Libur"
             End If
         Else
             StatusHari = "Libur"
         End If
     '-----Ambil ID -------------------
        GetNomorID ("tbltimesheet")
        StrKodeID = NewID
    '---------------------------------
       Command2.Caption = Lrow
    StrSQL = "Update Timesheet2 Set Keterangan = 'Ok' Where NIp ='" & .TextMatrix(Lrow, 2) & "' And Tanggal='" & Format(.TextMatrix(Lrow, 8), "MM/dd/yyyy") & "'"
'    PerintahExecute (StrSQL)
    CN.Execute StrSQL
        .TextMatrix(Lrow, 6) = Replace(.TextMatrix(Lrow, 6), "*", "")

    StrSQL = "Insert Into tbltimesheet (IDtimesheet,NIP,Hari,NoProject,Tanggal,JamAwal,JamAkhir,Total,StatusPM,StatusDivisi,Last_update ,last_user,kd_divisi,Masuk,Keluar,Status,Keterangan)"
    StrSQL = StrSQL & "Values('" & StrKodeID & "','" & .TextMatrix(Lrow, 2) & "','" & StatusHari & "','" & .TextMatrix(Lrow, 6) & "','" & Format(.TextMatrix(Lrow, 8), "yyyy/MM/dd") & "','" & Format(.TextMatrix(Lrow, 3), "hh:mm") & "','" & Format(.TextMatrix(Lrow, 4), "hh:mm") & "','" & .TextMatrix(Lrow, 14) & "','" & .TextMatrix(Lrow, 10) & "','" & .TextMatrix(Lrow, 11) & "','" & Now & "','" & StrUser & "','" & .TextMatrix(Lrow, 12) & "','" & .TextMatrix(Lrow, 9) & "','" & .TextMatrix(Lrow, 15) & "','" & .TextMatrix(Lrow, 5) & "','" & .TextMatrix(Lrow, 7) & "')"
    CN.Execute StrSQL
   
Next
End With
VsFlex.Rows = 1
MsgBox "Finish"
End Sub

Private Sub Command4_Click()
Dim Jam1, Jam2 As String
Dim Str As String
Dim RsTS As New ADODB.Recordset
With fg
If Rscek.State = adStateOpen Then Rscek.Close
StrSQL = "Select * From Timesheet2 Where Keterangan <> 'Ok' And Tanggal Between '" & Format(DTPicker1, "yyyy/MM/dd") & "' And '" & Format(DTPicker2, "yyyy/MM/dd") & "' Order By Tanggal, NIP, Status ,ID ASC"
Rscek.Open StrSQL, CN, adOpenStatic
Set fg.DataSource = Rscek

Sleep 1000
Set Rscek = Nothing
 If RsTS.State = adStateOpen Then RsTS.Close
    StrSQL = "Select IDtimesheet,NIP,JamAwal As [Jam Awal],JamAkhir AS [Jam Akhir],Status,NoProject As Project,Keterangan,Tanggal,Masuk,StatusPM,StatusDivisi,kd_divisi,Hari,Hari AS Total,Masuk AS Keluar From tbltimesheet Where Tanggal = '" & Format(DTPicker1, "yyyy/MM/dd") & "' And NIP = '" & StrNIPUser & "'  AND Status='Actual' Order By IDTimesheet Asc"
    RsTS.Open StrSQL, CN, adOpenStatic
    Set VsFlex.DataSource = RsTS
'    VsFlex.ColDataType(1) = flexDTBoolean
    Hari = Format(DTPicker1, "ddd")

.Cols = .Cols + 5
.ColWidth(7) = 300
.ColWidth(8) = 300
.ColWidth(9) = 300
.ColWidth(10) = 300
.ColWidth(11) = 300
.TextMatrix(0, 12) = "JamAwal"
.TextMatrix(0, 13) = "JamAkhir"
.TextMatrix(0, 14) = "Total"
.TextMatrix(0, 15) = "Masuk"
.TextMatrix(0, 16) = "Keluar"
For Lrow = 1 To .Rows - 1
    Command4.Caption = Lrow
        .TextMatrix(Lrow, 0) = Lrow
'        .TextMatrix(0, 0) = .Rows
        .TextMatrix(Lrow, 11) = "Timesheet"
        .TextMatrix(Lrow, 9) = Now
        .TextMatrix(Lrow, 10) = StrNIPUser
         Str = InStr(.TextMatrix(Lrow, 6), "*")

        If .TextMatrix(Lrow, 5) = "Actual" And Str = 0 Then
           .TextMatrix(Lrow, 7) = 1
           .TextMatrix(Lrow, 8) = 1
        End If
        .TextMatrix(Lrow, 6) = Replace(.TextMatrix(Lrow, 6), "*", "")
         Jam1 = Trim(.TextMatrix(Lrow, 1))
            Select Case Jam1
                Case 1
                     .TextMatrix(Lrow, 12) = "08:00"
                     .TextMatrix(Lrow, 13) = "08:30"
                Case 2
                    .TextMatrix(Lrow, 12) = "08:30"
                     .TextMatrix(Lrow, 13) = "09:00"
                Case 3
                     .TextMatrix(Lrow, 12) = "09:00"
                     .TextMatrix(Lrow, 13) = "09:30"
                Case 4
                    .TextMatrix(Lrow, 12) = "09:30"
                     .TextMatrix(Lrow, 13) = "10:00"
                Case 5
                     .TextMatrix(Lrow, 12) = "10:00"
                     .TextMatrix(Lrow, 13) = "10:30"
                Case 6
                     .TextMatrix(Lrow, 12) = "10:30"
                     .TextMatrix(Lrow, 13) = "11:00"
                Case 7
                     .TextMatrix(Lrow, 12) = "11:00"
                     .TextMatrix(Lrow, 13) = "11:30"
                Case 8
                    .TextMatrix(Lrow, 12) = "11:30"
                     .TextMatrix(Lrow, 13) = "12:00"
                Case 9
                      .TextMatrix(Lrow, 12) = "ISTIRAHAT"
                Case 10
                     .TextMatrix(Lrow, 13) = "ISTIRAHAT"
                Case 11
                      .TextMatrix(Lrow, 12) = "13:00"
                     .TextMatrix(Lrow, 13) = "13:30"
                Case 12
                     .TextMatrix(Lrow, 12) = "13:30"
                     .TextMatrix(Lrow, 13) = "14:00"
                Case 13
                    .TextMatrix(Lrow, 12) = "14:00"
                     .TextMatrix(Lrow, 13) = "14:30"
                Case 14
                    .TextMatrix(Lrow, 12) = "14:30"
                     .TextMatrix(Lrow, 13) = "15:00"
                Case 15
                    .TextMatrix(Lrow, 12) = "15:00"
                     .TextMatrix(Lrow, 13) = "15:30"
                Case 16
                     .TextMatrix(Lrow, 12) = "15:30"
                     .TextMatrix(Lrow, 13) = "16:00"
                Case 17
                     .TextMatrix(Lrow, 12) = "16:00"
                     .TextMatrix(Lrow, 13) = "16:30"
                Case 18
                    .TextMatrix(Lrow, 12) = "16:30"
                     .TextMatrix(Lrow, 13) = "17:00"
            End Select

        If .TextMatrix(Lrow, 13) <> "" Then
           DTPicker3 = Format(.TextMatrix(Lrow, 12), "HH:mm")
           DTPicker4 = Format(.TextMatrix(Lrow, 13), "HH:mm")
           .TextMatrix(Lrow, 14) = Abs(DateDiff("n", Format(DTPicker3, "hh:mm"), Format(DTPicker4, "hh:mm")))
           End If
        Str = Showdata(.TextMatrix(Lrow, 2), .TextMatrix(Lrow, 4), .TextMatrix(Lrow, 5), Lrow, .TextMatrix(Lrow, 6), .TextMatrix(Lrow - 1, 6))
'   If Lrow = 17 Then Exit Sub
Next
End With
End Sub
Function Showdata(NIP As String, Tanggal As String, status As String, Lrow As Integer, Project As String, LastProject As String) As String
Dim Row As Integer
Dim SudahAda As Boolean
With VsFlex
    SudahAda = False
    If Project = "Telat" Then Exit Function
    For Row = 1 To .Rows - 1
        If .TextMatrix(Row, 2) = NIP And .TextMatrix(Row, 8) = Tanggal And .TextMatrix(Row, 5) = status And .TextMatrix(Row, 6) = Project And LastProject = Project Then
'            If .Rows > 2 Then
'                If .TextMatrix(Row - 1, 6) = Project Then
'                    SudahAda = True
'                Else
'                    SudahAda = False
'                End If
'            Else
                SudahAda = True
'            End If
'        Else
'             SudahAda = False
        End If

    Next
'    If SudahAda = True Then
'        MsgBox Row
'    End If
    If SudahAda = False Then
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 14) = 0
        If Rscek.State = adStateOpen Then Rscek.Close
        Rscek.Open "SELECT * From Absensi Where NIP = '" & NIP & "' And Tgl = '" & Format(Tanggal, "MM/dd/yyyy") & "'", CN, adOpenStatic
         If Not Rscek.EOF Then
            .TextMatrix(.Rows - 1, 9) = Format(Rscek!masuk, "HH:mm")
            .TextMatrix(.Rows - 1, 15) = Format(Rscek!keluar, "HH:mm")
    
        End If
    End If
        .TextMatrix(.Rows - 1, 0) = .Rows - 1
        .TextMatrix(.Rows - 1, 2) = NIP
        If SudahAda = False Then
            .TextMatrix(.Rows - 1, 3) = fg.TextMatrix(Lrow, 12)
        End If
        If .TextMatrix(.Rows - 1, 14) = "" Then .TextMatrix(.Rows - 1, 14) = 0
        .TextMatrix(.Rows - 1, 4) = fg.TextMatrix(Lrow, 13)
        .TextMatrix(.Rows - 1, 5) = fg.TextMatrix(Lrow, 5)
        .TextMatrix(.Rows - 1, 6) = fg.TextMatrix(Lrow, 6)
        .TextMatrix(.Rows - 1, 7) = fg.TextMatrix(Lrow, 11)
        .TextMatrix(.Rows - 1, 8) = fg.TextMatrix(Lrow, 4)
        .TextMatrix(.Rows - 1, 10) = fg.TextMatrix(Lrow, 7)
        .TextMatrix(.Rows - 1, 11) = fg.TextMatrix(Lrow, 8)
        .TextMatrix(.Rows - 1, 12) = fg.TextMatrix(Lrow, 3)
        .TextMatrix(.Rows - 1, 13) = "Kerja"
        .TextMatrix(.Rows - 1, 14) = CCur(.TextMatrix(.Rows - 1, 14)) + CCur(fg.TextMatrix(Lrow, 14))
End With
End Function

Private Sub Command5_Click()
Dim Hari, StatusHari As String
If Rscek.State = adStateOpen Then Rscek.Close
Rscek.Open "Select * From TblTimesheet Where Hari = ''", CN, adOpenStatic
Set fg.DataSource = Rscek

With fg
    For Lrow = 1 To .Rows - 1
    Command6.Caption = Lrow
         Hari = Format(.TextMatrix(Lrow, 9), "ddd")
   If Hari <> "Sat" And Hari <> "Sun" And Hari <> "Sabtu" And Hari <> "Minggu" Then
       StatusHari = "Kerja"
        StrSQL = "select tanggallibur from kalender " & _
            "where tanggallibur = '" & Format(.TextMatrix(Lrow, 9), "MM/dd/yyyy") & "'"
        If Rscek.State = adStateOpen Then Rscek.Close
        Rscek.Open StrSQL, CN, adOpenStatic
        If Not Rscek.EOF Then
            StatusHari = "Libur"
        End If
    Else
        StatusHari = "Libur"
    End If
      .TextMatrix(Lrow, 6) = StatusHari
    Next
End With
End Sub

Private Sub Command6_Click()
'Dim Hari, StatusHari As String
With fg
    For Lrow = 1 To .Rows - 1
    Command6.Caption = Lrow
         Hari = Format(.TextMatrix(Lrow, 9), "ddd")
   If Hari <> "Sat" And Hari <> "Sun" Then
       StatusHari = "Kerja"
        StrSQL = "select tanggallibur from kalender " & _
            "where tanggallibur = '" & Format(.TextMatrix(Lrow, 9), "MM/dd/yyyy") & "'"
        If Rscek.State = adStateOpen Then Rscek.Close
        Rscek.Open StrSQL, CN, adOpenStatic
        If Not Rscek.EOF Then
            StatusHari = "Libur"
        End If
    Else
        StatusHari = "Libur"
    End If
        If CN.State = adStateClosed Then CN.Open
         StrSQL = "Update tbltimesheet Set Hari = '" & .TextMatrix(Lrow, 6) & "' Where IDtimesheet='" & .TextMatrix(Lrow, 1) & "' "
         CN.Execute StrSQL
    Next
    .Rows = 1
End With
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker1.Value = DateSerial(Year(Now), Month(Now), 1)
DTPicker1.Value = DateAdd("M", 0, DTPicker1.Value)
DTPicker2.Value = DateSerial(Year(Now), Month(Now), 1)
DTPicker2.Value = DateAdd("M", 1, DTPicker2.Value) - 1
DTPicker1.CustomFormat = "dd/MMM/yyyy"
DTPicker2.CustomFormat = "dd/MMM/yyyy"

End Sub


Private Sub Form_Resize()
 On Error Resume Next

    With fg
'    .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left
    .Height = Me.Height - 1500
    End With
    VsFlex.Height = Me.Height - 1500
End Sub

