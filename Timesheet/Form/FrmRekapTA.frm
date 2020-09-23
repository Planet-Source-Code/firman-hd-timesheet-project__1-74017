VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Begin VB.Form FrmRekapTA 
   Caption         =   "Rekap CV Tenaga Ahli"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   9465
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   9435
      TabIndex        =   2
      Top             =   0
      Width           =   9465
      Begin VB.ComboBox CboPendidikan 
         Height          =   315
         Left            =   5760
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.ComboBox CmbStatus 
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
         ItemData        =   "FrmRekapTA.frx":0000
         Left            =   960
         List            =   "FrmRekapTA.frx":000D
         TabIndex        =   10
         Text            =   "Karyawan Aktif"
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Export To Sheet"
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
         Left            =   9720
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
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
         Left            =   7920
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Print out"
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
         Left            =   9720
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VSFlex8Ctl.VSFlexGrid CboFlex 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   120
         Width           =   3225
         _cx             =   5689
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
         FormatString    =   $"FrmRekapTA.frx":003D
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
      Begin VSFlex8Ctl.VSFlexGrid CboAhli 
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   480
         Width           =   3225
         _cx             =   5689
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
         FormatString    =   $"FrmRekapTA.frx":0066
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pendidikan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   14
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Keahlian"
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
         Width           =   855
      End
      Begin VB.Label Label6 
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
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9465
      TabIndex        =   0
      Top             =   5055
      Width           =   9465
      Begin VB.CommandButton CmdClose 
         Caption         =   "Clos&e"
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
         Left            =   840
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6720
      OleObjectBlob   =   "FrmRekapTA.frx":008F
      Top             =   1080
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3375
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   6615
      _cx             =   11668
      _cy             =   5953
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
      SelectionMode   =   3
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
      AutoSearch      =   1
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
Attribute VB_Name = "FrmRekapTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub Command1_Click()

Showdata

End Sub
Sub Setgrid()
With fg
    .Rows = 1
    .RowHeight(0) = 800
    .Cols = 8
    .WordWrap = True
    .TextMatrix(0, 0) = "NO"
    .TextMatrix(0, 1) = "NAMA"
    .TextMatrix(0, 2) = "TGL/BLN LAHIR"
    .TextMatrix(0, 3) = "PENDIDIKAN"
    .TextMatrix(0, 4) = "JABATAN"
    .TextMatrix(0, 5) = "PENG. KERJA"
    .TextMatrix(0, 6) = "PROFESI/KEAHLIAN"
    .TextMatrix(0, 7) = "SERTIFIKAT/IJAZAH"
    .ColWidth(0) = 500
    .ColWidth(1) = 3000
    .ColWidth(2) = 2000
    .ColWidth(3) = 1500
    .ColWidth(4) = 0
    .ColWidth(5) = 1800
    .ColWidth(6) = 1800
    .ColWidth(7) = 2200

    For lCol = 1 To .Cols - 1
        .FixedAlignment(lCol) = flexAlignCenterCenter
    Next
End With
End Sub

Public Sub Showdata()
Dim Lrow, x As Long
Dim RsPendidikan As New ADODB.Recordset
Dim RsSka As New ADODB.Recordset
Dim KdPendidikan As Integer
fg.Rows = 1
On Error GoTo AdaError
With fg
    StrSQL = "Select Nip,Nama,Tgl_lahir,Keahlian,Sekolah,tmp_lahir,Sekolah From vKesehatan Where kdstatus <> 14 And Len(NIP) < 5"
    If Trim(cboFlex) <> "" Then
        StrSQL = StrSQL & " AND Divisi = '" & cboFlex & "'"
    ElseIf Trim(CboAhli) <> "" Then
        StrSQL = StrSQL & " AND Keahlian = '" & CboAhli & "'"
    End If
    
   
    If CmbStatus = "Tetap Aktif" Then
        StrSQL = StrSQL & " AND kdstatus = 1 Or kdStatus = 11"
    ElseIf CmbStatus = "Kontrak Aktif" Then
        StrSQL = StrSQL & " And kdstatus = 2 Or kdStatus = 12"
    End If
    If CboPendidikan <> "" Then
        Select Case CboPendidikan
            Case "-"
                KdPendidikan = 0
            Case "SD"
                  KdPendidikan = 1
            Case "SLTP"
                  KdPendidikan = 2
            Case "SLTA"
                 KdPendidikan = 3
            Case "STM"
                  KdPendidikan = 4
            Case "S1"
                  KdPendidikan = 5
            Case "S2"
                 KdPendidikan = 6
            Case "S3"
                 KdPendidikan = 7
            Case "D3"
                 KdPendidikan = 8
            Case "D1"
                 KdPendidikan = 9
        End Select
        StrSQL = StrSQL & " And Sekolah = '" & KdPendidikan & "'"
    End If
        StrSQL = StrSQL & " Order By Nama"
    If Rscek.State = adStateOpen Then Rscek.Close
    Rscek.Open StrSQL, CN, adOpenStatic
    Do Until Rscek.EOF
        .Rows = .Rows + 1
'        .TextMatrix(.Rows - 1, 0) = .Rows - 1
        .TextMatrix(.Rows - 1, 1) = Rscek!Nama
        .TextMatrix(.Rows - 1, 2) = Trim(Rscek!Tmp_lahir) & ", " & Format(Rscek!Tgl_Lahir, "dd MMM yyyy")
        .TextMatrix(.Rows - 1, 6) = IIf(IsNull(Rscek!keahlian), "", Rscek!keahlian)
        .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rscek!Sekolah), "", Rscek!Sekolah)
        If RsPendidikan.State = adStateOpen Then RsPendidikan.Close
            StrSQL = "Select * From Kar_Pendidikan Where NIP = '" & Rscek!NIP & "' And kd_sekolah = '" & Rscek!Sekolah & "' order by kd_Sekolah"
            RsPendidikan.Open StrSQL, CN, adOpenStatic
           If Not RsPendidikan.EOF Then
                .TextMatrix(.Rows - 1, 5) = RsPendidikan!Tgl_lulus
                .TextMatrix(.Rows - 1, 7) = RsPendidikan!Nama_sekolah & ", " & RsPendidikan!Jurusan & ", " & Format(RsPendidikan!Tgl_lulus, "dd MMM yyyy") & "," & RsPendidikan!Noijazah
            End If
        If .TextMatrix(.Rows - 1, 3) = 6 Or .TextMatrix(.Rows - 1, 3) = 7 Then
            If RsPendidikan.State = adStateOpen Then RsPendidikan.Close
            StrSQL = "Select * From Kar_Pendidikan Where NIP = '" & Rscek!NIP & "' And kd_sekolah >=5 order by kd_Sekolah"
            RsPendidikan.Open StrSQL, CN, adOpenStatic
            Do Until RsPendidikan.EOF
                .TextMatrix(.Rows - 1, 3) = RsPendidikan!kd_Sekolah
                .TextMatrix(.Rows - 1, 5) = RsPendidikan!Tgl_lulus
                .TextMatrix(.Rows - 1, 7) = RsPendidikan!Nama_sekolah & ", " & RsPendidikan!Jurusan & ", " & Format(RsPendidikan!Tgl_lulus, "dd MMM yyyy") & "," & RsPendidikan!Noijazah
                Select Case .TextMatrix(.Rows - 1, 3)
                    Case 0
                        .TextMatrix(.Rows - 1, 3) = "-"
                    Case 1
                         .TextMatrix(.Rows - 1, 3) = "SD"
                    Case 2
                         .TextMatrix(.Rows - 1, 3) = "SLTP"
                    Case 3
                         .TextMatrix(.Rows - 1, 3) = "SLTA"
                    Case 4
                         .TextMatrix(.Rows - 1, 3) = "STM"
                    Case 5
                         .TextMatrix(.Rows - 1, 3) = "S1"
                    Case 6
                         .TextMatrix(.Rows - 1, 3) = "S2"
                    Case 7
                         .TextMatrix(.Rows - 1, 3) = "S3"
                    Case 8
                         .TextMatrix(.Rows - 1, 3) = "D3"
                     Case 9
                         .TextMatrix(.Rows - 1, 3) = "D1"
                End Select
                
                RsPendidikan.MoveNext
                .Rows = .Rows + 1
            Loop
        End If
        If Trim(.TextMatrix(.Rows - 1, 3)) <> "" Then
             Select Case .TextMatrix(.Rows - 1, 3)
                Case 0
                    .TextMatrix(.Rows - 1, 3) = "-"
                Case 1
                     .TextMatrix(.Rows - 1, 3) = "SD"
                Case 2
                     .TextMatrix(.Rows - 1, 3) = "SLTP"
                Case 3
                     .TextMatrix(.Rows - 1, 3) = "SLTA"
                Case 4
                     .TextMatrix(.Rows - 1, 3) = "STM"
                Case 5
                     .TextMatrix(.Rows - 1, 3) = "S1"
                Case 6
                     .TextMatrix(.Rows - 1, 3) = "S2"
                Case 7
                     .TextMatrix(.Rows - 1, 3) = "S3"
                Case 8
                     .TextMatrix(.Rows - 1, 3) = "D3"
                 Case 9
                     .TextMatrix(.Rows - 1, 3) = "D1"
             
            End Select
        End If
              If RsSka.State = adStateOpen Then RsSka.Close
              StrSQL = "Select * from Kar_SKA Where NIP = '" & Rscek!NIP & "'"
              RsSka.Open StrSQL, CN, adOpenStatic
              Do Until RsSka.EOF
                If Trim(.TextMatrix(.Rows - 1, 7)) <> "" Then .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 7) = "* " & RsSka!Bidang_Keahlian & ", " & Format(RsSka!Tanggal, "dd MMM yyyy") & ", " & RsSka!No_SKA
               
                RsSka.MoveNext
              Loop
        Rscek.MoveNext
    Loop
        x = 0
        For Lrow = 1 To .Rows - 1
            If .TextMatrix(Lrow, 5) <> "" Then
              
                Dim D1, M1, Y1 As Integer
                Dim D2, M2, Y2 As Integer
                Dim d, m, y As Integer
                

        
                D1 = Day(.TextMatrix(Lrow, 5))
                M1 = Month(.TextMatrix(Lrow, 5))
                Y1 = Year(.TextMatrix(Lrow, 5))
                D2 = Day(Date)
                M2 = Month(Date)
                Y2 = Year(Date)
               
               
                If D2 < D1 Then M2 = M2 - 1: D2 = D2 + 30
                If M2 < M1 Then Y2 = Y2 - 1: M2 = 12 + M2
                d = Abs(D2 - D1)
                m = Abs(M2 - M1)
                y = Abs(Y2 - Y1)
          
                .TextMatrix(Lrow, 5) = y
              
         End If
                     .TextMatrix(Lrow, 0) = ""
                        If Trim(.TextMatrix(Lrow, 1)) <> "" Then
                            x = x + 1
                            .TextMatrix(Lrow, 0) = x
                        End If
                            
             
                    
        Next
End With

If fg.Rows = 1 Then
    MsgBox "Data Tidak Ditemukan", vbInformation
End If
Command1.Caption = "Refresh"
Command1.Enabled = True
Exit Sub
AdaError:
MsgBox err.Description
End Sub

Private Sub Command2_Click()
 On Error GoTo AdaError
'   fg.SaveGrid "C:\RekapTA.csv", flexFileCommaText, True
'    Shell PathOffice & "C:\RekapTA.csv", vbNormalFocus
'

With fg
If .Rows > 1 Then
     
     
        .AddItem "", 1
        .AddItem "", 2
        
        .Redraw = flexRDNone
        .Redraw = flexRDBuffered
        .TextMatrix(1, 1) = "Rekap Tenaga Ahli "
    For lCol = 1 To .Cols - 1
        
       .TextMatrix(2, lCol) = .TextMatrix(0, lCol)
        .Row = 2
        .Col = lCol
        .CellBackColor = vbGreen '&HE0E0E0
    Next
        .SaveGrid "C:\RekapTA.xls", flexFileExcel, False
        Call ShellExecute(Me.hwnd, "open", "C:\RekapTA.xls", vbNullString, "C:\RekapAbsen.xls", 1)
          .RemoveItem (1)
          .RemoveItem (1)
End If
End With
Exit Sub
AdaError:
MsgBox err.Description
End Sub

Private Sub Command3_Click()
On Error GoTo AdaError
If fg.Rows > 1 Then
 
        fg.PrintGrid "Rekap Tenaga Ahli " & DTPicker1.Value & " S/D " & DTPicker2.Value, , 2, 900, 500
 
End If
Exit Sub
AdaError:
MsgBox err.Description
End Sub

Private Sub Form_Load()
 If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path & "\Skins\" & skinsFileName
      Skin1.ApplySkin hwnd
     End If
AddDivisi
Setgrid
With CboPendidikan
        .AddItem ""
        .AddItem "SD"
        .AddItem "SLTP"
        .AddItem "SLTA"
        .AddItem "STM"
        .AddItem "D1"
        .AddItem "D3"
        .AddItem "S1"
        .AddItem "S2"
        .AddItem "S3"
End With

 End Sub
Private Sub AddDivisi()
    Dim Cboid     As String
    Dim cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    cboid1 = vbNullString
    StrSQL = "select Kd_divisi,nm_divisi from tb_divisi order by nm_Divisi"
    Rscek.Open StrSQL, CN, adOpenStatic
    cboid1 = " "
     
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("nm_divisi")
      cboid1 = cboid1 + Cboid
      Rscek.MoveNext
    Loop
     cboFlex.ColComboList(0) = cboid1
     cboFlex.CellAlignment = flexAlignLeftCenter
     Set Rscek = Nothing
     cboid1 = " "
     Cboid = vbNullString
 If Rscek.State = adStateOpen Then Rscek.Close
    
    StrSQL = "select * from TblKeahlian Order By Keahlian"
    Rscek.Open StrSQL, CN, adOpenStatic
 
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("keahlian")
      cboid1 = cboid1 + Cboid
      Rscek.MoveNext
    Loop
     CboAhli.ColComboList(0) = cboid1
     CboAhli.CellAlignment = flexAlignLeftCenter
End Sub
Private Sub Form_Resize()
 
CmdClose.Width = Picture3.Width
With fg
    .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture3.Height
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmRekapTA = Nothing
End Sub
 

