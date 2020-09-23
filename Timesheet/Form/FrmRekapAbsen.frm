VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmRekapAbsen 
   Caption         =   "Rekap Absen"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   10275
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   10245
      TabIndex        =   2
      Top             =   0
      Width           =   10275
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
         Left            =   6600
         TabIndex        =   5
         Top             =   600
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
         Left            =   5040
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
         Left            =   6600
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog dlg 
         Left            =   3000
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   960
         TabIndex        =   6
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
         Format          =   54067203
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3000
         TabIndex        =   7
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
         Format          =   54067203
         CurrentDate     =   39931
      End
      Begin VSFlex8Ctl.VSFlexGrid cboFlex 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   1080
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
         FormatString    =   $"FrmRekapAbsen.frx":0000
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
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   3120
         TabIndex        =   13
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   54067202
         CurrentDate     =   39940.34375
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   54067202
         CurrentDate     =   39940.3333333333
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   375
         Left            =   5040
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   54067202
         CurrentDate     =   39940.34375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "<="
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
         Left            =   2760
         TabIndex        =   16
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Jam         >="
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
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Divisi"
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
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label2 
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
         Left            =   2640
         TabIndex        =   9
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dari  :"
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
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   6405
      Width           =   10275
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
      Left            =   1320
      OleObjectBlob   =   "FrmRekapAbsen.frx":0029
      Top             =   6000
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlex 
      Height          =   4695
      Left            =   0
      TabIndex        =   18
      Top             =   1440
      Width           =   6735
      _cx             =   11880
      _cy             =   8281
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
      BackColorSel    =   16777215
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
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4695
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   6615
      _cx             =   11668
      _cy             =   8281
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
Attribute VB_Name = "FrmRekapAbsen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim Lrow, x As Long
Dim StatusTgl As Boolean
Dim RowFlex As Long
fg.Rows = 1
If Rscek.State = adStateOpen Then Rscek.Close
StrSQL = "SELECT Absensi.NIP, Absensi.Tgl, Absensi.masuk, Divisi.NM_DIV,Karyawan.Kd_Divisi FROM Absensi INNER JOIN Karyawan ON Absensi.NIP = Karyawan.NIP INNER JOIN Divisi ON Karyawan.Kd_Divisi = Divisi.KD_DIV where tgl between  '" & Format(DTPicker1, "yyyy/MM/dd") & "' And  '" & Format(DTPicker2, "yyyy/MM/dd") & "' And Masuk <>'01/01/2000'"
If Trim(cboFlex) <> "" Then StrSQL = StrSQL & " AND Karyawan.kd_Divisi = '" & cboFlex.Text & "'  "
StrSQL = StrSQL & " Order By Masuk"
Rscek.Open StrSQL, CN, adOpenStatic
Set fg.DataSource = Rscek
fg.ColFormat(3) = "HH:mm"
fg.ColFormat(1) = "dd/MMM/yyyy"
Command1.Enabled = False
With VsFlex
.Rows = 1
If Trim(cboFlex.Text) = "" Then
    .Cols = 3
    .TextMatrix(0, 0) = "NO"
    .TextMatrix(0, 1) = "Tanggal"
    .TextMatrix(0, 2) = "Jumlah"
    .ColWidth(0) = 500
    .ColWidth(1) = 1500
    .ColWidth(2) = 1000
    .ColFormat(1) = "dd/MMM/yyyy"
    .ColFormat(2) = "#,###"
Else
    .Cols = 4
    .TextMatrix(0, 0) = "NO"
    .TextMatrix(0, 1) = "Divisi"
    .TextMatrix(0, 2) = "Tanggal"
    .TextMatrix(0, 3) = "Jumlah"
    .ColWidth(0) = 500
    .ColWidth(1) = 2500
    .ColWidth(2) = 1500
    .ColWidth(3) = 1000
     .ColFormat(2) = "dd/MMM/yyyy"
End If
For Lrow = 1 To fg.Rows - 1
fg.TextMatrix(Lrow, 0) = Lrow
    DTPicker5 = Format(fg.TextMatrix(Lrow, 3), "HH:mm")

    If Format(DTPicker5.Value, "HH:mm") >= Format(DTPicker3.Value, "HH:mm") And Format(DTPicker5.Value, "HH:mm") <= Format(DTPicker4.Value, "HH:mm") Then
       StatusTgl = False
        If Trim(cboFlex) = "" Then
            
             If .Rows > 1 Then
                For x = 1 To .Rows - 1
                     If .TextMatrix(x, 1) = fg.TextMatrix(Lrow, 2) Then StatusTgl = True: Exit For
                Next
                 
                  If StatusTgl = True Then
'                       .TextMatrix(X, 0) = Lrow
                      .TextMatrix(x, 1) = fg.TextMatrix(Lrow, 2)
                      .TextMatrix(x, 2) = .TextMatrix(x, 2) + 1
                  Else
                      .Rows = .Rows + 1
'                     .TextMatrix(X, 0) = Lrow
                      .TextMatrix(.Rows - 1, 1) = fg.TextMatrix(Lrow, 2)
                      .TextMatrix(.Rows - 1, 2) = 1
                  End If
                  
             Else
               
               .Rows = .Rows + 1
'                .TextMatrix(X, 0) = Lrow
               .TextMatrix(.Rows - 1, 1) = fg.TextMatrix(Lrow, 2)
               .TextMatrix(.Rows - 1, 2) = 1
            End If
        Else
            
            If .Rows > 1 Then
                For x = 1 To .Rows - 1
                     If .TextMatrix(x, 2) = fg.TextMatrix(Lrow, 2) Then StatusTgl = True: Exit For
                Next
                 
                  If StatusTgl = True Then
'                       .TextMatrix(X, 0) = Lrow
                      .TextMatrix(.Rows - 1, 1) = fg.TextMatrix(Lrow, 4)
                      .TextMatrix(x, 2) = fg.TextMatrix(Lrow, 2)
                      .TextMatrix(x, 3) = .TextMatrix(x, 3) + 1
                  Else
                      .Rows = .Rows + 1
'                     .TextMatrix(X, 0) = Lrow
                      .TextMatrix(.Rows - 1, 1) = fg.TextMatrix(Lrow, 4)
                      .TextMatrix(.Rows - 1, 2) = fg.TextMatrix(Lrow, 2)
                      .TextMatrix(.Rows - 1, 3) = 1
                  End If
                  
             Else
                
               .Rows = .Rows + 1
'                .TextMatrix(.Rows - 1, 0) = Lrow
               .TextMatrix(.Rows - 1, 1) = fg.TextMatrix(Lrow, 4)
               .TextMatrix(.Rows - 1, 2) = fg.TextMatrix(Lrow, 2)
               .TextMatrix(.Rows - 1, 3) = 1
            End If
        End If
    End If
Next
For Lrow = 1 To .Rows - 1
    .TextMatrix(Lrow, 0) = Lrow
Next
Command1.Enabled = True
End With
End Sub

Private Sub Command2_Click()
Dim x As String
 VsFlex.SaveGrid "C:\RekapAbsen.xls", flexFileExcel, True
' Shell PathOffice & "C:\RekapAbsen.csv", vbNormalFocus
x = ShellExecute(Me.hwnd, "open", "C:\RekapAbsen.xls", vbNullString, "C:\RekapAbsen.xls", 1)
End Sub

Private Sub Command3_Click()
On Error GoTo Adaerror
If VsFlex.Rows > 1 Then VsFlex.PrintGrid "Rekap Absen " & DTPicker1.Value & " S/D " & DTPicker2.Value, 2, 2, 900, 500
Exit Sub
Adaerror:
MsgBox err.Description
End Sub

Private Sub Form_Load()
 If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path & "\Skins\" & skinsFileName
      Skin1.ApplySkin hwnd
     End If
AddDivisi
DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker1.Value = DateSerial(Year(Now), Month(Now), 1)
DTPicker1.Value = DateAdd("M", 0, DTPicker1.Value)
DTPicker2.Value = DateSerial(Year(Now), Month(Now), 1)
DTPicker2.Value = DateAdd("M", 1, DTPicker2.Value) - 1
DTPicker1.CustomFormat = "dd/MMM/yyyy"
DTPicker2.CustomFormat = "dd/MMM/yyyy"
End Sub
Private Sub AddDivisi()

    Dim Cboid     As String
    Dim cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    cboid1 = vbNullString
    StrSQL = "select * from Divisi Where kd_bid >= 2 and kd_bid <= 20 order by kd_bid"
    Rscek.Open StrSQL, CN, adOpenStatic
    cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("kd_Div") & vbTab & Rscek("NM_DIV")
      cboid1 = cboid1 + Cboid
      Rscek.MoveNext
    Loop
    cboFlex.ColComboList(0) = cboid1
End Sub
Private Sub Form_Resize()
 
CmdClose.Width = Picture3.Width
With VsFlex
    .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture3.Height
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmRekapAbsen = Nothing
End Sub
