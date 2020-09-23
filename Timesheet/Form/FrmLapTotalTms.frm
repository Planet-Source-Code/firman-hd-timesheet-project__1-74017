VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmLapTotalTms 
   Caption         =   "Rekap Total Timesheet"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9450
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9450
      TabIndex        =   13
      Top             =   6300
      Width           =   9450
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
         Left            =   0
         TabIndex        =   14
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   9420
      TabIndex        =   0
      Top             =   0
      Width           =   9450
      Begin VB.CommandButton Command3 
         Caption         =   "&Print out"
         Height          =   375
         Left            =   9960
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   8400
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Export To Sheet"
         Height          =   375
         Left            =   9960
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4440
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   6840
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   53870594
         CurrentDate     =   39940
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   720
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
         Format          =   53870595
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2880
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
         Format          =   53870595
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   5520
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   53870594
         CurrentDate     =   39940
      End
      Begin VSFlex8Ctl.VSFlexGrid CboDivisi 
         Height          =   315
         Left            =   720
         TabIndex        =   9
         Top             =   600
         Width           =   1425
         _cx             =   2514
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
         FormatString    =   $"FrmLapTotalTms.frx":0000
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
      Begin ACTIVESKINLibCtl.Skin Skin2 
         Left            =   0
         OleObjectBlob   =   "FrmLapTotalTms.frx":0029
         Top             =   0
      End
      Begin MSComDlg.CommonDialog dlg 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label LblTotaljam 
         BackStyle       =   0  'Transparent
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
         Left            =   4560
         TabIndex        =   19
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   12
         Top             =   120
         Width           =   855
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
         Left            =   2400
         TabIndex        =   11
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label4 
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
         TabIndex        =   10
         Top             =   600
         Width           =   495
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "FrmLapTotalTms.frx":025D
      Top             =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3735
      Left            =   0
      TabIndex        =   15
      ToolTipText     =   "Double Klik Kolom Nama  Untuk Melihat Detail Timesheet"
      Top             =   1320
      Width           =   8295
      _cx             =   14631
      _cy             =   6588
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
      BackColorSel    =   14737632
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
      FormatString    =   $"FrmLapTotalTms.frx":0491
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
      Editable        =   0
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
   Begin VSFlex8Ctl.VSFlexGrid VSFg 
      Height          =   3255
      Left            =   0
      TabIndex        =   16
      Top             =   5400
      Width           =   9975
      _cx             =   17595
      _cy             =   5741
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmLapTotalTms.frx":0570
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
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   3255
      Left            =   10080
      TabIndex        =   17
      Top             =   5400
      Width           =   6495
      _cx             =   11456
      _cy             =   5741
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmLapTotalTms.frx":064F
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
   Begin VSFlex8Ctl.VSFlexGrid VSJoin 
      Height          =   3975
      Left            =   8640
      TabIndex        =   18
      Top             =   1320
      Width           =   6735
      _cx             =   11880
      _cy             =   7011
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmLapTotalTms.frx":072E
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
Attribute VB_Name = "FrmLapTotalTms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalJam, Selisih As Double
Private Sub CmdClose_Click()
    Unload Me
End Sub
Private Sub Command1_Click()
Command1.Enabled = False
Setgrid
Showdata
Command1.Enabled = True
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
    CboDivisi.ColComboList(0) = cboid1
End Sub
Private Sub Command2_Click()
Dim x As String
With fg
    If fg.Rows > 1 Then
          .AddItem "", 1
            .AddItem "", 2
            
            .Redraw = flexRDNone
            .Redraw = flexRDBuffered
            .TextMatrix(1, 1) = "Rekap Total Jam Timesheet Periode " & Format(DTPicker1, "dd MMM yyyy") & " s/d " & Format(DTPicker2, "dd MMM yyyy")
        For lCol = 1 To .Cols - 1
            
           .TextMatrix(2, lCol) = .TextMatrix(0, lCol)
            .Row = 2
            .Col = lCol
            .CellBackColor = vbGreen
        Next
            .SaveGrid "C:\TotalJamTimesheet.xls", flexFileExcel, False
            x = ShellExecute(Me.hwnd, "open", "C:\TotalJamTimesheet.xls", vbNullString, "C:\TotalJamTimesheet.xls", 1)
              .RemoveItem (1)
              .RemoveItem (1)
    End If
End With
End Sub

Private Sub Command3_Click()
On Error Resume Next
If fg.Rows > 1 Then fg.PrintGrid "Total Jam - Periode " & DTPicker1.Value & " S/D " & DTPicker2.Value, 2, 2, 900, 500

End Sub
   
Private Sub fg_DblClick()
With fg
    If .Row > 0 Then
        frmRekapTimesheet.cboFlex = .TextMatrix(.Row, 1)
        frmRekapTimesheet.DTPicker1 = DTPicker1
        frmRekapTimesheet.DTPicker2 = DTPicker2
        
        frmRekapTimesheet.Showdata
        LoadForm frmRekapTimesheet

    End If
End With

End Sub

Private Sub Form_Load()
    AddDivisi
    Setgrid
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker1.Value = DateSerial(Year(Now), Month(Now), 26)
    DTPicker1.Value = DateAdd("M", -1, DTPicker1.Value)
    DTPicker2.Value = DateSerial(Year(Now), Month(Now), 25)
    DTPicker2.Value = DateAdd("M", 0, DTPicker2.Value)
    DTPicker1.CustomFormat = "dd/MMM/yyyy"
    DTPicker2.CustomFormat = "dd/MMM/yyyy"
     If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
      Skin1.ApplySkin hwnd
    End If
    Select Case UCase(strGroup)
        Case "IT", "PTW"
            CboDivisi.Enabled = True
        Case "ADMIN"
            CboDivisi.Text = NamaDivisi
            CboDivisi.Enabled = False
    End Select
End Sub
  
Sub Setgrid()
Dim x As Integer
 
With fg
    .Cols = 8
    .Rows = 2
    .FixedRows = 2
    
    .TextMatrix(0, 0) = "No"
    .TextMatrix(0, 1) = "NIP"
    .TextMatrix(0, 2) = "Nama"
    .TextMatrix(0, 3) = "Total Jam"
    .TextMatrix(0, 4) = "Total Jam"
    .TextMatrix(0, 5) = "Total Jam"
    .TextMatrix(0, 6) = "Total Jam"
    .TextMatrix(0, 7) = "Total Jam"
'    .TextMatrix(0, 8) = "Total Jam" '"Jam Lembur"
'    .TextMatrix(0, 9) = "Total Jam" '"Biaya Lembur"
    .TextMatrix(1, 0) = "No"
    .TextMatrix(1, 1) = "NIP"
    .TextMatrix(1, 2) = "Nama"
    .TextMatrix(1, 3) = "Kerja Standar"
    .TextMatrix(1, 4) = "Kerja + Lembur"
    .TextMatrix(1, 5) = "Persen"
    .TextMatrix(1, 6) = "Belum Diverifikasi"
    .TextMatrix(1, 7) = "Sudah Diverifikasi"
'    .TextMatrix(1, 8) = "Kerja + Lembur" '"Jam Lembur"
'    .TextMatrix(1, 9) = "Persen" '"Biaya Lembur"
    .RowHeight(1) = 700
    .ColWidth(0) = 500
    .ColWidth(1) = 700
    .ColWidth(2) = 3000
    .ColWidth(3) = 850
    
    .MergeCells = flexMergeFree
    .MergeRow(0) = True
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    
    For x = 0 To .Cols - 1
       .FixedAlignment(x) = flexAlignCenterCenter
    Next
End With
With VSJoin
.Rows = 1
.Cols = 19
.TextMatrix(.Rows - 1, 1) = "1. Tanggal"
.TextMatrix(.Rows - 1, 2) = "2.Project"
.TextMatrix(.Rows - 1, 3) = "3. NIP"
.TextMatrix(.Rows - 1, 4) = "4. Nama"
.TextMatrix(.Rows - 1, 5) = "5. Total / 2"
.TextMatrix(.Rows - 1, 6) = "6. masuk"
.TextMatrix(.Rows - 1, 7) = "7. keluar"
.TextMatrix(.Rows - 1, 8) = "8. NamaProject"
.TextMatrix(.Rows - 1, 9) = "9. StatusDivisi"
.TextMatrix(.Rows - 1, 10) = "10. TotalLembur"
.TextMatrix(.Rows - 1, 11) = "11. StatusPM"
.TextMatrix(.Rows - 1, 12) = "12. Keterangan"
.TextMatrix(.Rows - 1, 13) = "13. Hari"
.TextMatrix(.Rows - 1, 14) = "14. Divisi"
.TextMatrix(.Rows - 1, 15) = "15. gaji"
.TextMatrix(.Rows - 1, 16) = "16. Lembur"
.TextMatrix(.Rows - 1, 17) = "17. ProjetUmum"
.TextMatrix(.Rows - 1, 18) = "18. TotalKerja"
.ColDataType(6) = flexDTDate
.ColDataType(7) = flexDTDate
.ColFormat(6) = "HH:mm"
.ColFormat(7) = "HH:mm"
'.ColFormat(10) = "#.##"
.ColFormat(15) = "#,###"
.ColFormat(16) = "#,###"

End With
End Sub
Private Sub AddKaryawan(Divisi As String)
    Dim Cboid     As String
    Dim cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    cboid1 = vbNullString
    StrSQL = "select * from Karyawan Where kd_divisi = '" & Divisi & "' And Status <> '14' Order By Nama"
    Rscek.Open StrSQL, CN, adOpenStatic
    cboid1 = " "
    Combo1.Clear
    Do Until Rscek.EOF
       
      Combo1.AddItem Rscek!NIP
      Rscek.MoveNext
    Loop
    
End Sub
Private Sub Form_Resize()
    On Error Resume Next
        CmdClose.Width = Me.Width - 100
        With fg
             .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture2.Height

        End With
'
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set FrmLapTotalTms = Nothing
End Sub
Sub Showdata()
Dim x, TotalHari As Integer
Dim i, J As Integer
Dim Split, JmlLoop As Integer
Dim Jam, Hari As String
Dim IDWaktu As String
Dim Jam1, Jam2 As Date
Dim Jam3, Jam4 As Date
Dim TglAwal, TglAkhir As Date
Dim RsTS As New ADODB.Recordset
Dim searchDivisi As String


    Select Case UCase(strGroup)
        Case "IT", "PTW"
            searchDivisi = CboDivisi
           
        Case "ADMIN"
            searchDivisi = KodeDivisi
    End Select
  
  If Trim(CboDivisi) = "" Then
    MsgBox "Nama Divisi Belum Diisi", vbCritical
    Exit Sub
  End If
  AddKaryawan (searchDivisi)
'   TglAwal = DTPicker1
'    TglAwal = DateSerial(Year(DTPicker1), Month(DTPicker1), 1)
'    TglAwal = DateAdd("M", 0, DTPicker1)
'    TglAwal = Format(DTPicker1, "MM") & "/01/" & Format(DTPicker1, "yyyy")
'    TglAkhir = DTPicker1
'    TglAkhir = DateSerial(Year(DTPicker1), Month(DTPicker1), 1)
'    TglAkhir = DateAdd("M", 1, TglAkhir) - 1
'    Selisih = DateDiff("d", TglAwal, TglAkhir)
    TglAwal = DTPicker1
    TglAkhir = DTPicker2
    i = DateDiff("d", TglAwal, TglAkhir)
    TotalHari = 0

    For x = 0 To i
       If x > 0 Then TglAwal = DateAdd("d", 1, TglAwal)
       Hari = Format(TglAwal, "ddd")
       If Hari <> "Sat" And Hari <> "Sun" And Hari <> "Sabtu" And Hari <> "Minggu" Then
           
            StrSQL = "select tanggallibur from kalender " & _
                "where tanggallibur = '" & Format(TglAwal, "MM/dd/yyyy") & "'"
            If Rscek.State = adStateOpen Then Rscek.Close
            Rscek.Open StrSQL, CN, adOpenStatic
            If Rscek.EOF Then
               TotalHari = TotalHari + 1
            End If
        End If
    Next
     
    TotalJam = TotalHari * 8
    LblTotaljam = TotalJam & " Jam"
    TglAwal = DTPicker1
     For i = 1 To 30
        If x > 0 Then TglAwal = DateAdd("d", 1, TglAwal)
         Hari = Format(TglAwal, "ddd")
       If Hari <> "Sat" And Hari <> "Sun" And Hari <> "Sabtu" And Hari <> "Minggu" Then
            StrSQL = "select tanggallibur from kalender " & _
                "where tanggallibur = '" & Format(TglAwal, "MM/dd/yyyy") & "'"
            If Rscek.State = adStateOpen Then Rscek.Close
            Rscek.Open StrSQL, CN, adOpenStatic
            If Rscek.EOF Then
               If CDate(TglAwal) >= "08/11/2010" And CDate(TglAwal) <= "09/09/2010" Then
                TotalJam = TotalJam - 0.5
               End If
            End If
               
       End If
     Next
Command1.Enabled = False
With VSFlexGrid1
        .Rows = 1
        VSFg.Rows = 1
        VSFg.Cols = 17
    
    For x = 0 To Combo1.ListCount - 1
            Combo1.ListIndex = x
            Command1.Caption = Combo1.Text
            If RsTS.State = adStateOpen Then RsTS.Close
            StrSQL = "SELECT tbltimesheet.IDtimesheet,tbltimesheet.Tanggal,tbltimesheet.JamAwal As [Jam Awal],tbltimesheet.JamAkhir AS [Jam Akhir],tbltimesheet.Status,tbltimesheet.NoProject As Project,tbltimesheet.Keterangan,tbltimesheet.Tanggal,Absensi.Masuk,tbltimesheet.NIP,tbltimesheet.StatusDivisi, karyawan.Nama,tbltimesheet.StatusPM,Absensi.Keluar,Divisi.NM_DIV, karyawan.Nama,tbltimesheet.Hari,tbltimesheet.TotalKerja,tbltimesheet.ProjectUmum"
            StrSQL = StrSQL & " FROM tbltimesheet INNER JOIN karyawan ON tbltimesheet.NIP = karyawan.NIP INNER JOIN Divisi ON karyawan.Kd_Divisi = Divisi.KD_DIV INNER JOIN Absensi ON tbltimesheet.NIP = Absensi.NIP AND tbltimesheet.Tanggal = Absensi.Tgl"
'            StrSQL = "SELECT IDtimesheet,Tanggal,JamAwal As [Jam Awal],JamAkhir AS [Jam Akhir],Status,NoProject As Project,Keterangan,Tanggal,Masuk,StatusPM,StatusDivisi From tbltimesheet Where Tanggal Between '" & Format(TglAwal, "yyyy/MM/dd") & "' And '" & Format(TglAkhir, "yyyy/MM/dd") & "' And NIP = '" & StrNIPUser & "' Order By Tanggal DESC,IDTimesheet ASC"
            StrSQL = StrSQL & " Where tbltimesheet.Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And '" & Format(DTPicker2, "MM/dd/yyyy") & "'  And tbltimesheet.Status ='Actual'"
            StrSQL = StrSQL & " AND tbltimesheet.Nip = '" & Combo1.Text & "'"
            StrSQL = StrSQL & " And tbltimesheet.Kd_Divisi = '" & searchDivisi & "'"
            StrSQL = StrSQL & " Order By Divisi.NM_DIV,tbltimesheet.NoProject,tbltimesheet.Tanggal,tbltimesheet.NIP,tbltimesheet.IDTimesheet ASC"
            RsTS.Open StrSQL, CN, adOpenStatic
            Set .DataSource = RsTS
            .ColDataType(2) = flexDTDate
        
            For Lrow = 1 To .Rows - 1
              
                .TextMatrix(Lrow, 0) = Lrow
                .TextMatrix(0, 0) = .Rows
                If .TextMatrix(Lrow, 3) = "" Then .TextMatrix(Lrow, 3) = "00:00"
                 If .TextMatrix(Lrow, 4) = "" Then .TextMatrix(Lrow, 4) = "00:00"
        
                        JmlLoop = 0
                        Jam1 = CDate(.TextMatrix(Lrow, 3))
                        Jam2 = CDate(.TextMatrix(Lrow, 4))
                        DTPicker3.Value = Jam1
                        Do Until JmlLoop = 50
        
                             VSFg.Rows = VSFg.Rows + 1
                                DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                                VSFg.TextMatrix(VSFg.Rows - 1, 0) = VSFg.Rows - 1
                                IDWaktu = Format(DTPicker3.Value, "HH:mm")
                                VSFg.TextMatrix(VSFg.Rows - 1, 1) = IDWaktu
                                VSFg.TextMatrix(VSFg.Rows - 1, 2) = .TextMatrix(Lrow, 2) 'Format(.TextMatrix(Lrow, 2), "dd/MM/yyyy")
                                VSFg.TextMatrix(VSFg.Rows - 1, 3) = .TextMatrix(Lrow, 6)
                                VSFg.TextMatrix(VSFg.Rows - 1, 4) = .TextMatrix(Lrow, 10)
                                VSFg.TextMatrix(VSFg.Rows - 1, 5) = .TextMatrix(Lrow, 5)
                                VSFg.TextMatrix(VSFg.Rows - 1, 6) = Format(.TextMatrix(Lrow, 3), "HH:mm")
                                VSFg.TextMatrix(VSFg.Rows - 1, 7) = .TextMatrix(Lrow, 7)
                                VSFg.TextMatrix(VSFg.Rows - 1, 8) = .TextMatrix(Lrow, 11) 'StatusDivisi
                                VSFg.TextMatrix(VSFg.Rows - 1, 9) = .TextMatrix(Lrow, 12)
                                VSFg.TextMatrix(VSFg.Rows - 1, 10) = .TextMatrix(Lrow, 1)
                                VSFg.TextMatrix(VSFg.Rows - 1, 11) = .TextMatrix(Lrow, 4)
                                VSFg.TextMatrix(VSFg.Rows - 1, 12) = .TextMatrix(Lrow, 15)
                                VSFg.TextMatrix(VSFg.Rows - 1, 13) = .TextMatrix(Lrow, 13) 'StatusPM
                                VSFg.TextMatrix(VSFg.Rows - 1, 14) = .TextMatrix(Lrow, 17)
                                VSFg.TextMatrix(VSFg.Rows - 1, 15) = TotalJam '.TextMatrix(Lrow, 18)
                                VSFg.TextMatrix(VSFg.Rows - 1, 16) = .TextMatrix(Lrow, 19)
                                Jam3 = Format(.TextMatrix(Lrow, 9), "HH:mm")
        
                                Jam4 = DateAdd("h", 9, Jam3)
                                If Trim(.TextMatrix(Lrow, 17)) = "Kerja" Then
                                    If (VSFg.TextMatrix(VSFg.Rows - 1, 1) >= "08:00" And VSFg.TextMatrix(VSFg.Rows - 1, 1) <= "17:00") Or (VSFg.TextMatrix(VSFg.Rows - 1, 1) < Jam4 And VSFg.TextMatrix(VSFg.Rows - 1, 1) >= "08:00") Then
                                         VSFg.TextMatrix(VSFg.Rows - 1, 7) = "Timesheet"
                                      End If
        
                                End If
                                     If Format(DTPicker3.Value, "HH:mm") = Format(Jam2, "HH:mm") Then Exit Do
                                If Format(DTPicker3, "HH:mm") = "12:00" Then DTPicker3.Value = DateAdd("n", 60, DTPicker3)
                                JmlLoop = JmlLoop + 1
                        Loop
            Next
                VSFg.ColFormat(6) = "HH:mm"
                VSFg.ColFormat(11) = "HH:mm"
'            For lCol = 1 To VSFg.Cols - 1
'                VSFg.TextMatrix(0, lCol) = lCol
'            Next
        If .Rows = 1 Then
            If Rscek.State = adStateOpen Then Rscek.Close
            Rscek.Open "Select * From Vuser Where NIP = '" & Combo1 & "'", CN, adOpenStatic
            With fg
                If Not Rscek.EOF Then
                     .Rows = .Rows + 1
                     .TextMatrix(.Rows - 1, 0) = .Rows - 1
                     .TextMatrix(.Rows - 1, 1) = Rscek!NIP
                     .TextMatrix(.Rows - 1, 2) = Rscek!Nama
                     .TextMatrix(.Rows - 1, 3) = TotalJam
                     .TextMatrix(.Rows - 1, 4) = 0
                     .TextMatrix(.Rows - 1, 5) = 0
                     .TextMatrix(.Rows - 1, 6) = 0
                     .TextMatrix(.Rows - 1, 7) = 0
'                     .TextMatrix(.Rows - 1, 8) = 0
'                     .TextMatrix(.Rows - 1, 9) = 0
                End If
            End With
        Else
            With VSFg
                VSJoin.Rows = 1
                For Lrow = 1 To VSFg.Rows - 1
                         J = JoinGrid(.TextMatrix(Lrow, 3), .TextMatrix(Lrow, 2), Lrow, .TextMatrix(Lrow, 4), .TextMatrix(Lrow, 7))
                Next
            End With

            With VSJoin
                VSFg.Rows = 1
            For Lrow = 1 To VSJoin.Rows - 1
'                If Trim(.TextMatrix(Lrow, 5)) = "" Then .TextMatrix(Lrow, 5) = 0
'                .TextMatrix(Lrow, 0) = Lrow
'                .TextMatrix(0, 0) = .Rows
'
'                  If Trim(.TextMatrix(Lrow, 12)) = "Lembur" Then
'                   J = .TextMatrix(Lrow, 10)
'
'                  End If
            
                J = HitungGaji(.TextMatrix(Lrow, 3), .TextMatrix(Lrow, 9), .TextMatrix(Lrow, 14), Lrow, .TextMatrix(Lrow, 12))
            Next
              
            End With
        End If
        
    Next
End With

'Disiplit Perhari


With fg
 
    For Lrow = 2 To .Rows - 1
        .TextMatrix(Lrow, 0) = Lrow - 1
        If .TextMatrix(Lrow, 3) > 0 Then .TextMatrix(Lrow, 5) = Round(CDbl(.TextMatrix(Lrow, 4)) / CDbl(.TextMatrix(Lrow, 3)) * 100, 2)
        .TextMatrix(Lrow, 5) = .TextMatrix(Lrow, 5) & "%"
        If .TextMatrix(Lrow, 4) = 0 Then
            .Row = Lrow
            For x = 1 To .Cols - 1
                .Col = x
                .CellBackColor = vbRed
            Next
        End If
    Next
 
End With
 
Command1.Caption = "Refresh"
Command1.Enabled = True
End Sub

Function HitungGaji(NIP As String, StatusDivisi As String, Divisi As String, Row As Integer, Keterangan As String)
Dim StatusProject As Boolean
Dim FRow As Integer
Dim HRow As Integer
With fg
    StatusProject = False
    For FRow = 0 To .Rows - 1
    If Trim(.TextMatrix(FRow, 1)) = NIP Then
           StatusProject = True
           HRow = FRow
        End If
    Next
        If StatusProject = False Then
             .Rows = .Rows + 1
             HRow = .Rows - 1
             If .TextMatrix(HRow, 4) = "" Then .TextMatrix(HRow, 4) = 0
            If .TextMatrix(HRow, 5) = "" Then .TextMatrix(HRow, 5) = 0
            If .TextMatrix(HRow, 6) = "" Then .TextMatrix(HRow, 6) = 0
            If .TextMatrix(HRow, 7) = "" Then .TextMatrix(HRow, 7) = 0
 
'            If .TextMatrix(HRow, 9) = "" Then .TextMatrix(HRow, 9) = 0
 
        End If
            .TextMatrix(HRow, 1) = VSJoin.TextMatrix(Row, 3)
            .TextMatrix(HRow, 2) = VSJoin.TextMatrix(Row, 4)
            .TextMatrix(HRow, 3) = TotalJam
 
'            If Trim(Keterangan) = "Timesheet" Then
                If StatusDivisi = 0 Then
                   .TextMatrix(HRow, 6) = CCur(.TextMatrix(HRow, 6)) + CCur(VSJoin.TextMatrix(Row, 5)) + CDbl(VSJoin.TextMatrix(Row, 10))
                Else
                   .TextMatrix(HRow, 7) = CCur(.TextMatrix(HRow, 7)) + VSJoin.TextMatrix(Row, 5) + CDbl(VSJoin.TextMatrix(Row, 10))
                End If
                .TextMatrix(HRow, 4) = CCur(.TextMatrix(HRow, 4)) + CCur(VSJoin.TextMatrix(Row, 5)) + CDbl(VSJoin.TextMatrix(Row, 10))
'             Else
'                .TextMatrix(HRow, 7) = CDbl(.TextMatrix(HRow, 7)) + CDbl(VSJoin.TextMatrix(Row, 10))
'            End If
End With
 
End Function
Function JoinGrid(ByVal Project As String, ByVal tgl1 As String, ByVal Row As Integer, ByVal NIP As String, ByVal Keterangan As String)
Dim StatusNIP As Boolean
Dim FRow As Integer
Dim J As String
With VSJoin
    StatusNIP = False
    For FRow = 0 To .Rows - 1
        If .TextMatrix(FRow, 2) = Project And .TextMatrix(FRow, 1) = tgl1 And .TextMatrix(FRow, 3) = NIP And .TextMatrix(FRow, 12) = Keterangan Then
           StatusNIP = True
        End If
    Next
        If StatusNIP = False Then
             .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 5) = 0
'            .TextMatrix(.Rows - 1, 10) = 0
        End If
            .TextMatrix(.Rows - 1, 1) = VSFg.TextMatrix(Row, 2)
            .TextMatrix(.Rows - 1, 2) = VSFg.TextMatrix(Row, 3)
            .TextMatrix(.Rows - 1, 3) = VSFg.TextMatrix(Row, 4)
            .TextMatrix(.Rows - 1, 4) = VSFg.TextMatrix(Row, 9)
            If Trim(Keterangan) = "Timesheet" Then
               .TextMatrix(.Rows - 1, 5) = (.TextMatrix(.Rows - 1, 5) + 0.5)
            End If
            .TextMatrix(.Rows - 1, 6) = Format(VSFg.TextMatrix(Row, 6), "HH:mm")
            .TextMatrix(.Rows - 1, 7) = Format(VSFg.TextMatrix(Row, 11), "HH:mm")
            .TextMatrix(.Rows - 1, 8) = VSFg.TextMatrix(Row, 13)
            .TextMatrix(.Rows - 1, 9) = VSFg.TextMatrix(Row, 8)
            If .TextMatrix(.Rows - 1, 10) = "" Then .TextMatrix(.Rows - 1, 10) = 0
            If Trim(Keterangan) = "Lembur" Then
                .TextMatrix(.Rows - 1, 10) = .TextMatrix(.Rows - 1, 10) + 0.5
            End If
            .TextMatrix(.Rows - 1, 11) = VSFg.TextMatrix(Row, 13)
            .TextMatrix(.Rows - 1, 12) = VSFg.TextMatrix(Row, 7)
            .TextMatrix(.Rows - 1, 13) = VSFg.TextMatrix(Row, 14) 'Hari
            .TextMatrix(.Rows - 1, 14) = VSFg.TextMatrix(Row, 12)
            .TextMatrix(.Rows - 1, 16) = 0
            .TextMatrix(.Rows - 1, 17) = VSFg.TextMatrix(Row, 15)
            .TextMatrix(.Rows - 1, 18) = VSFg.TextMatrix(Row, 16)
End With
End Function
 


'-----BEDA TAMPILANNYA
'Option Explicit
'Dim TotalJam, Selisih As Double
'Private Sub CmdClose_Click()
'    Unload Me
'End Sub
'Private Sub Command1_Click()
'Command1.Enabled = False
'Setgrid
'Showdata
'Command1.Enabled = True
'End Sub
'Private Sub AddDivisi()
'
'    Dim Cboid     As String
'    Dim Cboid1    As String
'If Rscek.State = adStateOpen Then Rscek.Close
'    Cboid = vbNullString
'    Cboid1 = vbNullString
'    StrSQL = "select * from Divisi Where kd_bid >= 0 and kd_bid <= 20 order by kd_div"
'    Rscek.Open StrSQL, CN, adOpenStatic
'    Cboid1 = " "
'    Do Until Rscek.EOF
'      Cboid = "|" & Rscek("kd_Div") & vbTab & Rscek("NM_DIV")
'      Cboid1 = Cboid1 + Cboid
'
'      Rscek.MoveNext
'    Loop
'    CboDivisi.ColComboList(0) = Cboid1
'End Sub
'Private Sub Command2_Click()
'If fg.Rows > 1 Then
''   fg.Cols = 11
'   fg.SaveGrid "C:\TotalJamTimesheet.csv", flexFileCommaText, True
'    Shell PathOffice & "C:\TotalJamTimesheet.csv", vbNormalFocus
'End If
'End Sub
'
'Private Sub Command3_Click()
'On Error Resume Next
'If fg.Rows > 1 Then fg.PrintGrid "Total Jam - Periode " & DTPicker1.Value & " S/D " & DTPicker2.Value, , 2, 900, 500
'
'End Sub
'
'Private Sub fg_DblClick()
'With fg
'    If .Row > 0 Then
'        frmRekapTimesheet.CboFlex = .TextMatrix(.Row, 1)
'        frmRekapTimesheet.Showdata
'        LoadForm frmRekapTimesheet
'
'    End If
'End With
'
'End Sub
'
'Private Sub Form_Load()
'    AddDivisi
'    Setgrid
'    DTPicker1.Value = Date
'    DTPicker2.Value = Date
'    DTPicker1.Value = DateSerial(Year(Now), Month(Now), 1)
'    DTPicker1.Value = DateAdd("M", 0, DTPicker1.Value)
'    DTPicker2.Value = DateSerial(Year(Now), Month(Now), 1)
'    DTPicker2.Value = DateAdd("M", 1, DTPicker2.Value) - 1
'    DTPicker1.CustomFormat = "dd/MMM/yyyy"
'    DTPicker2.CustomFormat = "dd/MMM/yyyy"
'     If Len(skinsFileName) <> 0 Then
'      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
'      Skin1.ApplySkin hwnd
'    End If
'    Select Case UCase(strGroup)
'        Case "IT", "PTW"
'            CboDivisi.Enabled = True
'        Case "ADMIN"
'            CboDivisi.Text = NamaDivisi
'            CboDivisi.Enabled = False
'    End Select
'End Sub
'
'Sub Setgrid()
'Dim x As Integer
'
'With fg
'    .Cols = 8
'    .Rows = 2
'    .FixedRows = 2
'
'    .TextMatrix(0, 0) = "No"
'    .TextMatrix(0, 1) = "NIP"
'    .TextMatrix(0, 2) = "Nama"
'    .TextMatrix(0, 3) = "Total Jam"
'    .TextMatrix(0, 4) = "Total Jam"
'    .TextMatrix(0, 5) = "Total Jam"
'    .TextMatrix(0, 6) = "Total Jam"
'    .TextMatrix(0, 7) = "Total Jam"
''    .TextMatrix(0, 8) = "Total Jam" '"Jam Lembur"
''    .TextMatrix(0, 9) = "Total Jam" '"Biaya Lembur"
'    .TextMatrix(1, 0) = "No"
'    .TextMatrix(1, 1) = "NIP"
'    .TextMatrix(1, 2) = "Nama"
'    .TextMatrix(1, 3) = "Kerja Standar"
'    .TextMatrix(1, 4) = "Belum Diverifikasi"
'    .TextMatrix(1, 5) = "Sudah Diverifikasi"
'    .TextMatrix(1, 6) = "Kerja + Lembur"
'    .TextMatrix(1, 7) = "Persen"
''    .TextMatrix(1, 8) = "Kerja + Lembur" '"Jam Lembur"
''    .TextMatrix(1, 9) = "Persen" '"Biaya Lembur"
'    .RowHeight(1) = 700
'    .ColWidth(0) = 500
'    .ColWidth(1) = 700
'    .ColWidth(2) = 3000
'    .ColWidth(3) = 850
'
'    .MergeCells = flexMergeFree
'    .MergeRow(0) = True
'    .MergeCol(0) = True
'    .MergeCol(1) = True
'    .MergeCol(2) = True
'
'    For x = 0 To .Cols - 1
'       .FixedAlignment(x) = flexAlignCenterCenter
'    Next
'End With
'With VSJoin
'.Rows = 1
'.Cols = 19
'.TextMatrix(.Rows - 1, 1) = "1. Tanggal"
'.TextMatrix(.Rows - 1, 2) = "2.Project"
'.TextMatrix(.Rows - 1, 3) = "3. NIP"
'.TextMatrix(.Rows - 1, 4) = "4. Nama"
'.TextMatrix(.Rows - 1, 5) = "5. Total / 2"
'.TextMatrix(.Rows - 1, 6) = "6. masuk"
'.TextMatrix(.Rows - 1, 7) = "7. keluar"
'.TextMatrix(.Rows - 1, 8) = "8. NamaProject"
'.TextMatrix(.Rows - 1, 9) = "9. StatusDivisi"
'.TextMatrix(.Rows - 1, 10) = "10. TotalLembur"
'.TextMatrix(.Rows - 1, 11) = "11. StatusPM"
'.TextMatrix(.Rows - 1, 12) = "12. Keterangan"
'.TextMatrix(.Rows - 1, 13) = "13. Hari"
'.TextMatrix(.Rows - 1, 14) = "14. Divisi"
'.TextMatrix(.Rows - 1, 15) = "15. gaji"
'.TextMatrix(.Rows - 1, 16) = "16. Lembur"
'.TextMatrix(.Rows - 1, 17) = "17. ProjetUmum"
'.TextMatrix(.Rows - 1, 18) = "18. TotalKerja"
'.ColDataType(6) = flexDTDate
'.ColDataType(7) = flexDTDate
'.ColFormat(6) = "HH:mm"
'.ColFormat(7) = "HH:mm"
''.ColFormat(10) = "#.##"
'.ColFormat(15) = "#,###"
'.ColFormat(16) = "#,###"
'
'End With
'End Sub
'Private Sub AddKaryawan(Divisi As String)
'    Dim Cboid     As String
'    Dim Cboid1    As String
'If Rscek.State = adStateOpen Then Rscek.Close
'    Cboid = vbNullString
'    Cboid1 = vbNullString
'    StrSQL = "select * from Karyawan Where kd_divisi = '" & Divisi & "' And Status <> '14' Order By Nama"
'    Rscek.Open StrSQL, CN, adOpenStatic
'    Cboid1 = " "
'    Combo1.Clear
'    Do Until Rscek.EOF
'
'      Combo1.AddItem Rscek!NIP
'      Rscek.MoveNext
'    Loop
'
'End Sub
'Private Sub Form_Resize()
'    On Error Resume Next
'        CmdClose.Width = Me.Width - 100
'        With fg
'             .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture2.Height
'
'        End With
''
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'
'    Set frmLapBiayaProject = Nothing
'End Sub
'Sub Showdata()
'Dim x, TotalHari As Integer
'Dim i, J As Integer
'Dim Split, JmlLoop As Integer
'Dim Jam, Hari As String
'Dim IDWaktu As String
'Dim Jam1, Jam2 As Date
'Dim Jam3, Jam4 As Date
'Dim TglAwal, TglAkhir As Date
'Dim RsTS As New ADODB.Recordset
'Dim searchDivisi As String
'
'
'    Select Case UCase(strGroup)
'        Case "IT", "PTW"
'            searchDivisi = CboDivisi
'
'        Case "ADMIN"
'            searchDivisi = KodeDivisi
'    End Select
'
'  If Trim(CboDivisi) = "" Then
'    MsgBox "Nama Divisi Belum Diisi", vbCritical
'    Exit Sub
'  End If
'  AddKaryawan (searchDivisi)
'   TglAwal = DTPicker1
'    TglAwal = DateSerial(Year(DTPicker1), Month(DTPicker1), 1)
'    TglAwal = DateAdd("M", 0, DTPicker1)
'    TglAwal = Format(DTPicker1, "MM") & "/01/" & Format(DTPicker1, "yyyy")
'    TglAkhir = DTPicker1
'    TglAkhir = DateSerial(Year(DTPicker1), Month(DTPicker1), 1)
'    TglAkhir = DateAdd("M", 1, TglAkhir) - 1
'    Selisih = DateDiff("d", TglAwal, TglAkhir)
'    TotalHari = 0
'
'    For x = 0 To Selisih
'       If x > 0 Then TglAwal = DateAdd("d", 1, TglAwal)
'       Hari = Format(TglAwal, "ddd")
'       If Hari <> "Sat" And Hari <> "Sun" Then
'
'            StrSQL = "select tanggallibur from kalender " & _
'                "where tanggallibur = '" & Format(TglAwal, "MM/dd/yyyy") & "'"
'            If Rscek.State = adStateOpen Then Rscek.Close
'            Rscek.Open StrSQL, CN, adOpenStatic
'            If Rscek.EOF Then
'               TotalHari = TotalHari + 1
'            End If
'        End If
'    Next
'
'    TotalJam = TotalHari * 8
'Command1.Enabled = False
'With VSFlexGrid1
'        .Rows = 1
'        VSFg.Rows = 1
'        VSFg.Cols = 17
'
'    For x = 0 To Combo1.ListCount - 1
'            Combo1.ListIndex = x
'            Command1.Caption = Combo1.Text
'            If RsTS.State = adStateOpen Then RsTS.Close
'            StrSQL = "SELECT tbltimesheet.IDtimesheet,tbltimesheet.Tanggal,tbltimesheet.JamAwal As [Jam Awal],tbltimesheet.JamAkhir AS [Jam Akhir],tbltimesheet.Status,tbltimesheet.NoProject As Project,tbltimesheet.Keterangan,tbltimesheet.Tanggal,Absensi.Masuk,tbltimesheet.NIP,tbltimesheet.StatusDivisi, karyawan.Nama,tbltimesheet.StatusPM,Absensi.Keluar,Divisi.NM_DIV, karyawan.Nama,tbltimesheet.Hari,tbltimesheet.TotalKerja,tbltimesheet.ProjectUmum"
'            StrSQL = StrSQL & " FROM tbltimesheet INNER JOIN karyawan ON tbltimesheet.NIP = karyawan.NIP INNER JOIN Divisi ON karyawan.Kd_Divisi = Divisi.KD_DIV INNER JOIN Absensi ON tbltimesheet.NIP = Absensi.NIP AND tbltimesheet.Tanggal = Absensi.Tgl"
''            StrSQL = "SELECT IDtimesheet,Tanggal,JamAwal As [Jam Awal],JamAkhir AS [Jam Akhir],Status,NoProject As Project,Keterangan,Tanggal,Masuk,StatusPM,StatusDivisi From tbltimesheet Where Tanggal Between '" & Format(TglAwal, "yyyy/MM/dd") & "' And '" & Format(TglAkhir, "yyyy/MM/dd") & "' And NIP = '" & StrNIPUser & "' Order By Tanggal DESC,IDTimesheet ASC"
'            StrSQL = StrSQL & " Where tbltimesheet.Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And '" & Format(DTPicker2, "MM/dd/yyyy") & "'  And tbltimesheet.Status ='Actual'"
'            StrSQL = StrSQL & " AND tbltimesheet.Nip = '" & Combo1.Text & "'"
'            StrSQL = StrSQL & " And tbltimesheet.Kd_Divisi = '" & searchDivisi & "'"
'            StrSQL = StrSQL & " Order By Divisi.NM_DIV,tbltimesheet.NoProject,tbltimesheet.Tanggal,tbltimesheet.NIP,tbltimesheet.IDTimesheet ASC"
'            RsTS.Open StrSQL, CN, adOpenStatic
'            Set .DataSource = RsTS
'            .ColDataType(2) = flexDTDate
'
'            For Lrow = 1 To .Rows - 1
'
'                .TextMatrix(Lrow, 0) = Lrow
'                .TextMatrix(0, 0) = .Rows
'                If .TextMatrix(Lrow, 3) = "" Then .TextMatrix(Lrow, 3) = "00:00"
'                 If .TextMatrix(Lrow, 4) = "" Then .TextMatrix(Lrow, 4) = "00:00"
'
'                        JmlLoop = 0
'                        Jam1 = CDate(.TextMatrix(Lrow, 3))
'                        Jam2 = CDate(.TextMatrix(Lrow, 4))
'                        DTPicker3.Value = Jam1
'                        Do Until JmlLoop = 50
'
'                             VSFg.Rows = VSFg.Rows + 1
'                                DTPicker3.Value = DateAdd("n", 30, DTPicker3)
'                                VSFg.TextMatrix(VSFg.Rows - 1, 0) = VSFg.Rows - 1
'                                IDWaktu = Format(DTPicker3.Value, "HH:mm")
'                                VSFg.TextMatrix(VSFg.Rows - 1, 1) = IDWaktu
'                                VSFg.TextMatrix(VSFg.Rows - 1, 2) = .TextMatrix(Lrow, 2) 'Format(.TextMatrix(Lrow, 2), "dd/MM/yyyy")
'                                VSFg.TextMatrix(VSFg.Rows - 1, 3) = .TextMatrix(Lrow, 6)
'                                VSFg.TextMatrix(VSFg.Rows - 1, 4) = .TextMatrix(Lrow, 10)
'                                VSFg.TextMatrix(VSFg.Rows - 1, 5) = .TextMatrix(Lrow, 5)
'                                VSFg.TextMatrix(VSFg.Rows - 1, 6) = Format(.TextMatrix(Lrow, 3), "HH:mm")
'                                VSFg.TextMatrix(VSFg.Rows - 1, 7) = .TextMatrix(Lrow, 7)
'                                VSFg.TextMatrix(VSFg.Rows - 1, 8) = .TextMatrix(Lrow, 11) 'StatusDivisi
'                                VSFg.TextMatrix(VSFg.Rows - 1, 9) = .TextMatrix(Lrow, 12)
'                                VSFg.TextMatrix(VSFg.Rows - 1, 10) = .TextMatrix(Lrow, 1)
'                                VSFg.TextMatrix(VSFg.Rows - 1, 11) = .TextMatrix(Lrow, 4)
'                                VSFg.TextMatrix(VSFg.Rows - 1, 12) = .TextMatrix(Lrow, 15)
'                                VSFg.TextMatrix(VSFg.Rows - 1, 13) = .TextMatrix(Lrow, 13) 'StatusPM
'                                VSFg.TextMatrix(VSFg.Rows - 1, 14) = .TextMatrix(Lrow, 17)
'                                VSFg.TextMatrix(VSFg.Rows - 1, 15) = .TextMatrix(Lrow, 18)
'                                VSFg.TextMatrix(VSFg.Rows - 1, 16) = .TextMatrix(Lrow, 19)
'                                Jam3 = Format(.TextMatrix(Lrow, 9), "HH:mm")
'
'                                Jam4 = DateAdd("h", 9, Jam3)
'                                If Trim(.TextMatrix(Lrow, 17)) = "Kerja" Then
'                                    If (VSFg.TextMatrix(VSFg.Rows - 1, 1) >= "08:00" And VSFg.TextMatrix(VSFg.Rows - 1, 1) <= "17:00") Or (VSFg.TextMatrix(VSFg.Rows - 1, 1) < Jam4 And VSFg.TextMatrix(VSFg.Rows - 1, 1) >= "08:00") Then
'                                         VSFg.TextMatrix(VSFg.Rows - 1, 7) = "Timesheet"
'                                      End If
'
'                                End If
'                                    If Format(DTPicker3.Value, "HH:mm") = Format(Jam2, "hh:mm") Then Exit Do
'                                If Format(DTPicker3, "HH:mm") = "12:00" Then DTPicker3.Value = DateAdd("n", 60, DTPicker3)
'                                JmlLoop = JmlLoop + 1
'                        Loop
'            Next
'                VSFg.ColFormat(6) = "HH:mm"
'                VSFg.ColFormat(11) = "HH:mm"
''            For lCol = 1 To VSFg.Cols - 1
''                VSFg.TextMatrix(0, lCol) = lCol
''            Next
'        If .Rows = 1 Then
'            If Rscek.State = adStateOpen Then Rscek.Close
'            Rscek.Open "Select * From Vuser Where NIP = '" & Combo1 & "'", CN, adOpenStatic
'            With fg
'                If Not Rscek.EOF Then
'                     .Rows = .Rows + 1
'                     .TextMatrix(.Rows - 1, 0) = .Rows - 1
'                     .TextMatrix(.Rows - 1, 1) = Rscek!NIP
'                     .TextMatrix(.Rows - 1, 2) = Rscek!Nama
'                     .TextMatrix(.Rows - 1, 3) = TotalJam
'                     .TextMatrix(.Rows - 1, 4) = 0
'                     .TextMatrix(.Rows - 1, 5) = 0
'                     .TextMatrix(.Rows - 1, 6) = 0
'                     .TextMatrix(.Rows - 1, 7) = 0
''                     .TextMatrix(.Rows - 1, 8) = 0
''                     .TextMatrix(.Rows - 1, 9) = 0
'                End If
'            End With
'        Else
'            With VSFg
'                VSJoin.Rows = 1
'                For Lrow = 1 To VSFg.Rows - 1
'                         J = JoinGrid(.TextMatrix(Lrow, 3), .TextMatrix(Lrow, 2), Lrow, .TextMatrix(Lrow, 4), .TextMatrix(Lrow, 7))
'                Next
'            End With
'
'            With VSJoin
'                VSFg.Rows = 1
'            For Lrow = 1 To VSJoin.Rows - 1
'                If Trim(.TextMatrix(Lrow, 5)) = "" Then .TextMatrix(Lrow, 5) = 0
'                .TextMatrix(Lrow, 0) = Lrow
'                .TextMatrix(0, 0) = .Rows
'
'                  If Trim(.TextMatrix(Lrow, 12)) = "Lembur" Then
'                   J = .TextMatrix(Lrow, 10)
'
'                  End If
'
'                J = HitungGaji(.TextMatrix(Lrow, 3), .TextMatrix(Lrow, 9), .TextMatrix(Lrow, 14), Lrow, .TextMatrix(Lrow, 12))
'            Next
'
'            End With
'        End If
'
'    Next
'End With
'
''Disiplit Perhari
'
'
'With fg
'
'    For Lrow = 2 To .Rows - 1
'        .TextMatrix(Lrow, 0) = Lrow - 1
'        .TextMatrix(Lrow, 6) = CDbl(.TextMatrix(Lrow, 4)) + CDbl(.TextMatrix(Lrow, 5))
'        .TextMatrix(Lrow, 7) = Round(CDbl(.TextMatrix(Lrow, 6)) / CDbl(.TextMatrix(Lrow, 3)) * 100, 2)
'        .TextMatrix(Lrow, 7) = .TextMatrix(Lrow, 7) & "%"
'    Next
'
'End With
'
'Command1.Caption = "Refresh"
'Command1.Enabled = True
'End Sub
'
'Function HitungGaji(NIP As String, StatusDivisi As String, Divisi As String, Row As Integer, Keterangan As String)
'Dim StatusProject As Boolean
'Dim FRow As Integer
'Dim HRow As Integer
'With fg
'    StatusProject = False
'    For FRow = 0 To .Rows - 1
'    If Trim(.TextMatrix(FRow, 1)) = NIP Then
'           StatusProject = True
'           HRow = FRow
'        End If
'    Next
'        If StatusProject = False Then
'             .Rows = .Rows + 1
'             HRow = .Rows - 1
'             If .TextMatrix(HRow, 4) = "" Then .TextMatrix(HRow, 4) = 0
'            If .TextMatrix(HRow, 5) = "" Then .TextMatrix(HRow, 5) = 0
'            If .TextMatrix(HRow, 6) = "" Then .TextMatrix(HRow, 6) = 0
'            If .TextMatrix(HRow, 7) = "" Then .TextMatrix(HRow, 7) = 0
'
''            If .TextMatrix(HRow, 9) = "" Then .TextMatrix(HRow, 9) = 0
'
'        End If
'            .TextMatrix(HRow, 1) = VSJoin.TextMatrix(Row, 3)
'            .TextMatrix(HRow, 2) = VSJoin.TextMatrix(Row, 4)
'            .TextMatrix(HRow, 3) = TotalJam
'
''            If Trim(Keterangan) = "Timesheet" Then
'                If StatusDivisi = 0 Then
'                   .TextMatrix(HRow, 4) = CCur(.TextMatrix(HRow, 4)) + CCur(VSJoin.TextMatrix(Row, 5))
'                Else
'                   .TextMatrix(HRow, 5) = CCur(.TextMatrix(HRow, 5)) + VSJoin.TextMatrix(Row, 5)
'                End If
'                .TextMatrix(HRow, 6) = CCur(.TextMatrix(HRow, 6)) + CCur(VSJoin.TextMatrix(Row, 5))
''             Else
''                .TextMatrix(HRow, 7) = CDbl(.TextMatrix(HRow, 7)) + CDbl(VSJoin.TextMatrix(Row, 10))
''            End If
'End With
'
'End Function
'Function JoinGrid(ByVal Project As String, ByVal tgl1 As String, ByVal Row As Integer, ByVal NIP As String, ByVal Keterangan As String)
'Dim StatusNIP As Boolean
'Dim FRow As Integer
'Dim J As String
'With VSJoin
'    StatusNIP = False
'    For FRow = 0 To .Rows - 1
'        If .TextMatrix(FRow, 2) = Project And .TextMatrix(FRow, 1) = tgl1 And .TextMatrix(FRow, 3) = NIP And .TextMatrix(FRow, 12) = Keterangan Then
'           StatusNIP = True
'        End If
'    Next
'        If StatusNIP = False Then
'             .Rows = .Rows + 1
'            .TextMatrix(.Rows - 1, 5) = 0
''            .TextMatrix(.Rows - 1, 10) = 0
'        End If
'            .TextMatrix(.Rows - 1, 1) = VSFg.TextMatrix(Row, 2)
'            .TextMatrix(.Rows - 1, 2) = VSFg.TextMatrix(Row, 3)
'            .TextMatrix(.Rows - 1, 3) = VSFg.TextMatrix(Row, 4)
'            .TextMatrix(.Rows - 1, 4) = VSFg.TextMatrix(Row, 9)
'            If Trim(Keterangan) = "Timesheet" Then
'               .TextMatrix(.Rows - 1, 5) = (.TextMatrix(.Rows - 1, 5) + 0.5)
'            End If
'            .TextMatrix(.Rows - 1, 6) = Format(VSFg.TextMatrix(Row, 6), "HH:mm")
'            .TextMatrix(.Rows - 1, 7) = Format(VSFg.TextMatrix(Row, 11), "HH:mm")
'            .TextMatrix(.Rows - 1, 8) = VSFg.TextMatrix(Row, 13)
'            .TextMatrix(.Rows - 1, 9) = VSFg.TextMatrix(Row, 8)
'            If .TextMatrix(.Rows - 1, 10) = "" Then .TextMatrix(.Rows - 1, 10) = 0
'            If Trim(Keterangan) = "Lembur" Then
'                .TextMatrix(.Rows - 1, 10) = .TextMatrix(.Rows - 1, 10) + 0.5
'            End If
'            .TextMatrix(.Rows - 1, 11) = VSFg.TextMatrix(Row, 13)
'            .TextMatrix(.Rows - 1, 12) = VSFg.TextMatrix(Row, 7)
'            .TextMatrix(.Rows - 1, 13) = VSFg.TextMatrix(Row, 14) 'Hari
'            .TextMatrix(.Rows - 1, 14) = VSFg.TextMatrix(Row, 12)
''            .TextMatrix(.Rows - 1, 15) = GetSet(.TextMatrix(.Rows - 2, 3), .TextMatrix(.Rows - 1, 3)) = True
''             .TextMatrix(.Rows - 1, 15) = Round(CCur(((NilaiGaji) / 173)) * CDbl(.TextMatrix(.Rows - 1, 5)), 2)
'            .TextMatrix(.Rows - 1, 16) = 0
'            .TextMatrix(.Rows - 1, 17) = VSFg.TextMatrix(Row, 15)
'            .TextMatrix(.Rows - 1, 18) = VSFg.TextMatrix(Row, 16)
'End With
'End Function
'
'

