VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLapBiayaNonVerifikasi 
   Caption         =   "Rekap Biaya Project Non Verifikasi"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3255
      Left            =   0
      TabIndex        =   22
      Top             =   1800
      Width           =   8295
      _cx             =   14631
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
      FormatString    =   $"FrmLapBiayaNonVerifikasi.frx":0000
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
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1785
      ScaleWidth      =   19050
      TabIndex        =   2
      Top             =   0
      Width           =   19080
      Begin VB.CommandButton Command7 
         Caption         =   "Show Gaji"
         Height          =   375
         Left            =   8400
         TabIndex        =   31
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   11160
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Next Sheet"
         Height          =   375
         Left            =   12360
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Open Gaji"
         Height          =   375
         Left            =   8400
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Before Sheet"
         Height          =   375
         Left            =   13440
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Export To Sheet"
         Height          =   375
         Left            =   9840
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   6960
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Print out"
         Height          =   375
         Left            =   9840
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   10920
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   57802754
         CurrentDate     =   39940
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   720
         TabIndex        =   11
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
         Format          =   57802755
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2880
         TabIndex        =   12
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
         Format          =   57802755
         CurrentDate     =   39931
      End
      Begin VSFlex8Ctl.VSFlexGrid cboFlex 
         Height          =   315
         Left            =   3720
         TabIndex        =   13
         Top             =   600
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
         FormatString    =   $"FrmLapBiayaNonVerifikasi.frx":00DF
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
      Begin VSFlex8Ctl.VSFlexGrid CboKaryawan 
         Height          =   315
         Left            =   720
         TabIndex        =   14
         Top             =   960
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
         FormatString    =   $"FrmLapBiayaNonVerifikasi.frx":0108
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
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   11040
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   57802754
         CurrentDate     =   39940
      End
      Begin VSFlex8Ctl.VSFlexGrid CboDivisi 
         Height          =   315
         Left            =   720
         TabIndex        =   16
         Top             =   600
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
         FormatString    =   $"FrmLapBiayaNonVerifikasi.frx":0131
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
         TabIndex        =   30
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "* Total Biaya = (Nilai Gaji/Total Jam Periode Bulan) * Total Jam"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   29
         Top             =   1560
         Width           =   8895
      End
      Begin VB.Label Label9 
         Caption         =   "* Untuk Edit File Gaji Gunakan Microsoft Office Excel, Jika dengan Open Office File Gaji Tidak terbaca "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   28
         Top             =   1320
         Width           =   8895
      End
      Begin VB.Label Label8 
         Caption         =   "Note :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   27
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   495
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
         TabIndex        =   20
         Top             =   120
         Width           =   375
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
         TabIndex        =   19
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label6 
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
         Left            =   2400
         TabIndex        =   18
         Top             =   600
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
         TabIndex        =   17
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   19080
      TabIndex        =   0
      Top             =   10995
      Width           =   19080
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
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "FrmLapBiayaNonVerifikasi.frx":015A
      Top             =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid VSGaji 
      Height          =   1455
      Left            =   0
      TabIndex        =   26
      ToolTipText     =   "Double Klik Kolom Project Untuk Melihat PM"
      Top             =   1800
      Width           =   8295
      _cx             =   14631
      _cy             =   2566
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
      FormatString    =   $"FrmLapBiayaNonVerifikasi.frx":038E
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
      Left            =   8400
      TabIndex        =   25
      Top             =   1800
      Width           =   10575
      _cx             =   18653
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
      FormatString    =   $"FrmLapBiayaNonVerifikasi.frx":046D
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
   Begin VSFlex8Ctl.VSFlexGrid VSFg 
      Height          =   3255
      Left            =   8400
      TabIndex        =   23
      Top             =   5880
      Width           =   10575
      _cx             =   18653
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
      FormatString    =   $"FrmLapBiayaNonVerifikasi.frx":054C
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
      Height          =   3975
      Left            =   0
      TabIndex        =   24
      Top             =   5160
      Width           =   8295
      _cx             =   14631
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
      FormatString    =   $"FrmLapBiayaNonVerifikasi.frx":062B
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
Attribute VB_Name = "FrmLapBiayaNonVerifikasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NilaiGaji As Currency
Dim Tingkat As Integer
Dim TotalHari, StrTotalJam As Integer
Dim temptotaljam, tempbiayalembur1 As String
Dim TotalUM As Currency
Const APPNAME = "Excel"
Dim sheet%
Dim i As Integer, FileTitle As String
Private Sub CboDivisi_AfterEdit(ByVal Row As Long, ByVal Col As Long)
AddProject
AddKaryawan
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub
Private Sub Command1_Click()
Dim Tanggal As String
If VSGaji.Rows <= 2 Then
   MsgBox "Data Gaji Karyawan Masih Kosong", vbCritical
   Exit Sub
End If
'Tanggal = Format(DTPicker1, "MM")
'If Mid(FileTitle, 6, 2) <> Tanggal Then
'    If MsgBox("File Data Gaji Tidak Sama Dengan Tanggal Pencarian, Apakah Anda Akan Melanjutkan Proses ini ?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then
'           Exit Sub
'    End If
'End If

fg.Visible = True
VSGaji.Visible = False
Command7.Caption = "Show Gaji"
Command1.Enabled = False
Setgrid
Showdata
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Dim x As String
On Error GoTo Adaerror
With fg
If fg.Rows > 1 Then
     
     
        .AddItem "", 1
        .AddItem "", 2
        
        .Redraw = flexRDNone
        .Redraw = flexRDBuffered
        .TextMatrix(1, 1) = "Rekap Biaya Project Periode " & Format(DTPicker1, "dd MMM yyyy") & " s/d " & Format(DTPicker2, "dd MMM yyyy")
    For lCol = 1 To .Cols - 1
        
       .TextMatrix(2, lCol) = .TextMatrix(0, lCol)
        .Row = 2
        .Col = lCol
        .CellBackColor = vbGreen '&HE0E0E0
    Next
        fg.SaveGrid "C:\BiayaPerProject2.xls", flexFileExcel, False
'        Shell PathOffice & "C:\BiayaPerProject2.xls", vbNormalFocus
        x = ShellExecute(Me.hwnd, "open", "C:\BiayaPerProject2.xls", vbNullString, "C:\BiayaPerProject2.xls", 1)
  
          .RemoveItem (1)
          .RemoveItem (1)
End If
End With
Exit Sub
Adaerror:
MsgBox err.Description
End Sub

Private Sub Command3_Click()
On Error Resume Next
If fg.Rows > 1 Then fg.PrintGrid "Biaya Per Project2 - Periode " & DTPicker1.Value & " S/D " & DTPicker2.Value, 2, 2, 900, 500

End Sub

Private Sub Command4_Click()

Showgaji
'fg.Visible = False
'VSGaji.Visible = True
End Sub

Private Sub Command5_Click()
sheet = sheet - 1

    MousePointer = MousePointerConstants.vbHourglass

    On Error Resume Next
    VSGaji.LoadGrid dlg.FileName, flexFileCommaText, sheet
     With VSGaji
        Dim Lrow As Long
        For Lrow = 1 To .Rows - 1
          .TextMatrix(Lrow, 0) = Lrow
        Next
    End With
    If err <> 0 Then
        MsgBox "No More Sheets"
        sheet = 0
'        Command5.Enabled = False
    End If
    On Error GoTo 0

    MousePointer = MousePointerConstants.vbDefault
End Sub

Private Sub Command6_Click()
sheet = sheet + 1

    MousePointer = MousePointerConstants.vbHourglass

    On Error Resume Next
    VSGaji.LoadGrid dlg.FileName, flexFileCommaText, sheet
     With VSGaji
        Dim Lrow As Long
        For Lrow = 1 To .Rows - 1
          .TextMatrix(Lrow, 0) = Lrow
        Next
    End With
    If err <> 0 Then
        MsgBox "No More Sheets"
        sheet = 0
'        Command6.Enabled = False
    End If
    On Error GoTo 0

    MousePointer = MousePointerConstants.vbDefault

End Sub

Private Sub Command7_Click()
Select Case Command7.Caption
    Case "Show Gaji"
        VSGaji.Visible = True
        fg.Visible = False
        Command7.Caption = "Hide Gaji"

    Case "Hide Gaji"
        VSGaji.Visible = False
        fg.Visible = True
        Command7.Caption = "Show Gaji"
End Select

End Sub

Private Sub Form_Load()

    AddDivisi
    AddKaryawan
    AddProject
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
 
    If VSGaji.Rows = 1 Then Showgaji
End Sub
 
Sub Showgaji()
On Error GoTo Adaerror
    MsgBox "Silahkan Pilih File Data Gaji Karyawan", vbInformation

    dlg.FileName = ""
    dlg.Filter = "File Gaji (*.xls)|*.xls"
    dlg.DialogTitle = "Data Gaji Karyawan "
    dlg.ShowOpen
    If Len(dlg.FileName) = 0 Then Exit Sub
    FileTitle = dlg.FileTitle
    MousePointer = MousePointerConstants.vbHourglass
    VSGaji.LoadGrid dlg.FileName, flexFileExcel
    MousePointer = MousePointerConstants.vbDefault

    Caption = APPNAME + " " + dlg.FileName

    sheet = 0
    Command5.Enabled = True
    With VSGaji
        Dim Lrow As Long
        For Lrow = 1 To .Rows - 1
          .TextMatrix(Lrow, 0) = Lrow
        Next
    End With
Exit Sub
Adaerror:
MousePointer = MousePointerConstants.vbDefault
MsgBox err.Description & " Or File NO Match "

End Sub
Sub Setgrid()
With fg
.Cols = 15
.Rows = 1
.ColWidth(0) = 300
.ColWidth(1) = 2000
.ColWidth(3) = 2500
.ColWidth(2) = 700
.ColWidth(4) = 1000
.ColWidth(5) = 2000
.ColWidth(6) = 900
.ColDataType(6) = flexDTDouble
.MergeCells = flexMergeFree
.MergeCol(1) = True
.MergeCol(2) = True
.MergeCol(3) = True
.MergeCol(4) = True
.MergeCol(5) = True

.TextMatrix(0, 0) = "No"
.TextMatrix(0, 1) = "Nama Divisi"
.TextMatrix(0, 4) = "Kode Project"
.TextMatrix(0, 5) = "Nama Project"
.TextMatrix(0, 2) = "NIP"
.TextMatrix(0, 3) = "Nama"
.TextMatrix(0, 6) = "Total Jam"
.TextMatrix(0, 7) = "Total %"
.TextMatrix(0, 8) = "Total Biaya" '"Jam Lembur"
.TextMatrix(0, 9) = "Jam Lembur" '"Biaya Lembur"
.TextMatrix(0, 10) = "Biaya Lembur"
.TextMatrix(0, 11) = "Project UMUM"
.TextMatrix(0, 12) = "Total Kerja"
.TextMatrix(0, 13) = "Uang Makan"
.TextMatrix(0, 14) = "Transport"
.ColFormat(8) = "#,###"
'.ColFormat(9) = "#,###"
.ColFormat(10) = "#,###"
.ColFormat(13) = "#,###"
.ColWidth(10) = 1000
.ColWidth(6) = 1000
.ColWidth(7) = 1000
.ColWidth(8) = 1200
.ColWidth(9) = 1200
.ColWidth(10) = 0
.ColWidth(11) = 0
.ColWidth(12) = 0
.ColWidth(13) = 0
.ColWidth(14) = 0
.ColDataType(6) = flexDTDouble
.ColDataType(9) = flexDTDouble
.ColDataType(10) = flexDTCurrency
.ColDataType(13) = flexDTDouble
'.ColDataType(14) = flexDTDecimal
'.ColFormat(14) = "#,###.##"
End With
End Sub
Private Sub AddKaryawan()

    Dim Cboid     As String
    Dim cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    cboid1 = vbNullString
    If Trim(CboDivisi) = "" Then
        StrSQL = "select * from Karyawan Where Status <> '14' Order By NIP"
    Else
        StrSQL = "select * from Karyawan Where Status <> '14' And kd_divisi = '" & CboDivisi & "'  Order By NIP"

    End If
    Rscek.Open StrSQL, CN, adOpenStatic
    cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("NIP") & vbTab & Rscek("Nama")
      cboid1 = cboid1 + Cboid
      Combo1.AddItem Rscek!NIP
      Rscek.MoveNext
    Loop
    CboKaryawan.ColComboList(0) = cboid1
    CboKaryawan.CellAlignment = flexAlignLeftCenter
End Sub


Private Sub AddProject()

    Dim Cboid     As String
    Dim cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    cboid1 = vbNullString
'    If Trim(CboDivisi) = "" Then
        StrSQL = "select Kode,Nama from project " & _
             "group by kode,Nama " & _
             "order by kode"
'    Else
'        StrSQL = "select Kode,Nama from project " & _
'             "where  Kd_Divisi = '" & CboDivisi & "'" & _
'             "group by kode,Nama " & _
'             "order by kode"
'    End If
    Rscek.Open StrSQL, CN, adOpenStatic
    cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("Kode") & vbTab & Rscek("Nama")
      cboid1 = cboid1 + Cboid
      Rscek.MoveNext
    Loop
    cboFlex.ColComboList(0) = cboid1
    cboFlex.CellAlignment = flexAlignLeftCenter
End Sub
Sub AddDivisi()
Dim Cboid     As String
Dim cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
StrSQL = "select * from Divisi Where kd_bid >= 2 and kd_bid <= 20 order by kd_bid"
Rscek.Open StrSQL, CN, adOpenStatic
cboid1 = " "
Do Until Rscek.EOF
     Cboid = "|" & Rscek("Kd_div") & vbTab & Rscek("NM_DIV")
     cboid1 = cboid1 + Cboid
     Rscek.MoveNext
Loop
    CboDivisi.ColComboList(0) = cboid1
    CboDivisi.CellAlignment = flexAlignLeftCenter
End Sub
Private Sub Form_Resize()
    On Error Resume Next
        CmdClose.Width = Me.Width - 100
        With fg
             .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture2.Height - 150

        End With
        With VSGaji
             .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture2.Height - 150

        End With
         With VSJoin
             .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture2.Height

        End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set FrmLapBiayaNonVerifikasi = Nothing
End Sub
Sub Showdata()
Dim x, Menit, TMenit As Currency
Dim i, J, Tjam As Currency
Dim Split, JmlLoop As Currency
Dim Jam, Hari As String
Dim IDWaktu As String
Dim Jam1, Jam2 As Date
Dim Jam3, Jam4 As Date
Dim TglAwal, TglAkhir As Date
Dim RsTS As New ADODB.Recordset
Dim Uangmakan As Currency
On Error GoTo Adaerror
If Trim(CboDivisi.Text) = "" And Trim(CboKaryawan.Text) = "" Then
   MsgBox "Silahkan Pilih Divisi / Karyawan Terlebih Dahulu", vbCritical
   cboFlex.SetFocus
   Exit Sub
End If

CN.Execute "Delete From TblcetakTs"
Command1.Enabled = False
With VSFlexGrid1
        .Rows = 1
        VSFg.Rows = 1
        VSFg.Cols = 18
        VSJoin.Cols = 17
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
     
    StrTotalJam = TotalHari * 8
    LblTotaljam = StrTotalJam & " Jam"
     If RsTS.State = adStateOpen Then RsTS.Close
    StrSQL = "SELECT tbltimesheet.IDtimesheet,tbltimesheet.Tanggal,tbltimesheet.JamAwal As [Jam Awal],tbltimesheet.JamAkhir AS [Jam Akhir],tbltimesheet.Status,tbltimesheet.NoProject As Project,tbltimesheet.Keterangan,tbltimesheet.Tanggal,Absensi.Masuk,tbltimesheet.NIP,tbltimesheet.StatusDivisi, karyawan.Nama,tbltimesheet.StatusPM,Absensi.Keluar,Divisi.NM_DIV, tbltimesheet.Hari,tbltimesheet.Hari,tbltimesheet.TotalKerja,tbltimesheet.ProjectUmum"
    StrSQL = StrSQL & " FROM tbltimesheet INNER JOIN karyawan ON tbltimesheet.NIP = karyawan.NIP INNER JOIN Divisi ON karyawan.Kd_Divisi = Divisi.KD_DIV INNER JOIN Absensi ON tbltimesheet.NIP = Absensi.NIP AND tbltimesheet.Tanggal = Absensi.Tgl"
    StrSQL = StrSQL & " Where tbltimesheet.Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And '" & Format(DTPicker2, "MM/dd/yyyy") & "'  And tbltimesheet.Status ='Actual' And  tbltimesheet.NoProject <> ''"
    If Trim(cboFlex) <> "" Then StrSQL = StrSQL & " AND tbltimesheet.NoProject = '" & cboFlex & "'"
    If Trim(CboKaryawan) <> "" Then StrSQL = StrSQL & " AND tbltimesheet.Nip = '" & CboKaryawan & "'"
    If Trim(CboDivisi) <> "" Then StrSQL = StrSQL & " And tbltimesheet.Kd_Divisi = '" & CboDivisi & "'"
    StrSQL = StrSQL & " Order By tbltimesheet.NIP,Divisi.NM_DIV,tbltimesheet.NoProject,tbltimesheet.Tanggal,tbltimesheet.IDTimesheet ASC"
    RsTS.Open StrSQL, CN, adOpenStatic
    Set .DataSource = RsTS
  
    .ColDataType(2) = flexDTDate
 
    For Lrow = 1 To .Rows - 1
        .TextMatrix(Lrow, 0) = Lrow
        .TextMatrix(0, 0) = .Rows
        If .TextMatrix(Lrow, 3) = "" Then .TextMatrix(Lrow, 3) = "00:00"
         If .TextMatrix(Lrow, 4) = "" Then .TextMatrix(Lrow, 4) = "00:00"
            Command1.Caption = .TextMatrix(Lrow, 10)
                        JmlLoop = 0
                        Jam1 = CDate(.TextMatrix(Lrow, 3))
                        Jam2 = CDate(.TextMatrix(Lrow, 4))
                        DTPicker3.Value = Jam1
                        Do Until JmlLoop = 50
                                If Format(DTPicker3.Value, "HH:mm") = Format(Jam2, "HH:mm") Then Exit Do
                                DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                                VSFg.Rows = VSFg.Rows + 1
                                VSFg.TextMatrix(VSFg.Rows - 1, 0) = VSFg.Rows - 1
                                IDWaktu = Format(DTPicker3.Value, "HH:mm")
                                VSFg.TextMatrix(VSFg.Rows - 1, 1) = IDWaktu
                                VSFg.TextMatrix(VSFg.Rows - 1, 2) = .TextMatrix(Lrow, 2) 'Format(.TextMatrix(Lrow, 2), "dd/MM/yyyy")
                                VSFg.TextMatrix(VSFg.Rows - 1, 3) = .TextMatrix(Lrow, 6)
                                VSFg.TextMatrix(VSFg.Rows - 1, 4) = .TextMatrix(Lrow, 10)
                                VSFg.TextMatrix(VSFg.Rows - 1, 5) = .TextMatrix(Lrow, 5)
                                VSFg.TextMatrix(VSFg.Rows - 1, 6) = Format(.TextMatrix(Lrow, 3), "HH:mm")
                                VSFg.TextMatrix(VSFg.Rows - 1, 7) = .TextMatrix(Lrow, 7)
                                VSFg.TextMatrix(VSFg.Rows - 1, 8) = .TextMatrix(Lrow, 11)
                                VSFg.TextMatrix(VSFg.Rows - 1, 9) = .TextMatrix(Lrow, 12)
                                VSFg.TextMatrix(VSFg.Rows - 1, 10) = .TextMatrix(Lrow, 1)
                                VSFg.TextMatrix(VSFg.Rows - 1, 11) = .TextMatrix(Lrow, 4)
                                VSFg.TextMatrix(VSFg.Rows - 1, 12) = .TextMatrix(Lrow, 15)
                                VSFg.TextMatrix(VSFg.Rows - 1, 13) = .TextMatrix(Lrow, 16)
                                VSFg.TextMatrix(VSFg.Rows - 1, 14) = .TextMatrix(Lrow, 17)
                                VSFg.TextMatrix(VSFg.Rows - 1, 15) = StrTotalJam '.TextMatrix(Lrow, 18)
                                VSFg.TextMatrix(VSFg.Rows - 1, 16) = .TextMatrix(Lrow, 19)
                                Jam3 = Format(.TextMatrix(Lrow, 9), "HH:mm")
        
                                Jam4 = DateAdd("h", 9, Jam3)
                                If Trim(.TextMatrix(Lrow, 17)) = "Kerja" Then
                                    If (VSFg.TextMatrix(VSFg.Rows - 1, 1) >= "08:00" And VSFg.TextMatrix(VSFg.Rows - 1, 1) <= "17:00") Or (VSFg.TextMatrix(VSFg.Rows - 1, 1) < Jam4 And VSFg.TextMatrix(VSFg.Rows - 1, 1) >= "08:00") Then
                                         VSFg.TextMatrix(VSFg.Rows - 1, 7) = "Timesheet"
                                      End If
        
                                End If
                              VSFg.TextMatrix(VSFg.Rows - 1, 17) = Format(.TextMatrix(Lrow, 14), "HH:mm")
        '                        If Format(DTPicker3, "HH:mm") = "12:00" Then DTPicker3.Value = DateAdd("n", 60, DTPicker3)
                                JmlLoop = JmlLoop + 1
                        Loop
                       
    Next

        VSFg.ColFormat(6) = "HH:mm"
        VSFg.ColFormat(11) = "HH:mm"
'    For lCol = 1 To VSFg.Cols - 1
'        VSFg.TextMatrix(0, lCol) = lCol
'    Next
End With

'Disiplit Perhari
With VSJoin
.Rows = 1
.Cols = 20
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
.TextMatrix(.Rows - 1, 19) = "18. Makan"
.ColDataType(6) = flexDTDate
.ColDataType(7) = flexDTDate
.ColFormat(6) = "HH:mm"
.ColFormat(7) = "HH:mm"
'.ColFormat(10) = "#.##"
.ColFormat(15) = "#,###"
.ColFormat(16) = "#,###"

End With
With VSFg
    For Lrow = 1 To VSFg.Rows - 1
    Command1.Caption = .TextMatrix(Lrow, 4)
             J = JoinGrid(.TextMatrix(Lrow, 3), .TextMatrix(Lrow, 2), Lrow, .TextMatrix(Lrow, 4), .TextMatrix(Lrow, 7))
    Next
End With

With VSJoin
For Lrow = 1 To VSJoin.Rows - 1
    Command1.Caption = .TextMatrix(Lrow, 3)
    If Trim(.TextMatrix(Lrow, 5)) = "" Then .TextMatrix(Lrow, 5) = 0
    .TextMatrix(Lrow, 0) = Lrow
    .TextMatrix(0, 0) = .Rows
    
      If Trim(.TextMatrix(Lrow, 12)) = "Lembur" Then
       J = .TextMatrix(Lrow, 10)
       J = GetLembur(.TextMatrix(Lrow, 3), Lrow)
      End If

    J = HitungGaji(.TextMatrix(Lrow, 3), .TextMatrix(Lrow, 2), .TextMatrix(Lrow, 14), Lrow, .TextMatrix(Lrow, 12))
Next
'   If .Rows = 1 Then MsgBox "Data Tidak Ditemukan ", vbInformation
End With
'  Exit Sub
 With fg
Dim Gaji As Double
Dim tlembur, lembur As Double
Dim tPersen As Double
Dim RsCetak As New ADODB.Recordset
    .Rows = .Rows + 1
    .TextMatrix(.Rows - 1, 3) = "Grand Total"
    For Lrow = 1 To .Rows - 2
           .TextMatrix(Lrow, 5) = "???"
        If Rscek.State = adStateOpen Then Rscek.Close
        Rscek.Open "Select * From Project Where Kode = '" & .TextMatrix(Lrow, 4) & "'", CN, adOpenStatic
        If Not Rscek.EOF Then .TextMatrix(Lrow, 5) = Rscek!Nama
        Tjam = Tjam + CDbl(.TextMatrix(Lrow, 6))
        Gaji = Gaji + CDbl(.TextMatrix(Lrow, 8))
        If .TextMatrix(Lrow, 11) <> 0 Then .TextMatrix(Lrow, 7) = Round(.TextMatrix(Lrow, 6) / .TextMatrix(Lrow, 11) * 100, 2)
        tPersen = CDbl(tPersen) + CDbl(.TextMatrix(Lrow, 7))
        .TextMatrix(Lrow, 7) = .TextMatrix(Lrow, 7)
        .Row = Lrow
        .Col = 6
        .CellAlignment = flexAlignRightCenter
        
 
        tlembur = tlembur + CDbl(.TextMatrix(Lrow, 9))
 
        lembur = CDbl(lembur) + CDbl(.TextMatrix(Lrow, 10))
        .TextMatrix(Lrow, 0) = Lrow
'        StrSQL = GetSet(.TextMatrix(Lrow, 2), .TextMatrix(Lrow, 2))
        StrSQL = "Insert Into tblcetakts(Nama_Divisi,NIP,Nama,Kode_Project,Nama_Project,Total_Jam ,Total,Total_Biaya,Jam_Lembur,Biaya_Lembur,Tingkat,Uang_Makan)"
        StrSQL = StrSQL & "Values('" & .TextMatrix(Lrow, 1) & "','" & .TextMatrix(Lrow, 2) & "','" & .TextMatrix(Lrow, 3) & "','" & .TextMatrix(Lrow, 4) & "',"
        StrSQL = StrSQL & "'" & .TextMatrix(Lrow, 5) & "'," & CCur(.TextMatrix(Lrow, 6)) & "," & CCur(.TextMatrix(Lrow, 7)) & "," & CCur(.TextMatrix(Lrow, 8)) & "," & CCur(.TextMatrix(Lrow, 9)) & "," & CCur(.TextMatrix(Lrow, 10)) & ",'" & Tingkat & "'," & CCur(.TextMatrix(Lrow, 13)) & ")"
        CN.Execute StrSQL
    Next
    
    StrSQL = "Select * From tblcetakts Order By NIP"
    If RsCetak.State = adStateOpen Then RsCetak.Close
    RsCetak.Open StrSQL, CN, adOpenStatic
    Set VSFg.DataSource = RsCetak
    .Rows = 1
    .Rows = 2
    For Lrow = 1 To VSFg.Rows - 1
        If Lrow = 1 Then
            .TextMatrix(.Rows - 1, 1) = VSFg.TextMatrix(Lrow, 1)
            .TextMatrix(.Rows - 1, 2) = VSFg.TextMatrix(Lrow, 2) ' Gaji
            .TextMatrix(.Rows - 1, 3) = VSFg.TextMatrix(Lrow, 3)
            .TextMatrix(.Rows - 1, 4) = VSFg.TextMatrix(Lrow, 4)
            .TextMatrix(.Rows - 1, 5) = VSFg.TextMatrix(Lrow, 5)
            .TextMatrix(.Rows - 1, 6) = VSFg.TextMatrix(Lrow, 6)
            .TextMatrix(.Rows - 1, 7) = VSFg.TextMatrix(Lrow, 7) & "%"
            .TextMatrix(.Rows - 1, 8) = VSFg.TextMatrix(Lrow, 8)
            .TextMatrix(.Rows - 1, 9) = VSFg.TextMatrix(Lrow, 9)
            .TextMatrix(.Rows - 1, 10) = VSFg.TextMatrix(Lrow, 10)
            .TextMatrix(.Rows - 1, 13) = VSFg.TextMatrix(Lrow, 12)
            Tingkat = VSFg.TextMatrix(Lrow, 11)
            
            Tjam = 0
            tPersen = 0
            Gaji = 0
            tlembur = 0
            lembur = 0
            Uangmakan = 0
            Tjam = CCur(Tjam) + CCur(.TextMatrix(.Rows - 1, 6))
            tPersen = tPersen + CDbl(VSFg.TextMatrix(Lrow, 7))
            Gaji = Gaji + CDbl(VSFg.TextMatrix(Lrow, 8))
            tlembur = tlembur + CDbl(VSFg.TextMatrix(Lrow, 9))
            lembur = lembur + CDbl(VSFg.TextMatrix(Lrow, 10))
            Uangmakan = Uangmakan + CDbl(VSFg.TextMatrix(Lrow, 12))
        Else
            If VSFg.TextMatrix(Lrow, 2) <> VSFg.TextMatrix(Lrow - 1, 2) Then
               .Rows = .Rows + 1
                
                .TextMatrix(.Rows - 1, 3) = "TOTAL"
                .TextMatrix(.Rows - 1, 1) = VSFg.TextMatrix(Lrow, 1)
                .TextMatrix(.Rows - 1, 6) = Tjam
                .TextMatrix(.Rows - 1, 7) = tPersen & "%"
                .TextMatrix(.Rows - 1, 8) = Gaji
                .TextMatrix(.Rows - 1, 9) = tlembur
                .TextMatrix(.Rows - 1, 13) = Uangmakan
                .TextMatrix(.Rows - 1, 10) = lembur
'
                .TextMatrix(.Rows - 1, 10) = lembur
 
                .Row = .Rows - 1
                For x = 1 To .Cols - 1
                    .Col = x
                    .CellBackColor = vbGreen
                Next
                Tjam = 0
                tPersen = 0
                Gaji = 0
                tlembur = 0
                lembur = 0
                 Uangmakan = 0
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = VSFg.TextMatrix(Lrow, 1)
                .TextMatrix(.Rows - 1, 2) = VSFg.TextMatrix(Lrow, 2) ' Gaji
                .TextMatrix(.Rows - 1, 3) = VSFg.TextMatrix(Lrow, 3)
                .TextMatrix(.Rows - 1, 4) = VSFg.TextMatrix(Lrow, 4)
                .TextMatrix(.Rows - 1, 5) = VSFg.TextMatrix(Lrow, 5)
                .TextMatrix(.Rows - 1, 6) = VSFg.TextMatrix(Lrow, 6)
                .TextMatrix(.Rows - 1, 7) = VSFg.TextMatrix(Lrow, 7) & "%"
                .TextMatrix(.Rows - 1, 8) = VSFg.TextMatrix(Lrow, 8)
                .TextMatrix(.Rows - 1, 9) = VSFg.TextMatrix(Lrow, 9)
                .TextMatrix(.Rows - 1, 10) = VSFg.TextMatrix(Lrow, 10)
                 .TextMatrix(.Rows - 1, 13) = VSFg.TextMatrix(Lrow, 12)
                 
                Tjam = CDbl(Tjam) + CDbl(.TextMatrix(.Rows - 1, 6))
                tPersen = tPersen + CDbl(VSFg.TextMatrix(Lrow, 7))
                 Gaji = Gaji + CDbl(VSFg.TextMatrix(Lrow, 8))
                tlembur = tlembur + CDbl(VSFg.TextMatrix(Lrow, 9))
                lembur = lembur + CDbl(VSFg.TextMatrix(Lrow, 10))
                Uangmakan = Uangmakan + CDbl(VSFg.TextMatrix(Lrow, 12))
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = VSFg.TextMatrix(Lrow, 1)
                .TextMatrix(.Rows - 1, 2) = VSFg.TextMatrix(Lrow, 2) ' Gaji
                .TextMatrix(.Rows - 1, 3) = VSFg.TextMatrix(Lrow, 3)
                .TextMatrix(.Rows - 1, 4) = VSFg.TextMatrix(Lrow, 4)
                .TextMatrix(.Rows - 1, 5) = VSFg.TextMatrix(Lrow, 5)
                .TextMatrix(.Rows - 1, 6) = VSFg.TextMatrix(Lrow, 6)
                .TextMatrix(.Rows - 1, 7) = VSFg.TextMatrix(Lrow, 7) & "%"
                .TextMatrix(.Rows - 1, 8) = VSFg.TextMatrix(Lrow, 8)
                .TextMatrix(.Rows - 1, 9) = VSFg.TextMatrix(Lrow, 9)
                .TextMatrix(.Rows - 1, 10) = VSFg.TextMatrix(Lrow, 10)
                 .TextMatrix(.Rows - 1, 13) = VSFg.TextMatrix(Lrow, 12)
                 Tingkat = VSFg.TextMatrix(Lrow, 11)
                Tjam = CDbl(Tjam) + CDbl(.TextMatrix(.Rows - 1, 6))
                tPersen = tPersen + CDbl(VSFg.TextMatrix(Lrow, 7))
                 Gaji = Gaji + CDbl(VSFg.TextMatrix(Lrow, 8))
            tlembur = tlembur + CDbl(VSFg.TextMatrix(Lrow, 9))
            lembur = lembur + CDbl(VSFg.TextMatrix(Lrow, 10))
            Uangmakan = Uangmakan + CDbl(VSFg.TextMatrix(Lrow, 12))
            End If
        End If
    Next
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 3) = "TOTAL"
        .TextMatrix(.Rows - 1, 6) = Tjam
        .TextMatrix(.Rows - 1, 7) = tPersen & "%"
        .TextMatrix(.Rows - 1, 8) = Gaji
        .TextMatrix(.Rows - 1, 9) = tlembur
        .TextMatrix(.Rows - 1, 13) = Uangmakan
        .TextMatrix(.Rows - 1, 10) = lembur
 
        .Row = .Rows - 1
        For x = 1 To .Cols - 1
            .Col = x
            .CellBackColor = vbGreen
        Next
 
End With
Command1.Caption = "Refresh"
Command1.Enabled = True
Exit Sub
Adaerror:
   MsgBox "Silahkan Ulangi Lagi", vbCritical
End Sub

Function HitungGaji(NIP As String, Project As String, Divisi As String, Row As Integer, Keterangan As String)
Dim StatusProject As Boolean
Dim FRow As Integer
Dim HRow As Integer
On Error GoTo Adaerror
With fg
    StatusProject = False
    For FRow = 0 To .Rows - 1
    If Trim(.TextMatrix(FRow, 1)) = Divisi And Trim(.TextMatrix(FRow, 2)) = NIP And Trim(.TextMatrix(FRow, 4)) = Project Then
           StatusProject = True
           HRow = FRow
        End If
    Next
        If StatusProject = False Then
             .Rows = .Rows + 1
             HRow = .Rows - 1
            .TextMatrix(HRow, 6) = 0
            .TextMatrix(HRow, 7) = 0
            .TextMatrix(HRow, 8) = 0
            .TextMatrix(HRow, 9) = 0
            .TextMatrix(HRow, 10) = 0
            .TextMatrix(HRow, 13) = 0
        End If
            .TextMatrix(HRow, 1) = VSJoin.TextMatrix(Row, 14)
            .TextMatrix(HRow, 2) = VSJoin.TextMatrix(Row, 3)
            .TextMatrix(HRow, 3) = VSJoin.TextMatrix(Row, 4)
            .TextMatrix(HRow, 4) = VSJoin.TextMatrix(Row, 2)
            .TextMatrix(HRow, 5) = VSJoin.TextMatrix(Row, 8)
            .TextMatrix(HRow, 11) = VSJoin.TextMatrix(Row, 17)
             .TextMatrix(HRow, 12) = VSJoin.TextMatrix(Row, 18)
             If VSJoin.TextMatrix(Row, 19) = "" Then VSJoin.TextMatrix(Row, 19) = 0
              .TextMatrix(HRow, 13) = CCur(.TextMatrix(HRow, 13)) + VSJoin.TextMatrix(Row, 19)
            If Trim(Keterangan) = "Timesheet" Then
                .TextMatrix(HRow, 6) = CCur(.TextMatrix(HRow, 6)) + CCur(VSJoin.TextMatrix(Row, 5))
                .TextMatrix(HRow, 8) = CCur(.TextMatrix(HRow, 8)) + VSJoin.TextMatrix(Row, 15)
             Else
                .TextMatrix(HRow, 9) = CDbl(.TextMatrix(HRow, 9)) + CDbl(VSJoin.TextMatrix(Row, 10))
                .TextMatrix(HRow, 10) = CCur(.TextMatrix(HRow, 10)) + VSJoin.TextMatrix(Row, 16)
            End If
End With
Exit Function
Adaerror:
    MsgBox err.Description
End Function
Function JoinGrid(ByVal Project As String, ByVal tgl1 As String, ByVal Row As Integer, ByVal NIP As String, ByVal Keterangan As String)
Dim StatusNIP As Boolean
Dim FRow As Integer
Dim J As String
Dim JamTs As Date
On Error GoTo Adaerror
With VSJoin
    StatusNIP = False
   If Format(VSFg.TextMatrix(Row, 1), "HH:mm") = "12:30" Or Format(VSFg.TextMatrix(Row, 1), "HH:mm") = "13:00" Then Exit Function
'    For FRow = 0 To .Rows - 1
'        If .TextMatrix(FRow, 2) = Project And .TextMatrix(FRow, 1) = tgl1 And .TextMatrix(FRow, 3) = NIP And .TextMatrix(FRow, 12) = Keterangan Then
'           StatusNIP = True
'           Exit For
'        End If
'    Next
        If .TextMatrix(.Rows - 1, 2) = Project And .TextMatrix(.Rows - 1, 1) = tgl1 And .TextMatrix(.Rows - 1, 3) = NIP And .TextMatrix(.Rows - 1, 12) = Keterangan Then
           StatusNIP = True
'           Exit For
        End If
        If StatusNIP = False Then
             .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 5) = 0
            .TextMatrix(.Rows - 1, 10) = 0
        End If
            JamTs = VSFg.TextMatrix(Row, 1)
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
            .TextMatrix(.Rows - 1, 11) = VSFg.TextMatrix(Row, 8)
            .TextMatrix(.Rows - 1, 12) = VSFg.TextMatrix(Row, 7)
            .TextMatrix(.Rows - 1, 13) = VSFg.TextMatrix(Row, 14) 'Hari
            .TextMatrix(.Rows - 1, 14) = VSFg.TextMatrix(Row, 12)
            If .TextMatrix(.Rows - 2, 3) <> .TextMatrix(.Rows - 1, 3) Then Call GetSet(.TextMatrix(.Rows - 2, 3), .TextMatrix(.Rows - 1, 3))
            .TextMatrix(.Rows - 1, 15) = 0
            If VSFg.TextMatrix(Row, 15) > 0 Then .TextMatrix(.Rows - 1, 15) = Round(CCur(((NilaiGaji) / VSFg.TextMatrix(Row, 15))) * CDbl(.TextMatrix(.Rows - 1, 5)), 2)
            .TextMatrix(.Rows - 1, 16) = 0
            .TextMatrix(.Rows - 1, 17) = VSFg.TextMatrix(Row, 15)
            .TextMatrix(.Rows - 1, 18) = VSFg.TextMatrix(Row, 16)
End With
Exit Function
Adaerror:
    MsgBox err.Description

End Function
Function GetLembur(ByVal NIP As String, Lrow As Integer)
Dim jamMasuk, Hari As String
Dim JamKeluar As String
Dim TotalJam, TotalLemburBruto As String
Dim TotalLembur, Terlambat As String
Dim Gapok As Double
Dim rsAbsen As New ADODB.Recordset
Dim RsGaji As New ADODB.Recordset
Dim RsSetting As New ADODB.Recordset
Dim Istirahat As String
Dim Jam1, Jam2, Jam6 As Double
Dim Jam3, Jam4, Jam5 As Double
Dim JamMakan, UpahJam As Double
Dim Transport As String

Dim Upah1, Upah2, Upah3, UM As Currency
Dim JamLembur, SplitLembur As Double
Dim NilaiUpah As Currency, Menitlembur As Currency

With VSJoin
         TotalLembur = 0
         Istirahat = 0
         TotalLemburBruto = 0
         UM = 0
         JamLembur = 0
         Hari = .TextMatrix(Lrow, 13)
         DTPicker3 = Format(.TextMatrix(Lrow, 6), "HH:mm")
         DTPicker4 = Format(.TextMatrix(Lrow, 7), "HH:mm")
         If Format(DTPicker3, "HH:mm") = "00:00" Then
                jamMasuk = "TIDAK ABSEN"
             Else
                jamMasuk = Format(DTPicker3, "HH:mm")
             End If
             
             If Format(DTPicker4, "HH:mm") = "00:00" Then
                JamKeluar = "TIDAK ABSEN"
                TotalJam = 0
            Else
                JamKeluar = Format(DTPicker4, "HH:mm")
                TotalJam = .TextMatrix(Lrow, 10)
        End If
       
               Transport = GetSet(NIP, NIP)
         If RsSetting.State = adStateOpen Then RsSetting.Close
         RsSetting.Open "SELECT * FROM TblSETTING WHERE TINGKAT = '" & Tingkat & "' AND HARI = '" & Hari & "' AND StatusAktif = 1 And Berlaku_SD >= '" & Format(Date, "mm/dd/yyyy") & "'", CN, adOpenStatic
         If Not RsSetting.EOF Then
             Jam1 = RsSetting!jamlembur1
             Jam2 = RsSetting!jamlembur2
             Jam3 = RsSetting!jamlembur3
             Jam4 = RsSetting!jamlembur4
             UM = RsSetting!upahmakan
             JamMakan = RsSetting!jammakan1
            If JamKeluar = "TIDAK ABSEN" Then
                TotalLembur = 0
             Else
                Dim Jam As Integer
                Dim Menit As Integer

                  Select Case Hari
                      Case "Kerja"
                            TotalLemburBruto = TotalJam
                            TotalLembur = TotalJam
                      Case "Libur"
                            TotalLembur = TotalJam
                            TotalLemburBruto = TotalJam
                  End Select
 
                  TotalLembur = Replace(TotalLembur, ":", ".")
                  TotalLembur = CDbl(TotalLembur)
                  If TotalLembur >= RsSetting!ist1 And TotalLembur <= RsSetting!ist2 Then
                      Istirahat = 0.5
                      TotalLembur = TotalLembur - Istirahat
 
                   ElseIf TotalLembur >= RsSetting!ist3 And TotalLembur <= RsSetting!ist4 Then
                      Istirahat = 1
                      TotalLembur = TotalLembur - Istirahat
 
                    ElseIf TotalLembur >= RsSetting!ist5 And TotalLembur <= RsSetting!ist6 Then
                      Istirahat = 1.5
                      TotalLembur = TotalLembur - Istirahat
 
                    ElseIf TotalLembur >= RsSetting!ist7 And TotalLembur <= RsSetting!ist8 Then
                      Istirahat = 2
                      TotalLembur = TotalLembur - Istirahat
                    
                    End If
            End If
         End If
         If Left(TotalLembur, 2) < JamMakan Then UM = 0
         
         'pembulatan menit
         If TotalLembur > 0 Then
'             DTPicker4 = TotalLembur
              'upah
              JamLembur = Replace(TotalLembur, ":", ".")
               Upah1 = RsSetting!Upah1
               Upah2 = RsSetting!Upah2
               Upah3 = RsSetting!Upah3
                 UpahJam = RsSetting!UpahPerjam
             UM = RsSetting!upahmakan
             JamMakan = RsSetting!jammakan1
               NilaiUpah = 0
'             JamLembur = "09.30"
                 If Hari = "Kerja" Then
                        If CDbl(JamLembur) >= Jam1 And CDbl(JamLembur) <= Jam2 Then
                            NilaiUpah = (CDbl(JamLembur) * NilaiGaji / UpahJam) * Upah1
'                            .TextMatrix(Grow, 15) = Upah1
                        ElseIf CDbl(JamLembur) >= Jam3 And CDbl(JamLembur) <= Jam4 Then
                            NilaiUpah = (Jam2 * NilaiGaji / UpahJam) * Upah1
                            SplitLembur = CDbl(JamLembur) - Jam2
                            If SplitLembur > 0 Then
                                NilaiUpah = NilaiUpah + (SplitLembur * NilaiGaji / UpahJam) * Upah2
                            End If
'                            .TextMatrix(Grow, 15) = Upah2
                        End If
                Else
'                         JamLembur = "10.30"
                        If CDbl(JamLembur) >= Jam1 And CDbl(JamLembur) <= Jam2 Then
                            NilaiUpah = (CDbl(JamLembur) * NilaiGaji / UpahJam) * Upah1
'                            .TextMatrix(Grow, 15) = Upah1
                        ElseIf CDbl(JamLembur) >= Jam3 And CDbl(JamLembur) <= Jam4 Then
                            NilaiUpah = (Jam2 * NilaiGaji / UpahJam) * Upah1
                            SplitLembur = CDbl(JamLembur) - 8
                            NilaiUpah = NilaiUpah + (SplitLembur * NilaiGaji / UpahJam) * Upah2
'                            .TextMatrix(Grow, 15) = Upah2
                         ElseIf CDbl(JamLembur) >= Jam5 And CDbl(JamLembur) <= Jam6 Then
                            NilaiUpah = (Jam2 * NilaiGaji / UpahJam) * Upah1
                            SplitLembur = CDbl(JamLembur) - 8
                            NilaiUpah = NilaiUpah + (1 * NilaiGaji / UpahJam) * Upah2
                            SplitLembur = CDbl(JamLembur) - 8 - 1
                            NilaiUpah = NilaiUpah + (SplitLembur * NilaiGaji / UpahJam) * Upah3
'
                        End If
                End If
                .TextMatrix(Lrow, 10) = TotalLembur
                .TextMatrix(Lrow, 16) = NilaiUpah
                If TotalLembur >= JamMakan Then
                    .TextMatrix(Lrow, 19) = UM
                End If
         End If
        
End With
End Function
Function GetSet(Oldnip As String, NIP As String) As String
Dim i As Integer
GetSet = False
'If Oldnip = NIP Then Exit Function
With VSGaji
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 25) = "" Then .TextMatrix(i, 25) = 0
        If .TextMatrix(i, 6) = "" Then .TextMatrix(i, 6) = 0
        If Trim(.TextMatrix(i, 2)) = Trim(NIP) Then
            GetSet = True
            NilaiGaji = .TextMatrix(i, 25)
            Tingkat = .TextMatrix(i, 6)
            Exit For
        Else
             NilaiGaji = 0
             Tingkat = 0
        End If
    Next
End With
End Function
  

Private Sub VSGaji_DblClick()
MsgBox VSGaji.Col
End Sub
