VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLapBiayaProject 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Rekap Biaya Project"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10155
   ScaleWidth      =   13410
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   13410
      TabIndex        =   10
      Top             =   9660
      Width           =   13410
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
         TabIndex        =   11
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2385
      ScaleWidth      =   13380
      TabIndex        =   0
      Top             =   0
      Width           =   13410
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   5640
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   53936130
         CurrentDate     =   39940
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Print out"
         Height          =   375
         Left            =   8400
         TabIndex        =   24
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   6960
         TabIndex        =   23
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Export To Sheet"
         Height          =   375
         Left            =   8400
         TabIndex        =   22
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Before Sheet"
         Height          =   375
         Left            =   10080
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Open XLS"
         Height          =   375
         Left            =   6960
         TabIndex        =   20
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Next Sheet"
         Height          =   375
         Left            =   10080
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4440
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   720
         TabIndex        =   1
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
         Format          =   53936131
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2880
         TabIndex        =   2
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
         Format          =   53936131
         CurrentDate     =   39931
      End
      Begin VSFlex8Ctl.VSFlexGrid cboFlex 
         Height          =   315
         Left            =   3720
         TabIndex        =   3
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
         FormatString    =   $"frmLapBiayaProject.frx":0000
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
         TabIndex        =   4
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
         FormatString    =   $"frmLapBiayaProject.frx":0029
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
         Left            =   5640
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   53936130
         CurrentDate     =   39940
      End
      Begin VSFlex8Ctl.VSFlexGrid CboDivisi 
         Height          =   315
         Left            =   720
         TabIndex        =   15
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
         FormatString    =   $"frmLapBiayaProject.frx":0052
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
      Begin VB.Label Label10 
         Caption         =   "* Biaya Lembur = (Nilai Gaji/173) * Total Jam Lembur * Upah Perjam"
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
         Left            =   2760
         TabIndex        =   31
         Top             =   2040
         Width           =   8895
      End
      Begin VB.Label Label9 
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
         Left            =   2760
         TabIndex        =   30
         Top             =   1800
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
         Left            =   2520
         TabIndex        =   29
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "* Total Jam Kerja Akan Tampil Apabila Sudah Di Verifikasi Divisi"
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
         Left            =   2760
         TabIndex        =   28
         Top             =   1320
         Width           =   5775
      End
      Begin VB.Label Label7 
         Caption         =   "* Untuk Telat, Tidak Mengisi dan Belum Diverifikasi PM / Divisi Total Jam Dimasukan ke Project Umum"
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
         Left            =   2760
         TabIndex        =   27
         Top             =   1560
         Width           =   8895
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
         TabIndex        =   16
         Top             =   600
         Width           =   495
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
         TabIndex        =   9
         Top             =   600
         Width           =   1455
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   120
         Width           =   375
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
         TabIndex        =   6
         Top             =   960
         Width           =   495
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "frmLapBiayaProject.frx":007B
      Top             =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3255
      Left            =   -120
      TabIndex        =   17
      Top             =   2400
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
      FormatString    =   $"frmLapBiayaProject.frx":02AF
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
      Height          =   2775
      Left            =   8400
      TabIndex        =   18
      Top             =   5640
      Width           =   6735
      _cx             =   11880
      _cy             =   4895
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
      FormatString    =   $"frmLapBiayaProject.frx":038E
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
      Height          =   2415
      Left            =   0
      TabIndex        =   13
      Top             =   5760
      Width           =   8295
      _cx             =   14631
      _cy             =   4260
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
      FormatString    =   $"frmLapBiayaProject.frx":046D
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
      TabIndex        =   12
      Top             =   1440
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
      FormatString    =   $"frmLapBiayaProject.frx":054C
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
   Begin VSFlex8Ctl.VSFlexGrid VSGaji 
      Height          =   1455
      Left            =   0
      TabIndex        =   25
      ToolTipText     =   "Double Klik Kolom Project Untuk Melihat PM"
      Top             =   2400
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
      FormatString    =   $"frmLapBiayaProject.frx":062B
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
Attribute VB_Name = "frmLapBiayaProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NilaiGaji As Currency
Dim Tingkat As Integer
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
Tanggal = Format(DTPicker1, "MM")
If Mid(FileTitle, 6, 2) <> Tanggal Then
    If MsgBox("File Data Gaji Tidak Sama Dengan Tanggal Pencarian, Apakah Anda Akan Melanjutkan Proses ini ?", vbQuestion + vbYesNo, "Konfirmasi hapus") = vbNo Then
           Exit Sub
    End If
End If
fg.Visible = True
VSGaji.Visible = False
Command1.Enabled = False
Setgrid
Showdata
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Dim x As String
If fg.Rows > 1 Then
   fg.SaveGrid "C:\BiayaPerProject.xls", flexFileExcel, True
'    Shell PathOffice & "C:\BiayaPerProject.csv", vbNormalFocus
x = ShellExecute(Me.hwnd, "open", "C:\BiayaPerproject.xls", vbNullString, "C:\BiayaPerProject.xls", 1)
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
If fg.Rows > 1 Then fg.PrintGrid "Biaya Per Project - Periode " & DTPicker1.Value & " S/D " & DTPicker2.Value, 2, 2, 900, 500

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
    dlg.Filter = "File Gaji (*.Csv)|*.Csv"
    dlg.DialogTitle = "Data Gaji Karyawan "
    dlg.ShowOpen
    If Len(dlg.FileName) = 0 Then Exit Sub
    FileTitle = dlg.FileTitle
    MousePointer = MousePointerConstants.vbHourglass
    VSGaji.LoadGrid dlg.FileName, flexFileCommaText
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
.Cols = 13
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
'.TextMatrix(0, 13) = "TempTotal"
'.TextMatrix(0, 14) = "TempPersen"
.ColFormat(8) = "#,###"
'.ColFormat(9) = "#,###"
.ColFormat(10) = "#,###"
.ColWidth(6) = 1000
.ColWidth(10) = 1000
.ColWidth(7) = 1000
.ColWidth(8) = 1200
.ColWidth(9) = 1200
.ColWidth(11) = 0
.ColWidth(12) = 0
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
    If Trim(CboDivisi) = "" Then
        StrSQL = "select Kode,Nama from project " & _
             "group by kode,Nama " & _
             "order by kode"
    Else
        StrSQL = "select Kode,Nama from project " & _
             "where  Kd_Divisi = '" & CboDivisi & "'" & _
             "group by kode,Nama " & _
             "order by kode"
    End If
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
             .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture2.Height

        End With
       
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmLapBiayaProject = Nothing
End Sub
Sub Showdata()
Dim x As Integer
Dim i, J As Integer
Dim Split, JmlLoop As Integer
Dim Jam, Hari As String
Dim IDWaktu As String
Dim Jam1, Jam2 As Date
Dim Jam3, Jam4 As Date
Dim TglAwal, TglAkhir As Date
Dim RsTS As New ADODB.Recordset
Dim TotalHari, TotalJam As Integer
 
Command1.Enabled = False
With VSFlexGrid1
        .Rows = 1
        VSFg.Rows = 1
        VSFg.Cols = 17
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
     
    TotalJam = TotalHari * 8
    
    If RsTS.State = adStateOpen Then RsTS.Close
    StrSQL = "SELECT tbltimesheet.IDtimesheet,tbltimesheet.Tanggal,tbltimesheet.JamAwal As [Jam Awal],tbltimesheet.JamAkhir AS [Jam Akhir],tbltimesheet.Status,tbltimesheet.NoProject As Project,tbltimesheet.Keterangan,tbltimesheet.Tanggal,Absensi.Masuk,tbltimesheet.NIP,tbltimesheet.StatusDivisi, karyawan.Nama,tbltimesheet.StatusPM,Absensi.Keluar,Divisi.NM_DIV, tbltimesheet.Hari,tbltimesheet.Hari,tbltimesheet.TotalKerja,tbltimesheet.ProjectUmum"
    StrSQL = StrSQL & " FROM tbltimesheet INNER JOIN karyawan ON tbltimesheet.NIP = karyawan.NIP INNER JOIN Divisi ON karyawan.Kd_Divisi = Divisi.KD_DIV INNER JOIN Absensi ON tbltimesheet.NIP = Absensi.NIP AND tbltimesheet.Tanggal = Absensi.Tgl"
    StrSQL = StrSQL & " Where tbltimesheet.Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And '" & Format(DTPicker2, "MM/dd/yyyy") & "'  And tbltimesheet.Status ='Actual' And  tbltimesheet.NoProject <> '' "
    If Trim(cboFlex) <> "" Then StrSQL = StrSQL & " AND tbltimesheet.NoProject = '" & cboFlex & "'"
    If Trim(CboKaryawan) <> "" Then StrSQL = StrSQL & " AND tbltimesheet.Nip = '" & CboKaryawan & "'"
    If Trim(CboDivisi) <> "" Then StrSQL = StrSQL & " And tbltimesheet.Kd_Divisi = '" & CboDivisi & "'"
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
                     If Format(DTPicker3.Value, "HH:mm") = Format(Jam2, "HH:mm") Then Exit Do
                        DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                        VSFg.Rows = VSFg.Rows + 1
'                     VSFg.Rows = VSFg.Rows + 1
'                        DTPicker3.Value = DateAdd("n", 30, DTPicker3)
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
                        VSFg.TextMatrix(VSFg.Rows - 1, 15) = TotalJam '.TextMatrix(Lrow, 18)
                        VSFg.TextMatrix(VSFg.Rows - 1, 16) = .TextMatrix(Lrow, 19)
                        Jam3 = Format(.TextMatrix(Lrow, 9), "HH:mm")

                        Jam4 = DateAdd("h", 9, Jam3)
                        If Trim(.TextMatrix(Lrow, 17)) = "Kerja" Then
                            If (VSFg.TextMatrix(VSFg.Rows - 1, 1) >= "08:00" And VSFg.TextMatrix(VSFg.Rows - 1, 1) <= "17:00") Or (VSFg.TextMatrix(VSFg.Rows - 1, 1) < Jam4 And VSFg.TextMatrix(VSFg.Rows - 1, 1) >= "08:00") Then
                                 VSFg.TextMatrix(VSFg.Rows - 1, 7) = "Timesheet"
                              End If

                        End If
'                         If Format(DTPicker3.Value, "HH:mm") = Format(Jam2, "HH:mm") Then Exit Do
'                        If Format(DTPicker3, "HH:mm") = "12:00" Then DTPicker3.Value = DateAdd("n", 60, DTPicker3)
                        JmlLoop = JmlLoop + 1
                Loop
    Next
        VSFg.ColFormat(6) = "HH:mm"
        VSFg.ColFormat(11) = "HH:mm"
    For lCol = 1 To VSFg.Cols - 1
        VSFg.TextMatrix(0, lCol) = lCol
    Next
End With

'Disiplit Perhari
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
With VSFg
    For Lrow = 1 To VSFg.Rows - 1
             J = JoinGrid(.TextMatrix(Lrow, 3), .TextMatrix(Lrow, 2), Lrow, .TextMatrix(Lrow, 4), .TextMatrix(Lrow, 7))
    Next
End With

With VSJoin
For Lrow = 1 To VSJoin.Rows - 1
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

 With fg
Dim Tjam, Gaji As Double
Dim tlembur, lembur As Double
Dim tPersen As Double
 
    .Rows = .Rows + 1
    For Lrow = 1 To .Rows - 2
        .TextMatrix(Lrow, 5) = "???"
        If Rscek.State = adStateOpen Then Rscek.Close
        Rscek.Open "Select * From Project Where Kode = '" & .TextMatrix(Lrow, 4) & "'", CN, adOpenStatic
        If Not Rscek.EOF Then .TextMatrix(Lrow, 5) = Rscek!Nama
        Tjam = Tjam + CDbl(.TextMatrix(Lrow, 6))
        Gaji = Gaji + CDbl(.TextMatrix(Lrow, 8))
        .TextMatrix(Lrow, 7) = Round(.TextMatrix(Lrow, 6) / .TextMatrix(Lrow, 11) * 100, 2)
        tPersen = CDbl(tPersen) + CDbl(.TextMatrix(Lrow, 7))
        .TextMatrix(Lrow, 7) = .TextMatrix(Lrow, 7) & "%"
        .Row = Lrow
        .Col = 6
        .CellAlignment = flexAlignRightCenter
        
'        .TextMatrix(Lrow, 9) = Round(.TextMatrix(Lrow, 9) / .TextMatrix(Lrow, 11) * 100, 2)
        tlembur = tlembur + CDbl(.TextMatrix(Lrow, 9))
'         .TextMatrix(Lrow, 9) = .TextMatrix(Lrow, 9) & "%"
        lembur = CDbl(lembur) + CDbl(.TextMatrix(Lrow, 10))
        .TextMatrix(Lrow, 0) = Lrow
    Next
    If .TextMatrix(1, 1) <> "" Then
     
    .TextMatrix(.Rows - 1, 6) = Tjam
    .TextMatrix(.Rows - 1, 7) = tPersen & "%" ' Gaji
    .TextMatrix(.Rows - 1, 9) = tlembur
    .TextMatrix(.Rows - 1, 8) = Gaji
     .TextMatrix(.Rows - 1, 10) = lembur
    End If
    Dim SisaTotal As Double
    Dim StatusSisa As Boolean
    If Trim(CboKaryawan) <> "" And Trim(cboFlex) = "" Then
          StrSQL = "SELECT tbltimesheet.NIP, tbltimesheet.kd_divisi, Karyawan.Nama, "
            StrSQL = StrSQL & " tbltimesheet.TotalKerja, tbltimesheet.ProjectUmum,Divisi.NM_DIV,Project.Nama AS NamaProject FROM Divisi INNER JOIN"
            StrSQL = StrSQL & " tbltimesheet ON Divisi.KD_DIV = tbltimesheet.kd_divisi INNER JOIN Karyawan ON tbltimesheet.NIP = Karyawan.NIP INNER JOIN Project ON tbltimesheet.ProjectUmum = Project.Kode Where tbltimesheet.NIP = '" & CboKaryawan & "' And Tanggal = '" & Format(DTPicker1, "MM/dd/yyyy") & "'"
            If Rscek.State = adStateOpen Then Rscek.Close
            Rscek.Open StrSQL, CN, adOpenStatic
        If VSJoin.Rows = 1 Then

            If Not Rscek.EOF Then
                .TextMatrix(.Rows - 1, 1) = Rscek!NM_DIV 'NamaDivisi
                .TextMatrix(.Rows - 1, 4) = Rscek!ProjectUmum
                .TextMatrix(.Rows - 1, 5) = Rscek!NamaProject
                .TextMatrix(.Rows - 1, 2) = Rscek!NIP
                .TextMatrix(.Rows - 1, 3) = Rscek!Nama
                .TextMatrix(.Rows - 1, 6) = Rscek!TotalKerja
                .TextMatrix(.Rows - 1, 7) = "100%"
                .TextMatrix(.Rows - 1, 9) = "0"
                .TextMatrix(.Rows - 1, 10) = "0"
                .TextMatrix(.Rows - 1, 8) = GetSet(.TextMatrix(.Rows - 2, 2), .TextMatrix(.Rows - 1, 2)) = True
                .TextMatrix(.Rows - 1, 8) = Round(CCur(((NilaiGaji) / .TextMatrix(.Rows - 1, 6))) * CDbl(.TextMatrix(.Rows - 1, 6)), 2)
                Exit Sub
            End If
        End If
        .TextMatrix(.Rows - 1, 6) = Trim(.TextMatrix(1, 11))
        .TextMatrix(.Rows - 1, 7) = "100%"
        SisaTotal = .TextMatrix(1, 11) - Tjam

       If SisaTotal > 0 Then
       StatusSisa = False
          For Lrow = 1 To .Rows - 2
             If Trim(.TextMatrix(Lrow, 4)) = Trim(.TextMatrix(Lrow, 12)) Then
                StatusSisa = True
                Exit For
            End If
          Next
          If StatusSisa = True Then
                .TextMatrix(Lrow, 6) = .TextMatrix(Lrow, 6) + SisaTotal
                .TextMatrix(Lrow, 7) = Round(.TextMatrix(Lrow, 6) / .TextMatrix(Lrow, 11) * 100, 2) & "%"
          Else
                .TextMatrix(.Rows - 1, 1) = Rscek!NM_DIV 'NamaDivisi
                .TextMatrix(.Rows - 1, 4) = Rscek!ProjectUmum
                .TextMatrix(.Rows - 1, 5) = Rscek!NamaProject
                .TextMatrix(.Rows - 1, 2) = Rscek!NIP
                .TextMatrix(.Rows - 1, 3) = Rscek!Nama
                .TextMatrix(.Rows - 1, 6) = SisaTotal
                .TextMatrix(.Rows - 1, 7) = "100%"
                .TextMatrix(.Rows - 1, 9) = "0"
                .TextMatrix(.Rows - 1, 10) = "0"
                .TextMatrix(.Rows - 1, 8) = GetSet(.TextMatrix(.Rows - 1, 2), .TextMatrix(.Rows - 1, 2)) = True
                .TextMatrix(.Rows - 1, 8) = Round(CCur(((NilaiGaji) / Trim(.TextMatrix(1, 11)))) * CDbl(.TextMatrix(.Rows - 1, 6)), 2)
                .Rows = .Rows + 1
          End If
       End If
         .TextMatrix(.Rows - 1, 6) = 0
         .TextMatrix(.Rows - 1, 7) = "100%"
         .TextMatrix(.Rows - 1, 8) = 0

          For Lrow = 1 To .Rows - 2
          .TextMatrix(.Rows - 1, 6) = CCur(.TextMatrix(.Rows - 1, 6)) + CCur(.TextMatrix(Lrow, 6))
          .TextMatrix(.Rows - 1, 8) = CCur(.TextMatrix(.Rows - 1, 8)) + CCur(.TextMatrix(Lrow, 8))
          Next
    End If
End With
Command1.Enabled = True
End Sub

Function HitungGaji(NIP As String, Project As String, Divisi As String, Row As Integer, Keterangan As String)
Dim StatusProject As Boolean
Dim FRow As Integer
Dim HRow As Integer
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
        End If
            .TextMatrix(HRow, 1) = VSJoin.TextMatrix(Row, 14)
            .TextMatrix(HRow, 2) = VSJoin.TextMatrix(Row, 3)
            .TextMatrix(HRow, 3) = VSJoin.TextMatrix(Row, 4)
            .TextMatrix(HRow, 4) = VSJoin.TextMatrix(Row, 2)
            .TextMatrix(HRow, 5) = VSJoin.TextMatrix(Row, 8)
            .TextMatrix(HRow, 11) = VSJoin.TextMatrix(Row, 17)
             .TextMatrix(HRow, 12) = VSJoin.TextMatrix(Row, 18)
            If Trim(Keterangan) = "Timesheet" Then
                .TextMatrix(HRow, 6) = CCur(.TextMatrix(HRow, 6)) + CCur(VSJoin.TextMatrix(Row, 5))
                .TextMatrix(HRow, 8) = CCur(.TextMatrix(HRow, 8)) + VSJoin.TextMatrix(Row, 15)
             Else
                .TextMatrix(HRow, 9) = CDbl(.TextMatrix(HRow, 9)) + CDbl(VSJoin.TextMatrix(Row, 10))
                .TextMatrix(HRow, 10) = CCur(.TextMatrix(HRow, 10)) + VSJoin.TextMatrix(Row, 16)
            End If
End With
 
End Function
Function JoinGrid(ByVal Project As String, ByVal tgl1 As String, ByVal Row As Integer, ByVal NIP As String, ByVal Keterangan As String)
Dim StatusNIP As Boolean
Dim FRow As Integer
Dim J As String
With VSJoin
    StatusNIP = False
     If Format(VSFg.TextMatrix(Row, 1), "HH:mm") = "12:30" Or Format(VSFg.TextMatrix(Row, 1), "HH:mm") = "13:00" Then Exit Function
    For FRow = 0 To .Rows - 1
        If .TextMatrix(FRow, 2) = Project And .TextMatrix(FRow, 1) = tgl1 And .TextMatrix(FRow, 3) = NIP And .TextMatrix(FRow, 12) = Keterangan Then
           StatusNIP = True
        End If
    Next
        If StatusNIP = False Then
             .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 5) = 0
            .TextMatrix(.Rows - 1, 10) = 0
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
            .TextMatrix(.Rows - 1, 11) = VSFg.TextMatrix(Row, 8)
            .TextMatrix(.Rows - 1, 12) = VSFg.TextMatrix(Row, 7)
            .TextMatrix(.Rows - 1, 13) = VSFg.TextMatrix(Row, 14) 'Hari
            .TextMatrix(.Rows - 1, 14) = VSFg.TextMatrix(Row, 12)
            .TextMatrix(.Rows - 1, 15) = GetSet(.TextMatrix(.Rows - 2, 3), .TextMatrix(.Rows - 1, 3)) = True
             .TextMatrix(.Rows - 1, 15) = Round(CCur(((NilaiGaji) / VSFg.TextMatrix(Row, 15))) * CDbl(.TextMatrix(.Rows - 1, 5)), 2)
            .TextMatrix(.Rows - 1, 16) = 0
            .TextMatrix(.Rows - 1, 17) = VSFg.TextMatrix(Row, 15)
            .TextMatrix(.Rows - 1, 18) = VSFg.TextMatrix(Row, 16)
End With
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
Dim Jam1, Jam2 As Double
Dim Jam3, Jam4 As Double
Dim JamMakan As Double
Dim Transport As String
Dim Upah1, Upah2, UM As Currency
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
                      Istirahat = RsSetting!jamist_1
                      TotalLembur = TotalLembur - Istirahat
 
                   ElseIf TotalLembur >= RsSetting!ist3 And TotalLembur <= RsSetting!ist4 Then
                      Istirahat = RsSetting!jamist_2
                      TotalLembur = TotalLembur - Istirahat
 
                    ElseIf TotalLembur >= RsSetting!ist5 And TotalLembur <= RsSetting!ist6 Then
                      Istirahat = RsSetting!jamist_3
                      TotalLembur = TotalLembur - Istirahat
 
                    ElseIf TotalLembur >= RsSetting!ist7 And TotalLembur <= RsSetting!ist8 Then
                      Istirahat = RsSetting!jamist_4
                      TotalLembur = TotalLembur - Istirahat
                    
                    End If
            End If
         End If
         If Left(TotalLembur, 2) < JamMakan Then UM = 0
         
         'pembulatan menit
         If TotalLembur > 0 Then

         
         If JamKeluar <> "TIDAK ABSEN" Then
            If JamKeluar >= "00:00" And JamKeluar <= "05:00" Then
               Transport = "Ya"
            Else
               Transport = "Tidak"
            End If
         Else
            Transport = "Tidak"
         End If
         'upah
           JamLembur = Replace(TotalLembur, ":", ".")
            Upah1 = RsSetting!Upah1
            Upah2 = RsSetting!Upah2
            NilaiUpah = 0
          
            If JamLembur >= Jam1 And JamLembur <= Jam2 Then
                NilaiUpah = (JamLembur * NilaiGaji / 173) * Upah1
            ElseIf JamLembur >= Jam3 And JamLembur <= Jam4 Then
                NilaiUpah = (Jam2 * NilaiGaji / 173) * Upah1
                SplitLembur = JamLembur - Jam2
                NilaiUpah = NilaiUpah + (SplitLembur * NilaiGaji / 173) * Upah2
            End If
         Else
            UM = 0
            Transport = "Tidak"
            NilaiUpah = 0
         End If
        If TotalLembur > 0 Then
            .TextMatrix(Lrow, 10) = TotalLembur
            .TextMatrix(Lrow, 16) = NilaiUpah
        End If
End With
End Function
Function GetSet(Oldnip As String, NIP As String) As String
Dim i As Integer
GetSet = False
'If Oldnip = NIP Then Exit Function
With VSGaji
    For i = 1 To .Rows - 1
        If Trim(.TextMatrix(i, 1)) = Trim(NIP) Then
            GetSet = True
            NilaiGaji = .TextMatrix(i, 3)
            Tingkat = .TextMatrix(i, 4)
            Exit For
        Else
             NilaiGaji = 0
             Tingkat = 0
        End If
    Next
End With
End Function
