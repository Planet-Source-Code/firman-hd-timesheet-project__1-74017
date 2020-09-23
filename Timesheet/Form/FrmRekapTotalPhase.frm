VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmRekapTotalPhase 
   Caption         =   "Rekap Total Jam Kerja"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   10095
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   10095
      TabIndex        =   16
      Top             =   6120
      Width           =   10095
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
         TabIndex        =   17
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   10065
      TabIndex        =   0
      Top             =   0
      Width           =   10095
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
         Left            =   4200
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   960
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
      Begin VSFlex8Ctl.VSFlexGrid cboFlex 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
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
         FormatString    =   $"FrmRekapTotalPhase.frx":0000
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
         Left            =   1440
         TabIndex        =   9
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
         FormatString    =   $"FrmRekapTotalPhase.frx":0029
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
         Left            =   6960
         TabIndex        =   10
         Top             =   960
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
         Left            =   1440
         TabIndex        =   11
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
         FormatString    =   $"FrmRekapTotalPhase.frx":0052
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
         OleObjectBlob   =   "FrmRekapTotalPhase.frx":007B
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
         Left            =   4680
         TabIndex        =   22
         Top             =   120
         Width           =   1575
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
         Left            =   120
         TabIndex        =   15
         Top             =   960
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         Left            =   960
         TabIndex        =   12
         Top             =   600
         Width           =   495
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "FrmRekapTotalPhase.frx":02AF
      Top             =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3735
      Left            =   0
      TabIndex        =   18
      ToolTipText     =   "Double Klik Kolom Project Untuk Melihat PM"
      Top             =   1440
      Width           =   6015
      _cx             =   10610
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
      FormatString    =   $"FrmRekapTotalPhase.frx":04E3
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
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   3255
      Left            =   0
      TabIndex        =   20
      Top             =   5520
      Width           =   14655
      _cx             =   25850
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
      FormatString    =   $"FrmRekapTotalPhase.frx":05C2
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
      Left            =   6360
      TabIndex        =   19
      Top             =   1680
      Width           =   8415
      _cx             =   14843
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
      FormatString    =   $"FrmRekapTotalPhase.frx":06A1
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
      Height          =   3735
      Left            =   6120
      TabIndex        =   21
      Top             =   1440
      Width           =   8535
      _cx             =   15055
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
      FormatString    =   $"FrmRekapTotalPhase.frx":0780
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
Attribute VB_Name = "FrmRekapTotalPhase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub CmdClose_Click()
Unload Me
End Sub
Private Sub Command1_Click()
Dim Selisih, x As Integer
Dim TotalHari As Integer
Dim Hari As String
Dim TglAwal, TglAkhir As Date
'If Trim(CboKaryawan) = "" And Trim(cboFlex) = "" Then
'   MsgBox "Kriteria Pencarian Belum Diisi", vbCritical
'   Exit Sub
'End If
Command1.Enabled = False
Setgrid
Selisih = DateDiff("d", DTPicker1, DTPicker2)
    TotalHari = 0
    TglAwal = DTPicker1
    For x = 0 To Selisih
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
     
    TotalJamPUmum = TotalHari * 8
    LblTotaljam = TotalJamPUmum & " Jam"
Showdata
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Dim x As String
If fg.Rows > 1 Then
   fg.Cols = 11
   fg.SaveGrid "C:\TotalJam.xls", flexFileExcel, True
'    Shell PathOffice & "C:\TotalJam.csv", vbNormalFocus
   x = ShellExecute(Me.hwnd, "open", "C:\TotalJam.xls", vbNullString, "C:\TotalJam.xls", 1)
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
If fg.Rows > 1 Then fg.PrintGrid "Total Jam - Periode " & DTPicker1.Value & " S/D " & DTPicker2.Value, 2, 2, 1000, 500

End Sub
   
Private Sub fg_DblClick()
Dim J As String
With fg
If .Row < 0 Then Exit Sub
If .Col = 4 And .TextMatrix(.Row, 4) <> "" Then
    J = FrmPM.Showdata(.TextMatrix(.Row, 4), KodeDivisi)
    FrmPM.NIP = .TextMatrix(.Row, 2)
    FrmPM.show vbModal
End If
End With
End Sub

Private Sub Form_Load()
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
   Select Case UCase(strGroup)
       Case "USER", "PM"
            CboKaryawan.Enabled = False
            
   End Select
   CboKaryawan.Text = StrNIPUser
End Sub
  
  
Sub Setgrid()
If cboFlex = "" And Trim(CboKaryawan) <> "" Then
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
'    .ColDataType(6) = flexDTDouble
    .MergeCells = flexMergeFree
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    .TextMatrix(0, 1) = "Nama Divisi"
    .TextMatrix(0, 2) = "NIP"
    .TextMatrix(0, 3) = "Nama"
    .TextMatrix(0, 4) = "Kode Project"
    .TextMatrix(0, 5) = "Nama Project"
    .TextMatrix(0, 6) = "Phase"
    .TextMatrix(0, 7) = "Total Jam"
    .TextMatrix(0, 8) = "Total %"
    .TextMatrix(0, 9) = "Jam Lembur"
    .TextMatrix(0, 10) = ""
    .TextMatrix(0, 11) = ""
    .TextMatrix(0, 12) = ""
    .ColFormat(10) = "#,###"
    .ColWidth(6) = 1000
    .ColWidth(10) = 0
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1200
    .ColWidth(11) = 0
    .ColWidth(12) = 0
    End With
Else
     With fg
    .Cols = 13
    .Rows = 1
    
    .TextMatrix(0, 1) = "Nama Divisi"
    .TextMatrix(0, 2) = "Kode Project"
    .TextMatrix(0, 3) = "Nama Project"
    .TextMatrix(0, 4) = "Phase"
    .TextMatrix(0, 5) = "NIP"
    .TextMatrix(0, 6) = "Nama"
    .TextMatrix(0, 7) = "Total Jam"
    .TextMatrix(0, 8) = "Total %"
    .TextMatrix(0, 9) = "Jam Lembur"
    .TextMatrix(0, 10) = ""
    .TextMatrix(0, 11) = ""
    .TextMatrix(0, 12) = ""
    .ColFormat(10) = "#,###"
    .ColWidth(0) = 300
    .ColWidth(1) = 2000
    .ColWidth(3) = 2000
    .ColWidth(2) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 700
    .ColWidth(6) = 2500
    .ColWidth(10) = 0
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1200
    .ColWidth(11) = 0
    .ColWidth(12) = 0
    .MergeCells = flexMergeFree
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = False
    .MergeCol(5) = False
    .MergeCol(6) = False
    End With
End If
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
.TextMatrix(.Rows - 1, 15) = "15. Phase"
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
Private Sub AddKaryawan()
    Dim Cboid     As String
    Dim cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    cboid1 = vbNullString
    Select Case UCase(strGroup)
        Case "PTW", "IT", "HRD"
                StrSQL = "select * from Karyawan Where Status <> '14' And Len(NIP) < 5 Order By Nama"
        Case "ADMIN", "PM"
            StrSQL = "select * from Karyawan Where Status <> '14' And Len(NIP) < 5 And kd_divisi = '" & KodeDivisi & "'  Order By Nama"
  
    
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
 End Select
End Sub
 
Private Sub AddProject()

    Dim Cboid     As String
    Dim cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    cboid1 = vbNullString
    Select Case UCase(strGroup)
        Case "IT", "PTW", "HRD"
            StrSQL = "select Kode,Nama from project " & _
                 "group by kode,Nama " & _
                 "order by kode"
        Case "USER", "PM", "ADMIN"
             StrSQL = "select Kode,Nama from project " & _
             "where  Kd_Divisi = '" & KodeDivisi & "'" & _
             "group by kode,Nama " & _
             "order by kode"
    End Select
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
 
Private Sub Form_Resize()
    On Error Resume Next
        CmdClose.Width = Me.Width - 100
        With fg
             .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture2.Height

        End With
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set FrmRekapTotalPhase = Nothing
End Sub
Sub Showdata()
Dim x As Integer
Dim i, J As Integer
Dim Split, JmlLoop As Integer
Dim Jam As String
Dim IDWaktu As String
Dim Jam1, Jam2 As Date
Dim Jam3, Jam4 As Date
Dim TglAwal, TglAkhir As Date
Dim RsTS As New ADODB.Recordset
Dim TotalJamBulan As Integer
 Setgrid
Command1.Enabled = False
With VSFlexGrid1
        .Rows = 1
        VSFg.Rows = 1
        VSFg.Cols = 18
    If RsTS.State = adStateOpen Then RsTS.Close
    StrSQL = "SELECT tbltimesheet.IDtimesheet,tbltimesheet.Tanggal,tbltimesheet.JamAwal As [Jam Awal],tbltimesheet.JamAkhir AS [Jam Akhir],tbltimesheet.Status,tbltimesheet.NoProject As Project,tbltimesheet.Keterangan,tbltimesheet.Tanggal,Absensi.Masuk,tbltimesheet.NIP,tbltimesheet.StatusDivisi, karyawan.Nama,tbltimesheet.StatusPM,Absensi.Keluar,Divisi.NM_DIV, tbltimesheet.Hari,tbltimesheet.Hari,tbltimesheet.TotalKerja,tbltimesheet.ProjectUmum,tbltimesheet.Phase"
    StrSQL = StrSQL & " FROM tbltimesheet INNER JOIN karyawan ON tbltimesheet.NIP = karyawan.NIP INNER JOIN Divisi ON karyawan.Kd_Divisi = Divisi.KD_DIV INNER JOIN Absensi ON tbltimesheet.NIP = Absensi.NIP AND tbltimesheet.Tanggal = Absensi.Tgl"
    StrSQL = StrSQL & " Where tbltimesheet.Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And '" & Format(DTPicker2, "MM/dd/yyyy") & "'  And tbltimesheet.Status ='Actual' And  tbltimesheet.NoProject <> '' "
    If Trim(cboFlex) <> "" Then StrSQL = StrSQL & " AND tbltimesheet.NoProject = '" & cboFlex & "'"
    If Trim(CboKaryawan) <> "" Then StrSQL = StrSQL & " AND tbltimesheet.Nip = '" & CboKaryawan & "'"
    If Trim(CboKaryawan) = "" And Trim(cboFlex) = "" Then StrSQL = StrSQL & " And tbltimesheet.Kd_Divisi = '" & KodeDivisi & "'"
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
                        VSFg.TextMatrix(VSFg.Rows - 1, 15) = TotalJamPUmum '.TextMatrix(Lrow, 18)
                        VSFg.TextMatrix(VSFg.Rows - 1, 16) = .TextMatrix(Lrow, 19)
                        VSFg.TextMatrix(VSFg.Rows - 1, 17) = .TextMatrix(Lrow, 20)
                        Jam3 = Format(.TextMatrix(Lrow, 9), "HH:mm")

                        Jam4 = DateAdd("h", 9, Jam3)
                        If Trim(.TextMatrix(Lrow, 17)) = "Kerja" Then
                            If (VSFg.TextMatrix(VSFg.Rows - 1, 1) >= "08:00" And VSFg.TextMatrix(VSFg.Rows - 1, 1) <= "17:00") Or (VSFg.TextMatrix(VSFg.Rows - 1, 1) < Jam4 And VSFg.TextMatrix(VSFg.Rows - 1, 1) >= "08:00") Then
                                 VSFg.TextMatrix(VSFg.Rows - 1, 7) = "Timesheet"
                              End If

                        End If
'                          If Format(DTPicker3.Value, "HH:mm") = Format(Jam2, "HH:mm") Then Exit Do
'                        If Format(DTPicker3, "HH:mm") = "12:00" Then DTPicker3.Value = DateAdd("n", 60, DTPicker3)
                        JmlLoop = JmlLoop + 1
                Loop
                Command1.Caption = .TextMatrix(Lrow, 1)
    Next
        VSFg.ColFormat(6) = "HH:mm"
        VSFg.ColFormat(11) = "HH:mm"
    For lCol = 1 To VSFg.Cols - 1
        VSFg.TextMatrix(0, lCol) = lCol
    Next
End With
 
With VSFg
    For Lrow = 1 To VSFg.Rows - 1
             J = JoinGrid(.TextMatrix(Lrow, 3), .TextMatrix(Lrow, 2), Lrow, .TextMatrix(Lrow, 4), .TextMatrix(Lrow, 7))
        Command1.Caption = .TextMatrix(Lrow, 1)
    Next
End With

With VSJoin
For Lrow = 1 To VSJoin.Rows - 1
   
    If cboFlex = "" And Trim(CboKaryawan) <> "" Then
        J = HitungGaji(.TextMatrix(Lrow, 3), .TextMatrix(Lrow, 2), .TextMatrix(Lrow, 14), Lrow, .TextMatrix(Lrow, 12), .TextMatrix(Lrow, 15))
    Else
        J = HitungGaji2(.TextMatrix(Lrow, 3), .TextMatrix(Lrow, 2), .TextMatrix(Lrow, 14), Lrow, .TextMatrix(Lrow, 12), .TextMatrix(Lrow, 15))

    End If
    Command1.Caption = .TextMatrix(Lrow, 1)
Next
  
End With
With fg
Dim Tjam, Gaji As Double
Dim tlembur, lembur As Double
Dim tPersen As Double
 
    .Rows = .Rows + 1
    For Lrow = 1 To .Rows - 2
'           .TextMatrix(Lrow, 5) = "???"
        If Rscek.State = adStateOpen Then Rscek.Close
        If cboFlex = "" And Trim(CboKaryawan) <> "" Then
        
            Rscek.Open "Select * From Project Where Kode = '" & .TextMatrix(Lrow, 4) & "'", CN, adOpenStatic
            
            If Not Rscek.EOF Then .TextMatrix(Lrow, 5) = Rscek!Nama
        Else
             Rscek.Open "Select * From Project Where Kode = '" & .TextMatrix(Lrow, 2) & "'", CN, adOpenStatic
            
            If Not Rscek.EOF Then .TextMatrix(Lrow, 3) = Rscek!Nama
        End If
        Tjam = Tjam + CDbl(.TextMatrix(Lrow, 7))
        Gaji = Gaji + CDbl(.TextMatrix(Lrow, 8))
        .TextMatrix(Lrow, 8) = Round(.TextMatrix(Lrow, 7) / .TextMatrix(Lrow, 11) * 100, 2)
        tPersen = CDbl(tPersen) + CDbl(.TextMatrix(Lrow, 8))
        .TextMatrix(Lrow, 8) = .TextMatrix(Lrow, 8) & "%"
'        .Row = Lrow
'        .Col = 6
'        .CellAlignment = flexAlignRightCenter
        
        tlembur = tlembur + CDbl(.TextMatrix(Lrow, 9))
        .TextMatrix(Lrow, 0) = Lrow
    Next
    If .TextMatrix(1, 1) <> "" Then
     
    .TextMatrix(.Rows - 1, 7) = Tjam
    .TextMatrix(.Rows - 1, 8) = tPersen & "%" ' Gaji
    .TextMatrix(.Rows - 1, 9) = tlembur
    End If

'         .TextMatrix(.Rows - 1, 7) = 0
'          For Lrow = 1 To .Rows - 2
'          .TextMatrix(.Rows - 1, 7) = CCur(.TextMatrix(.Rows - 1, 7)) + CCur(.TextMatrix(Lrow, 7))
'
'          Next
End With
 
Command1.Caption = "Refresh"
Command1.Enabled = True
End Sub

Function HitungGaji(NIP As String, Project As String, Divisi As String, Row As Integer, Keterangan As String, ByVal Phase As String)
Dim StatusProject As Boolean
Dim FRow As Integer
Dim HRow As Integer
With fg
    StatusProject = False
    For FRow = 0 To .Rows - 1
    If Trim(.TextMatrix(FRow, 1)) = Divisi And Trim(.TextMatrix(FRow, 2)) = NIP And Trim(.TextMatrix(FRow, 4)) = Project And Trim(.TextMatrix(FRow, 6)) = Phase Then
           StatusProject = True
           HRow = FRow
        End If
    Next
        If StatusProject = False Then
             .Rows = .Rows + 1
             HRow = .Rows - 1
'            .TextMatrix(HRow, 6) = 0
            .TextMatrix(HRow, 7) = 0
            .TextMatrix(HRow, 8) = 0
            .TextMatrix(HRow, 9) = 0
'            .TextMatrix(HRow, 10) = 0
        End If
            .TextMatrix(HRow, 1) = VSJoin.TextMatrix(Row, 14)
            .TextMatrix(HRow, 2) = VSJoin.TextMatrix(Row, 3)
            .TextMatrix(HRow, 3) = VSJoin.TextMatrix(Row, 4)
            .TextMatrix(HRow, 4) = VSJoin.TextMatrix(Row, 2)
            .TextMatrix(HRow, 6) = Phase
            .TextMatrix(HRow, 11) = VSJoin.TextMatrix(Row, 17)
             .TextMatrix(HRow, 12) = VSJoin.TextMatrix(Row, 18)
            If Trim(Keterangan) = "Timesheet" Then
                .TextMatrix(HRow, 7) = CCur(.TextMatrix(HRow, 7)) + CCur(VSJoin.TextMatrix(Row, 5))
'                .TextMatrix(HRow, 8) = CCur(.TextMatrix(HRow, 8)) + VSJoin.TextMatrix(Row, 15)
             Else
                .TextMatrix(HRow, 9) = CDbl(.TextMatrix(HRow, 9)) + CDbl(VSJoin.TextMatrix(Row, 10))
'                .TextMatrix(HRow, 10) = CCur(.TextMatrix(HRow, 10)) + VSJoin.TextMatrix(Row, 16)
            End If
End With
 
End Function
Function HitungGaji2(NIP As String, Project As String, Divisi As String, Row As Integer, Keterangan As String, ByVal Phase As String)
Dim StatusProject As Boolean
Dim FRow As Integer
Dim HRow As Integer
With fg
    StatusProject = False
    For FRow = 0 To .Rows - 1
    If Trim(.TextMatrix(FRow, 1)) = Divisi And Trim(.TextMatrix(FRow, 5)) = NIP And Trim(.TextMatrix(FRow, 2)) = Project And Trim(.TextMatrix(FRow, 4)) = Phase Then
           StatusProject = True
           HRow = FRow
        End If
    Next
        If StatusProject = False Then
             .Rows = .Rows + 1
             HRow = .Rows - 1
            .TextMatrix(HRow, 7) = 0
            .TextMatrix(HRow, 8) = 0
            .TextMatrix(HRow, 9) = 0
         End If
            .TextMatrix(HRow, 1) = VSJoin.TextMatrix(Row, 14)
            .TextMatrix(HRow, 5) = VSJoin.TextMatrix(Row, 3)
            .TextMatrix(HRow, 6) = VSJoin.TextMatrix(Row, 4)
            .TextMatrix(HRow, 2) = VSJoin.TextMatrix(Row, 2)
            .TextMatrix(HRow, 4) = Phase
            .TextMatrix(HRow, 11) = VSJoin.TextMatrix(Row, 17)
             .TextMatrix(HRow, 12) = VSJoin.TextMatrix(Row, 18)
            If Trim(Keterangan) = "Timesheet" Then
                .TextMatrix(HRow, 7) = CCur(.TextMatrix(HRow, 7)) + CCur(VSJoin.TextMatrix(Row, 5))
              Else
                .TextMatrix(HRow, 9) = CDbl(.TextMatrix(HRow, 9)) + CDbl(VSJoin.TextMatrix(Row, 10))
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
            .TextMatrix(.Rows - 1, 11) = VSFg.TextMatrix(Row, 8)
            .TextMatrix(.Rows - 1, 12) = VSFg.TextMatrix(Row, 7)
            .TextMatrix(.Rows - 1, 13) = VSFg.TextMatrix(Row, 14) 'Hari
            .TextMatrix(.Rows - 1, 14) = VSFg.TextMatrix(Row, 12)
            .TextMatrix(.Rows - 1, 15) = VSFg.TextMatrix(Row, 17)
            .TextMatrix(.Rows - 1, 16) = 0
            .TextMatrix(.Rows - 1, 17) = VSFg.TextMatrix(Row, 15)
            .TextMatrix(.Rows - 1, 18) = VSFg.TextMatrix(Row, 16)
End With
End Function


Private Sub VSFlexGrid1_DblClick()
MsgBox VSFlexGrid1.Col
End Sub
