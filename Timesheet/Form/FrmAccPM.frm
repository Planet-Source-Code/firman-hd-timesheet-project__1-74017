VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAccPM 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Verifikasi Timesheet & Lembur PM"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   10230
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   7095
      Left            =   4200
      TabIndex        =   16
      Top             =   1320
      Width           =   6015
      _cx             =   10610
      _cy             =   12515
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   15648682
      ForeColorFixed  =   -2147483630
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
      SelectionMode   =   0
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
      ExplorerBar     =   1
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
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10230
      TabIndex        =   0
      Top             =   7590
      Width           =   10230
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
         Top             =   0
         Width           =   1215
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin2 
      Left            =   8400
      OleObjectBlob   =   "FrmAccPM.frx":0000
      Top             =   2760
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FrmAccPM.frx":0234
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   10200
      TabIndex        =   2
      Top             =   0
      Width           =   10230
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   960
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Simpan"
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
         Left            =   4920
         TabIndex        =   15
         ToolTipText     =   "Verifikasi / Unverifikasi"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   7680
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   240
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
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   7800
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   55836674
         CurrentDate     =   39940
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   240
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
      Begin VSFlex8Ctl.VSFlexGrid CboFlex 
         Height          =   315
         Left            =   7800
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   2865
         _cx             =   5054
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
         FormatString    =   $"FrmAccPM.frx":0468
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "List Actual Timesheet "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   11
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal :"
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
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "SD"
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
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "List Project Muncul Berdasarkan Tgl Akhir Project"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   4095
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSProject 
      Height          =   4455
      Left            =   0
      TabIndex        =   17
      ToolTipText     =   "Silahkan Pilih Projectnya"
      Top             =   1320
      Width           =   4095
      _cx             =   7223
      _cy             =   7858
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   15648682
      ForeColorFixed  =   -2147483630
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
      AutoSearch      =   2
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
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   2895
      Left            =   10560
      TabIndex        =   12
      Top             =   6720
      Width           =   4575
      _cx             =   8070
      _cy             =   5106
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
      FormatString    =   $"FrmAccPM.frx":0491
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
   Begin VSFlex8Ctl.VSFlexGrid VSfg 
      Height          =   2535
      Left            =   10560
      TabIndex        =   13
      Top             =   1320
      Width           =   4575
      _cx             =   8070
      _cy             =   4471
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
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
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
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
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
      Height          =   2175
      Left            =   10560
      TabIndex        =   14
      Top             =   4320
      Width           =   4575
      _cx             =   8070
      _cy             =   3836
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
      FormatString    =   $"FrmAccPM.frx":0570
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
Attribute VB_Name = "FrmAccPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTimesheet As New Recordset
Dim StatusHari As String
Public Sub Perintah(ByVal What As String)
Dim Lrow As Long
Dim lCol As Long
    On Error GoTo err
    Select Case What
        Case "New"
          FrmTimesheetPlan.show
         Case "Search"
           
         Case "Select"
            With fg
                For Lrow = 1 To .Rows - 2
                    .TextMatrix(Lrow, 1) = "-1"
                Next
            End With
        Case "Delete"
            Call Hapus
        Case "Close"
            Unload Me
    End Select
    Exit Sub
    'Trap the error
err:
    If err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it was used by other records! If you want to delete this record" & vbCrLf & _
               "you will first have to delete or change the records that currenly used this record as shown bellow." & vbCrLf & vbCrLf & _
               err.Description, , "Delete Operation Failed!"
    End If
MsgBox err.Description
    Me.MousePointer = vbDefault
End Sub

Private Sub Check1_Click()
AddProject
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim Lrow As Long
Dim StrSQL As String
On Error GoTo Adaerror
 If CN.State = adStateClosed Then CN.Open
        With VSFlexGrid2
            If MsgBox("Apakah Anda yakin ingin Menyimpan  Data ?", vbQuestion + vbYesNo, "Konfirmasi Simpan Data") = vbNo Then
               Exit Sub
            Else
             Do Until Lrow = .Rows
               If Lrow > 0 Then
                    
                     If .TextMatrix(Lrow, 9) <> .TextMatrix(Lrow, 11) Then
                            StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Data Verifikasi Timesheet, " & .TextMatrix(Lrow, 3) & "','Acc PM')"
                            PerintahExecute (StrSQL)
                            
'                            StrSQL = "Update Tbltimesheet Set StatusPM = '" & .TextMatrix(Lrow, 9) & "',last_update='" & Now & "',last_user='" & StrUser & "' Where Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And '" & Format(DTPicker2, "MM/dd/yyyy") & "' And NIP = '" & .TextMatrix(Lrow, 3) & "' And NoProject = '" & VSProject.TextMatrix(VSProject.Row, 1) & "'"
                            StrSQL = "Update Tbltimesheet Set StatusPM = '" & .TextMatrix(Lrow, 9) & "',last_update='" & Now & "',last_user='" & StrUser & "' Where Tanggal = '" & Format(.TextMatrix(Lrow, 1), "MM/dd/yyyy") & "' And NIP = '" & .TextMatrix(Lrow, 3) & "' And NoProject = '" & VSProject.TextMatrix(VSProject.Row, 1) & "'"
                            CN.Execute StrSQL
            
                     End If
                 End If
                 Lrow = Lrow + 1
            Loop
                    MsgBox "Data Berhasil Disimpan", vbInformation
                   
                End If
            End With
            With fg
                For Lrow = 1 To .Rows - 1
                        For lCol = 3 To .Cols - 1
                            If .TextMatrix(Lrow, lCol - 1) = "-1" Then
                               .Row = Lrow
                               .Col = lCol
                               .CellBackColor = vbGreen
                            Else
                               .Row = Lrow
                               .Col = lCol
                               .CellBackColor = vbWhite
                            End If
                        Next
                    Next
                    .Col = 0
                    .Row = 0
            End With
Exit Sub
Adaerror:
MsgBox err.Description
End Sub
'
'Private Sub Command1_Click()
'Command1.Enabled = False
'Showdata
'Command1.Enabled = True
'End Sub

Private Sub DTPicker1_Change()
AddProject
End Sub

Private Sub DTPicker1_Click()
'AddProject
End Sub

Private Sub DTPicker1_LostFocus()
'AddProject
End Sub

Private Sub DTPicker2_Change()
AddProject
End Sub

Private Sub DTPicker2_Click()
'AddProject
End Sub

Private Sub DTPicker2_LostFocus()
'AddProject
End Sub

Private Sub Form_Load()
     If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
      Skin1.ApplySkin hwnd
    End If
    Setgrid
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker1.Value = DateSerial(Year(Now), Month(Now), 26)
    DTPicker1.Value = DateAdd("M", -1, DTPicker1.Value)
    DTPicker2.Value = DateSerial(Year(Now), Month(Now), 25)
    DTPicker2.Value = DateAdd("M", 0, DTPicker2.Value)
    DTPicker1.CustomFormat = "dd/MMM/yyyy"
    DTPicker2.CustomFormat = "dd/MMM/yyyy"
    AddProject
    
'    Showdata

End Sub
Sub Setgrid()
fg.Rows = 1
fg.Cols = 3
fg.ColWidth(0) = 500
fg.ColWidth(1) = 700
fg.ColWidth(2) = 3000
fg.TextMatrix(0, 0) = "No"
fg.TextMatrix(0, 1) = "NIP"
fg.TextMatrix(0, 2) = "Nama"
End Sub
Private Sub AddProject()
 '    Dim Cboid     As String
'    Dim Cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
'    Cboid = vbNullString
'    Cboid1 = vbNullString
    VSProject.Rows = 1
    If StrUser = "3578" Then
        StrSQL = "select Kode,Nama from project group by kode,Nama order by kode"
    Else
        If Check1.Value = 1 Then
            StrSQL = "select Kode,Nama from project Where nip_pm = '" & StrNIPUser & "' And Tgl_Akhir >= '" & Format(DTPicker1, "MM/dd/yyyy") & "' Order by kode"
        Else
            StrSQL = "select Kode,Nama from project Where nip_pm = '" & StrNIPUser & "' order by kode"
        End If
    End If
    
    Rscek.Open StrSQL, CN, adOpenStatic
'    Cboid1 = " "
'    Do Until Rscek.EOF
'      Cboid = "|" & Rscek("Kode") & vbTab & Rscek("Nama")
'      Cboid1 = Cboid1 + Cboid
'      Rscek.MoveNext
'    Loop
'    CboFlex.ColComboList(0) = Cboid1
    Set VSProject.DataSource = Rscek
With VSProject
    .TextMatrix(0, 0) = "No"
    .ColWidth(0) = 500
    For Lrow = 1 To .Rows - 1
        .TextMatrix(Lrow, 0) = Lrow
    Next
End With
    If Rscek.State = adStateOpen Then Rscek.Close
    Set VSProject.DataSource = Nothing
    Set Rscek = Nothing
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        CmdClose.Width = Me.Width - 100
        With fg
'             .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture3.Height
              .Width = Me.Width - 4200
              .Height = Me.Height - 1800
        End With

        VSProject.Height = Me.Height - 1800
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
    Set FrmListPlan = Nothing
End Sub


Private Sub Hapus()
Dim StrSQL As String
Dim Lrow As Integer
Dim Tanya As String
Dim ErrConn As Long
         With fg
                .Rows = .Rows + 1
             If CekCurek("hapus", fg) = False Then .Rows = .Rows - 1: Exit Sub
              
            If MsgBox("Apakah Anda yakin ingin menghapus Data ?", vbQuestion + vbYesNo, "Konfirmasi hapus") = vbNo Then
               Exit Sub
            Else
             Do Until Lrow = .Rows - 1
             
                If .TextMatrix(Lrow, 1) = "-1" Then
                         StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Hapus Timesheet, " & StrNIPUser & " &  " & .TextMatrix(Lrow, 3) & "','Acc PM')"
                         PerintahExecute (StrSQL)
                                
                        StrSQL = "Delete From tbltimesheet"
                        StrSQL = StrSQL & " where Tanggal= '" & Format(.TextMatrix(Lrow, 2), "yyyy/MM/dd") & "' And NIP = '" & .TextMatrix(Lrow, 4) & "' And Status = 'Plan'"
                        CN.Execute StrSQL
                        .RemoveItem (Lrow)
                        Lrow = Lrow - 1
                End If
                Lrow = Lrow + 1
            Loop
                    
                End If
                  .Rows = .Rows - 1
            End With
Exit Sub
Adaerror:
If ErrConn > 0 Then CN.RollbackTrans
MsgBox err.Description
End Sub

Sub Showdata(Project As String)
Dim x As Integer
Dim i, J As Integer
Dim Split, JmlLoop As Integer
Dim Jam As String
Dim IDWaktu As String
Dim Jam1, Jam2 As Date
Dim Jam3, Jam4 As Date
Dim TglAwal, TglAkhir As Date
Dim RsTS As New ADODB.Recordset
   
'If Trim(CboFlex.Text) = "" Then
'   MsgBox "Silahkan Pilih Project Terlebih Dahulu", vbCritical
'   CboFlex.SetFocus
'   Exit Sub
'End If
fg.Rows = 1
fg.Cols = 3
fg.ColWidth(0) = 500
fg.ColWidth(1) = 700
fg.ColWidth(2) = 3000
fg.TextMatrix(0, 0) = "No"
fg.TextMatrix(0, 1) = "NIP"
fg.TextMatrix(0, 2) = "Nama"
fg.FrozenCols = 2
With VSFlexGrid1
        .Rows = 1
        VSFg.Rows = 1
        fg.Rows = 1
        VSFg.Cols = 14
        VSFlexGrid2.Cols = 11
    If RsTS.State = adStateOpen Then RsTS.Close
    Set RsTS = Nothing
    StrSQL = "SELECT tbltimesheet.IDtimesheet,tbltimesheet.Tanggal,tbltimesheet.JamAwal As [Jam Awal],tbltimesheet.JamAkhir AS [Jam Akhir],tbltimesheet.Status,tbltimesheet.NoProject As Project,tbltimesheet.Keterangan,tbltimesheet.Tanggal,tbltimesheet.Masuk,tbltimesheet.NIP,tbltimesheet.StatusPM, karyawan.Nama,tbltimesheet.Hari,tbltimesheet.Masuk,tbltimesheet.keluar FROM tbltimesheet INNER JOIN  karyawan ON tbltimesheet.NIP = karyawan.NIP "
    StrSQL = StrSQL & " Where tbltimesheet.Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And '" & Format(DTPicker2, "MM/dd/yyyy") & "' And tbltimesheet.Status ='Actual'"
    StrSQL = StrSQL & " AND tbltimesheet.NoProject = '" & Project & "'"
    StrSQL = StrSQL & " Order By tbltimesheet.Tanggal,tbltimesheet.NIP,tbltimesheet.IdTimesheet ASC"
    
    RsTS.Open StrSQL, CN, adOpenStatic
    Set .DataSource = RsTS
    .ColDataType(2) = flexDTDate
'    .ColFormat(2) = "dd/MM/yyyy"

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
                        VSFg.TextMatrix(VSFg.Rows - 1, 6) = Format(.TextMatrix(Lrow, 9), "HH:mm")
                        VSFg.TextMatrix(VSFg.Rows - 1, 7) = .TextMatrix(Lrow, 7)
                        VSFg.TextMatrix(VSFg.Rows - 1, 8) = .TextMatrix(Lrow, 11)
                        VSFg.TextMatrix(VSFg.Rows - 1, 9) = .TextMatrix(Lrow, 12)
                        VSFg.TextMatrix(VSFg.Rows - 1, 10) = .TextMatrix(Lrow, 1)
                        VSFg.TextMatrix(VSFg.Rows - 1, 11) = .TextMatrix(Lrow, 13)
                        Jam4 = Format(.TextMatrix(Lrow, 9), "HH:mm")
                       
                        VSFg.TextMatrix(VSFg.Rows - 1, 12) = DateAdd("h", 9, Jam4)
                        VSFg.TextMatrix(VSFg.Rows - 1, 12) = Format(VSFg.TextMatrix(VSFg.Rows - 1, 12), "HH:mm")
                        If Trim(.TextMatrix(Lrow, 13)) = "Kerja" Then
                              If (VSFg.TextMatrix(VSFg.Rows - 1, 1) >= "08:00" And VSFg.TextMatrix(VSFg.Rows - 1, 1) <= "17:00") Or (VSFg.TextMatrix(VSFg.Rows - 1, 1) < VSFg.TextMatrix(VSFg.Rows - 1, 12) And VSFg.TextMatrix(VSFg.Rows - 1, 1) >= "08:00") Then
                                 VSFg.TextMatrix(VSFg.Rows - 1, 7) = "Timesheet"
                              End If
                              
                        End If
                        VSFg.TextMatrix(VSFg.Rows - 1, 13) = Format(.TextMatrix(Lrow, 15), "HH:mm")
                         
                        If Format(DTPicker3, "HH:mm") = "12:00" Then
                            DTPicker3.Value = DateAdd("n", 60, DTPicker3)
                        End If
                          If Format(DTPicker3.Value, "HH:mm") >= Format(Jam2, "HH:mm") Then
                              Exit Do
                           End If
                        JmlLoop = JmlLoop + 1
                    
                Loop
    Next
End With
    
'Disiplit Perhari
With VSFlexGrid2
.Rows = 1
.Cols = 14
.TextMatrix(.Rows - 1, 1) = "Tanggal"
.TextMatrix(.Rows - 1, 2) = "Project"
.TextMatrix(.Rows - 1, 3) = "NIP"
.TextMatrix(.Rows - 1, 4) = "Nama"
.TextMatrix(.Rows - 1, 5) = "Total / 2"
.TextMatrix(.Rows - 1, 6) = "masuk"
.TextMatrix(.Rows - 1, 7) = "keluar"
.TextMatrix(.Rows - 1, 8) = "Project"
.TextMatrix(.Rows - 1, 9) = "StatusPM"
.TextMatrix(.Rows - 1, 10) = "TotalLembur"
.TextMatrix(.Rows - 1, 11) = "StatusPM"
.TextMatrix(.Rows - 1, 12) = "Keterangn"
End With
With VSFg
For Lrow = 1 To VSFg.Rows - 1
    
        J = JoinGrid(.TextMatrix(Lrow - 1, 2), .TextMatrix(Lrow, 2), Lrow, .TextMatrix(Lrow, 4), .TextMatrix(Lrow, 7))

Next
End With
With VSFlexGrid2
For Lrow = 1 To VSFlexGrid2.Rows - 1
    If Trim(.TextMatrix(Lrow, 5)) = "" Then .TextMatrix(Lrow, 5) = 0
    .TextMatrix(Lrow, 0) = Lrow
    .TextMatrix(0, 0) = .Rows
    J = Tampil(Format(VSFlexGrid2.TextMatrix(Lrow, 1), "dd/MM/yy"), Lrow, .TextMatrix(Lrow, 3), .TextMatrix(Lrow, 9))
Next
      If .Rows = 1 Then MsgBox "Data Tidak Ditemukan", vbInformation
fg.Col = 1
End With
End Sub
Function JoinGrid(ByVal Tgl As String, ByVal tgl1 As String, ByVal Row As Integer, ByVal NIP As String, ByVal Keterangan As String)
Dim x, Kolom As Long
Dim StatusNIP As Boolean
Dim FRow As Integer
With VSFlexGrid2
    StatusNIP = False
    For FRow = 0 To .Rows - 1
        If .TextMatrix(FRow, 1) = tgl1 And .TextMatrix(FRow, 3) = NIP And .TextMatrix(FRow, 12) = Keterangan Then
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
                    .TextMatrix(.Rows - 1, 6) = VSFg.TextMatrix(Row, 1) 'jam kerja/ts
                    .TextMatrix(.Rows - 1, 7) = VSFg.TextMatrix(Row, 13) 'jam keluar
                    .TextMatrix(.Rows - 1, 8) = VSFg.TextMatrix(Row, 3)
                    .TextMatrix(.Rows - 1, 9) = VSFg.TextMatrix(Row, 8)
                    If Trim(Keterangan) = "Lembur" And .TextMatrix(.Rows - 1, 6) < .TextMatrix(.Rows - 1, 7) Then
                        .TextMatrix(.Rows - 1, 10) = (.TextMatrix(.Rows - 1, 10) + 0.5)
                    End If
                    .TextMatrix(.Rows - 1, 11) = VSFg.TextMatrix(Row, 8)
                    .TextMatrix(.Rows - 1, 12) = VSFg.TextMatrix(Row, 7)
                    .TextMatrix(.Rows - 1, 13) = VSFg.TextMatrix(Row, 10)
               
End With
End Function
Function Tampil(ByVal tgl1 As String, ByVal Row As Integer, ByVal NIP As String, ByVal status As String)
Dim x, Kolom As Long
Dim StatusNIP, StatusTgl As Boolean
Dim VsFlexRow As Integer
With fg
    'Cek di Vsflex NIP Sudah Ada \ Blm
    StatusNIP = False
    For x = 1 To .Rows - 1
        If .TextMatrix(x, 1) = NIP Then
            StatusNIP = True
            VsFlexRow = x
            Exit For
        End If
         
    Next
    If StatusNIP = False Then
      .Rows = .Rows + 1
       VsFlexRow = fg.Rows - 1
    End If
    '-------------------Cek Tanggal
    For x = 3 To .Cols - 1
         If .TextMatrix(0, x) = tgl1 Then
             StatusTgl = True
             Kolom = x
             Exit For
         End If
    Next
    If StatusTgl = False Then
        .Cols = .Cols + 1
        .ColDataType(.Cols - 1) = flexDTBoolean
        .ColWidth(.Cols - 1) = 300
        .Cols = .Cols + 1
        Kolom = .Cols - 1
    End If
        If status = 1 Then
            .Col = Kolom
            .Row = VsFlexRow
            .CellBackColor = vbGreen
            .TextMatrix(VsFlexRow, Kolom - 1) = "-1"
        End If
        .TextMatrix(VsFlexRow, 0) = VsFlexRow
        If VSFlexGrid2.TextMatrix(Row - 1, 1) = VSFlexGrid2.TextMatrix(Row, 1) And VSFlexGrid2.TextMatrix(Row - 1, 3) = VSFlexGrid2.TextMatrix(Row, 3) Then
           VSFlexGrid2.TextMatrix(Row, 5) = CCur(VSFlexGrid2.TextMatrix(Row, 5)) + CCur(VSFlexGrid2.TextMatrix(Row - 1, 5))
           VSFlexGrid2.TextMatrix(Row, 10) = CCur(VSFlexGrid2.TextMatrix(Row, 10)) + CCur(VSFlexGrid2.TextMatrix(Row - 1, 10))
        End If
        .TextMatrix(VsFlexRow, Kolom) = VSFlexGrid2.TextMatrix(Row, 5) & " + " & VSFlexGrid2.TextMatrix(Row, 10)
       .TextMatrix(VsFlexRow, 1) = VSFlexGrid2.TextMatrix(Row, 3)
       .TextMatrix(VsFlexRow, 2) = VSFlexGrid2.TextMatrix(Row, 4)
       .TextMatrix(0, Kolom) = tgl1
       
       .Col = Kolom
       .Row = VsFlexRow
       .CellAlignment = flexAlignCenterCenter

End With
End Function


Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim J As String
With fg
On Error GoTo Adaerror
If .TextMatrix(Row, Col + 1) = "" Then .TextMatrix(Row, Col) = 0
.Col = Col + 1
.Row = Row
If .CellBackColor = vbGreen Then
   If MsgBox("Apakah Anda Akan Membatalkan Verifikasi ?", vbQuestion + vbYesNo, "Konfirmasi Verifikasi") = vbNo Then
      .TextMatrix(Row, Col) = "-1"
       Exit Sub
    Else
        .CellBackColor = vbWhite
        J = Cekfg(.TextMatrix(0, Col + 1), .TextMatrix(Row, 1), .TextMatrix(Row, Col))
    End If
Else
    J = Cekfg(.TextMatrix(0, Col + 1), .TextMatrix(Row, 1), .TextMatrix(Row, Col))
End If
End With
Exit Sub
Adaerror:
Exit Sub
End Sub
Private Sub fg_DblClick()
'Dim J As String
'With fg
'If .Col = 2 Then
'        J = FrmAccPMdetail.ShowDetail1(DTPicker1, DTPicker2, .TextMatrix(.Row, 1), cboFlex.Text)
'          FrmAccPMdetail.Show vbModal
'        Exit Sub
'    End If
'End With
End Sub
Private Sub fg_Click()
Dim J As String
With fg
    On Error Resume Next
    
    If .TextMatrix(.Row, .Col + 1) = "" Then .Editable = flexEDNone: Exit Sub
    If .ColDataType(.Col) = flexDTBoolean Then
       .Editable = flexEDKbdMouse
    Else
       .Editable = flexEDNone
    End If
End With
End Sub

Function Cekfg(Tgl As String, NIP As String, Nilai As String) As String
Dim x As Integer
With VSFlexGrid2
  For x = 1 To .Rows - 1
        If Format(.TextMatrix(x, 1), "dd/MM/yy") = Tgl And .TextMatrix(x, 3) = NIP Then
            .TextMatrix(x, 9) = Abs(Nilai)
        End If
  Next
End With
End Function

Private Sub VSProject_Click()
Showdata (VSProject.TextMatrix(VSProject.Row, 1))
End Sub

