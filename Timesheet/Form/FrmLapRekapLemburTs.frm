VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLapRekapLemburTs 
   Caption         =   "Rekap Lembur"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   10275
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   9495
      _cx             =   16748
      _cy             =   6165
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
      FormatString    =   $"FrmLapRekapLemburTs.frx":0000
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
   Begin VSFlex8Ctl.VSFlexGrid VsDetail 
      Height          =   3495
      Left            =   0
      TabIndex        =   32
      Top             =   1800
      Width           =   6495
      _cx             =   11456
      _cy             =   6165
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
      FormatString    =   $"FrmLapRekapLemburTs.frx":00DF
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
      ExplorerBar     =   0
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
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   10275
      TabIndex        =   24
      Top             =   6555
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
         Left            =   0
         TabIndex        =   25
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1785
      ScaleWidth      =   10245
      TabIndex        =   1
      Top             =   0
      Width           =   10275
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option2"
         Height          =   255
         Left            =   3600
         TabIndex        =   34
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3600
         TabIndex        =   33
         Top             =   720
         Value           =   -1  'True
         Width           =   255
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   375
         Left            =   9840
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20709378
         CurrentDate     =   39940
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Print out"
         Height          =   375
         Left            =   8400
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   8400
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Export To Sheet"
         Height          =   375
         Left            =   6960
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Before Sheet"
         Height          =   375
         Left            =   13440
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Open File Gaji"
         Height          =   375
         Left            =   6960
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Next Sheet"
         Height          =   375
         Left            =   12360
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4800
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   9840
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20709378
         CurrentDate     =   39940
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   720
         TabIndex        =   10
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
         Format          =   20709379
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2880
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
         Format          =   20709379
         CurrentDate     =   39931
      End
      Begin VSFlex8Ctl.VSFlexGrid cboFlex 
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   1320
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
         FormatString    =   $"FrmLapRekapLemburTs.frx":01BE
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
         Left            =   1560
         TabIndex        =   13
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
         FormatString    =   $"FrmLapRekapLemburTs.frx":01E7
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
         Left            =   9840
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20709378
         CurrentDate     =   39940
      End
      Begin VSFlex8Ctl.VSFlexGrid CboDivisi 
         Height          =   315
         Left            =   1560
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
         FormatString    =   $"FrmLapRekapLemburTs.frx":0210
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
         Left            =   360
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker DTPicker6 
         Height          =   375
         Left            =   9840
         TabIndex        =   31
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20709378
         CurrentDate     =   39940
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   36
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Detail"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   35
         Top             =   720
         Width           =   615
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
         TabIndex        =   23
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
         Left            =   120
         TabIndex        =   22
         Top             =   1320
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
         Left            =   120
         TabIndex        =   21
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
         TabIndex        =   20
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
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   495
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
         Left            =   5880
         TabIndex        =   18
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "* Total Bersih = Jam Mulai - Jam Akhir - Istirahat"
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
         Left            =   5880
         TabIndex        =   17
         Top             =   1320
         Width           =   8895
      End
      Begin VB.Label Label10 
         Caption         =   "* Biaya Lembur = (Nilai Gaji/173) * Total Bersih * Upah Perjam"
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
         Left            =   5880
         TabIndex        =   16
         Top             =   1560
         Width           =   8895
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "FrmLapRekapLemburTs.frx":0239
      Top             =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid VSGaji 
      Height          =   1455
      Left            =   0
      TabIndex        =   29
      ToolTipText     =   "Double Klik Kolom Project Untuk Melihat PM"
      Top             =   1800
      Width           =   5655
      _cx             =   9975
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
      FormatString    =   $"FrmLapRekapLemburTs.frx":046D
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
      Height          =   2655
      Left            =   0
      TabIndex        =   27
      Top             =   6360
      Width           =   15135
      _cx             =   26696
      _cy             =   4683
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
      FormatString    =   $"FrmLapRekapLemburTs.frx":054C
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
      Height          =   3495
      Left            =   9600
      TabIndex        =   28
      Top             =   2640
      Width           =   5295
      _cx             =   9340
      _cy             =   6165
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
      FormatString    =   $"FrmLapRekapLemburTs.frx":062B
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
      Height          =   855
      Left            =   9600
      TabIndex        =   26
      Top             =   1800
      Width           =   4215
      _cx             =   7435
      _cy             =   1508
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
      FormatString    =   $"FrmLapRekapLemburTs.frx":070A
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
Attribute VB_Name = "FrmLapRekapLemburTs"
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
Dim JmlGanda As Integer
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
'    If MsgBox("File Name Data Gaji Tidak Sama Dengan Tanggal Pencarian, Apakah Anda Akan Melanjutkan Proses ini ?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then
'           Exit Sub
'    End If
'End If
'
Command1.Enabled = False
Setgrid
Showdata
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Dim x As String
On Error GoTo Adaerror
If Option1.Value = True Then
    With fg
    
    If .Rows > 1 Then
         
         
            .AddItem "", 1
            .AddItem "", 2
            
            .Redraw = flexRDNone
            .Redraw = flexRDBuffered
            .TextMatrix(1, 1) = "Rekap Biaya Lembur Periode " & Format(DTPicker1, "dd MMM yyyy") & " s/d " & Format(DTPicker2, "dd MMM yyyy")
        For lCol = 1 To .Cols - 1
            
           .TextMatrix(2, lCol) = .TextMatrix(0, lCol)
            .Row = 2
            .Col = lCol
            .CellBackColor = vbGreen '&HE0E0E0
        Next
            .SaveGrid "C:\BiayaLembur.xls", flexFileExcel, False
            x = ShellExecute(Me.hwnd, "open", "C:\BiayaLembur.xls", vbNullString, "C:\BiayaLembur.xls", 1)
              .RemoveItem (1)
              .RemoveItem (1)
    End If
    End With
Else
    With VsDetail
  
    If .Rows > 1 Then
         
         
            .AddItem "", 1
            .AddItem "", 2
            
            .Redraw = flexRDNone
            .Redraw = flexRDBuffered
            .TextMatrix(1, 1) = "Rekap Biaya Lembur Periode " & Format(DTPicker1, "dd MMM yyyy") & " s/d " & Format(DTPicker2, "dd MMM yyyy")
        For lCol = 1 To .Cols - 1
            
           .TextMatrix(2, lCol) = .TextMatrix(0, lCol)
            .Row = 2
            .Col = lCol
            .CellBackColor = vbGreen '&HE0E0E0
        Next
            .SaveGrid "C:\BiayaLemburTotal.xls", flexFileExcel, False
    '        Shell PathOffice & "C:\BiayaPerProject2.xls", vbNormalFocus
    x = ShellExecute(Me.hwnd, "open", "C:\BiayaLemburTotal.xls", vbNullString, "C:\BiayaLemburTotal.xls", 1)
              .RemoveItem (1)
              .RemoveItem (1)
    End If
    End With
End If
Exit Sub
Adaerror:
MsgBox err.Description
End Sub

Private Sub Command3_Click()
On Error Resume Next

If Option1 = True Then If fg.Rows > 1 Then fg.PrintGrid "Biaya Per Project2 - Periode " & DTPicker1.Value & " S/D " & DTPicker2.Value, 2, 2, 900, 500: Exit Sub
If Option2 = True Then If VsDetail.Rows > 1 Then VsDetail.PrintGrid "Biaya Per Project2 - Periode " & DTPicker1.Value & " S/D " & DTPicker2.Value, 2, 2, 900, 500

End Sub

Private Sub Command4_Click()

Showgaji
'fg.Visible = False
'VSGaji.Visible = True
'MsgBox CDate(DTPicker4) + CDate(DTPicker3)
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
    DTPicker1.Value = DateSerial(Year(Now), Month(Now), 1)
    DTPicker1.Value = DateAdd("M", -1, DTPicker1.Value)
    DTPicker2.Value = DateSerial(Year(Now), Month(Now), 0)
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
.Cols = 16
.Rows = 1
.TextMatrix(0, 0) = "No"
.TextMatrix(0, 1) = "NIP"
.TextMatrix(0, 2) = "Nama"
.TextMatrix(0, 3) = "Tanggal"
.TextMatrix(0, 4) = "L/K"
.TextMatrix(0, 5) = "No.Proyek"
.TextMatrix(0, 6) = "Jam Awal"
.TextMatrix(0, 7) = "Jam Akhir"
.TextMatrix(0, 8) = "Masuk"
.TextMatrix(0, 9) = "Keluar"
.TextMatrix(0, 10) = "Jam Mulai"
.TextMatrix(0, 11) = "Total Jam"
.TextMatrix(0, 12) = "Total Bersih"
.TextMatrix(0, 13) = "Biaya Lembur"
.TextMatrix(0, 14) = "Upah M/T"
.TextMatrix(0, 15) = "Grand Total"
.ColDataType(3) = flexDTDate

.ColFormat(3) = "dd/MMM/yyyy"
'.ColFormat(12) = "#,##"
.ColFormat(13) = "#,###"
.ColFormat(14) = "#,###"
.ColFormat(15) = "#,###"
.ColFormat(6) = "HH:mm"
.ColFormat(7) = "HH:mm"
.ColFormat(8) = "HH:mm"
.ColFormat(9) = "HH:mm"
.ColFormat(10) = "HH:mm"
.ColFormat(11) = "HH:mm"
.ColFormat(12) = "HH:mm"
.ColWidth(0) = 500
.ColWidth(1) = 700
.ColWidth(2) = 1500
.ColWidth(4) = 1000
.ColDataType(13) = flexDTDouble
.ColDataType(14) = flexDTDouble
.ColDataType(15) = flexDTDouble
.MergeCells = flexMergeFree
.MergeCol(1) = True
.MergeCol(2) = True
.FrozenCols = 5
End With
With VsDetail
    .Rows = 1
    .Cols = 7
    .TextMatrix(0, 0) = "No"
    .TextMatrix(0, 1) = "NIP"
    .TextMatrix(0, 2) = "Nama"
    .TextMatrix(0, 3) = "Kode Proyek"
    .TextMatrix(0, 4) = "Biaya Lembur"
    .TextMatrix(0, 5) = "Upah"
    .TextMatrix(0, 6) = "Grand Total"
    .ColFormat(4) = "#,###"
    .ColFormat(5) = "#,###"
    .ColFormat(6) = "#,###"
   
    .ColWidth(0) = 500
    .ColWidth(1) = 700
    .ColWidth(2) = 1500
    .ColWidth(4) = 1000
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .MergeCells = flexMergeFree
    .MergeCol(1) = True
    .MergeCol(2) = True
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
             .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture2.Height

        End With
        With VSGaji
             .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture2.Height

        End With
         With VsDetail
             .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture2.Height

        End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set FrmLapRekapLemburTs = Nothing
End Sub
Sub Showdata()
Dim x As Integer
Dim Menit, TMenit, Tjam As Currency
Dim Split, JmlLoop As Currency
Dim Jam, Hari As String
Dim IDWaktu, BJam As String
Dim Jam1, Jam2 As Date
Dim Jam3, Jam4 As Date
Dim TglAwal, TglAkhir As Date
Dim RsTS As New ADODB.Recordset

On Error GoTo Adaerror
'On Error Resume Next
If Trim(CboDivisi.Text) = "" And Trim(CboKaryawan.Text) = "" Then
   MsgBox "Silahkan Pilih Divisi / Karyawan Terlebih Dahulu", vbCritical
   cboFlex.SetFocus
   Exit Sub
End If

If StrUser = 3578 Then
    CN.Execute "Delete From Tblcetaklembur2"
Else
    CN.Execute "Delete From Tblcetaklembur"
End If

Command1.Enabled = False
VSJoin.Rows = 1
With VSFlexGrid1

 
     If RsTS.State = adStateOpen Then RsTS.Close
    StrSQL = "SELECT tbltimesheet.NIP, Karyawan.Nama,tbltimesheet.Tanggal,tbltimesheet.NoProject,tbltimesheet.Hari,"
    StrSQL = StrSQL & " tbltimesheet.JamAwal, tbltimesheet.JamAkhir,Absensi.masuk, Absensi.keluar,tbltimesheet.TotalKerja AS JamMulai,tbltimesheet.TotalKerja AS Total_Jam,tbltimesheet.TotalKerja AS TotalBersih,tbltimesheet.TotalKerja AS Biaya,tbltimesheet.TotalKerja AS Upah "
    StrSQL = StrSQL & " FROM tbltimesheet INNER JOIN Karyawan ON tbltimesheet.NIP = Karyawan.NIP INNER JOIN Absensi ON tbltimesheet.NIP = Absensi.NIP AND tbltimesheet.Tanggal = Absensi.Tgl"
    StrSQL = StrSQL & " Where tbltimesheet.Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And '" & Format(DTPicker2, "MM/dd/yyyy") & "'  And tbltimesheet.Status ='Actual' And  tbltimesheet.NoProject <> '' And tbltimesheet.Keterangan ='Lembur'"
    If Trim(cboFlex) <> "" Then StrSQL = StrSQL & " AND tbltimesheet.NoProject = '" & cboFlex & "'"
    If Trim(CboKaryawan) <> "" Then StrSQL = StrSQL & " AND tbltimesheet.Nip = '" & CboKaryawan & "'"
    If Trim(CboDivisi) <> "" Then StrSQL = StrSQL & " And  karyawan.Kd_Divisi = '" & CboDivisi & "'"
    StrSQL = StrSQL & " Order By tbltimesheet.NIP,tbltimesheet.Tanggal,tbltimesheet.IDTimesheet ASC"
    RsTS.Open StrSQL, CN, adOpenStatic
    Set .DataSource = RsTS
        .Rows = .Rows + 1
    .Cols = .Cols + 5
    .TextMatrix(0, .Cols - 1) = "warna"
    For Lrow = 1 To .Rows - 1
'        If .TextMatrix(Lrow, 1) = "3387" Then
'           MsgBox "aa"
'        End If
           
        If .TextMatrix(Lrow, 1) = "" Then .RemoveItem (Lrow): Exit For
        .TextMatrix(Lrow, 0) = Lrow
        .TextMatrix(0, 0) = .Rows
        .TextMatrix(Lrow, 12) = 0
        .TextMatrix(Lrow, 13) = 0
        .TextMatrix(Lrow, 14) = 0
        If .TextMatrix(Lrow, 5) = "Kerja" Then
            Jam3 = Format(.TextMatrix(Lrow, 8), "HH:mm")
           
            If Format(.TextMatrix(Lrow, 8), "mm/dd/yyyy") > "08/11/2010" And Format(.TextMatrix(Lrow, 8), "mm/dd/yyyy") < "09/08/2010" Then
               Jam4 = DateAdd("h", 9, Jam3)
               Jam4 = DateAdd("n", -30, Jam4)
               .TextMatrix(Lrow, 10) = Jam4
                If Format(.TextMatrix(Lrow, 10), "HH:mm") < "16:30" Then .TextMatrix(Lrow, 10) = "16:30"
               .TextMatrix(Lrow, 6) = "16:30"
            Else
                Jam4 = DateAdd("h", 9, Jam3)
                .TextMatrix(Lrow, 10) = Jam4
                If Format(.TextMatrix(Lrow, 10), "HH:mm") < "17:00" Then .TextMatrix(Lrow, 10) = "17:00"
            End If
        Else
            
                .TextMatrix(Lrow, 10) = Format(.TextMatrix(Lrow, 8), "HH:mm")
           
        End If
        
        Menit = 60
        Jam2 = Format(.TextMatrix(Lrow, 9), "hh:mm")
        
    '-----------------Cek Total Jam-----------------
      If Lrow = 1 Then
        If .Rows >= 3 Then
            If .TextMatrix(Lrow, 1) = .TextMatrix(Lrow + 1, 1) And .TextMatrix(Lrow, 3) = .TextMatrix(Lrow + 1, 3) Then
                .TextMatrix(Lrow, 11) = ((DateDiff("n", CDate(.TextMatrix(Lrow, 6)), CDate(.TextMatrix(Lrow, 7))) / 60)) 'berdasarkan timeshhett
              Else
                .TextMatrix(Lrow, 11) = ((DateDiff("n", CDate(.TextMatrix(Lrow, 10)), CDate(Jam2)) / 60)) 'berdasarkan Absen
    
              End If
        Else
'            If CDate(Jam2) > CDate(.TextMatrix(Lrow, 10)) Then
                    .TextMatrix(Lrow, 11) = ((DateDiff("n", CDate(.TextMatrix(Lrow, 10)), CDate(Jam2)) / 60)) 'berdasarkan Absen
'                Else
'                    .TextMatrix(Lrow, 11) = 0
'            End If
        End If
        
      Else
          If .TextMatrix(Lrow, 1) = .TextMatrix(Lrow - 1, 1) And .TextMatrix(Lrow, 3) = .TextMatrix(Lrow - 1, 3) Then
                .TextMatrix(Lrow, 11) = ((DateDiff("n", CDate(.TextMatrix(Lrow, 6)), CDate(.TextMatrix(Lrow, 7))) / 60)) 'berdasarkan timeshhett
          ElseIf .TextMatrix(Lrow, 1) = .TextMatrix(Lrow + 1, 1) And .TextMatrix(Lrow, 3) = .TextMatrix(Lrow + 1, 3) Then
                .TextMatrix(Lrow, 11) = ((DateDiff("n", CDate(.TextMatrix(Lrow, 6)), CDate(.TextMatrix(Lrow, 7))) / 60)) 'berdasarkan timeshhett

          Else
'                If CDate(Jam2) > CDate(.TextMatrix(Lrow, 10)) Then
                    .TextMatrix(Lrow, 11) = ((DateDiff("n", CDate(.TextMatrix(Lrow, 10)), CDate(Jam2)) / 60)) 'berdasarkan Absen
'                Else
'                    .TextMatrix(Lrow, 11) = 0
'                End If
          End If
      End If
   '----------------------------------
            .TextMatrix(Lrow, 17) = 0
         If .TextMatrix(Lrow, 11) < 0 Then .TextMatrix(Lrow, 11) = 24 + .TextMatrix(Lrow, 11)
        .TextMatrix(Lrow, 11) = Round(.TextMatrix(Lrow, 11), 3) * 60
       .TextMatrix(Lrow, 11) = Round(.TextMatrix(Lrow, 11), 0)
'       If .TextMatrix(Lrow, 11) = 60 Then .TextMatrix(Lrow, 11) = "1.00"
       If .TextMatrix(Lrow, 11) >= 60 Then
            Tjam = 0
            
            Do
                TMenit = .TextMatrix(Lrow, 11) - 60
                If TMenit < 60 Then
                    Tjam = Tjam + 1
                    If Len(TMenit) = 1 Then TMenit = "0" & TMenit
                    .TextMatrix(Lrow, 11) = Tjam & "." & TMenit
                    Exit Do
                End If
                Tjam = Tjam + 1
                .TextMatrix(Lrow, 11) = TMenit
            Loop
         Else
            .TextMatrix(Lrow, 11) = 0
         End If
         
         If .TextMatrix(Lrow, 11) < 10 Then .TextMatrix(Lrow, 11) = "0" & .TextMatrix(Lrow, 11)
         If .TextMatrix(Lrow, 11) = 60 Then .TextMatrix(Lrow, 11) = "01.00"
         If Len(.TextMatrix(Lrow, 11)) = 2 Then .TextMatrix(Lrow, 11) = "00." & .TextMatrix(Lrow, 11)
        .TextMatrix(Lrow, 14) = 0
        
        If Lrow = 1 Then
           .TextMatrix(Lrow, 15) = 0
           .TextMatrix(Lrow, 16) = .TextMatrix(Lrow, 11)
           If CDbl(.TextMatrix(Lrow, 11)) > 0 Then GetLembur (.TextMatrix(Lrow, 1)), Lrow
        Else
            
            If CekGanda(Lrow) = False Then
'                If .TextMatrix(Lrow, 15) = 0 Then
'                  Jam2 = Format(.TextMatrix(Lrow, 9), "hh:mm")
'                  If Format(.TextMatrix(Lrow, 7), "hh:mm") > Format(.TextMatrix(Lrow, 7), "hh:mm") Then
                                    
                    If .TextMatrix(Lrow, 1) = .TextMatrix(Lrow + 1, 1) Then
                        If CDate(.TextMatrix(Lrow, 3)) = CDate(.TextMatrix(Lrow + 1, 3)) Then
                            .TextMatrix(Lrow, 11) = ((DateDiff("n", CDate(.TextMatrix(Lrow, 10)), CDate(.TextMatrix(Lrow, 7))) / 60))
        '                  Else
        '                    .TextMatrix(Lrow, 11) = ((DateDiff("n", CDate(.TextMatrix(Lrow, 10)), CDate(Jam2)) / 60)) 'berdasarkan Absen
        '                  End If
                           If .TextMatrix(Lrow, 11) < 0 Then .TextMatrix(Lrow, 11) = 24 + .TextMatrix(Lrow, 11)
                                .TextMatrix(Lrow, 11) = Round(.TextMatrix(Lrow, 11), 3) * 60
                                .TextMatrix(Lrow, 11) = Round(.TextMatrix(Lrow, 11), 0)
                            If .TextMatrix(Lrow, 11) >= 60 Then
                              Tjam = 0
        
                                Do
                                    TMenit = .TextMatrix(Lrow, 11) - 60
                                    If TMenit < 60 Then
                                        Tjam = Tjam + 1
                                        If Len(TMenit) = 1 Then TMenit = "0" & TMenit
                                        .TextMatrix(Lrow, 11) = Tjam & "." & TMenit
                                        Exit Do
                                    End If
                                    Tjam = Tjam + 1
                                    .TextMatrix(Lrow, 11) = TMenit
                                Loop
                             Else
'                                .TextMatrix(Lrow, 11) = 0
                            End If
                         
                                If .TextMatrix(Lrow, 11) < 10 Then .TextMatrix(Lrow, 11) = "0" & .TextMatrix(Lrow, 11)
                                If .TextMatrix(Lrow, 11) = 60 Then .TextMatrix(Lrow, 11) = "01.00"
                                If Len(.TextMatrix(Lrow, 11)) = 2 Then .TextMatrix(Lrow, 11) = "00." & .TextMatrix(Lrow, 11)
                            End If
                        End If
                           If CDbl(.TextMatrix(Lrow, 11)) > 0 Then GetLembur (.TextMatrix(Lrow, 1)), Lrow
              

                End If

            End If
          
          
          Command1.Caption = .TextMatrix(Lrow, 1)
  
    Next
End With
'Exit Sub
 
    
    If VSFlexGrid1.Rows > 2 Then
      For Lrow = 1 To VSFlexGrid1.Rows - 1
        If VSFlexGrid1.TextMatrix(Lrow, 15) >= 1 Then
'                GetLembur (VSFlexGrid1.TextMatrix(Lrow, 1)), Lrow
                VSJoin.Rows = VSJoin.Rows + 1
                VSJoin.TextMatrix(VSJoin.Rows - 1, 0) = Lrow
                VSJoin.TextMatrix(VSJoin.Rows - 1, 1) = VSFlexGrid1.TextMatrix(Lrow, 1)
                VSJoin.TextMatrix(VSJoin.Rows - 1, 2) = VSFlexGrid1.TextMatrix(Lrow, 3)
                VSJoin.TextMatrix(VSJoin.Rows - 1, 3) = VSFlexGrid1.TextMatrix(Lrow, 5)
                VSJoin.TextMatrix(VSJoin.Rows - 1, 4) = VSFlexGrid1.TextMatrix(Lrow, 11)
                VSJoin.TextMatrix(VSJoin.Rows - 1, 5) = VSFlexGrid1.TextMatrix(Lrow, 11)
                VSJoin.TextMatrix(VSJoin.Rows - 1, 9) = VSFlexGrid1.TextMatrix(Lrow, 17)
                VSFlexGrid1.TextMatrix(Lrow, 12) = "00:00"
                If VSJoin.Rows = 2 Then
                    JmlGanda = 1
                Else
                    If VSFlexGrid1.TextMatrix(Lrow, 1) <> VSFlexGrid1.TextMatrix(Lrow - 1, 1) Or VSFlexGrid1.TextMatrix(Lrow - 1, 3) <> VSFlexGrid1.TextMatrix(Lrow, 3) Then
                        JmlGanda = JmlGanda + 1
                    End If
                End If
                 VSJoin.TextMatrix(VSJoin.Rows - 1, 6) = JmlGanda
                VSFlexGrid1.TextMatrix(Lrow, 19) = "Kuning"
                 If VSJoin.Rows > 2 Then
                    If VSJoin.TextMatrix(VSJoin.Rows - 1, 1) <> "" Then
                         If VSJoin.TextMatrix(VSJoin.Rows - 1, 6) <> VSJoin.TextMatrix(VSJoin.Rows - 2, 6) Then
                
                            VSJoin.AddItem "Total", VSJoin.Rows - 1
                            VSJoin.TextMatrix(VSJoin.Rows - 2, 4) = Format(DTPicker5, "hh:mm")
                            VSJoin.TextMatrix(VSJoin.Rows - 2, 1) = VSJoin.TextMatrix(VSJoin.Rows - 3, 1)
                            VSJoin.TextMatrix(VSJoin.Rows - 2, 2) = VSJoin.TextMatrix(VSJoin.Rows - 3, 2)
                            DTPicker5 = VSJoin.TextMatrix(VSJoin.Rows - 1, 4)
                            DTPicker6 = "00:00"
                            GetLembur2 VSJoin.TextMatrix(VSJoin.Rows - 3, 1), VSJoin.Rows - 2
                            DTPicker5 = VSJoin.TextMatrix(VSJoin.Rows - 1, 4)
                         Else
                            DTPicker6 = VSJoin.TextMatrix(VSJoin.Rows - 1, 4)
                            DTPicker5 = DTPicker5 + DTPicker6
                        End If

                    End If
                 Else
                    DTPicker5 = VSJoin.TextMatrix(VSJoin.Rows - 1, 4)
                End If
          End If
    Next

        If VSJoin.Rows > 2 Then
            VSJoin.Rows = VSJoin.Rows + 1
            VSJoin.TextMatrix(VSJoin.Rows - 1, 0) = "Total"
            VSJoin.TextMatrix(VSJoin.Rows - 1, 1) = VSJoin.TextMatrix(VSJoin.Rows - 2, 1)
            VSJoin.TextMatrix(VSJoin.Rows - 1, 2) = VSJoin.TextMatrix(VSJoin.Rows - 2, 2)
             VSJoin.TextMatrix(VSJoin.Rows - 1, 4) = Format(DTPicker5, "hh:mm")
             VSJoin.Rows = VSJoin.Rows + 1
             If Len(VSJoin.TextMatrix(VSJoin.Rows - 2, 4)) <> 0 Then
                GetLembur2 VSJoin.TextMatrix(VSJoin.Rows - 3, 1), VSJoin.Rows - 2
             End If
            VSJoin.Rows = VSJoin.Rows - 1
        End If
     End If
 
 '-------------hitung split projek grrrrrrrrrrr--------------
Dim jmlnip As Integer
Dim HRow As Integer
jmlnip = 0
For Lrow = 1 To VSJoin.Rows - 1
        If Lrow = 1 Then jmlnip = jmlnip + 1
        If VSJoin.TextMatrix(Lrow, 1) = VSJoin.TextMatrix(Lrow - 1, 1) And VSJoin.TextMatrix(Lrow, 0) <> "Total" Then
            jmlnip = jmlnip + 1
        ElseIf VSJoin.TextMatrix(Lrow, 1) = VSJoin.TextMatrix(Lrow - 1, 1) And VSJoin.TextMatrix(Lrow, 0) = "Total" Then
            If VSJoin.TextMatrix(Lrow - jmlnip, 0) = "Total" Then
               jmlnip = jmlnip - 1
            End If
            HRow = VSJoin.TextMatrix(Lrow - jmlnip, 0)
            VSFlexGrid1.TextMatrix(HRow, 14) = VSJoin.TextMatrix(Lrow, 7)
            Call JmlLembur(VSJoin.TextMatrix(Lrow, 1), Lrow, jmlnip)
            jmlnip = 1
        End If


Next
' Exit Sub

 With VSFlexGrid1
    For Lrow = 1 To .Rows - 1
        If .TextMatrix(Lrow, 14) = "" Then .TextMatrix(Lrow, 14) = 0
          .TextMatrix(Lrow, 13) = Round(.TextMatrix(Lrow, 13), 0)
        If CCur(.TextMatrix(Lrow, 11)) >= 1 Then
                If StrUser = 3578 Then
                    .TextMatrix(Lrow, 11) = Replace(.TextMatrix(Lrow, 11), ".", ":")
                    StrSQL = "Insert Into tblcetaklembur2(NIP,Nama,Tanggal,Keterangan,Kode_Proyek,JamAwal,JamAkhir,Masuk,Keluar,JamMulai,TotalJam,TotalBersih,BiayaLembur,Upah,Warna)"
                    StrSQL = StrSQL & "Values('" & .TextMatrix(Lrow, 1) & "','" & .TextMatrix(Lrow, 2) & "','" & Format(.TextMatrix(Lrow, 3), "MM/dd/yyyy") & "','" & .TextMatrix(Lrow, 5) & "','" & .TextMatrix(Lrow, 4) & "',"
                    StrSQL = StrSQL & "'" & Format(.TextMatrix(Lrow, 6), "hh:mm") & "','" & Format(.TextMatrix(Lrow, 7), "hh:mm") & "','" & Format(.TextMatrix(Lrow, 8), "hh:mm") & "','" & Format(.TextMatrix(Lrow, 9), "hh:mm") & "','" & Format(.TextMatrix(Lrow, 10), "hh:mm") & "',"
                    StrSQL = StrSQL & "'" & .TextMatrix(Lrow, 11) & "','" & .TextMatrix(Lrow, 12) & "'," & CCur(.TextMatrix(Lrow, 13)) & "," & CCur(.TextMatrix(Lrow, 14)) & ",'" & .TextMatrix(Lrow, 19) & "')"
                    CN.Execute StrSQL
                Else
                     .TextMatrix(Lrow, 11) = Replace(.TextMatrix(Lrow, 11), ".", ":")
                    StrSQL = "Insert Into tblcetaklembur(NIP,Nama,Tanggal,Keterangan,Kode_Proyek,JamAwal,JamAkhir,Masuk,Keluar,JamMulai,TotalJam,TotalBersih,BiayaLembur,Upah,Warna)"
                    StrSQL = StrSQL & "Values('" & .TextMatrix(Lrow, 1) & "','" & .TextMatrix(Lrow, 2) & "','" & Format(.TextMatrix(Lrow, 3), "MM/dd/yyyy") & "','" & .TextMatrix(Lrow, 5) & "','" & .TextMatrix(Lrow, 4) & "',"
                    StrSQL = StrSQL & "'" & Format(.TextMatrix(Lrow, 6), "hh:mm") & "','" & Format(.TextMatrix(Lrow, 7), "hh:mm") & "','" & Format(.TextMatrix(Lrow, 8), "hh:mm") & "','" & Format(.TextMatrix(Lrow, 9), "hh:mm") & "','" & Format(.TextMatrix(Lrow, 10), "hh:mm") & "',"
                    StrSQL = StrSQL & "'" & .TextMatrix(Lrow, 11) & "','" & .TextMatrix(Lrow, 12) & "'," & CCur(.TextMatrix(Lrow, 13)) & "," & CCur(.TextMatrix(Lrow, 14)) & ",'" & .TextMatrix(Lrow, 19) & "')"
                    CN.Execute StrSQL
                End If
        End If
    Next
End With
   

ShowDetail
ShowTotal

Command1.Caption = "Refresh"
Command1.Enabled = True
Exit Sub

Adaerror:
'Showdata
Command1.Caption = "Refresh"
Command1.Enabled = True
    MsgBox err.Description & " " & Lrow
End Sub
Sub ShowTotal()
With VsDetail
'Exit Sub
Dim Gaji As Double
Dim lembur As Double
Dim tPersen As Double
Dim Nomor, x As Integer
Dim GrandTotal As Double
Dim Uangmakan As Currency
Dim Menit, TMenit, Tjam As Integer
Dim RsCetak As New ADODB.Recordset
    '-------------Proses Pegurutan
    If StrUser = 3578 Then
        StrSQL = "SELECT NIP, Nama, Kode_Proyek, SUM(BiayaLembur) AS BiayaLembur, SUM(Upah) AS Upah from dbo.TblCetakLembur2 GROUP BY NIP, Nama, Kode_Proyek ORDER BY NIP"
    Else
        StrSQL = "SELECT NIP, Nama, Kode_Proyek, SUM(BiayaLembur) AS BiayaLembur, SUM(Upah) AS Upah from dbo.TblCetakLembur GROUP BY NIP, Nama, Kode_Proyek ORDER BY NIP"
    End If
    If RsCetak.State = adStateOpen Then RsCetak.Close
    RsCetak.Open StrSQL, CN, adOpenStatic
    Set VSFg.DataSource = RsCetak
    Set RsCetak = Nothing
    .Rows = 1
    .Rows = 2
    Nomor = 0
     For Lrow = 1 To VSFg.Rows - 1
        If Lrow = 1 Then
            Nomor = Nomor + 1
            .TextMatrix(.Rows - 1, 0) = Nomor
            .TextMatrix(.Rows - 1, 1) = VSFg.TextMatrix(Lrow, 1)
            .TextMatrix(.Rows - 1, 2) = VSFg.TextMatrix(Lrow, 2)
            .TextMatrix(.Rows - 1, 3) = VSFg.TextMatrix(Lrow, 3)
            .TextMatrix(.Rows - 1, 4) = VSFg.TextMatrix(Lrow, 4)
            .TextMatrix(.Rows - 1, 5) = VSFg.TextMatrix(Lrow, 5)
            .TextMatrix(.Rows - 1, 6) = CCur(.TextMatrix(.Rows - 1, 4)) + CCur(.TextMatrix(.Rows - 1, 5))
           
            lembur = 0
            Uangmakan = 0
            GrandTotal = 0
            lembur = lembur + CDbl(VSFg.TextMatrix(Lrow, 4))
            Uangmakan = Uangmakan + CDbl(VSFg.TextMatrix(Lrow, 5))
            GrandTotal = GrandTotal + CCur(.TextMatrix(.Rows - 1, 6))
        Else
            If VSFg.TextMatrix(Lrow, 1) <> VSFg.TextMatrix(Lrow - 1, 1) Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 4) = lembur
                .TextMatrix(.Rows - 1, 5) = Uangmakan
                .TextMatrix(.Rows - 1, 6) = GrandTotal
                .Row = .Rows - 1
                For x = 1 To .Cols - 1
                    .Col = x
                    .CellBackColor = vbGreen
                Next
                Nomor = Nomor + 1
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = Nomor
                .TextMatrix(.Rows - 1, 1) = VSFg.TextMatrix(Lrow, 1)
                .TextMatrix(.Rows - 1, 2) = VSFg.TextMatrix(Lrow, 2)
                .TextMatrix(.Rows - 1, 3) = VSFg.TextMatrix(Lrow, 3)
                .TextMatrix(.Rows - 1, 4) = VSFg.TextMatrix(Lrow, 4)
                .TextMatrix(.Rows - 1, 5) = VSFg.TextMatrix(Lrow, 5)
                .TextMatrix(.Rows - 1, 6) = CCur(.TextMatrix(.Rows - 1, 4)) + CCur(.TextMatrix(.Rows - 1, 5))

                lembur = 0
                Uangmakan = 0
                GrandTotal = 0
                lembur = lembur + CDbl(VSFg.TextMatrix(Lrow, 4))
                Uangmakan = Uangmakan + CDbl(VSFg.TextMatrix(Lrow, 5))
                GrandTotal = GrandTotal + CCur(.TextMatrix(.Rows - 1, 6))

            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = VSFg.TextMatrix(Lrow, 1)
                .TextMatrix(.Rows - 1, 2) = VSFg.TextMatrix(Lrow, 2)
                .TextMatrix(.Rows - 1, 3) = VSFg.TextMatrix(Lrow, 3)
                .TextMatrix(.Rows - 1, 4) = VSFg.TextMatrix(Lrow, 4)
                .TextMatrix(.Rows - 1, 5) = VSFg.TextMatrix(Lrow, 5)
                .TextMatrix(.Rows - 1, 6) = CCur(.TextMatrix(.Rows - 1, 4)) + CCur(.TextMatrix(.Rows - 1, 5))
                lembur = lembur + CDbl(VSFg.TextMatrix(Lrow, 4))
                Uangmakan = Uangmakan + CDbl(VSFg.TextMatrix(Lrow, 5))
                GrandTotal = GrandTotal + CCur(.TextMatrix(.Rows - 1, 6))

            End If
        End If
    Next
        
        
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 4) = lembur
            .TextMatrix(.Rows - 1, 5) = Uangmakan
            .TextMatrix(.Rows - 1, 6) = GrandTotal
            .Row = .Rows - 1
            For x = 1 To .Cols - 1
                .Col = x
                .CellBackColor = vbGreen
            Next
 
End With
End Sub
Sub ShowDetail()
With fg
Dim Gaji As Double
Dim lembur As Double
Dim TotalJam, TotalMenit As Double
Dim tPersen As Double
Dim Nomor, x As Integer
Dim GrandTotal As Double
Dim Uangmakan As Currency
Dim Menit, TMenit, Tjam As Integer
Dim RsCetak As New ADODB.Recordset
    '-------------Proses Pegurutan
'    Exit Sub
    If StrUser = 3578 Then
        StrSQL = "Select * From tblcetakLembur2 Order By NIP,Tanggal"
    Else
        StrSQL = "Select * From tblcetakLembur Order By NIP,Tanggal"
    End If
    If RsCetak.State = adStateOpen Then RsCetak.Close
    RsCetak.Open StrSQL, CN, adOpenStatic
    Set VSFg.DataSource = RsCetak
    Set RsCetak = Nothing
    .Rows = 1
    .Rows = 2
    Nomor = 0
     For Lrow = 1 To VSFg.Rows - 1
        If Lrow = 1 Then
            Nomor = Nomor + 1
            .TextMatrix(.Rows - 1, 0) = Nomor
            .TextMatrix(.Rows - 1, 1) = VSFg.TextMatrix(Lrow, 1)
            .TextMatrix(.Rows - 1, 2) = VSFg.TextMatrix(Lrow, 2)
            .TextMatrix(.Rows - 1, 3) = Format(VSFg.TextMatrix(Lrow, 3), "dd/MMM/yyyy")
            .TextMatrix(.Rows - 1, 4) = VSFg.TextMatrix(Lrow, 4)
            .TextMatrix(.Rows - 1, 5) = VSFg.TextMatrix(Lrow, 5)
            .TextMatrix(.Rows - 1, 6) = VSFg.TextMatrix(Lrow, 6)
            .TextMatrix(.Rows - 1, 7) = VSFg.TextMatrix(Lrow, 7)
            .TextMatrix(.Rows - 1, 8) = VSFg.TextMatrix(Lrow, 8)
            .TextMatrix(.Rows - 1, 9) = VSFg.TextMatrix(Lrow, 9)
            .TextMatrix(.Rows - 1, 10) = VSFg.TextMatrix(Lrow, 10)
            .TextMatrix(.Rows - 1, 11) = VSFg.TextMatrix(Lrow, 11)
            .TextMatrix(.Rows - 1, 12) = Format(VSFg.TextMatrix(Lrow, 12), "hh:mm")
            .TextMatrix(.Rows - 1, 13) = VSFg.TextMatrix(Lrow, 13)
            .TextMatrix(.Rows - 1, 14) = VSFg.TextMatrix(Lrow, 14)
            .TextMatrix(.Rows - 1, 15) = CCur(.TextMatrix(.Rows - 1, 13)) + CCur(.TextMatrix(.Rows - 1, 14))
          
          If VSFg.TextMatrix(Lrow, 15) = "Kuning" Then
            .Row = .Rows - 1
            For x = 6 To .Cols - 1
               .Col = x
               .CellBackColor = vbYellow
            Next
         End If
            
            DTPicker3 = "00:00"
            TotalJam = 0
            TotalMenit = 0
            lembur = 0
            Uangmakan = 0
            GrandTotal = 0
            If Len(VSFg.TextMatrix(Lrow, 12)) = 1 Then VSFg.TextMatrix(Lrow, 12) = "00:00"
            TotalJam = TotalJam + Left(VSFg.TextMatrix(Lrow, 12), 2)
            TotalMenit = TotalMenit + Mid(VSFg.TextMatrix(Lrow, 12), 4, 2)
            lembur = lembur + CDbl(VSFg.TextMatrix(Lrow, 13))
            Uangmakan = Uangmakan + CDbl(VSFg.TextMatrix(Lrow, 14))
            GrandTotal = GrandTotal + CCur(.TextMatrix(.Rows - 1, 15))
        Else
            If VSFg.TextMatrix(Lrow, 1) <> VSFg.TextMatrix(Lrow - 1, 1) Then
                Tjam = 0
                .Rows = .Rows + 1
                 If TotalMenit > 60 Then
                    Do
                        TMenit = TotalMenit - 60
                        If TMenit < 60 Then
                            Tjam = Tjam + 1
                            If Len(TMenit) = 1 Then TMenit = "0" & TMenit
                            TotalJam = TotalJam + Tjam & "." & TMenit
                            Exit Do
                        End If
                        Tjam = Tjam + 1
                        TotalMenit = TMenit
                    Loop
                    .TextMatrix(.Rows - 1, 12) = Replace(TotalJam, ".", ":")
                End If
             
                .TextMatrix(.Rows - 1, 3) = "TOTAL"
               If TotalMenit < 60 Then
                     .TextMatrix(.Rows - 1, 12) = TotalJam & ":" & TotalMenit
                End If
                .TextMatrix(.Rows - 1, 13) = lembur
                .TextMatrix(.Rows - 1, 14) = Uangmakan
                .TextMatrix(.Rows - 1, 15) = GrandTotal
                .Row = .Rows - 1
                For x = 1 To .Cols - 1
                    .Col = x
                    .CellBackColor = vbGreen
                Next
                Nomor = Nomor + 1
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = Nomor
                .TextMatrix(.Rows - 1, 1) = VSFg.TextMatrix(Lrow, 1)
                .TextMatrix(.Rows - 1, 2) = VSFg.TextMatrix(Lrow, 2)
                .TextMatrix(.Rows - 1, 3) = Format(VSFg.TextMatrix(Lrow, 3), "dd/MMM/yyyy")
                .TextMatrix(.Rows - 1, 4) = VSFg.TextMatrix(Lrow, 4)
                .TextMatrix(.Rows - 1, 5) = VSFg.TextMatrix(Lrow, 5)
                .TextMatrix(.Rows - 1, 6) = VSFg.TextMatrix(Lrow, 6)
                .TextMatrix(.Rows - 1, 7) = VSFg.TextMatrix(Lrow, 7)
                .TextMatrix(.Rows - 1, 8) = VSFg.TextMatrix(Lrow, 8)
                .TextMatrix(.Rows - 1, 9) = VSFg.TextMatrix(Lrow, 9)
                .TextMatrix(.Rows - 1, 10) = VSFg.TextMatrix(Lrow, 10)
                .TextMatrix(.Rows - 1, 11) = VSFg.TextMatrix(Lrow, 11)
                .TextMatrix(.Rows - 1, 12) = Format(VSFg.TextMatrix(Lrow, 12), "hh:mm")
                .TextMatrix(.Rows - 1, 13) = VSFg.TextMatrix(Lrow, 13)
                .TextMatrix(.Rows - 1, 14) = VSFg.TextMatrix(Lrow, 14)
                .TextMatrix(.Rows - 1, 15) = CCur(.TextMatrix(.Rows - 1, 13)) + CCur(.TextMatrix(.Rows - 1, 14))
                
                If VSFg.TextMatrix(Lrow, 15) = "Kuning" Then
                .Row = .Rows - 1
                .Row = .Rows - 1
                    For x = 6 To .Cols - 1
                       .Col = x
                       .CellBackColor = vbYellow
                    Next
                End If
              
              If Len(VSFg.TextMatrix(Lrow, 12)) = 1 Then VSFg.TextMatrix(Lrow, 12) = "00:00"
               DTPicker3 = "00:00"
                lembur = 0
                Uangmakan = 0
                TotalJam = 0
                TotalMenit = 0
                GrandTotal = 0
                TotalMenit = TotalMenit + Mid(VSFg.TextMatrix(Lrow, 12), 4, 2)
                TotalJam = TotalJam + Left(VSFg.TextMatrix(Lrow, 12), 2)
                lembur = lembur + CDbl(VSFg.TextMatrix(Lrow, 13))
                Uangmakan = Uangmakan + CDbl(VSFg.TextMatrix(Lrow, 14))
                GrandTotal = GrandTotal + CCur(.TextMatrix(.Rows - 1, 15))

            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = VSFg.TextMatrix(Lrow, 1)
                .TextMatrix(.Rows - 1, 2) = VSFg.TextMatrix(Lrow, 2)
                .TextMatrix(.Rows - 1, 3) = Format(VSFg.TextMatrix(Lrow, 3), "dd/MMM/yyyy")
                .TextMatrix(.Rows - 1, 4) = VSFg.TextMatrix(Lrow, 4)
                .TextMatrix(.Rows - 1, 5) = VSFg.TextMatrix(Lrow, 5)
                .TextMatrix(.Rows - 1, 6) = VSFg.TextMatrix(Lrow, 6)
                .TextMatrix(.Rows - 1, 7) = VSFg.TextMatrix(Lrow, 7)
                .TextMatrix(.Rows - 1, 8) = VSFg.TextMatrix(Lrow, 8)
                .TextMatrix(.Rows - 1, 9) = VSFg.TextMatrix(Lrow, 9)
                .TextMatrix(.Rows - 1, 10) = VSFg.TextMatrix(Lrow, 10)
                .TextMatrix(.Rows - 1, 11) = VSFg.TextMatrix(Lrow, 11)
                .TextMatrix(.Rows - 1, 12) = Format(VSFg.TextMatrix(Lrow, 12), "hh:mm")
                .TextMatrix(.Rows - 1, 13) = VSFg.TextMatrix(Lrow, 13)
                .TextMatrix(.Rows - 1, 14) = VSFg.TextMatrix(Lrow, 14)
                .TextMatrix(.Rows - 1, 15) = CCur(.TextMatrix(.Rows - 1, 13)) + CCur(.TextMatrix(.Rows - 1, 14))
                
                If VSFg.TextMatrix(Lrow, 15) = "Kuning" Then
                    .Row = .Rows - 1
               .Row = .Rows - 1
                    For x = 6 To .Cols - 1
                       .Col = x
                       .CellBackColor = vbYellow
                    Next
                End If
             
             If Len(VSFg.TextMatrix(Lrow, 12)) = 1 Then VSFg.TextMatrix(Lrow, 12) = "00:00"

                TotalMenit = TotalMenit + Mid(VSFg.TextMatrix(Lrow, 12), 4, 2)
                TotalJam = TotalJam + Left(VSFg.TextMatrix(Lrow, 12), 2)
                lembur = lembur + CDbl(VSFg.TextMatrix(Lrow, 13))
                Uangmakan = Uangmakan + CDbl(VSFg.TextMatrix(Lrow, 14))
                GrandTotal = GrandTotal + CCur(.TextMatrix(.Rows - 1, 15))

            End If
        End If
    Next
        
        Tjam = 0
        .Rows = .Rows + 1
         If TotalMenit > 60 Then
            Do
                TMenit = TotalMenit - 60
                If TMenit < 60 Then
                    Tjam = Tjam + 1
                    If Len(TMenit) = 1 Then TMenit = "0" & TMenit
                    TotalJam = TotalJam + Tjam & "." & TMenit
                    Exit Do
                End If
                Tjam = Tjam + 1
                TotalMenit = TMenit
            Loop
            .TextMatrix(.Rows - 1, 12) = Replace(TotalJam, ".", ":")
        End If
             
            .TextMatrix(.Rows - 1, 3) = "TOTAL"
            If TotalMenit < 60 Then
                  .TextMatrix(.Rows - 1, 12) = TotalJam & ":" & TotalMenit
             End If
        .TextMatrix(.Rows - 1, 13) = lembur
        .TextMatrix(.Rows - 1, 14) = Uangmakan
        .TextMatrix(.Rows - 1, 15) = GrandTotal
        .Row = .Rows - 1
        For x = 1 To .Cols - 1
            .Col = x
            .CellBackColor = vbGreen
        Next
 
End With
End Sub
Function CekGanda(Row As Integer) As Boolean
Dim x As Integer
Dim Menit, TMenit, Tjam As Integer
    CekGanda = False
    With VSFlexGrid1
       
            If .TextMatrix(Row, 1) = .TextMatrix(Row - 1, 1) And .TextMatrix(Row, 3) = .TextMatrix(Row - 1, 3) Then
               
                 
                If .TextMatrix(Row - 1, 15) = 0 Then .TextMatrix(Row - 1, 15) = .TextMatrix(Row - 1, 15) + 1
                .TextMatrix(Row, 15) = .TextMatrix(Row - 1, 15) + 1
                   
                
                DTPicker3 = .TextMatrix(Row - 1, 16)
                DTPicker4 = .TextMatrix(Row, 11)
                .TextMatrix(Row, 16) = DTPicker3 + DTPicker4
                .TextMatrix(Row, 16) = Format(.TextMatrix(Row, 16), "hh:mm")
                .TextMatrix(Row, 16) = Replace(.TextMatrix(Row, 16), ":", ".")
                CekGanda = True
            Else
                .TextMatrix(Row, 15) = 0
                .TextMatrix(Row, 16) = .TextMatrix(Row, 11)
            End If
         
       
    End With
End Function
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
            If Trim(Keterangan) = "Lembur" And Format(JamTs, "HH:mm") < Format(VSFg.TextMatrix(Row, 17), "HH:mm") Then
                .TextMatrix(.Rows - 1, 10) = .TextMatrix(.Rows - 1, 10) + 0.5
            End If
            .TextMatrix(.Rows - 1, 11) = VSFg.TextMatrix(Row, 8)
            .TextMatrix(.Rows - 1, 12) = VSFg.TextMatrix(Row, 7)
            .TextMatrix(.Rows - 1, 13) = VSFg.TextMatrix(Row, 14) 'Hari
            .TextMatrix(.Rows - 1, 14) = VSFg.TextMatrix(Row, 12)
            If .TextMatrix(.Rows - 2, 3) <> .TextMatrix(.Rows - 1, 3) Then Call GetSet(.TextMatrix(.Rows - 2, 3), .TextMatrix(.Rows - 1, 3))
'            .TextMatrix(.Rows - 1, 15) = GetSet(.TextMatrix(.Rows - 2, 3), .TextMatrix(.Rows - 1, 3)) = True
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
Function GetLembur(ByVal NIP As String, Grow As Integer)
Dim jamMasuk, Hari As String
Dim JamKeluar, SisaLembur As String
Dim TotalJam, TotalLemburBruto As String
Dim TotalLembur, Terlambat As String
Dim Gapok As Double
Dim rsAbsen As New ADODB.Recordset
Dim RsGaji As New ADODB.Recordset
Dim RsSetting As New ADODB.Recordset
Dim Istirahat As String
Dim Jam1, Jam2 As Double
Dim Jam3, Jam4 As Double
Dim Jam5, Jam6 As Double
Dim JamMakan As Double
Dim Transport As String
Dim Upah1, Upah2, UM, Upah3 As Currency
Dim JamLembur, SplitLembur As Double
Dim NilaiUpah As Currency, Menitlembur As Currency
Dim UpahJam As Double
Dim Pembulatan As Integer
Dim JamB As Integer
'Exit Function
With VSFlexGrid1
         TotalLembur = 0
         Istirahat = 0
         TotalLemburBruto = 0
         UM = 0
         JamLembur = 0
         Hari = .TextMatrix(Grow, 5)
         DTPicker3 = Format(.TextMatrix(Grow, 8), "HH:mm") 'totalLembur
         DTPicker4 = Format(.TextMatrix(Grow, 9), "HH:mm")
         
        
         If Format(DTPicker3, "HH:mm") = "00:00" Then
                jamMasuk = "TIDAK ABSEN"
             Else
                jamMasuk = Format(DTPicker3, "HH:mm")
             End If
             
             If Format(DTPicker4, "HH:mm") = "00:00" Then
                If CDate(.TextMatrix(Grow, 9)) < CDate(.TextMatrix(Grow, 3)) Then
                    JamKeluar = "TIDAK ABSEN"
                    TotalJam = 0
                Else
                    JamKeluar = Format(DTPicker4, "HH:mm")
                    TotalJam = .TextMatrix(Grow, 11)
                End If
            Else
                JamKeluar = Format(DTPicker4, "HH:mm")
                TotalJam = .TextMatrix(Grow, 11)
        End If
'        If NIP = 2925 and  Then MsgBox NIP
'        .TextMatrix(Grow, 17) = DateDiff("n", CDate(.TextMatrix(Grow, 16)), CDate(.TextMatrix(Grow, 11))) / 60
               Transport = GetSet(NIP, NIP)
         If RsSetting.State = adStateOpen Then RsSetting.Close
         RsSetting.Open "SELECT * FROM TblSETTING WHERE TINGKAT = '" & Tingkat & "' AND HARI = '" & Hari & "' AND StatusAktif = 1 And Berlaku_SD >= '" & Format(Date, "mm/dd/yyyy") & "'", CN, adOpenStatic
         If Not RsSetting.EOF Then
             Jam1 = RsSetting!jamlembur1
             Jam2 = RsSetting!jamlembur2
             Jam3 = RsSetting!jamlembur3
             Jam4 = RsSetting!jamlembur4
             Jam5 = RsSetting!Jamlembur5
             Jam6 = RsSetting!Jamlembur6
             UpahJam = RsSetting!UpahPerjam
             UM = RsSetting!upahmakan
             JamMakan = RsSetting!jammakan1
            If JamKeluar = "TIDAK ABSEN" Then
                TotalLembur = 0
             Else
                Dim Jam As Integer
                Dim Menit As Integer
                   .TextMatrix(Grow, 17) = 0
                    .TextMatrix(Grow, 18) = 0
                  Select Case Hari
                      Case "Kerja"
                             TotalLembur = TotalJam
                            TotalLemburBruto = TotalJam
                             DTPicker3 = TotalJam
                            If JamKeluar > "18:00" And JamKeluar <= "18:15" Then
                               TotalLembur = "01.00"
                                DTPicker3 = "01:00"
                                Istirahat = RsSetting!jamist_1
                                TotalLembur = DateAdd("n", Istirahat, DTPicker3)
                                DTPicker3 = DateAdd("n", Istirahat, DTPicker3)
                             ElseIf JamKeluar > "18:15" And JamKeluar <= "18:30" Then
                               TotalLembur = "01.30"
                                DTPicker3 = "01.30"
                                Istirahat = "-" & RsSetting!jamist_1
                              TotalLembur = DateAdd("n", Istirahat, DTPicker3)
                            End If
'                            If CDbl(TotalLembur) >= RsSetting!ist1 And CDbl(TotalLembur) <= RsSetting!ist2 And JamKeluar > "18:15" Or JamKeluar < "05:00" Then
                          If CDbl(TotalLembur) >= RsSetting!ist1 And CDbl(TotalLembur) <= RsSetting!ist2 Then
   
                              Istirahat = "-" & RsSetting!jamist_1
                              TotalLembur = DateAdd("n", Istirahat, DTPicker3)
        
                           ElseIf CDbl(TotalLembur) >= RsSetting!ist3 And CDbl(TotalLembur) <= RsSetting!ist4 Then
                              Istirahat = "-" & RsSetting!jamist_2
                              TotalLembur = DateAdd("n", Istirahat, DTPicker3)
        
                            ElseIf CDbl(TotalLembur) >= RsSetting!ist5 And CDbl(TotalLembur) <= RsSetting!ist6 Then
                              Istirahat = "-" & RsSetting!jamist_3
                              TotalLembur = DateAdd("n", Istirahat, DTPicker3)
        
                            ElseIf CDbl(TotalLembur) >= RsSetting!ist7 And TotalLembur <= RsSetting!ist8 Then
                              Istirahat = "-" & RsSetting!jamist_4
                              TotalLembur = DateAdd("n", Istirahat, DTPicker3)
                            End If
                        Case "Libur"
                                TotalLembur = TotalJam
                                TotalLemburBruto = TotalJam
                                DTPicker3 = TotalJam
                                JamB = Left(TotalLembur, 2)
                                Pembulatan = Right(TotalLembur, 2)
                                If JamB = 4 Then
                                    If Pembulatan <= 30 Then TotalLembur = Left(TotalLembur, 3) & "00"
                                     
                                End If
                                If CDbl(TotalLembur) >= RsSetting!ist1 And CDbl(TotalLembur) <= RsSetting!ist2 Then
                                  Istirahat = "-" & RsSetting!jamist_1
                                  TotalLembur = DateAdd("n", Istirahat, DTPicker3)
            
                               ElseIf CDbl(TotalLembur) >= RsSetting!ist3 And CDbl(TotalLembur) <= RsSetting!ist4 Then
                                  Istirahat = "-" & RsSetting!jamist_2
                                  TotalLembur = DateAdd("n", Istirahat, DTPicker3)
            
                                ElseIf CDbl(TotalLembur) >= RsSetting!ist5 And CDbl(TotalLembur) <= RsSetting!ist6 Then
                                  Istirahat = "-" & RsSetting!jamist_3
                                  TotalLembur = DateAdd("n", Istirahat, DTPicker3)
            
                                ElseIf CDbl(TotalLembur) >= RsSetting!ist7 And TotalLembur <= RsSetting!ist8 Then
                                  Istirahat = "-" & RsSetting!jamist_4
                                  TotalLembur = DateAdd("n", Istirahat, DTPicker3)
            
                                End If
                        End Select
                   
                    .TextMatrix(Grow, 17) = Istirahat
                    .TextMatrix(Grow, 18) = Format(TotalLembur, "hh:mm")
            End If
            TotalLembur = Left(TotalLembur, 5)
            If TotalLembur <> 0 Then DTPicker4 = TotalLembur
            TotalLembur = Replace(TotalLembur, ":", ".")
'            If CDbl(Left(TotalLembur, 2)) < JamMakan Then UM = 0
         End If
         
        
         
         'pembulatan menit
         If TotalLembur > 0 Then
            
            
                Pembulatan = Right(TotalLembur, 2)
            
                If Pembulatan >= 0 And Pembulatan <= 7 Then
                   TotalLembur = Left(TotalLembur, 3) & "00"
                ElseIf Pembulatan >= 8 And Pembulatan <= 22 Then
                   TotalLembur = Left(TotalLembur, 3) & 15
                ElseIf Pembulatan >= 23 And Pembulatan <= 37 Then
                   TotalLembur = Left(TotalLembur, 3) & 30
                ElseIf Pembulatan >= 38 And Pembulatan <= 52 Then
                   TotalLembur = Left(TotalLembur, 3) & 45
                ElseIf Pembulatan >= 53 And Pembulatan <= 60 Then
                       If DTPicker4.Hour <> 23 Then
                        DTPicker4.Hour = DTPicker4.Hour + 1
                        TotalLembur = "0" & DTPicker4.Hour & ".00"
                       End If
                End If
            
              'upah
              JamLembur = Replace(TotalLembur, ":", ".")
               Upah1 = RsSetting!Upah1
               Upah2 = RsSetting!Upah2
               Upah3 = RsSetting!Upah3
               NilaiUpah = 0
'             JamLembur = "09.30"
           
                 If Hari = "Kerja" Then
                        If CDbl(JamLembur) >= Jam1 And CDbl(JamLembur) <= Jam2 Then
                            NilaiUpah = (CDbl(JamLembur) * NilaiGaji / UpahJam) * Upah1
'                            .TextMatrix(Grow, 15) = Upah1
                        ElseIf CDbl(JamLembur) >= Jam3 And CDbl(JamLembur) <= Jam4 Then
                            NilaiUpah = (1 * NilaiGaji / UpahJam) * Upah1
                            SplitLembur = CDbl(JamLembur) - 1
                            If SplitLembur > 0 Then
                               SisaLembur = Replace(JamLembur, ".", ":")
                               DTPicker5 = CDate(SisaLembur)
                               Jam1 = DTPicker5.Hour - 1
                               SisaLembur = DTPicker5.Minute
                               SisaLembur = SisaLembur / 60
                               SplitLembur = Jam1 + SisaLembur
                               NilaiUpah = NilaiUpah + (SplitLembur * NilaiGaji / UpahJam) * Upah2
                            End If
'                            .TextMatrix(Grow, 15) = Upah2
                        End If
                Else

                        If CDbl(JamLembur) >= Jam1 And CDbl(JamLembur) <= Jam2 Then
                               SisaLembur = Replace(JamLembur, ".", ":")
                               DTPicker5 = CDate(SisaLembur)
                               SisaLembur = DTPicker5.Minute
                               SisaLembur = SisaLembur / 60
                               SplitLembur = DTPicker5.Hour + SisaLembur
                               NilaiUpah = (CDbl(SplitLembur) * NilaiGaji / UpahJam) * Upah1
'                            .TextMatrix(Grow, 15) = Upah1
                        ElseIf CDbl(JamLembur) >= Jam3 And CDbl(JamLembur) <= Jam4 Then
                            NilaiUpah = (8 * NilaiGaji / UpahJam) * Upah1
                            SisaLembur = Replace(JamLembur, ".", ":")
                            DTPicker5 = CDate(SisaLembur)
                            Jam1 = DTPicker5.Hour - 8
                            SisaLembur = DTPicker5.Minute
                            SisaLembur = SisaLembur / 60
                            SplitLembur = Jam1 + SisaLembur
'                            SplitLembur = CDbl(JamLembur) - 8
                            NilaiUpah = NilaiUpah + (SplitLembur * NilaiGaji / UpahJam) * Upah2
'                            .TextMatrix(Grow, 15) = Upah2
                         ElseIf CDbl(JamLembur) >= Jam5 And CDbl(JamLembur) <= Jam6 Then
                            NilaiUpah = (8 * NilaiGaji / UpahJam) * Upah1
                            
                            SplitLembur = CDbl(JamLembur) - 8
                            NilaiUpah = NilaiUpah + (1 * NilaiGaji / UpahJam) * Upah2
                            
                            SisaLembur = Replace(JamLembur, ".", ":")
                            DTPicker5 = CDate(SisaLembur)
                            Jam1 = DTPicker5.Hour - 9
                            SisaLembur = DTPicker5.Minute
                            SisaLembur = SisaLembur / 60
                            SplitLembur = Jam1 + SisaLembur
                            NilaiUpah = NilaiUpah + (SplitLembur * NilaiGaji / UpahJam) * Upah3
                        End If
                End If
                .TextMatrix(Grow, 12) = Format(TotalLembur, "hh:mm")
                .TextMatrix(Grow, 18) = Format(TotalLembur, "hh:mm")
                .TextMatrix(Grow, 13) = NilaiUpah
                If TotalLembur >= JamMakan Then
                    .TextMatrix(Grow, 14) = UM
                End If

         End If
        
End With
End Function
Function GetLembur2(ByVal NIP As String, Grow As Integer)
Dim jamMasuk, Hari As String
Dim JamKeluar, SisaLembur As String
Dim TotalJam, TotalLemburBruto As String
Dim TotalLembur, Terlambat As String
Dim Gapok As Double
Dim rsAbsen As New ADODB.Recordset
Dim RsGaji As New ADODB.Recordset
Dim RsSetting As New ADODB.Recordset
Dim Istirahat As String
Dim Jam1, Jam2 As Double
Dim Jam3, Jam4 As Double
Dim Jam5, Jam6 As Double
Dim JamMakan As Double
Dim Transport As String
Dim Upah1, Upah2, UM, Upah3 As Currency
Dim JamLembur, SplitLembur As Double
Dim NilaiUpah As Currency, Menitlembur As Currency
Dim UpahJam As Double
With VSJoin
         TotalLembur = 0
         Istirahat = 0
         TotalLemburBruto = 0
         UM = 0
         JamLembur = 0
       
         Hari = .TextMatrix(Grow - 1, 3)
      
'         DTPicker3 = Format(.TextMatrix(Grow, 8), "HH:mm") 'totalLembur
'         DTPicker4 = Format(.TextMatrix(Grow, 9), "HH:mm")
'         If Format(DTPicker3, "HH:mm") = "00:00" Then
'                jamMasuk = "TIDAK ABSEN"
'             Else
'                jamMasuk = Format(DTPicker3, "HH:mm")
'             End If
'
'             If Format(DTPicker4, "HH:mm") = "00:00" Then
'                JamKeluar = "TIDAK ABSEN"
'                TotalJam = 0
'            Else
'                JamKeluar = Format(DTPicker4, "HH:mm")
                TotalJam = .TextMatrix(Grow, 4)
'        End If
'        .TextMatrix(Grow, 17) = DateDiff("n", CDate(.TextMatrix(Grow, 16)), CDate(.TextMatrix(Grow, 11))) / 60
               Transport = GetSet(NIP, NIP)
          If Tingkat = 0 Then
            MsgBox "Untuk Data Gaji dan Tingkat NIP = " & NIP & " Belum ada", vbCritical
            Exit Function
          End If
         If RsSetting.State = adStateOpen Then RsSetting.Close
         RsSetting.Open "SELECT * FROM TblSETTING WHERE TINGKAT = '" & Tingkat & "' AND HARI = '" & Hari & "' AND StatusAktif = 1 And Berlaku_SD >= '" & Format(Date, "mm/dd/yyyy") & "'", CN, adOpenStatic
         If Not RsSetting.EOF Then
             Jam1 = RsSetting!jamlembur1
             Jam2 = RsSetting!jamlembur2
             Jam3 = RsSetting!jamlembur3
             Jam4 = RsSetting!jamlembur4
             Jam5 = RsSetting!Jamlembur5
             Jam6 = RsSetting!Jamlembur6
             UpahJam = RsSetting!UpahPerjam
             UM = RsSetting!upahmakan
             JamMakan = RsSetting!jammakan1
'            If JamKeluar = "TIDAK ABSEN" Then
'                TotalLembur = 0
'             Else
                Dim Jam As Integer
                Dim Menit As Integer
'
'   If Hari = "Libur" Then
                   TotalLembur = TotalJam
                    TotalLemburBruto = TotalJam
                    DTPicker3 = TotalJam
                  TotalLembur = Replace(TotalLembur, ":", ".")
''                  TotalLembur = CDbl(TotalLembur)
 Select Case Hari
                      Case "Kerja"
'                             TotalLembur = TotalJam
                            TotalLemburBruto = TotalJam
                             DTPicker3 = TotalJam
                            If JamKeluar > "18:00" And JamKeluar <= "18:15" Then
                               TotalLembur = "01.00"
                                DTPicker3 = "01:00"
                                Istirahat = RsSetting!jamist_1
                                TotalLembur = DateAdd("n", Istirahat, DTPicker3)
                                DTPicker3 = DateAdd("n", Istirahat, DTPicker3)
                             ElseIf JamKeluar > "18:15" And JamKeluar <= "18:30" Then
                               TotalLembur = "01.30"
                                DTPicker3 = "01.30"
                                Istirahat = "-" & RsSetting!jamist_1
                              TotalLembur = DateAdd("n", Istirahat, DTPicker3)
                            End If
'                            If CDbl(TotalLembur) >= RsSetting!ist1 And CDbl(TotalLembur) <= RsSetting!ist2 And JamKeluar > "18:15" Or JamKeluar < "05:00" Then
                          If CDbl(TotalLembur) >= RsSetting!ist1 And CDbl(TotalLembur) <= RsSetting!ist2 Then
   
                              Istirahat = "-" & RsSetting!jamist_1
                              TotalLembur = DateAdd("n", Istirahat, DTPicker3)
        
                           ElseIf CDbl(TotalLembur) >= RsSetting!ist3 And CDbl(TotalLembur) <= RsSetting!ist4 Then
                              Istirahat = "-" & RsSetting!jamist_2
                              TotalLembur = DateAdd("n", Istirahat, DTPicker3)
        
                            ElseIf CDbl(TotalLembur) >= RsSetting!ist5 And CDbl(TotalLembur) <= RsSetting!ist6 Then
                              Istirahat = "-" & RsSetting!jamist_3
                              TotalLembur = DateAdd("n", Istirahat, DTPicker3)
        
                            ElseIf CDbl(TotalLembur) >= RsSetting!ist7 And TotalLembur <= RsSetting!ist8 Then
                              Istirahat = "-" & RsSetting!jamist_4
                              TotalLembur = DateAdd("n", Istirahat, DTPicker3)
                            End If
                        Case "Libur"
'                                TotalLembur = TotalJam
                                TotalLemburBruto = TotalJam
                                DTPicker3 = TotalJam
                 
                                If CDbl(TotalLembur) >= RsSetting!ist1 And CDbl(TotalLembur) <= RsSetting!ist2 Then
                                  Istirahat = "-" & RsSetting!jamist_1
                                  TotalLembur = DateAdd("n", Istirahat, DTPicker3)
            
                               ElseIf CDbl(TotalLembur) >= RsSetting!ist3 And CDbl(TotalLembur) <= RsSetting!ist4 Then
                                  Istirahat = "-" & RsSetting!jamist_2
                                  TotalLembur = DateAdd("n", Istirahat, DTPicker3)
            
                                ElseIf CDbl(TotalLembur) >= RsSetting!ist5 And CDbl(TotalLembur) <= RsSetting!ist6 Then
                                  Istirahat = "-" & RsSetting!jamist_3
                                  TotalLembur = DateAdd("n", Istirahat, DTPicker3)
            
                                ElseIf CDbl(TotalLembur) >= RsSetting!ist7 And TotalLembur <= RsSetting!ist8 Then
                                  Istirahat = "-" & RsSetting!jamist_4
                                  TotalLembur = DateAdd("n", Istirahat, DTPicker3)
            
                                End If
                        End Select
             End If
             TotalLembur = Left(TotalLembur, 5)
             If TotalLembur <> 0 Then DTPicker4 = TotalLembur
            TotalLembur = Replace(TotalLembur, ":", ".")
 
         
         'pembulatan menit
         If TotalLembur > 0 Then
'             DTPicker4 = TotalLembur
              'upah
              Dim Pembulatan As Integer
                Pembulatan = Right(TotalLembur, 2)
                If Pembulatan >= 0 And Pembulatan <= 7 Then
                   TotalLembur = Left(TotalLembur, 3) & "00"
                ElseIf Pembulatan >= 8 And Pembulatan <= 22 Then
                   TotalLembur = Left(TotalLembur, 3) & 15
                ElseIf Pembulatan >= 23 And Pembulatan <= 37 Then
                   TotalLembur = Left(TotalLembur, 3) & 30
                ElseIf Pembulatan >= 38 And Pembulatan <= 52 Then
                   TotalLembur = Left(TotalLembur, 3) & 45
                ElseIf Pembulatan >= 53 And Pembulatan <= 60 Then
                       DTPicker4.Hour = DTPicker4.Hour + 1
                       TotalLembur = "0" & DTPicker4.Hour & ".00"
                End If
              JamLembur = Replace(TotalLembur, ":", ".")
               Upah1 = RsSetting!Upah1
               Upah2 = RsSetting!Upah2
               Upah3 = RsSetting!Upah3
               NilaiUpah = 0
'             JamLembur = "09.30"
                  If Hari = "Kerja" Then
                        If CDbl(JamLembur) >= Jam1 And CDbl(JamLembur) <= Jam2 Then
                            NilaiUpah = (CDbl(JamLembur) * NilaiGaji / UpahJam) * Upah1
'                            .TextMatrix(Grow, 15) = Upah1
                        ElseIf CDbl(JamLembur) >= Jam3 And CDbl(JamLembur) <= Jam4 Then
                            NilaiUpah = (1 * NilaiGaji / UpahJam) * Upah1
                            SplitLembur = CDbl(JamLembur) - 1
                            If SplitLembur > 0 Then
                               SisaLembur = Replace(JamLembur, ".", ":")
                               DTPicker5 = CDate(SisaLembur)
                               Jam1 = DTPicker5.Hour - 1
                               SisaLembur = DTPicker5.Minute
                               SisaLembur = SisaLembur / 60
                               SplitLembur = Jam1 + SisaLembur
                               NilaiUpah = NilaiUpah + (SplitLembur * NilaiGaji / UpahJam) * Upah2
                            End If
'                            .TextMatrix(Grow, 15) = Upah2
                        End If
                Else
'                         JamLembur = "10.30"
                        If CDbl(JamLembur) >= Jam1 And CDbl(JamLembur) <= Jam2 Then
                               SisaLembur = Replace(JamLembur, ".", ":")
                               DTPicker5 = CDate(SisaLembur)
                               SisaLembur = DTPicker5.Minute
                               SisaLembur = SisaLembur / 60
                               SplitLembur = DTPicker5.Hour + SisaLembur
                               NilaiUpah = (CDbl(SplitLembur) * NilaiGaji / UpahJam) * Upah1
'                            .TextMatrix(Grow, 15) = Upah1
                        ElseIf CDbl(JamLembur) >= Jam3 And CDbl(JamLembur) <= Jam4 Then
                            NilaiUpah = (8 * NilaiGaji / UpahJam) * Upah1
                            SisaLembur = Replace(JamLembur, ".", ":")
                            DTPicker5 = CDate(SisaLembur)
                            Jam1 = DTPicker5.Hour - 8
                            SisaLembur = DTPicker5.Minute
                            SisaLembur = SisaLembur / 60
                            SplitLembur = Jam1 + SisaLembur
'                            SplitLembur = CDbl(JamLembur) - 8
                            NilaiUpah = NilaiUpah + (SplitLembur * NilaiGaji / UpahJam) * Upah2
'                            .TextMatrix(Grow, 15) = Upah2
                         ElseIf CDbl(JamLembur) >= Jam5 And CDbl(JamLembur) <= Jam6 Then
                            NilaiUpah = (8 * NilaiGaji / UpahJam) * Upah1
                            
                            SplitLembur = CDbl(JamLembur) - 8
                            NilaiUpah = NilaiUpah + (1 * NilaiGaji / UpahJam) * Upah2
                            
                            SisaLembur = Replace(JamLembur, ".", ":")
                            DTPicker5 = CDate(SisaLembur)
                            Jam1 = DTPicker5.Hour - 9
                            SisaLembur = DTPicker5.Minute
                            SisaLembur = SisaLembur / 60
                            SplitLembur = Jam1 + SisaLembur
                            NilaiUpah = NilaiUpah + (SplitLembur * NilaiGaji / UpahJam) * Upah3
                        End If
                End If
                .TextMatrix(Grow, 5) = Format(TotalLembur, "hh:mm")
                .TextMatrix(Grow, 6) = Round(NilaiUpah, 2)
                .TextMatrix(Grow, 9) = 0
                DTPicker3 = .TextMatrix(Grow, 4)
                DTPicker4 = .TextMatrix(Grow, 5)
                .TextMatrix(Grow, 8) = Format(DTPicker3 - DTPicker4, "hh:mm")
                DTPicker3 = "00:00"
                DTPicker4 = "00:00"
'                DTPicker5 = "00:00"
                If TotalLembur >= JamMakan Then
                    .TextMatrix(Grow, 7) = UM
                End If

         End If
        
End With
End Function
Function JmlLembur(ByVal NIP As String, Grow As Integer, Jml As Integer)
Dim jamMasuk, Hari As String
Dim JamKeluar, SisaLembur As String
Dim TotalJam, TotalLemburBruto As String
Dim TotalLembur, Terlambat As String
Dim Gapok As Double
Dim rsAbsen As New ADODB.Recordset
Dim RsGaji As New ADODB.Recordset
Dim RsSetting As New ADODB.Recordset
Dim Istirahat As String
Dim Jam1, Jam2 As Double
Dim Jam3, Jam4 As Double
Dim Jam5, Jam6 As Double
Dim JamMakan As Double
Dim Transport As String
Dim Upah1, Upah2, UM, Upah3 As Currency
Dim JamLembur, SplitLembur, SplitLembur2 As Double
Dim NilaiUpah As Currency, Menitlembur As Currency
Dim UpahJam As Double
With VSJoin
         TotalLembur = 0
         Istirahat = 0
         TotalLemburBruto = 0
         UM = 0
         JamLembur = 0
'       If .TextMatrix(Grow - Jml, 3) = "" Then
            Hari = .TextMatrix(Grow - Jml, 3)
'       Else
'
'       End If
                TotalJam = .TextMatrix(Grow, 5)
                Transport = GetSet(NIP, NIP)
          If Tingkat = 0 Then
'            MsgBox "Untuk Gaji NIP = " & NIP & " Belum ada", vbCritical
            Exit Function
          End If
         If RsSetting.State = adStateOpen Then RsSetting.Close
         RsSetting.Open "SELECT * FROM TblSETTING WHERE TINGKAT = '" & Tingkat & "' AND HARI = '" & Hari & "' AND StatusAktif = 1 And Berlaku_SD >= '" & Format(Date, "mm/dd/yyyy") & "'", CN, adOpenStatic
         If Not RsSetting.EOF Then
             Jam1 = RsSetting!jamlembur1
             Jam2 = RsSetting!jamlembur2
             Jam3 = RsSetting!jamlembur3
             Jam4 = RsSetting!jamlembur4
             Jam5 = RsSetting!Jamlembur5
             Jam6 = RsSetting!Jamlembur6
             UpahJam = RsSetting!UpahPerjam
             UM = RsSetting!upahmakan
             JamMakan = RsSetting!jammakan1
          End If
                 Dim Jam As Integer
                Dim Menit As Integer
                Dim d, jmlasal As Integer
 
                d = 1
                jmlasal = Jml
                Do Until jmlasal = 0
                  If jmlasal = 0 Then Exit Do
                   If Len(.TextMatrix(Grow - jmlasal - 1, 3)) > 0 Then
                      DTPicker6 = .TextMatrix(Grow - jmlasal, 5)
                      DTPicker6 = DTPicker5 + DTPicker6
                      TotalJam = Replace(Format(DTPicker6, "hh:mm"), ":", ".")
                      
                   Else
                        DTPicker5 = .TextMatrix(Grow - jmlasal, 5)
                        TotalJam = .TextMatrix(Grow - jmlasal, 5)
                   End If
                   Terlambat = .TextMatrix(Grow - jmlasal, 5)
                   TotalLembur = TotalJam
                    TotalLemburBruto = TotalJam
                    DTPicker3 = TotalJam
                  TotalLembur = Replace(TotalLembur, ":", ".")
          
                  If CDbl(TotalLembur) >= RsSetting!ist1 And CDbl(TotalLembur) <= RsSetting!ist2 Then
'                    If Len(.TextMatrix(Grow - jmlasal, 9)) = 0 Then
                        Istirahat = "-" & RsSetting!jamist_1
                        TotalLembur = DateAdd("n", Istirahat, DTPicker3)
                        Terlambat = DateAdd("n", Istirahat, Terlambat)
'                    End If
                   ElseIf CDbl(TotalLembur) >= RsSetting!ist3 And CDbl(TotalLembur) <= RsSetting!ist4 Then
                       
                          Istirahat = "-" & RsSetting!jamist_2
                          TotalLembur = DateAdd("n", Istirahat, DTPicker3)
                          Terlambat = DateAdd("n", Istirahat, Terlambat)
                      
                    ElseIf CDbl(TotalLembur) >= RsSetting!ist5 And CDbl(TotalLembur) <= RsSetting!ist6 Then
                    
                        Istirahat = "-" & RsSetting!jamist_3
                        TotalLembur = DateAdd("n", Istirahat, DTPicker3)
                        Terlambat = DateAdd("n", Istirahat, Terlambat)
                   
                    ElseIf CDbl(TotalLembur) >= RsSetting!ist7 And TotalLembur <= RsSetting!ist8 Then
                     
                         Istirahat = "-" & RsSetting!jamist_4
                         TotalLembur = DateAdd("n", Istirahat, DTPicker3)
                        Terlambat = DateAdd("n", Istirahat, Terlambat)
                      End If
             If Grow > 2 Then
                If .TextMatrix(Grow - jmlasal, 9) < 0 Then
                     Istirahat = Abs(Istirahat)
                     TotalLembur = DateAdd("n", Istirahat, DTPicker3)
                     Terlambat = DateAdd("n", Istirahat, Terlambat)
                End If
            End If
            TotalLembur = Terlambat
             If TotalLembur <> 0 Then DTPicker4 = TotalLembur
            TotalLembur = Replace(TotalLembur, ":", ".")
             TotalLembur = Left(TotalLembur, 5)
         
'        74844.99
         
         'pembulatan menit
         If TotalLembur > 0 Then
'             DTPicker4 = TotalLembur
              'upah
              Dim Pembulatan As Integer
                Pembulatan = Right(TotalLembur, 2)
                If Pembulatan >= 0 And Pembulatan <= 7 Then
                   TotalLembur = Left(TotalLembur, 3) & "00"
                ElseIf Pembulatan >= 8 And Pembulatan <= 22 Then
                   TotalLembur = Left(TotalLembur, 3) & 15
                ElseIf Pembulatan >= 23 And Pembulatan <= 37 Then
                   TotalLembur = Left(TotalLembur, 3) & 30
                ElseIf Pembulatan >= 38 And Pembulatan <= 52 Then
                   TotalLembur = Left(TotalLembur, 3) & 45
                ElseIf Pembulatan >= 53 And Pembulatan <= 60 Then
                       DTPicker4.Hour = DTPicker4.Hour + 1
                       TotalLembur = "0" & DTPicker4.Hour & ".00"
                End If
                    JamLembur = Replace(TotalLembur, ":", ".")
                    SisaLembur = Replace(JamLembur, ".", ":")
                    DTPicker3 = CDate(SisaLembur)
                    Jam1 = DTPicker3.Hour
                    SisaLembur = DTPicker3.Minute
                    SisaLembur = SisaLembur / 60
                    SplitLembur = Jam1 + SisaLembur
                    
                    
                    '------------totaljam ke desimal
                    JamLembur = Replace(.TextMatrix(Grow, 5), ":", ".")
                    SisaLembur = Replace(JamLembur, ".", ":")
                    DTPicker3 = CDate(SisaLembur)
                    Jam1 = DTPicker3.Hour
                    SisaLembur = DTPicker3.Minute
                    SisaLembur = SisaLembur / 60
                    SplitLembur2 = Jam1 + SisaLembur
                     
                    NilaiUpah = SplitLembur / SplitLembur2 * .TextMatrix(Grow, 6)
                    
'                .TextMatrix(Grow - jmlasal, 4) = Format(TotalLembur, "hh:mm")
                Gapok = .TextMatrix(Grow - jmlasal, 0)
                VSFlexGrid1.TextMatrix(Gapok, 12) = TotalLembur
                VSFlexGrid1.TextMatrix(Gapok, 13) = Round(NilaiUpah, 2)
                
                .TextMatrix(Grow - jmlasal, 9) = Istirahat
                .TextMatrix(Grow - jmlasal, 5) = TotalLembur
                .TextMatrix(Grow - jmlasal, 6) = Round(NilaiUpah, 2)
                DTPicker3 = .TextMatrix(Grow, 4)
                DTPicker4 = .TextMatrix(Grow, 5)
                .TextMatrix(Grow, 8) = Format(DTPicker3 - DTPicker4, "hh:mm")
                DTPicker3 = "00:00"
                DTPicker4 = "00:00"
 
         End If
            d = d + 1
            jmlasal = jmlasal - 1
         Loop
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

Private Sub Label3_DblClick()
VSGaji.Visible = True
fg.Visible = False
VsDetail.Visible = False
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    VSGaji.Visible = False
    fg.Visible = True
    VsDetail.Visible = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
    VSGaji.Visible = False
    fg.Visible = False
    VsDetail.Visible = True
End If
End Sub
 

Private Sub VSFlexGrid1_DblClick()
MsgBox VSFlexGrid1.Col
End Sub

Private Sub VSGaji_DblClick()
VSGaji.Visible = False
fg.Visible = True
VsDetail.Visible = True
End Sub




