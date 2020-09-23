VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRekapTimesheet 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Rekap Timesheet"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4575
      Left            =   0
      TabIndex        =   15
      ToolTipText     =   "Double Klik Kolom Project Untuk Melihat PM Dan Mengirim Pesan"
      Top             =   2640
      Width           =   6735
      _cx             =   11880
      _cy             =   8070
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
      FormatString    =   $"frmRekapTimesheet.frx":0000
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
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   11685
      TabIndex        =   9
      Top             =   8145
      Width           =   11685
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
         TabIndex        =   10
         Top             =   0
         Width           =   1215
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5520
      OleObjectBlob   =   "frmRekapTimesheet.frx":00DF
      Top             =   1200
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   2775
      Left            =   6840
      TabIndex        =   14
      Top             =   2640
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
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRekapTimesheet.frx":0313
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
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2625
      ScaleWidth      =   11655
      TabIndex        =   0
      Top             =   0
      Width           =   11685
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Format 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   32
         Top             =   2280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Format 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   31
         Top             =   1920
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmRekapTimesheet.frx":03F2
         Left            =   1680
         List            =   "frmRekapTimesheet.frx":03FF
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2160
         Width           =   1575
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   9840
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   26
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   9840
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   22
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   9840
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   21
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   9840
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   20
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmRekapTimesheet.frx":0416
         Left            =   1680
         List            =   "frmRekapTimesheet.frx":0426
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   3855
      End
      Begin VSFlex8Ctl.VSFlexGrid cboFlex 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   1800
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
         FormatString    =   $"frmRekapTimesheet.frx":0488
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
      Begin VB.CommandButton Command2 
         Caption         =   "&Export"
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
         Left            =   6960
         TabIndex        =   7
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
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
         Left            =   6960
         MaskColor       =   &H00C0FFC0&
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
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
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   20643843
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
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
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   20643843
         CurrentDate     =   39931
      End
      Begin VSFlex8Ctl.VSFlexGrid Combo2 
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Top             =   1440
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
         FormatString    =   $"frmRekapTimesheet.frx":04B1
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
         Left            =   6480
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20643842
         CurrentDate     =   39940
      End
      Begin VSFlex8Ctl.VSFlexGrid CboPM 
         Height          =   315
         Left            =   1680
         TabIndex        =   18
         Top             =   1080
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
         FormatString    =   $"frmRekapTimesheet.frx":04DA
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
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Status TS"
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
         TabIndex        =   30
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Tidak Terdaftar Di Master Project / Tidak Diisi"
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
         Left            =   10440
         TabIndex        =   27
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Belum Diverifikasi PM / Divisi"
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
         Left            =   10440
         TabIndex        =   25
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Diverifikasi PM Dan Belum Diverifikasi Divisi"
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
         Left            =   10440
         TabIndex        =   24
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Diverifikasi PM Dan Diverifikasi Divisi"
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
         Left            =   10440
         TabIndex        =   23
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "NIP PM"
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
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "No Project"
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
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "NIP Karyawan"
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
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Status Verifikasi"
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
         TabIndex        =   5
         Top             =   600
         Width           =   1455
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
         Left            =   2520
         TabIndex        =   2
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
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   3255
      Left            =   6840
      TabIndex        =   28
      Top             =   5520
      Width           =   7935
      _cx             =   13996
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
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRekapTimesheet.frx":0503
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
Attribute VB_Name = "frmRekapTimesheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTimesheet As New Recordset
Dim StatusHari As String


Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
If Option1.Value = True Then
    Setgrid
    Showdata
Else
    Setgrid2
    Showdata2
End If
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Dim x As String
With fg
If .Rows > 1 Then
        .AddItem "", 1
        .AddItem "", 2
        .Redraw = flexRDNone
        .Redraw = flexRDBuffered
        
    For lCol = 1 To .Cols - 1
         
       .TextMatrix(2, lCol) = .TextMatrix(0, lCol)
        .Row = 2
        .Col = lCol
        .CellBackColor = vbGreen '&HE0E0E0
    Next
        .TextMatrix(1, 2) = "Rekap Timesheet Dari " & Format(DTPicker1, "dd MMM yyyy") & " - " & Format(DTPicker2, "dd MMM yyyy") & "  "
        .SaveGrid "C:\RekapTS.xls", flexFileExcel, False
        x = ShellExecute(Me.hwnd, "open", "C:\RekapTS.xls", vbNullString, "C:\RekapTS.xls", 1)
       
          .RemoveItem (1)
          .RemoveItem (1)
End If
End With
 
End Sub

Private Sub fg_DblClick()
Dim J As String
With fg
If .TextMatrix(.Row, .Col) = "Telat" Then Exit Sub
If .Col >= 6 And .TextMatrix(.Row, .Col) <> "" Then
    J = FrmPM.Showdata(.TextMatrix(.Row, .Col), KodeDivisi)
    FrmPM.NIP = .TextMatrix(.Row, 4)
    FrmPM.show vbModal
End If
End With
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker1.Value = DateSerial(Year(Now), Month(Now), 26)
    DTPicker1.Value = DateAdd("M", -1, DTPicker1.Value)
    DTPicker2.Value = DateSerial(Year(Now), Month(Now), 25)
    DTPicker2.Value = DateAdd("M", 0, DTPicker2.Value)
    
    Setgrid
     If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
      Skin1.ApplySkin hwnd
    End If
    AddKaryawan
    AddProject
    AddPM
    Combo1.ListIndex = 0
    Select Case UCase(strGroup)
        Case "USER"
                CboPM.Enabled = False
                cboFlex.Enabled = False
                cboFlex.Text = StrUser
        Case "PM"
                CboPM.Enabled = False
                CboPM.Text = StrUser
    End Select
    Combo3.ListIndex = 1
End Sub
Private Sub AddKaryawan()

    Dim Cboid     As String
    Dim Cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    Cboid1 = vbNullString
    If StrUser = "3578" Or strGroup = "Keuangan" Then
        StrSQL = "select * from Karyawan Where Status <> '14' Order By NIP"
    Else
        StrSQL = "select * from Karyawan Where kd_divisi = '" & KodeDivisi & "' And Status <> '14' Order By NIP"
    End If
    Rscek.Open StrSQL, CN, adOpenStatic
    Cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("NIP") & vbTab & Rscek("Nama")
      Cboid1 = Cboid1 + Cboid
      Rscek.MoveNext
    Loop
    cboFlex.ColComboList(0) = Cboid1
    cboFlex.CellAlignment = flexAlignLeftCenter
End Sub
Private Sub AddProject()

    Dim Cboid     As String
    Dim Cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    Cboid1 = vbNullString
     StrSQL = "select Kode,Nama from project " & _
        "where  kd_divisi = '" & KodeDivisi & "'" & _
        "group by kode,Nama " & _
        "order by kode"

    Rscek.Open StrSQL, CN, adOpenStatic
    Cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("Kode") & vbTab & Rscek("Nama")
      Cboid1 = Cboid1 + Cboid
      Rscek.MoveNext
    Loop
    Combo2.ColComboList(0) = Cboid1
End Sub
Private Sub AddPM()

    Dim Cboid     As String
    Dim Cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    Cboid1 = vbNullString
    StrSQL = " SELECT Distinct Project.Nip_PM, karyawan.Nama FROM Project INNER JOIN karyawan ON Project.Nip_PM = karyawan.NIP Where Project.Kd_Divisi = '" & KodeDivisi & "' Order By Karyawan.Nama"
    Rscek.Open StrSQL, CN, adOpenStatic
    Cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("Nip_PM") & vbTab & Rscek("Nama")
      Cboid1 = Cboid1 + Cboid
      Rscek.MoveNext
    Loop
    CboPM.ColComboList(0) = Cboid1
    CboPM.CellAlignment = flexAlignLeftCenter
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        CmdClose.Width = Me.Width - 100
With fg
    .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture3.Height
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
    Set frmRekapTimesheet = Nothing
End Sub


Sub Showdata()
Dim x As Integer
Dim i, J As Integer
Dim Split, JmlLoop As Integer
Dim Jam As String
Dim IDWaktu As String
Dim Jam1, Jam2 As Date
Dim TglAwal, TglAkhir As Date
Dim RsTS As New ADODB.Recordset
    TglAwal = DTPicker1
    TglAwal = DateSerial(Year(DTPicker1), Month(DTPicker1), 1)
    TglAwal = DateAdd("M", 0, DTPicker1)
    TglAwal = Format(DTPicker1, "MM") & "/01/" & Format(DTPicker1, "yyyy")
    TglAkhir = DTPicker1
    TglAkhir = DateSerial(Year(DTPicker1), Month(DTPicker1), 1)
    TglAkhir = DateAdd("M", 1, TglAkhir) - 1
  VSFlexGrid1.Rows = 1
VSFg.Rows = 1
VSFg.Cols = 10
With VSFlexGrid1
        .Rows = 1
    If RsTS.State = adStateOpen Then RsTS.Close
    StrSQL = "SELECT tbltimesheet.IDtimesheet,tbltimesheet.Tanggal,tbltimesheet.JamAwal As [Jam Awal],JamAkhir AS [Jam Akhir],tbltimesheet.Status,tbltimesheet.NoProject,tbltimesheet.Keterangan,tbltimesheet.Tanggal,tbltimesheet.Masuk,tbltimesheet.NIP,tbltimesheet.StatusPM,tbltimesheet.StatusDivisi,Project.Nip_PM,Project.Status AS StatusPr"
    StrSQL = StrSQL & " FROM tbltimesheet INNER JOIN Project ON tbltimesheet.NoProject = Project.Kode Where tbltimesheet.Tanggal Between '" & Format(DTPicker1, "yyyy/MM/dd") & "' And '" & Format(DTPicker2, "yyyy/MM/dd") & "'"
    If Trim(cboFlex) <> "" Then StrSQL = StrSQL & " And tbltimesheet.NIP = '" & cboFlex & "' "
    If Trim(Combo2) <> "" Then StrSQL = StrSQL & " And tbltimesheet.NoProject = '" & Combo2 & "' "
    Select Case Combo1.Text
       Case "Belum Diverifikasi PM"
            StrSQL = StrSQL & " And tbltimesheet.StatusPM = 0 And tbltimesheet.StatusDivisi =0"
       Case "Diverifikasi PM & Belum Diverifikasi Divisi"
            StrSQL = StrSQL & " And tbltimesheet.StatusPM = 1 And tbltimesheet.StatusDivisi =0"
       Case "Diverifikasi Divisi"
            StrSQL = StrSQL & " And tbltimesheet.StatusPM = 1 And tbltimesheet.StatusDivisi =1"
    End Select
    If Combo3.Text = "Actual" Then
        StrSQL = StrSQL & " And tbltimesheet.Status = 'Actual'"
    ElseIf Combo3.Text = "Plan" Then
        StrSQL = StrSQL & " And tbltimesheet.Status = 'Plan'"
    End If
    If Trim(CboPM) <> "" Then StrSQL = StrSQL & " And Project.NIP_PM = '" & CboPM & "'"
    StrSQL = StrSQL & " Order By tbltimesheet.Tanggal DESC,tbltimesheet.NIP,tbltimesheet.Status,tbltimesheet.IDtimesheet ASC"
    RsTS.Open StrSQL, CN, adOpenStatic
    Set .DataSource = RsTS
    .ColDataType(2) = flexDTDate
    .ColFormat(2) = "dd/MM/yyyy"
    For Lrow = 1 To .Rows - 1
        If .TextMatrix(Lrow, 3) = "" Then .TextMatrix(Lrow, 3) = "00:00"
         If .TextMatrix(Lrow, 4) = "" Then .TextMatrix(Lrow, 4) = "00:00"
 
                JmlLoop = 0
                Jam1 = CDate(.TextMatrix(Lrow, 3))
                If .TextMatrix(Lrow, 4) = "24:00" Then .TextMatrix(Lrow, 4) = "00:00"
                Jam2 = CDate(.TextMatrix(Lrow, 4))
                DTPicker3.Value = Jam1
                Do Until JmlLoop = 50
                    If Format(DTPicker3.Value, "hh:mm") = Format(Jam2, "hh:mm") Then Exit Do
                    VSFg.Rows = VSFg.Rows + 1
                     
                    VSFg.TextMatrix(VSFg.Rows - 1, 0) = VSFg.Rows - 1
                    IDWaktu = Format(DTPicker3.Value, "HH:mm")
'                    If Trim(.TextMatrix(Lrow, 7)) = "Timesheet" And IDWaktu = "17:00" Then IDWaktu = "16:30"
                    VSFg.TextMatrix(VSFg.Rows - 1, 1) = IDWaktu
                    VSFg.TextMatrix(VSFg.Rows - 1, 2) = Format(.TextMatrix(Lrow, 2), "dd/MM/yyyy")
                    VSFg.TextMatrix(VSFg.Rows - 1, 3) = .TextMatrix(Lrow, 6)
                    VSFg.TextMatrix(VSFg.Rows - 1, 4) = .TextMatrix(Lrow, 10)
                    VSFg.TextMatrix(VSFg.Rows - 1, 5) = .TextMatrix(Lrow, 5)
                    VSFg.TextMatrix(VSFg.Rows - 1, 6) = Format(.TextMatrix(Lrow, 9), "HH:mm")
                    VSFg.TextMatrix(VSFg.Rows - 1, 7) = .TextMatrix(Lrow, 7)
                    VSFg.TextMatrix(VSFg.Rows - 1, 8) = .TextMatrix(Lrow, 11)
                    VSFg.TextMatrix(VSFg.Rows - 1, 9) = .TextMatrix(Lrow, 12)
                    JmlLoop = JmlLoop + 1
                    DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                    If Format(DTPicker3, "HH:mm") = "12:30" Then DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                Loop
    Next
End With
    
'Disiplit Perhari
Dim Project As String, AdaKolom As Boolean
Dim JmlCol, PosisiKol As Integer
Dim PosisiJam As String

With fg
.Rows = 1
If VSFg.Rows >= 2 Then .Rows = 2

For i = 1 To VSFg.Rows - 1
    Jam = VSFg.TextMatrix(i, 1)

    
'    ------tampilkan di fg
    If i <> VSFg.Rows - 1 Then
        If VSFg.TextMatrix(i, 4) = VSFg.TextMatrix(i + 1, 4) Then

            If VSFg.TextMatrix(i, 2) = VSFg.TextMatrix(i + 1, 2) And VSFg.TextMatrix(i, 5) = VSFg.TextMatrix(i + 1, 5) Then
                J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
            Else
                 J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
                 .Rows = .Rows + 1
           End If
        Else
                
                J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
                .Rows = .Rows + 1
         End If
    Else
        If VSFg.TextMatrix(i, 4) <> .TextMatrix(.Rows - 1, 4) Then
             J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
        Else
             J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
      End If
    End If
   
Next
     
    For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = i
         If Trim(.TextMatrix(i, 5)) = "Plan" Then
            For J = 1 To .Cols - 1
                .Row = i
                .Col = J
                .CellBackColor = &HC0FFFF
            Next
       End If
          For J = 6 To .Cols - 1
                .ColWidth(J) = 1300
                    Jam1 = Format(.TextMatrix(i, 3), "HH:mm")
                    Jam2 = Left(.TextMatrix(0, J), 5)
                    Jam2 = Format(Jam2, "HH:mm")
                    If Jam2 < Jam1 And .TextMatrix(i, J) = "" Then
                        .TextMatrix(i, J) = "Telat"
                        .Col = J
                        .Row = i
                        .CellBackColor = &HE0E0E0
                    ElseIf .TextMatrix(i, J) = "" Then
                        .Col = J
                        .Row = i
                        .CellBackColor = &HFFC0C0
                        .TextMatrix(i, J) = "?"
                    End If
                    
                    If Jam2 < "08:00" Then
                            If .TextMatrix(i, J) = "Telat" Then
                                .TextMatrix(i, J) = "?"
                                .Col = J
                                .Row = i
                                .CellBackColor = &HFFC0C0
                            End If
                     End If
                    
            Next

       
    Next
    .ColDataType(2) = flexDTDate
    .ColFormat(2) = "dd/MM/yyyy"
    .ColWidth(3) = 800
    .ColWidth(14) = 0
    .Row = 0
End With
End Sub

Function Tampil(ByVal Jam As String, ByVal i As Integer, status As String, StatusPM As String, StatusDivisi As String)
Dim Project As String, AdaKolom As Boolean
Dim JmlCol, PosisiKol As Integer
Dim PosisiJam As String

With fg
 
    .TextMatrix(.Rows - 1, 2) = Format(VSFg.TextMatrix(i, 2), "dd/MM/yyyy")
    .TextMatrix(.Rows - 1, 3) = VSFg.TextMatrix(i, 6)
    .TextMatrix(.Rows - 1, 4) = VSFg.TextMatrix(i, 4)
    Project = VSFg.TextMatrix(i, 3)
    If status <> "" Then .TextMatrix(.Rows - 1, 5) = status
    If Project = "Telat*" Then Project = "Telat"
    If Jam <> "12:00" Then
        AdaKolom = False
        For JmlCol = 6 To .Cols - 1
            PosisiJam = Left(.TextMatrix(0, JmlCol), 5)
            If PosisiJam = "" Then Exit For
            If PosisiJam = Jam Then
                AdaKolom = True
                PosisiKol = JmlCol
            End If
        Next

            If AdaKolom = False Then
                .Cols = .Cols + 1
                 PosisiKol = .Cols - 1
                DTPicker3.Value = DateAdd("n", 30, Jam)
                .TextMatrix(0, PosisiKol) = Jam & "-" & Format(DTPicker3, "HH:mm")
            End If
        
         DTPicker3.Value = DateAdd("n", 30, Jam)
         .TextMatrix(.Rows - 1, PosisiKol) = Project
         If Trim(status) = "Actual" Then
            If StatusPM = 1 And StatusDivisi = 0 Then
                  .Row = .Rows - 1
                  .Col = PosisiKol
                  .CellBackColor = &HC0FFC0
            ElseIf StatusPM = 1 And StatusDivisi = 1 Then
                  .Row = .Rows - 1
                  .Col = PosisiKol
                  .CellBackColor = &HFF&
            End If
        End If
    End If
  End With
End Function
Sub Setgrid()
Dim Cboid, Cboid1 As String
Dim i As Integer

With fg

    .Rows = 1
    .Cols = 23
    .TextMatrix(0, 0) = "No"
    .TextMatrix(0, 1) = "Do"
    .TextMatrix(0, 2) = "Tanggal"
    .TextMatrix(0, 3) = "Masuk"
    .TextMatrix(0, 4) = "NIP"
    .TextMatrix(0, 5) = "Status"
    .TextMatrix(0, 6) = "08:00-08:30"
    .TextMatrix(0, 7) = "08:30-09:00"
    .TextMatrix(0, 8) = "09:00-09:30"
    .TextMatrix(0, 9) = "09:30-10:00"
    .TextMatrix(0, 10) = "10:00-10:30"
    .TextMatrix(0, 11) = "10:30-11:00"
    .TextMatrix(0, 12) = "11:00-11:30"
    .TextMatrix(0, 13) = "11:30-12:00"
    .TextMatrix(0, 14) = "12:00-13:00"
    .TextMatrix(0, 15) = "13:00-13:30"
    .TextMatrix(0, 16) = "13:30-14:00"
    .TextMatrix(0, 17) = "14:00-14:30"
    .TextMatrix(0, 18) = "14:30-15:00"
    .TextMatrix(0, 19) = "15:00-15:30"
    .TextMatrix(0, 20) = "15:30-16:00"
    .TextMatrix(0, 21) = "16:00-16:30"
    .TextMatrix(0, 22) = "16:30-17:00"
    .ColWidth(0) = 500
    .ColWidth(1) = 0
    .ColWidth(2) = 1200
    .ColWidth(3) = 700
    .ColWidth(4) = 700
     .ColWidth(5) = 800
     .ColWidth(14) = 0
    .ColWidth(15) = 0
    For i = 6 To 21
       .ColWidth(i) = 1200
       .ColComboList(i) = Cboid1
    Next
     .ColDataType(1) = flexDTBoolean
    .ColDataType(2) = flexDTDate
    .ColFormat(2) = "dd/MM/yyyy"
    .ColWidth(14) = 0
    .Editable = flexEDKbdMouse
    .FrozenCols = 5
    .Editable = flexEDNone
End With
End Sub
Sub Setgrid2()
Dim Cboid, Cboid1 As String
Dim i As Integer
'VSFlexGrid1.Rows = 1
'VSFG.Rows = 1
'VSFG.Cols = 10
With fg
   
    .Rows = 18
    .RowHeight(0) = 500
    .Cols = 4
    .TextMatrix(0, 0) = "Jam/Tanggal"
    .TextMatrix(0, 1) = "Do"
    
    .TextMatrix(0, 2) = "NIP"
    .TextMatrix(0, 3) = "Status"
    .FixedAlignment(2) = flexAlignCenterCenter
    .TextMatrix(1, 0) = "08:00-08:30"
    .TextMatrix(2, 0) = "08:30-09:00"
    .TextMatrix(3, 0) = "09:00-09:30"
    .TextMatrix(4, 0) = "09:30-10:00"
    .TextMatrix(5, 0) = "10:00-10:30"
    .TextMatrix(6, 0) = "10:30-11:00"
    .TextMatrix(7, 0) = "11:00-11:30"
    .TextMatrix(8, 0) = "11:30-12:00"
    .TextMatrix(9, 0) = "12:00-13:00"
    .TextMatrix(10, 0) = "13:00-13:30"
    .TextMatrix(11, 0) = "13:30-14:00"
    .TextMatrix(12, 0) = "14:00-14:30"
    .TextMatrix(13, 0) = "14:30-15:00"
    .TextMatrix(14, 0) = "15:00-15:30"
    .TextMatrix(15, 0) = "15:30-16:00"
    .TextMatrix(16, 0) = "16:00-16:30"
    .TextMatrix(17, 0) = "16:30-17:00"
    .ColWidth(0) = 1500
    .ColWidth(1) = 0
    .ColWidth(2) = 900
    .ColWidth(3) = 700
    
         
    .ColDataType(1) = flexDTBoolean
    .ColDataType(2) = flexDTDate
    .ColFormat(2) = "dd/MM/yyyy"
      
End With
End Sub
Sub Showdata2()
Dim x As Integer
Dim i, J As Integer
Dim Split, JmlLoop As Integer
Dim Jam As String
Dim IDWaktu As String
Dim Jam1, Jam2 As Date
Dim TglAwal, TglAkhir As Date
Dim RsTS As New ADODB.Recordset
    TglAwal = DTPicker1
    TglAwal = DateSerial(Year(DTPicker1), Month(DTPicker1), 1)
    TglAwal = DateAdd("M", 0, DTPicker1)
    TglAwal = Format(DTPicker1, "MM") & "/01/" & Format(DTPicker1, "yyyy")
    TglAkhir = DTPicker1
    TglAkhir = DateSerial(Year(DTPicker1), Month(DTPicker1), 1)
    TglAkhir = DateAdd("M", 1, TglAkhir) - 1
VSFlexGrid1.Rows = 1
VSFg.Rows = 1
VSFg.Cols = 10
With VSFlexGrid1
        .Rows = 1
    If RsTS.State = adStateOpen Then RsTS.Close
    StrSQL = "SELECT tbltimesheet.IDtimesheet,tbltimesheet.Tanggal,tbltimesheet.JamAwal As [Jam Awal],JamAkhir AS [Jam Akhir],tbltimesheet.Status,tbltimesheet.NoProject,tbltimesheet.Keterangan,tbltimesheet.Tanggal,tbltimesheet.Masuk,tbltimesheet.NIP,tbltimesheet.StatusPM,tbltimesheet.StatusDivisi,Project.Nip_PM,Project.Status AS StatusPr"
    StrSQL = StrSQL & " FROM tbltimesheet INNER JOIN Project ON tbltimesheet.NoProject = Project.Kode Where tbltimesheet.Tanggal Between '" & Format(DTPicker1, "yyyy/MM/dd") & "' And '" & Format(DTPicker2, "yyyy/MM/dd") & "'"
    If Trim(cboFlex) <> "" Then StrSQL = StrSQL & " And tbltimesheet.NIP = '" & cboFlex & "' "
    If Trim(Combo2) <> "" Then StrSQL = StrSQL & " And tbltimesheet.NoProject = '" & Combo2 & "' "
    Select Case Combo1.Text
       Case "Belum Diverifikasi PM"
            StrSQL = StrSQL & " And tbltimesheet.StatusPM = 0 And tbltimesheet.StatusDivisi =0"
       Case "Diverifikasi PM & Belum Diverifikasi Divisi"
            StrSQL = StrSQL & " And tbltimesheet.StatusPM = 1 And tbltimesheet.StatusDivisi =0"
       Case "Diverifikasi Divisi"
            StrSQL = StrSQL & " And tbltimesheet.StatusPM = 1 And tbltimesheet.StatusDivisi =1"
    End Select
    If Combo3.Text = "Actual" Then
        StrSQL = StrSQL & " And tbltimesheet.Status = 'Actual'"
    ElseIf Combo3.Text = "Plan" Then
        StrSQL = StrSQL & " And tbltimesheet.Status = 'Plan'"
    End If
    If Trim(CboPM) <> "" Then StrSQL = StrSQL & " And Project.NIP_PM = '" & CboPM & "'"
    StrSQL = StrSQL & " Order By tbltimesheet.Tanggal DESC,tbltimesheet.NIP,tbltimesheet.Status,tbltimesheet.IDtimesheet ASC"
    RsTS.Open StrSQL, CN, adOpenStatic
    Set .DataSource = RsTS
    .ColDataType(2) = flexDTDate
    .ColFormat(2) = "dd/MM/yyyy"
    For Lrow = 1 To .Rows - 1
        If .TextMatrix(Lrow, 3) = "" Then .TextMatrix(Lrow, 3) = "00:00"
         If .TextMatrix(Lrow, 4) = "" Then .TextMatrix(Lrow, 4) = "00:00"
 
                JmlLoop = 0
                Jam1 = CDate(.TextMatrix(Lrow, 3))
                If .TextMatrix(Lrow, 4) = "24:00" Then .TextMatrix(Lrow, 4) = "00:00"
                Jam2 = CDate(.TextMatrix(Lrow, 4))
                DTPicker3.Value = Jam1
                Do Until JmlLoop = 50
                    If Format(DTPicker3.Value, "hh:mm") = Format(Jam2, "hh:mm") Then Exit Do
                    VSFg.Rows = VSFg.Rows + 1
                     
                    VSFg.TextMatrix(VSFg.Rows - 1, 0) = VSFg.Rows - 1
                    IDWaktu = Format(DTPicker3.Value, "HH:mm")
'                    If Trim(.TextMatrix(Lrow, 7)) = "Timesheet" And IDWaktu = "17:00" Then IDWaktu = "16:30"
                    VSFg.TextMatrix(VSFg.Rows - 1, 1) = IDWaktu
                    VSFg.TextMatrix(VSFg.Rows - 1, 2) = Format(.TextMatrix(Lrow, 2), "dd/MM/yyyy")
                    VSFg.TextMatrix(VSFg.Rows - 1, 3) = .TextMatrix(Lrow, 6)
                    VSFg.TextMatrix(VSFg.Rows - 1, 4) = .TextMatrix(Lrow, 10)
                    VSFg.TextMatrix(VSFg.Rows - 1, 5) = .TextMatrix(Lrow, 5)
                    VSFg.TextMatrix(VSFg.Rows - 1, 6) = Format(.TextMatrix(Lrow, 9), "HH:mm")
                    VSFg.TextMatrix(VSFg.Rows - 1, 7) = .TextMatrix(Lrow, 7)
                    VSFg.TextMatrix(VSFg.Rows - 1, 8) = .TextMatrix(Lrow, 11)
                    VSFg.TextMatrix(VSFg.Rows - 1, 9) = .TextMatrix(Lrow, 12)
                    JmlLoop = JmlLoop + 1
                    DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                    If Format(DTPicker3, "HH:mm") = "12:30" Then DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                Loop
    Next
End With
    
'Disiplit Perhari
Dim Project As String, AdaKolom As Boolean
Dim JmlCol, PosisiKol As Integer
Dim PosisiJam As String

With fg
Setgrid2
 
For i = 1 To VSFg.Rows - 1
    Jam = VSFg.TextMatrix(i, 1)

    
'    ------tampilkan di fg
    If i <> VSFg.Rows - 1 Then
        If VSFg.TextMatrix(i, 4) = VSFg.TextMatrix(i + 1, 4) Then

            If VSFg.TextMatrix(i, 2) = VSFg.TextMatrix(i + 1, 2) And VSFg.TextMatrix(i, 5) = VSFg.TextMatrix(i + 1, 5) Then
                J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
            Else
                 J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
                 .Rows = .Rows + 1
           End If
        Else
                
                J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
                .Rows = .Rows + 1
         End If
    Else
        If VSFg.TextMatrix(i, 4) <> .TextMatrix(.Rows - 1, 4) Then
             J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
        Else
             J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
      End If
    End If
   
Next
     
    For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = i
         If Trim(.TextMatrix(i, 5)) = "Plan" Then
            For J = 1 To .Cols - 1
                .Row = i
                .Col = J
                .CellBackColor = &HC0FFFF
            Next
       End If
          For J = 6 To .Cols - 1
                .ColWidth(J) = 1300
                    Jam1 = Format(.TextMatrix(i, 3), "HH:mm")
                    Jam2 = Left(.TextMatrix(0, J), 5)
                    Jam2 = Format(Jam2, "HH:mm")
                    If Jam2 < Jam1 And .TextMatrix(i, J) = "" Then
                        .TextMatrix(i, J) = "Telat"
                        .Col = J
                        .Row = i
                        .CellBackColor = &HE0E0E0
                    ElseIf .TextMatrix(i, J) = "" Then
                        .Col = J
                        .Row = i
                        .CellBackColor = &HFFC0C0
                        .TextMatrix(i, J) = "?"
                    End If
                    
                    If Jam2 < "08:00" Then
                            If .TextMatrix(i, J) = "Telat" Then
                                .TextMatrix(i, J) = "?"
                                .Col = J
                                .Row = i
                                .CellBackColor = &HFFC0C0
                            End If
                     End If
                    
            Next

       
    Next
    .ColDataType(2) = flexDTDate
    .ColFormat(2) = "dd/MM/yyyy"
    .ColWidth(3) = 800
    .ColWidth(14) = 0
    .Row = 0
End With
End Sub

Function Tampil2(ByVal Jam As String, ByVal i As Integer, status As String, StatusPM As String, StatusDivisi As String)
Dim Project As String, AdaKolom As Boolean
Dim JmlCol, PosisiKol As Integer
Dim PosisiJam As String

With fg
 
    .TextMatrix(.Rows - 1, 2) = Format(VSFg.TextMatrix(i, 2), "dd/MM/yyyy")
    .TextMatrix(.Rows - 1, 3) = VSFg.TextMatrix(i, 6)
    .TextMatrix(.Rows - 1, 4) = VSFg.TextMatrix(i, 4)
    Project = VSFg.TextMatrix(i, 3)
    If status <> "" Then .TextMatrix(.Rows - 1, 5) = status
    If Project = "Telat*" Then Project = "Telat"
    If Jam <> "12:00" Then
        AdaKolom = False
        For JmlCol = 6 To .Cols - 1
            PosisiJam = Left(.TextMatrix(0, JmlCol), 5)
            If PosisiJam = "" Then Exit For
            If PosisiJam = Jam Then
                AdaKolom = True
                PosisiKol = JmlCol
            End If
        Next

            If AdaKolom = False Then
                .Cols = .Cols + 1
                 PosisiKol = .Cols - 1
                DTPicker3.Value = DateAdd("n", 30, Jam)
                .TextMatrix(0, PosisiKol) = Jam & "-" & Format(DTPicker3, "HH:mm")
            End If
        
         DTPicker3.Value = DateAdd("n", 30, Jam)
         .TextMatrix(.Rows - 1, PosisiKol) = Project
         If Trim(status) = "Actual" Then
            If StatusPM = 1 And StatusDivisi = 0 Then
                  .Row = .Rows - 1
                  .Col = PosisiKol
                  .CellBackColor = &HC0FFC0
            ElseIf StatusPM = 1 And StatusDivisi = 1 Then
                  .Row = .Rows - 1
                  .Col = PosisiKol
                  .CellBackColor = &HFF&
            End If
        End If
    End If
  End With
End Function
Private Sub Option1_Click()
Setgrid
End Sub

Private Sub Option2_Click()
Setgrid2
End Sub
