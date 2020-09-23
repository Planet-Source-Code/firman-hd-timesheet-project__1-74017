VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmAccDivisi2 
   Caption         =   "Verifikasi Divisi"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   9855
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   1695
      Left            =   5640
      TabIndex        =   15
      ToolTipText     =   "Double Klik Kolom Project Untuk Melihat PM Dan Mengirim Pesan"
      Top             =   720
      Width           =   6735
      _cx             =   11880
      _cy             =   2990
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
      FormatString    =   $"FrmAccDivisi2.frx":0000
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
   Begin VB.Frame Frame3 
      Caption         =   "Belum Diverifikasi PM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   17
      Top             =   3480
      Width           =   4935
      Begin VSFlex8Ctl.VSFlexGrid VSgrid 
         Height          =   2295
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Klik Untuk Melihat Detail Timesheet"
         Top             =   240
         Width           =   4695
         _cx             =   8281
         _cy             =   4048
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tidak Mengisi Timesheet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      TabIndex        =   10
      Top             =   6120
      Width           =   4935
      Begin VSFlex8Ctl.VSFlexGrid VsIsi 
         Height          =   2175
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Double Klik Untuk Membuat Manual Timesheet"
         Top             =   240
         Width           =   4695
         _cx             =   8281
         _cy             =   3836
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
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9855
      TabIndex        =   8
      Top             =   8115
      Width           =   9855
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
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7560
      OleObjectBlob   =   "FrmAccDivisi2.frx":00DF
      Top             =   2880
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   9825
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton Command2 
         Caption         =   "&Unverifikasi"
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
         Left            =   7560
         TabIndex        =   28
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Verifikasi"
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
         Left            =   6240
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
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
         Left            =   4680
         TabIndex        =   1
         Top             =   120
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   720
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
         Format          =   63176707
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2880
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
         Format          =   63176707
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   9120
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   16761024
         Format          =   63176706
         CurrentDate     =   39940
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sudah Diverifikasi PM / Divisi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   4935
      Begin VSFlex8Ctl.VSFlexGrid fgIsi 
         Height          =   2415
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Checklist Kemudian Klik Untuk Verifikasi/Unverifikasi"
         Top             =   240
         Width           =   4695
         _cx             =   8281
         _cy             =   4260
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
   End
   Begin VSFlex8Ctl.VSFlexGrid VSfg 
      Height          =   735
      Left            =   5520
      TabIndex        =   16
      Top             =   2880
      Width           =   6735
      _cx             =   11880
      _cy             =   1296
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
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   5520
      ScaleHeight     =   855
      ScaleWidth      =   9735
      TabIndex        =   19
      Top             =   9600
      Width           =   9735
      Begin VB.PictureBox Picture7 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4920
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   25
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4920
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   24
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   22
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   20
         Top             =   120
         Width           =   495
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
         Left            =   5520
         TabIndex        =   27
         Top             =   120
         Width           =   4215
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
         Left            =   5520
         TabIndex        =   26
         Top             =   480
         Width           =   4215
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
         Left            =   720
         TabIndex        =   23
         Top             =   480
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
         Left            =   720
         TabIndex        =   21
         Top             =   120
         Width           =   3735
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   6015
      Left            =   5520
      TabIndex        =   14
      Top             =   4680
      Visible         =   0   'False
      Width           =   6735
      _cx             =   11880
      _cy             =   10610
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
      FormatString    =   $"FrmAccDivisi2.frx":0313
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
Attribute VB_Name = "FrmAccDivisi2"
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

Private Sub cmdSave_Click()
Dim Lrow As Long
Dim StrSQL As String
On Error GoTo Adaerror
        With fgIsi
            .Rows = .Rows + 1
        If CekCurek(" Verifikasi Divisi", fgIsi) = False Then Exit Sub

            If MsgBox("Apakah Anda yakin ingin Mengverifikasi Divisi ?", vbQuestion + vbYesNo, "Konfirmasi Simpan Data") = vbNo Then
               Exit Sub
            Else
            Lrow = 1
             Do Until Lrow = .Rows
                    
                     If .TextMatrix(Lrow, 1) = "-1" Then
                            StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Data Verifikasi Timesheet Divisi, " & .TextMatrix(Lrow, 3) & "','ACC Divisi')"
                            PerintahExecute (StrSQL)
                            
                            StrSQL = "Update Tbltimesheet Set StatusDivisi = 1,last_update='" & Now & "',last_user='" & StrUser & "' Where NIP = '" & .TextMatrix(Lrow, 2) & "' And Tanggal Between  '" & Format(DTPicker1, "MM/dd/yyyy") & "' And  '" & Format(DTPicker2, "MM/dd/yyyy") & "'"
                            PerintahExecute (StrSQL)
                             
                             .TextMatrix(Lrow, 1) = ""
                     End If
                 
                 Lrow = Lrow + 1
            Loop
                    MsgBox "Data Berhasil Diverifikasi", vbInformation
                   
                End If
                .Rows = .Rows - 1
            End With
           Showdata (fgIsi.TextMatrix(fgIsi.Row, 2))
Exit Sub
Adaerror:
MsgBox err.Description
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
ShowTS
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Dim Lrow As Long
Dim StrSQL As String
On Error GoTo Adaerror
        With fgIsi
            .Rows = .Rows + 1
            If CekCurek(" Pembatalan Verifikasi Divisi", fgIsi) = False Then Exit Sub

            If MsgBox("Apakah Anda yakin ingin Membatalkan Verifikasi Divisi ?", vbQuestion + vbYesNo, "Konfirmasi Simpan Data") = vbNo Then
               Exit Sub
            Else
            Lrow = 1
             Do Until Lrow = .Rows
                    
                     If .TextMatrix(Lrow, 1) = "-1" Then
                            StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Data Unverifikasi Timesheet Divisi, " & .TextMatrix(Lrow, 3) & "','ACC Divisi')"
                            PerintahExecute (StrSQL)
                                     
                            StrSQL = "Update Tbltimesheet Set StatusDivisi = 0,last_update='" & Now & "',last_user='" & StrUser & "' Where NIP = '" & .TextMatrix(Lrow, 2) & "' And Tanggal Between  '" & Format(DTPicker1, "MM/dd/yyyy") & "' And  '" & Format(DTPicker2, "MM/dd/yyyy") & "'"
                            PerintahExecute (StrSQL)
                            .TextMatrix(Lrow, 1) = ""
                     End If
                 
                 Lrow = Lrow + 1
            Loop
                    MsgBox "Data Berhasil Di Unverifikasi", vbInformation
                   
                End If
                .Rows = .Rows - 1
            End With
           Showdata (fgIsi.TextMatrix(fgIsi.Row, 2))
Exit Sub
Adaerror:
MsgBox err.Description
End Sub

Private Sub fg_DblClick()
Dim J As String
With fg

If .Col >= 4 And .TextMatrix(.Row, .Col) <> "" And .TextMatrix(.Row, .Col) <> "Telat" Then
    J = FrmPM.Showdata(.TextMatrix(.Row, .Col), KodeDivisi)
    FrmPM.NIP = .TextMatrix(.Row, 4)
    FrmPM.show vbModal
End If
End With

End Sub

Private Sub fgIsi_Click()
If fgIsi.Rows > 1 And fgIsi.Col > 1 Then Showdata (fgIsi.TextMatrix(fgIsi.Row, 2))
If fgIsi.Col = 1 Then
    fgIsi.Editable = flexEDKbdMouse
Else
    fgIsi.Editable = flexEDNone
End If
End Sub

Private Sub fgIsi_GotFocus()
If fgIsi.Col = 1 Then
    fgIsi.Editable = flexEDKbdMouse
Else
    fgIsi.Editable = flexEDNone
End If
End Sub

Private Sub fgIsi_KeyPress(KeyAscii As Integer)
fgIsi_Click
End Sub

Private Sub fgIsi_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
fgIsi_Click
End Sub

Private Sub Form_Load()
  
    SetIsi
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
End Sub
Sub ShowTS()
Dim StatusIsi As Boolean
Dim RsTS As New ADODB.Recordset
Dim RsTS1 As New ADODB.Recordset
Dim x As Integer
VsIsi.Rows = 1
If Rscek.State = adStateOpen Then Rscek.Close
Set Rscek = Nothing

Rscek.Open "Select NIP,Nama From Karyawan Where Kd_divisi = '" & KodeDivisi & "'  And (Status = 1 Or status = 2) Order By NIP", CN, adOpenStatic
Set VsIsi.DataSource = Rscek
Frame1.Caption = "Tidak Mengisi Timesheet Dari " & Format(DTPicker1, "dd/MM/yyyy") & " - " & Format(DTPicker2, "dd/MM/yyyy")
Set Rscek = Nothing
VsIsi.ColWidth(2) = 3000
fgIsi.Rows = 1
fgIsi.Cols = 3
VSFlexGrid1.Rows = 1
fg.Rows = 1
If RsTS.State = adStateOpen Then RsTS.Close
Set RsTS = Nothing

StrSQL = "Select Distinct NIP As Do,NIP  From Tbltimesheet Where statusPM = 1 And status='Actual' And Kd_divisi = '" & KodeDivisi & "' And Tanggal Between  '" & Format(DTPicker1, "MM/dd/yyyy") & "' And  '" & Format(DTPicker2, "MM/dd/yyyy") & "' Order BY NIP"
RsTS.Open StrSQL, CN, adOpenStatic
Set fgIsi.DataSource = RsTS
fgIsi.ColDataType(1) = flexDTBoolean
fgIsi.ColWidth(1) = 500
fgIsi.Cols = fgIsi.Cols + 1
fgIsi.ColWidth(3) = 2500
fgIsi.TextMatrix(0, 3) = "Nama"
'Frame2.Caption = "Mengisi Timesheet Dari " & Format(DTPicker1, "dd/MM/yyyy") & " - " & Format(DTPicker2, "dd/MM/yyyy")

Set RsTS = Nothing
StrSQL = "Select Distinct NIP As Do,NIP  From Tbltimesheet Where statusPM = 0 And status='Actual' And Kd_divisi = '" & KodeDivisi & "' And Tanggal Between  '" & Format(DTPicker1, "MM/dd/yyyy") & "' And  '" & Format(DTPicker2, "MM/dd/yyyy") & "' Order BY NIP"
RsTS1.Open StrSQL, CN, adOpenStatic
Set VSgrid.DataSource = RsTS1
VSgrid.ColDataType(1) = flexDTBoolean
VSgrid.ColWidth(1) = 500
VSgrid.Cols = fgIsi.Cols
VSgrid.ColWidth(1) = 0
VSgrid.ColWidth(3) = 3000
VSgrid.TextMatrix(0, 3) = "Nama"

Set RsTS1 = Nothing
With VsIsi
Lrow = 0
Do Until Lrow = .Rows
    Set RsTS = Nothing
    Set fgIsi.DataSource = Nothing
    x = 0
    StatusIsi = False
    For x = 1 To fgIsi.Rows - 1
        fgIsi.TextMatrix(x, 1) = ""
        If Trim(.TextMatrix(Lrow, 1)) = Trim(fgIsi.TextMatrix(x, 2)) Then
            fgIsi.TextMatrix(x, 3) = .TextMatrix(Lrow, 2)
            StatusIsi = True
            Exit For
        End If
    Next
    
    x = 0
    For x = 1 To VSgrid.Rows - 1
        VSgrid.TextMatrix(x, 1) = ""
        If Trim(.TextMatrix(Lrow, 1)) = Trim(VSgrid.TextMatrix(x, 2)) Then
            VSgrid.TextMatrix(x, 3) = .TextMatrix(Lrow, 2)
            StatusIsi = True
            Exit For
        End If
    Next
        If StatusIsi = True Then
            .RemoveItem (Lrow)
            Lrow = Lrow - 1
        End If
         Lrow = Lrow + 1
Loop
    x = 0
    For x = 1 To .Rows - 1
        .TextMatrix(x, 0) = x
    Next
    Lrow = 0
    Do Until Lrow = VSgrid.Rows
     
         
        If Trim(VSgrid.TextMatrix(Lrow, 3)) = "" Then
             VSgrid.RemoveItem (Lrow)
            Lrow = Lrow - 1
        End If
         Lrow = Lrow + 1
   Loop
With fgIsi
    Lrow = 0
    Do Until Lrow = .Rows
        
        x = 0

        For x = 1 To VSgrid.Rows - 1
            If Trim(.TextMatrix(Lrow, 2)) = Trim(VSgrid.TextMatrix(x, 2)) Then
                VSgrid.TextMatrix(x, 3) = .TextMatrix(Lrow, 3)
                .RemoveItem (Lrow)
                Lrow = Lrow - 1
                Exit For
            End If
        Next
        
             Lrow = Lrow + 1
    Loop
End With
    For x = 1 To fgIsi.Rows - 1
        fgIsi.TextMatrix(x, 0) = x
    Next
    For x = 1 To VSgrid.Rows - 1
        VSgrid.TextMatrix(x, 0) = x
    Next
End With

End Sub

Sub Setgrid()
fgIsi.Rows = 1
fgIsi.Cols = 3
fgIsi.ColWidth(0) = 500
fgIsi.ColWidth(1) = 700
fgIsi.ColWidth(2) = 3000
fgIsi.TextMatrix(0, 0) = "No"
fgIsi.TextMatrix(0, 1) = "NIP"
fgIsi.TextMatrix(0, 2) = "Nama"
With VSgrid
.Rows = 1
.Cols = 3
.ColWidth(0) = 500
.ColWidth(1) = 700
.ColWidth(2) = 3000
.TextMatrix(0, 0) = "No"
.TextMatrix(0, 1) = "NIP"
.TextMatrix(0, 2) = "Nama"
End With
Dim Cboid, cboid1 As String
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
    .ColWidth(4) = 0
     .ColWidth(5) = 800
     .ColWidth(14) = 0
    .ColWidth(15) = 0
    For i = 6 To 21
       .ColWidth(i) = 1200
       .ColComboList(i) = cboid1
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
Sub SetIsi()
With VsIsi
    .Rows = 1
    .Cols = 3
    .TextMatrix(0, 0) = "No"
'    .TextMatrix(0, 1) = "Tidak Mengisi Timesheet"
'    .TextMatrix(0, 2) = "Tidak Mengisi Timesheet"
    .TextMatrix(0, 0) = "No"
    .TextMatrix(0, 1) = "NIP"
    .TextMatrix(0, 2) = "Nama"
    .FixedAlignment(0) = flexAlignCenterCenter
    .ColWidth(0) = 0
    .ColWidth(1) = 1000
    .ColWidth(2) = 3000
End With
End Sub
Private Sub Form_Resize()
    On Error Resume Next
        CmdClose.Width = Me.Width - 100

        Frame2.Height = Me.Height / 3.5
'        fgIsi.Height = Frame2.Height - 500
        Frame3.Top = Frame2.Height + 800
        Frame1.Top = Frame2.Height + Frame2.Height + 800
'        VsIsi.Height = Frame1.Height - 500
       
        fg.Width = Me.ScaleWidth / 1.6 + 500
        fg.Left = Frame1.Width + 200
        fg.Height = Me.Height - Picture2.Height - 1000 - Picture3.Height
        Picture3.Top = fg.Height + 800
        Picture3.Width = Me.ScaleWidth / 1.6 + 500
        Picture3.Left = Frame1.Width + 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmAccDivisi2 = Nothing
End Sub

Function Showdata(SNIP As String)
 
Dim x As Integer
Dim i, J As Integer
Dim Split, JmlLoop As Integer
Dim Jam As String
Dim IDWaktu, Hari As String
Dim Jam1, Jam2 As Date
Dim TglAwal, TglAkhir As Date
Dim TotalHari, Selisih As Double
Dim RsTS As New ADODB.Recordset
    fg.Rows = 1
    TglAwal = DTPicker1
    TglAkhir = DTPicker1
    TglAwal = DateSerial(Year(DTPicker1), Month(DTPicker1), 26)
      TglAkhir = DateSerial(Year(DTPicker1), Month(DTPicker1), 25)
    If Format(DTPicker1, "dd") < 26 Then
        TglAwal = DateAdd("M", -1, TglAwal)
        TglAkhir = DateAdd("M", 0, TglAkhir)
    Else
        TglAwal = DateAdd("M", 0, TglAwal)
        TglAkhir = DateAdd("M", 1, TglAkhir)
    End If
     VSFg.Rows = 1
    VSFg.Cols = 12
With VSFlexGrid1
        .Rows = 1
    If RsTS.State = adStateOpen Then RsTS.Close
    StrSQL = "SELECT IDtimesheet,Tanggal,JamAwal As [Jam Awal],JamAkhir AS [Jam Akhir],Status,NoProject As Project,Keterangan,Tanggal,Masuk,StatusPM,StatusDivisi,Keluar,Phase From tbltimesheet Where Tanggal Between '" & Format(TglAwal, "yyyy/MM/dd") & "' And '" & Format(TglAkhir, "yyyy/MM/dd") & "' And NIP = '" & SNIP & "' And Status='Actual' Order By Tanggal DESC,IDTimesheet ASC"
     RsTS.Open StrSQL, CN, adOpenStatic
    Set .DataSource = RsTS
    Selisih = DateDiff("d", TglAwal, TglAkhir)
    TotalHari = 0

    For x = 0 To Selisih
       If x > 0 Then TglAwal = DateAdd("d", 1, TglAwal)
       Hari = Format(TglAwal, "ddd")
        fg.Rows = fg.Rows + 1
        fg.TextMatrix(fg.Rows - 1, 2) = TglAwal
        fg.TextMatrix(fg.Rows - 1, 0) = fg.Rows - 1
       If Hari <> "Sat" And Hari <> "Sun" And Hari <> "Sabtu" And Hari <> "Minggu" Then
'           StatusHari = "Kerja"
            StrSQL = "select tanggallibur from kalender " & _
                "where tanggallibur = '" & Format(TglAwal, "MM/dd/yyyy") & "'"
            If Rscek.State = adStateOpen Then Rscek.Close
            Rscek.Open StrSQL, CN, adOpenStatic
            If Rscek.EOF Then
               TotalHari = TotalHari + 1
            Else
                For i = 1 To 5
                    fg.Col = i
                    fg.Row = fg.Rows - 1
                    fg.CellBackColor = &HC0C0FF
                Next
            End If
        Else
                For i = 1 To 5
                    fg.Col = i
                    fg.Row = fg.Rows - 1
                    fg.CellBackColor = &HC0C0FF
                Next
        End If
            If Rscek.State = adStateOpen Then Rscek.Close
            Rscek.Open "Select * from absensi where nip = '" & SNIP & "' And tgl= '" & Format(TglAwal, "MM/dd/yyyy") & "'", CN, adOpenStatic
            If Not Rscek.EOF Then
                fg.TextMatrix(fg.Rows - 1, 4) = Rscek!Kd_absensi
                If Rscek!Kd_absensi = 4 Then
                    fg.TextMatrix(fg.Rows - 1, 3) = "08:00"
                    fg.TextMatrix(fg.Rows - 1, 5) = "17:00"
                Else
                    fg.TextMatrix(fg.Rows - 1, 3) = Format(Rscek!masuk, "HH:mm")
                    fg.TextMatrix(fg.Rows - 1, 5) = Format(Rscek!keluar, "HH:mm")
                End If
            End If
    Next
'     Exit Sub
    TotalJamPUmum = TotalHari * 8

    
    .ColDataType(2) = flexDTDate
    .ColFormat(2) = "dd/MM/yyyy"
    For Lrow = 1 To .Rows - 1
        If .TextMatrix(Lrow, 3) = "" Then .TextMatrix(Lrow, 3) = "00:00"
         If .TextMatrix(Lrow, 4) = "" Then .TextMatrix(Lrow, 4) = "00:00"
       
                JmlLoop = 0
                Jam1 = CDate(.TextMatrix(Lrow, 3))
                Jam2 = CDate(.TextMatrix(Lrow, 4))
                DTPicker3.Value = Jam1
                Do Until JmlLoop = 50
                    If Format(DTPicker3.Value, "hh:mm") = Format(Jam2, "hh:mm") Then Exit Do
                    VSFg.Rows = VSFg.Rows + 1
                     
                    VSFg.TextMatrix(VSFg.Rows - 1, 0) = VSFg.Rows - 1
                    IDWaktu = Format(DTPicker3.Value, "HH:mm")
'                    If Trim(.TextMatrix(Lrow, 7)) = "Timesheet" And IDWaktu = "17:00" Then IDWaktu = "16:30"
                    VSFg.TextMatrix(VSFg.Rows - 1, 1) = IDWaktu
                    VSFg.TextMatrix(VSFg.Rows - 1, 2) = Format(.TextMatrix(Lrow, 2), "dd/MMM/yyyy")
                    VSFg.TextMatrix(VSFg.Rows - 1, 3) = .TextMatrix(Lrow, 6)
                    VSFg.TextMatrix(VSFg.Rows - 1, 4) = SNIP
                    VSFg.TextMatrix(VSFg.Rows - 1, 5) = .TextMatrix(Lrow, 5)
                    VSFg.TextMatrix(VSFg.Rows - 1, 6) = Format(.TextMatrix(Lrow, 9), "HH:mm")
                     VSFg.TextMatrix(VSFg.Rows - 1, 7) = .TextMatrix(Lrow, 7)
                     VSFg.TextMatrix(VSFg.Rows - 1, 8) = .TextMatrix(Lrow, 10)
                     VSFg.TextMatrix(VSFg.Rows - 1, 9) = .TextMatrix(Lrow, 11)
                     VSFg.TextMatrix(VSFg.Rows - 1, 10) = Format(.TextMatrix(Lrow, 12), "HH:mm")
                     VSFg.TextMatrix(VSFg.Rows - 1, 11) = Format(.TextMatrix(Lrow, 13), "HH:mm")
                    JmlLoop = JmlLoop + 1
                    DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                    If Format(DTPicker3, "HH:mm") = "12:30" Then DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                Loop
    Next
End With
    
'Disiplit Perhari
With fg
' .Rows = 1
For i = 1 To VSFg.Rows - 1
    Jam = VSFg.TextMatrix(i, 1)
    J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
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
                            .Col = 1
                            .Row = i
                        If .CellBackColor = &HC0C0FF And (.TextMatrix(i, J) = "?" Or .TextMatrix(i, J) = "Telat") Then
                                .TextMatrix(i, J) = "LIBUR"
                                .Col = J
                                .Row = i
                                .CellBackColor = &HC0C0FF
                        ElseIf .CellBackColor = &HC0C0FF And .TextMatrix(i, J) <> "?" Then
                                .Col = J
                                .Row = i
                                .CellBackColor = &HC0C0FF
                        End If
                        
                        If Trim(.TextMatrix(i, 3)) = "" And .TextMatrix(i, 6) <> "LIBUR" Then
                            .TextMatrix(i, J) = ""
                        End If
                    Select Case .TextMatrix(i, 4)
                        Case 2
                                .TextMatrix(i, J) = "Ijin/Sakit/dll"
                        Case 3
                                .TextMatrix(i, J) = "Cuti"
                    End Select
            Next
       
    Next
    .ColDataType(2) = flexDTDate
    .ColFormat(2) = "dd/MM/yyyy"
    .ColWidth(3) = 800
    .ColWidth(14) = 0
    .Row = 0
End With
End Function

Function Tampil(ByVal Jam As String, ByVal i As Integer, status As String, StatusPM As String, StatusDivisi As String)
Dim Project As String, AdaKolom As Boolean
Dim JmlCol, PosisiKol As Integer
Dim PosisiJam As String
Dim PosRow As Integer
Dim Adabaris As Boolean
With fg
    Adabaris = False
    For PosRow = 1 To .Rows - 1
        If Format(.TextMatrix(PosRow, 2), "dd/MMM/yyyy") = Format(VSFg.TextMatrix(i, 2), "dd/MMM/yyyy") Then: Adabaris = True: Exit For
    Next
    If Adabaris = True Then
'        .TextMatrix(PosRow, 2) = Format(VSFg.TextMatrix(i, 2), "dd/MM/yyyy")
        .TextMatrix(PosRow, 3) = VSFg.TextMatrix(i, 6)
        .TextMatrix(PosRow, 4) = VSFg.TextMatrix(i, 4)
        Project = VSFg.TextMatrix(i, 3)
        If KodeDivisi = 46 Then Project = VSFg.TextMatrix(i, 3) & "/" & VSFg.TextMatrix(i, 11)
'        If status <> "" Then .TextMatrix(PosRow, 5) = VSfg.TextMatrix(i, 10)
    '    .TextMatrix(.Rows - 1, 14) = "ISTIRAHAT"
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
             .TextMatrix(PosRow, PosisiKol) = Project
             If Trim(status) = "Actual" Then
                If StatusPM = 1 And StatusDivisi = 0 Then
                      .Row = PosRow
                      .Col = PosisiKol
                      .CellBackColor = &HC0FFC0
                ElseIf StatusPM = 1 And StatusDivisi = 1 Then
                      .Row = PosRow
                      .Col = PosisiKol
                      .CellBackColor = &HFF&
                End If
            End If
        End If
    End If
  End With
End Function



Private Sub VSgrid_Click()
If VSgrid.Rows > 1 Then Showdata (VSgrid.TextMatrix(VSgrid.Row, 2))

'Showdata (VSgrid.TextMatrix(VSgrid.Row, 2))
End Sub

Private Sub VsIsi_DblClick()
With VsIsi
    If .Rows = 1 Then Exit Sub
    If .TextMatrix(.Row, .Col) <> "" Then
        FrmManualTM.TxtNIP.Tag = .Row
        FrmManualTM.TxtNIP = .TextMatrix(.Row, 1)
        FrmManualTM.TxtNama = .TextMatrix(.Row, 2)
        FrmManualTM.show vbModal
    End If
End With
End Sub
