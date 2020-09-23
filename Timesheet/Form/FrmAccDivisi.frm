VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmAccDivisi 
   Caption         =   "Verifikasi Divisi"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   9960
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4695
      Left            =   0
      TabIndex        =   16
      ToolTipText     =   "Double Klik Kolom Nama, Untuk Melihat Detail Timesheet"
      Top             =   1440
      Width           =   6495
      _cx             =   11456
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
      BackColorFixed  =   15648682
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
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   9930
      TabIndex        =   2
      Top             =   0
      Width           =   9960
      Begin VB.Frame Frame1 
         Caption         =   "Tidak Mengisi Timesheet"
         Height          =   1215
         Left            =   10080
         TabIndex        =   14
         Top             =   120
         Width           =   4695
         Begin VSFlex8Ctl.VSFlexGrid VsflexIsi 
            Height          =   855
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "Double Klik Untuk Membuat Manual Timesheet"
            Top             =   240
            Width           =   4335
            _cx             =   7646
            _cy             =   1508
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
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3600
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
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
         Left            =   6840
         TabIndex        =   4
         Top             =   360
         Width           =   1335
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
         Left            =   8280
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
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
         Format          =   16515075
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3600
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
         Format          =   16515075
         CurrentDate     =   39931
      End
      Begin VSFlex8Ctl.VSFlexGrid cboFlex 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
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
         FormatString    =   $"FrmAccDivisi.frx":0000
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
         TabIndex        =   11
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
         FormatString    =   $"FrmAccDivisi.frx":0029
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
         Left            =   6840
         TabIndex        =   20
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515074
         CurrentDate     =   39940
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
         Left            =   3120
         TabIndex        =   10
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
         Left            =   960
         TabIndex        =   9
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
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9960
      TabIndex        =   0
      Top             =   7050
      Width           =   9960
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
   Begin ACTIVESKINLibCtl.Skin Skin2 
      Left            =   8880
      OleObjectBlob   =   "FrmAccDivisi.frx":0052
      Top             =   2160
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "FrmAccDivisi.frx":0286
      Top             =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid VSfg 
      Height          =   3015
      Left            =   7080
      TabIndex        =   17
      Top             =   1440
      Width           =   7815
      _cx             =   13785
      _cy             =   5318
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
      Height          =   3975
      Left            =   0
      TabIndex        =   18
      Top             =   5280
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
      FormatString    =   $"FrmAccDivisi.frx":04BA
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
      Left            =   7440
      TabIndex        =   19
      Top             =   5040
      Visible         =   0   'False
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
      FormatString    =   $"FrmAccDivisi.frx":0599
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
Attribute VB_Name = "FrmAccDivisi"
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
        Case "Refresh"
           Showdata
                
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
'                             If .TextMatrix(Lrow, 9) = "1" Then
                                      StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Data Verifikasi Timesheet Divisi, " & .TextMatrix(Lrow, 3) & "','ACC Divisi')"
                                     PerintahExecute (StrSQL)
                                     
                                     StrSQL = "Update Tbltimesheet Set StatusDivisi = '" & .TextMatrix(Lrow, 9) & "',last_update='" & Now & "',last_user='" & StrUser & "' Where IdTimesheet = '" & .TextMatrix(Lrow, 13) & "'"
                                     PerintahExecute (StrSQL)
            
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
            End With
Exit Sub
Adaerror:
MsgBox err.Description
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
Showdata
Command1.Enabled = True
End Sub

Private Sub Form_Load()
    AddKaryawan
    AddProject
    SetIsi
    Setgrid
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker1.Value = DateSerial(Year(Now), Month(Now), 1)
    DTPicker1.Value = DateAdd("M", 0, DTPicker1.Value)
    DTPicker2.Value = DateSerial(Year(Now), Month(Now), 1)
    DTPicker2.Value = DateAdd("M", 1, DTPicker2.Value) - 1
    DTPicker1.CustomFormat = "dd/MMM/yyyy"
    DTPicker2.CustomFormat = "dd/MMM/yyyy"
     If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
      Skin1.ApplySkin hwnd
    End If
End Sub
Private Sub AddKaryawan()

    Dim Cboid     As String
    Dim cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    cboid1 = vbNullString
    StrSQL = "select * from Karyawan Where kd_divisi = '" & KodeDivisi & "' And Status <> '14' Order By Nama"
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
Sub SetIsi()
With VsflexIsi

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
Private Sub AddProject()

    Dim Cboid     As String
    Dim cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    cboid1 = vbNullString
   StrSQL = "select Kode,Nama from project " & _
        "where status = 'Terpakai' And kd_divisi = '" & KodeDivisi & "'" & _
        "group by kode,Nama " & _
        "order by kode"
    Rscek.Open StrSQL, CN, adOpenStatic
    cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("Kode") & vbTab & Rscek("Nama")
      cboid1 = cboid1 + Cboid
      Rscek.MoveNext
    Loop
    cboFlex.ColComboList(0) = cboid1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        CmdClose.Width = Me.Width - 100
        With fg
             .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture2.Height

        End With
 
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
                         StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Hapus Timesheet, " & StrNIPUser & " &  " & .TextMatrix(Lrow, 3) & "','ACC Divisi')"
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

Sub Showdata()
Dim x As Integer
Dim i, J As Integer
Dim Split, JmlLoop As Integer
Dim Jam As String
Dim IDWaktu As String
Dim Jam1, Jam2 As Date
Dim TglAwal, TglAkhir As Date
Dim RsTS As New ADODB.Recordset
   
If Trim(cboFlex.Text) = "" Then
   MsgBox "Silahkan Pilih Project Terlebih Dahulu", vbCritical
   cboFlex.SetFocus
   Exit Sub
End If
SetIsi
For i = 0 To Combo1.ListCount - 1
    If Rscek.State = adStateOpen Then Rscek.Close
    Set Rscek = Nothing
    Combo1.ListIndex = i
    J = Combo1.Text
    Rscek.Open "SELECT * From Tbltimesheet where tbltimesheet.NIP ='" & J & "' And TANGGAL BETWEEN '" & Format(DTPicker1, "mm/dd/yyyy") & "' And '" & Format(DTPicker2, "mm/dd/yyyy") & "' And Status ='Actual'", CN, adOpenStatic
    If Rscek.EOF Then
       VsflexIsi.Rows = VsflexIsi.Rows + 1
       VsflexIsi.TextMatrix(VsflexIsi.Rows - 1, 1) = J
        If Rscek.State = adStateOpen Then Rscek.Close
        Rscek.Open "Select * From Karyawan Where NIP = '" & J & "'", CN, adOpenStatic
        If Not Rscek.EOF Then
            VsflexIsi.TextMatrix(VsflexIsi.Rows - 1, 2) = Rscek!Nama
            Frame1.Caption = "Tidak Mengisi Timesheet Dari " & Format(DTPicker1, "dd/MM/yyyy") & " - " & Format(DTPicker2, "dd/MM/yyyy")
        End If
    End If
Next
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
        VSfg.Rows = 1
        fg.Rows = 1
        VSfg.Cols = 11
        VSFlexGrid2.Cols = 13
    If RsTS.State = adStateOpen Then RsTS.Close
    StrSQL = "SELECT tbltimesheet.IDtimesheet,tbltimesheet.Tanggal,tbltimesheet.JamAwal As [Jam Awal],tbltimesheet.JamAkhir AS [Jam Akhir],tbltimesheet.Status,tbltimesheet.NoProject As Project,tbltimesheet.Keterangan,tbltimesheet.Tanggal,tbltimesheet.Masuk,tbltimesheet.NIP,tbltimesheet.StatusDivisi, karyawan.Nama,tbltimesheet.StatusPM FROM tbltimesheet INNER JOIN  karyawan ON tbltimesheet.NIP = karyawan.NIP "
    StrSQL = StrSQL & " Where tbltimesheet.Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And '" & Format(DTPicker2, "MM/dd/yyyy") & "' And tbltimesheet.Kd_Divisi = '" & KodeDivisi & "' And tbltimesheet.Status ='Actual' And StatusPM = 1"
    
     StrSQL = StrSQL & " AND tbltimesheet.NoProject = '" & cboFlex & "'"
    If CboKaryawan <> "" Then StrSQL = StrSQL & " AND tbltimesheet.Nip = '" & CboKaryawan & "'"

    StrSQL = StrSQL & " Order By tbltimesheet.Tanggal,tbltimesheet.NIP,tbltimesheet.Keterangan ASC"
    
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
                Do Until JmlLoop = 30
                    If Format(DTPicker3.Value, "hh:mm") = Format(Jam2, "hh:mm") Then
                       Exit Do
                    Else
                        VSfg.Rows = VSfg.Rows + 1
                         
                        VSfg.TextMatrix(VSfg.Rows - 1, 0) = VSfg.Rows - 1
                        IDWaktu = Format(DTPicker3.Value, "HH:mm")
                        If Trim(.TextMatrix(Lrow, 7)) = "Timesheet" And IDWaktu = "17:00" Then IDWaktu = "16:30"
                        VSfg.TextMatrix(VSfg.Rows - 1, 1) = IDWaktu
                        VSfg.TextMatrix(VSfg.Rows - 1, 2) = .TextMatrix(Lrow, 2) 'Format(.TextMatrix(Lrow, 2), "dd/MM/yyyy")
                        VSfg.TextMatrix(VSfg.Rows - 1, 3) = .TextMatrix(Lrow, 6)
                        VSfg.TextMatrix(VSfg.Rows - 1, 4) = .TextMatrix(Lrow, 10)
                        VSfg.TextMatrix(VSfg.Rows - 1, 5) = .TextMatrix(Lrow, 5)
                        VSfg.TextMatrix(VSfg.Rows - 1, 6) = Format(.TextMatrix(Lrow, 9), "HH:mm")
                        VSfg.TextMatrix(VSfg.Rows - 1, 7) = .TextMatrix(Lrow, 7)
                        VSfg.TextMatrix(VSfg.Rows - 1, 8) = .TextMatrix(Lrow, 11)
                        VSfg.TextMatrix(VSfg.Rows - 1, 9) = .TextMatrix(Lrow, 12)
                        VSfg.TextMatrix(VSfg.Rows - 1, 10) = .TextMatrix(Lrow, 1)
'                        VSfg.TextMatrix(VSfg.Rows - 1, 11) = .TextMatrix(Lrow, 13)
                       
                        DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                        If Format(DTPicker3, "HH:mm") = "12:30" Then DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                        JmlLoop = JmlLoop + 1
                    End If
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
.TextMatrix(.Rows - 1, 9) = "StatusDivisi"
.TextMatrix(.Rows - 1, 10) = "TotalLembur"
.TextMatrix(.Rows - 1, 11) = "StatusPM"
.TextMatrix(.Rows - 1, 12) = "Keterangan"
End With
With VSfg
For Lrow = 1 To VSfg.Rows - 1
    
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
   If .Rows = 1 Then MsgBox "Data Tidak Ditemukan ", vbInformation
End With
fg.Col = 1
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
        If StatusNIP = True Then
                    .TextMatrix(.Rows - 1, 1) = VSfg.TextMatrix(Row, 2)
                    .TextMatrix(.Rows - 1, 2) = VSfg.TextMatrix(Row, 3)
                    .TextMatrix(.Rows - 1, 3) = VSfg.TextMatrix(Row, 4)
                    .TextMatrix(.Rows - 1, 4) = VSfg.TextMatrix(Row, 9)
                    If VSfg.TextMatrix(Row, 1) <> "12:00" And Trim(Keterangan) = "Timesheet" Then
                    .TextMatrix(.Rows - 1, 5) = (.TextMatrix(.Rows - 1, 5) + 0.5)
                    End If
                    .TextMatrix(.Rows - 1, 6) = "masuk"
                    .TextMatrix(.Rows - 1, 7) = "keluar"
                    .TextMatrix(.Rows - 1, 8) = VSfg.TextMatrix(Row, 3)
                    .TextMatrix(.Rows - 1, 9) = VSfg.TextMatrix(Row, 8)
                    If Trim(VSfg.TextMatrix(Row, 7)) = "Lembur" Then
                        .TextMatrix(.Rows - 1, 10) = (.TextMatrix(.Rows - 1, 10) + 0.5)
                    End If
                    .TextMatrix(.Rows - 1, 11) = VSfg.TextMatrix(Row, 8)
                    .TextMatrix(.Rows - 1, 12) = VSfg.TextMatrix(Row, 7)
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 5) = 0
                    .TextMatrix(.Rows - 1, 10) = 0
                    .TextMatrix(.Rows - 1, 1) = VSfg.TextMatrix(Row, 2)
                    .TextMatrix(.Rows - 1, 2) = VSfg.TextMatrix(Row, 3)
                    .TextMatrix(.Rows - 1, 3) = VSfg.TextMatrix(Row, 4)
                 .TextMatrix(.Rows - 1, 4) = VSfg.TextMatrix(Row, 9)
                    If VSfg.TextMatrix(Row, 1) <> "12:00" And Trim(Keterangan) = "Timesheet" Then
                       .TextMatrix(.Rows - 1, 5) = (.TextMatrix(.Rows - 1, 5) + 0.5)
                    End If
                    .TextMatrix(.Rows - 1, 6) = "masuk"
                    .TextMatrix(.Rows - 1, 7) = "keluar"
                    .TextMatrix(.Rows - 1, 8) = VSfg.TextMatrix(Row, 3)
                    .TextMatrix(.Rows - 1, 9) = VSfg.TextMatrix(Row, 8)
                    If Trim(Keterangan) = "Lembur" Then
                        .TextMatrix(.Rows - 1, 10) = (.TextMatrix(.Rows - 1, 10) + 0.5)
                    End If
                    .TextMatrix(.Rows - 1, 11) = VSfg.TextMatrix(Row, 8)
                    .TextMatrix(.Rows - 1, 12) = VSfg.TextMatrix(Row, 7)
                    .TextMatrix(.Rows - 1, 13) = VSfg.TextMatrix(Row, 10)
                End If
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
'       .TextMatrix(0, Kolom) = Format(.TextMatrix(0, Kolom), "dd/MM/yyyy")

        
       .Col = Kolom
       .Row = VsFlexRow
       .CellAlignment = flexAlignCenterCenter

End With
End Function


Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim J As String
With fg
If .TextMatrix(Row, Col + 1) = "" Then .TextMatrix(Row, Col) = 0
.Col = Col + 1
.Row = Row
If .CellBackColor = vbGreen Then
   If MsgBox("Apakah Anda Akan Membatalkan Verifikasi ?", vbQuestion + vbYesNo, "Konfirmasi Verifikasi") = vbNo Then
      .TextMatrix(Row, Col) = "-1"
'      .Col = Col + 1
'      .Row = Row
'
       Exit Sub
    Else
        .CellBackColor = vbWhite
        J = Cekfg(.TextMatrix(0, Col + 1), .TextMatrix(Row, 1), .TextMatrix(Row, Col))
    End If
Else
    J = Cekfg(.TextMatrix(0, Col + 1), .TextMatrix(Row, 1), .TextMatrix(Row, Col))
End If
End With
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

Private Sub VsflexIsi_DblClick()
With VsflexIsi
    If .TextMatrix(.Row, .Col) <> "" Then
        FrmManualTM.TxtNIP.Tag = .Row
        FrmManualTM.TxtNIP = .TextMatrix(.Row, 1)
        FrmManualTM.TxtNama = .TextMatrix(.Row, 2)
        FrmManualTM.show vbModal
    End If
End With
End Sub
