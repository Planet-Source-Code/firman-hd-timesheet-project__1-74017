VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CYBER_~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C9680CB9-8919-4ED0-A47D-8DC07382CB7B}#1.0#0"; "StyleButtonX.ocx"
Begin VB.MDIForm MDIMENU 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Timesheet"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   -1335
   ClientWidth     =   11610
   Icon            =   "MDIMENU.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00800000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   11550
      TabIndex        =   10
      Top             =   0
      Width           =   11610
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   11550
      TabIndex        =   6
      Top             =   7710
      Width           =   11610
      Begin VB.Label Label4 
         Caption         =   "Label3"
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
         Left            =   10200
         TabIndex        =   14
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
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
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
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
         Left            =   1800
         TabIndex        =   8
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
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
         Left            =   5520
         TabIndex        =   7
         Top             =   120
         Width           =   3375
      End
   End
   Begin VB.PictureBox picSeparator 
      Align           =   4  'Align Right
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   7380
      MousePointer    =   9  'Size W E
      ScaleHeight     =   7455
      ScaleWidth      =   120
      TabIndex        =   4
      Top             =   255
      Width           =   120
      Begin StyleButtonX.StyleButton StyleButton2 
         Height          =   1095
         Left            =   0
         TabIndex        =   5
         Top             =   3720
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   1931
         UpColorTop1     =   -2147483633
         UpColorTop2     =   -2147483633
         UpColorTop3     =   -2147483633
         UpColorTop4     =   -2147483633
         UpColorButtom1  =   -2147483633
         UpColorButtom2  =   -2147483633
         UpColorButtom3  =   -2147483633
         UpColorButtom4  =   -2147483633
         UpColorLeft1    =   -2147483633
         UpColorLeft2    =   -2147483633
         UpColorLeft3    =   -2147483633
         UpColorLeft4    =   -2147483633
         UpColorRight1   =   -2147483633
         UpColorRight2   =   -2147483633
         UpColorRight3   =   -2147483633
         UpColorRight4   =   -2147483633
         DownColorTop1   =   7021576
         DownColorTop2   =   -2147483633
         DownColorTop3   =   -2147483633
         DownColorTop4   =   -2147483633
         DownColorButtom1=   7021576
         DownColorButtom2=   -2147483633
         DownColorButtom3=   -2147483633
         DownColorButtom4=   -2147483633
         DownColorLeft1  =   7021576
         DownColorLeft2  =   -2147483633
         DownColorLeft3  =   -2147483633
         DownColorLeft4  =   -2147483633
         DownColorRight1 =   7021576
         DownColorRight2 =   -2147483633
         DownColorRight3 =   -2147483633
         DownColorRight4 =   -2147483633
         HoverColorTop1  =   7021576
         HoverColorTop2  =   -2147483633
         HoverColorTop3  =   -2147483633
         HoverColorTop4  =   -2147483633
         HoverColorButtom1=   7021576
         HoverColorButtom2=   -2147483633
         HoverColorButtom3=   -2147483633
         HoverColorButtom4=   -2147483633
         HoverColorLeft1 =   7021576
         HoverColorLeft2 =   -2147483633
         HoverColorLeft3 =   -2147483633
         HoverColorLeft4 =   -2147483633
         HoverColorRight1=   7021576
         HoverColorRight2=   -2147483633
         HoverColorRight3=   -2147483633
         HoverColorRight4=   -2147483633
         FocusColorTop1  =   7021576
         FocusColorTop2  =   -2147483633
         FocusColorTop3  =   -2147483633
         FocusColorTop4  =   -2147483633
         FocusColorButtom1=   7021576
         FocusColorButtom2=   -2147483633
         FocusColorButtom3=   -2147483633
         FocusColorButtom4=   -2147483633
         FocusColorLeft1 =   7021576
         FocusColorLeft2 =   -2147483633
         FocusColorLeft3 =   -2147483633
         FocusColorLeft4 =   -2147483633
         FocusColorRight1=   7021576
         FocusColorRight2=   -2147483633
         FocusColorRight3=   -2147483633
         FocusColorRight4=   -2147483633
         DisabledColorTop1=   -2147483633
         DisabledColorTop2=   -2147483633
         DisabledColorTop3=   -2147483633
         DisabledColorTop4=   -2147483633
         DisabledColorButtom1=   -2147483633
         DisabledColorButtom2=   -2147483633
         DisabledColorButtom3=   -2147483633
         DisabledColorButtom4=   -2147483633
         DisabledColorLeft1=   -2147483633
         DisabledColorLeft2=   -2147483633
         DisabledColorLeft3=   -2147483633
         DisabledColorLeft4=   -2147483633
         DisabledColorRight1=   -2147483633
         DisabledColorRight2=   -2147483633
         DisabledColorRight3=   -2147483633
         DisabledColorRight4=   -2147483633
         Caption         =   ""
         MousePointer    =   1
         BackColorUp     =   -2147483633
         BackColorDown   =   11899524
         BackColorHover  =   14073525
         BackColorFocus  =   14604246
         BackColorDisabled=   -2147483633
         DotsInCornerColor=   16777215
         MoveWhenClick   =   0   'False
         ForeColorUp     =   -2147483630
         ForeColorDown   =   -2147483634
         ForeColorHover  =   -2147483630
         ForeColorFocus  =   -2147483630
         ForeColorDisabled=   12632256
         BeginProperty FontUp {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFocus {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowBorderLevel2=   0   'False
         DistanceBetweenPictureAndCaption=   -50
      End
   End
   Begin VB.PictureBox picLeft 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   7500
      ScaleHeight     =   7455
      ScaleWidth      =   4110
      TabIndex        =   0
      Top             =   255
      Width           =   4110
      Begin VB.CommandButton CmdChat 
         Caption         =   "Send Message"
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
         Left            =   0
         TabIndex        =   11
         Top             =   4560
         Width           =   4095
      End
      Begin VB.Frame Frame1 
         Height          =   465
         Left            =   0
         TabIndex        =   1
         Top             =   -75
         Width           =   4050
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   150
            OleObjectBlob   =   "MDIMENU.frx":08CA
            TabIndex        =   2
            Top             =   120
            Width           =   2655
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid lvWin 
         Height          =   3975
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "Untuk Menghapus Klik Kanan Mouse, Baca Pesan Klik Pada Kolom Date"
         Top             =   480
         Width           =   4095
         _cx             =   7223
         _cy             =   7011
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"MDIMENU.frx":0939
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
         TabBehavior     =   0
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
         Height          =   3975
         Left            =   0
         TabIndex        =   12
         ToolTipText     =   "Baca Pesan Klik Pada Kolom Date"
         Top             =   5280
         Width           =   4095
         _cx             =   7223
         _cy             =   7011
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"MDIMENU.frx":09DA
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
         TabBehavior     =   0
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "MDIMENU.frx":0A79
         TabIndex        =   13
         Top             =   5040
         Width           =   2655
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   1920
         Picture         =   "MDIMENU.frx":0AE8
         Top             =   4950
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Image5 
         Height          =   960
         Left            =   1950
         Picture         =   "MDIMENU.frx":1832
         Top             =   6030
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   360
   End
   Begin VB.Timer TimerUser 
      Interval        =   59200
      Left            =   600
      Top             =   360
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   960
      OleObjectBlob   =   "MDIMENU.frx":257C
      Top             =   1320
   End
   Begin MSComctlLib.ImageList i16x16g 
      Left            =   2880
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":27B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":2D4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":32E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":367E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":3A18
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":3DB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ig24x24 
      Left            =   1560
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":414C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   2205
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":4379
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":4D8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":579D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":5B37
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":5ED1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":626B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":6605
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":7017
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":7A29
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":843B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":8E4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":985F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":A271
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":AC83
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":B21F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   3480
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":B7BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":D14D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":EADF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":10471
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":11E03
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":13795
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":15127
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":16AB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":1844B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":19DDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":1AABB
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":1B39B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":1C077
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":1CD53
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":1DA2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":1E70B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":1F3E7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":1FCC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":21655
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":22331
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":23CC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":25655
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":26FE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":28979
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":29653
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":2A32D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":2B007
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":2BCE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":2C9BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":2D29B
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":2DF77
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":2EC53
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":2F92F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":30213
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":30EEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":317CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":324A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":33E3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMENU.frx":357CF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuKepFile 
      Caption         =   "File"
      Begin VB.Menu MnuDataUser 
         Caption         =   "Data User"
      End
      Begin VB.Menu MnuGroupUser 
         Caption         =   "Group User"
      End
      Begin VB.Menu MnuChangePasswor 
         Caption         =   "Change Password"
      End
      Begin VB.Menu q 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu MnuLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MnuKepSetting 
      Caption         =   "Setup"
      Begin VB.Menu mnuKaryawan 
         Caption         =   "Karyawan"
      End
      Begin VB.Menu MnuGaji 
         Caption         =   "Gaji"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuSettingLembur 
         Caption         =   "Setting Lembur"
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu MnuProject 
         Caption         =   "Project"
      End
      Begin VB.Menu MnuPhase 
         Caption         =   "Phase"
      End
   End
   Begin VB.Menu MnuKeptimesheet 
      Caption         =   "Timesheet/Lembur"
      Begin VB.Menu MnuTimesheet 
         Caption         =   "Timesheet / Lembur"
      End
      Begin VB.Menu f 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlan 
         Caption         =   "Planning Timesheet"
      End
      Begin VB.Menu MnuListPlan 
         Caption         =   "List Planning Timesheet"
      End
      Begin VB.Menu MnuEditActualTimesheet 
         Caption         =   "Edit Actual Timesheet"
      End
   End
   Begin VB.Menu MnuKepVerifikasi 
      Caption         =   "Verifikasi"
      Begin VB.Menu MnuVerifikasitimesheetPM 
         Caption         =   "Verifikasi Timesheet / Lembur PM"
      End
      Begin VB.Menu w 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVerifikasiTimesheetDivisi 
         Caption         =   "Verifikasi Timesheet / Lembur Divisi"
      End
   End
   Begin VB.Menu MnuKepRekapKaryawan 
      Caption         =   "Rekap"
      Begin VB.Menu MnuRekapAbsen 
         Caption         =   "Rekap Absensi"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuRkpKaryTimesheet 
         Caption         =   "Rekap Timesheet / Lembur "
      End
      Begin VB.Menu MnuRkpTs 
         Caption         =   "Rekap Total Jam Timesheet/Lembur"
      End
      Begin VB.Menu mnurdivisi 
         Caption         =   "Kepala Divisi"
         Begin VB.Menu MnuTotaljam 
            Caption         =   "Rekap Total Timesheet / Project Verifikasi"
         End
         Begin VB.Menu MnuLapTotalTs 
            Caption         =   "Rekap Total Timesheet / Divisi Verifikasi"
         End
         Begin VB.Menu mnurekapdivisi 
            Caption         =   "Rekap  Proyek per Divisi "
         End
      End
      Begin VB.Menu e 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRekapKeuangan 
         Caption         =   "Keuangan/HRD"
         Begin VB.Menu MnuLapRekapLembur 
            Caption         =   "Rekap Biaya Lembur"
         End
         Begin VB.Menu MnuRekapBiayaPerproject 
            Caption         =   "Rekap Biaya Project / Verifikasi"
            Visible         =   0   'False
         End
         Begin VB.Menu MnuRekapBiayaLembur 
            Caption         =   "Rekap Biaya Detail Project  / Verifikasi"
            Visible         =   0   'False
         End
         Begin VB.Menu z 
            Caption         =   "-"
         End
         Begin VB.Menu MnuRkpTotal 
            Caption         =   "Rekap Status Timesheet Per Divisi"
         End
         Begin VB.Menu MnuRekapBiayaPerproject2 
            Caption         =   "Rekap Biaya Project / Non Verifikasi"
         End
         Begin VB.Menu MnuRekapTotalNon 
            Caption         =   "Rekap Total Jam Project (Beban Divisi)"
         End
         Begin VB.Menu MnuRekapRandom 
            Caption         =   "Rekap Total Jam Project PerOrang"
         End
      End
   End
   Begin VB.Menu MnuKepUtility 
      Caption         =   "Utility"
      Begin VB.Menu Mnusetkoneksi 
         Caption         =   "Set Koneksi"
      End
      Begin VB.Menu mnuUC 
         Caption         =   "&Calculator"
      End
      Begin VB.Menu mnuUN 
         Caption         =   "&Notepad"
      End
      Begin VB.Menu mnuUWE 
         Caption         =   "Windows Explorer"
      End
      Begin VB.Menu x 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLoguser 
         Caption         =   "Log Users"
      End
   End
   Begin VB.Menu mnuRecA 
      Caption         =   "A&ction"
      Visible         =   0   'False
      Begin VB.Menu mnuRNew 
         Caption         =   "&New"
      End
      Begin VB.Menu MnuRSelect 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuRSearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu mnuRDelete 
         Caption         =   "&Delete Selected"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuRPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu MnuRView 
         Caption         =   "&View"
      End
      Begin VB.Menu MnuRBatal 
         Caption         =   "&Cancel"
      End
      Begin VB.Menu mnuRAC 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu MnuActionMDI 
      Caption         =   "ActionMDI"
      Visible         =   0   'False
      Begin VB.Menu MDIMenu1 
         Caption         =   "Select &All"
      End
      Begin VB.Menu MDIMenu2 
         Caption         =   "&Delete Selected"
      End
      Begin VB.Menu MDIMenu3 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu MnuKepHelp 
      Caption         =   "Help"
      Begin VB.Menu MnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu r 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu Mnukrepair 
      Caption         =   "Repair Timesheet"
      Visible         =   0   'False
      Begin VB.Menu Mnurepair 
         Caption         =   "Repair Timesheet"
      End
      Begin VB.Menu MnuRetingkat 
         Caption         =   "Repair Tingkat Karyawan"
      End
      Begin VB.Menu Mnulogin 
         Caption         =   "Datauser"
      End
      Begin VB.Menu MnuTotalhari 
         Caption         =   "Repair Total Hari"
      End
   End
End
Attribute VB_Name = "MDIMENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim show_mnu        As Boolean
Dim cursor_pos As POINTAPI
Dim RsEmail As New ADODB.Recordset
Dim SQLEmail As String

Private Sub CmdChat_Click()
' FrmMessageReply.TxtFrom.Tag = StrNIPUser
' FrmMessageReply.TxtFrom = StrNamaUser
' FrmMessageReply.TxtDate = Now
' FrmMessageReply.show vbModal
With FrmMessage
    .TxtFrom = StrNamaUser
    .TxtFrom.Tag = StrUser
    .show vbModal
End With
End Sub

Private Sub lvWin_Click()
Dim J As String
 
With lvWin
If .Col = 1 Then
   .Editable = flexEDKbdMouse
Else
    .Editable = flexEDNone
End If
    If .Row > 0 And .Col > 1 Then
        J = FrmMessageReply.Showdata(.TextMatrix(.Row, 5), .TextMatrix(.Row, 2))
        StrSQL = "Update Tblemail Set statusBaca = 1 Where NIP = '" & .TextMatrix(.Row, 5) & "' And Tanggal = '" & .TextMatrix(.Row, 2) & "' And NIPto = '" & StrUser & "' "
        CN.Execute StrSQL
        FrmMessageReply.show vbModal
        GetEmail
    End If
    
End With
End Sub

Private Sub lvWin_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    MDIMENU.MDIMenu1.Enabled = True
    MDIMENU.MDIMenu2.Enabled = True
    MDIMENU.MDIMenu3.Enabled = True
    PopupMenu MDIMENU.MnuActionMDI
End If
End Sub

Private Sub MDIForm_Activate()
  CmdChat.Enabled = True
  VSFlexGrid1.Enabled = True
'  show_menu (True)
  
End Sub

Private Sub MDIForm_Load()
 
If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path & "\Skins\" & skinsFileName
      Skin1.ApplySkin hwnd
    End If
 
    If lvWin.Rows > 1 Then
        show_mnu = True
        show_menu (show_mnu)
    Else
        show_mnu = False
        show_menu (show_mnu)
    End If
    With lvWin
        For Lrow = 1 To .Rows - 1
            If .TextMatrix(Lrow, 4) = 1 Then
                show_mnu = False
                show_menu (show_mnu)
            Else
                show_mnu = True
                show_menu (show_mnu)
            End If
        Next
    End With
     If StrUser = "3578" Then MDIMENU.Mnukrepair.Visible = True
   
    GetEmail
  
End Sub
Private Sub Hapus()
Dim StrSQL As String
Dim Lrow As Integer
Dim Tanya As String
Dim ErrConn As Long

            With lvWin
            .Rows = .Rows + 1
             If CekCurek("hapus", lvWin) = False Then GetEmail: Exit Sub
            If MsgBox("Apakah Anda yakin ingin menghapus Data ?", vbQuestion + vbYesNo, "Konfirmasi hapus") = vbNo Then
               Exit Sub
            Else
             Do Until Lrow = .Rows - 1
             
                If .TextMatrix(Lrow, 1) = "-1" Then
                        StrSQL = "Delete From tblEmail where IDEmail = '" & .TextMatrix(Lrow, 6) & "'"
                        CN.Execute StrSQL
                        .RemoveItem (Lrow)
                        Lrow = Lrow - 1
                End If
                Lrow = Lrow + 1
            Loop
        
                End If
                GetEmail
            End With
Exit Sub
Adaerror:
If ErrConn > 0 Then CN.RollbackTrans
MsgBox err.Description
End Sub
    
Private Sub MDIMenu1_Click()
With lvWin
    For Lrow = 1 To .Rows - 1
        .TextMatrix(Lrow, 1) = "-1"
    Next
End With
End Sub

Private Sub MDIMenu2_Click()
Hapus
End Sub

Private Sub MDIMenu3_Click()
GetEmail
End Sub

Private Sub MnuHelp_Click()
Dim strhfile As String
 strhfile = App.Path & "\timesheet.chm"
 ShellExecute Me.hwnd, "open", strhfile, "", "", vbMaximizedFocus
End Sub

Private Sub mnuKaryawan_Click()
LoadForm FrmListKaryawan
End Sub

Private Sub MnuLapRekapLembur_Click()
LoadForm FrmLapRekapLemburTs
End Sub

Private Sub MnuLoguser_Click()
LoadForm FrmLog
End Sub

Private Sub MnuPhase_Click()
LoadForm FrmPhase
End Sub

Private Sub MnuRekapBiayaPerproject2_Click()
LoadForm FrmLapBiayaNonVerifikasi
End Sub

Private Sub mnurekapdivisi_Click()
LoadForm FrmTotalNonverifikasi
End Sub

Private Sub MnuRekapRandom_Click()
LoadForm FrmRekapTotalRandom
End Sub

Private Sub MnuRekapTotalNon_Click()
LoadForm FrmTotalNonverifikasi
End Sub

Private Sub MnuRkpTotal_Click()
LoadForm FrmLapTotalTms
End Sub

Private Sub MnuRkpTs_Click()
LoadForm FrmRekapTskary
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    Frame1.Width = picLeft.ScaleWidth
    lvWin.Width = picLeft.ScaleWidth
    lvWin.Height = picLeft.ScaleHeight - lvWin.Top - 5000
 
     
End Sub

Private Sub StyleButton2_Click()
    show_mnu = Not show_mnu
    show_menu (show_mnu)
    
End Sub
Sub show_menu(ByVal show As Boolean)
    Dim img As Image
    If show = True Then
        Set img = Image2
    Else
        Set img = Image5
    End If
    'Set the style button graphics
    With StyleButton2
        Set .PictureDown = img.Picture
        Set .PictureFocus = img.Picture
        Set .PictureHover = img.Picture
        Set .PictureUp = img.Picture
    End With

    'Set picture visibility
    picLeft.Visible = show

    If show = True Then StyleButton2.ToolTipText = "Hide": picSeparator.MousePointer = vbSizeWE Else picSeparator.MousePointer = vbArrow: StyleButton2.ToolTipText = "Expand"
     Set img = Nothing
End Sub
Private Sub TimerUser_Timer()
GetEmail
End Sub
Sub GetEmail()
Dim Xemail As Integer
Dim RsNama As New ADODB.Recordset
On Error GoTo Adaerror
With lvWin

If RsEmail.State = adStateOpen Then RsEmail.Close
SQLEmail = "Select * From tblEmail Where NIPto = " & StrUser & " Order By Tanggal DESC"
RsEmail.Open SQLEmail, CN, adOpenStatic
lvWin.Redraw = flexRDNone
.Rows = 1
.Cols = 7
.ColWidth(2) = 2000
.ColWidth(3) = 1600
.ColWidth(4) = 0
.ColWidth(5) = 0
.ColWidth(6) = 0
'If Not RsEmail.EOF Then
    
    Do Until RsEmail.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 2) = RsEmail!Tanggal
        .TextMatrix(.Rows - 1, 3) = RsEmail!Nama
        .TextMatrix(.Rows - 1, 4) = RsEmail!StatusBaca
        .TextMatrix(.Rows - 1, 5) = RsEmail!NIP
        .TextMatrix(.Rows - 1, 6) = RsEmail!IDEmail
        If RsEmail!StatusBaca = 0 Then
            For Xemail = 1 To .Cols - 1
                .Col = Xemail
                .Row = .Rows - 1
                .CellBackColor = vbYellow
            Next
            show_menu (True)
        End If
        RsEmail.MoveNext
        
    Loop
'End If
lvWin.Redraw = flexRDBuffered
End With

With VSFlexGrid1

If RsEmail.State = adStateOpen Then RsEmail.Close
SQLEmail = "Select * From tblEmail Where NIP = " & StrUser & " Order By Tanggal DESC"
RsEmail.Open SQLEmail, CN, adOpenStatic
 .Redraw = flexRDNone
.Rows = 1
.Cols = 7
.ColWidth(2) = 2000
.ColWidth(3) = 1600
.ColWidth(4) = 0
.ColWidth(5) = 0
.ColWidth(6) = 0
'If Not RsEmail.EOF Then
    
    Do Until RsEmail.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 2) = RsEmail!Tanggal
        .TextMatrix(.Rows - 1, 3) = RsEmail!NIPto
        If RsNama.State = adStateOpen Then RsNama.Close
        RsNama.Open "Select Nama From Karyawan where Nip ='" & RsEmail!NIPto & "'", CN, adOpenStatic
        If Not RsNama.EOF Then .TextMatrix(.Rows - 1, 3) = RsNama!Nama
         If RsNama.State = adStateOpen Then RsNama.Close
        .TextMatrix(.Rows - 1, 4) = RsEmail!StatusBaca
        .TextMatrix(.Rows - 1, 5) = RsEmail!Subjeck
        .TextMatrix(.Rows - 1, 6) = RsEmail!Email
        If RsEmail!StatusBaca = 0 Then
            For Xemail = 1 To .Cols - 1
                .Col = Xemail
                .Row = .Rows - 1
                .CellBackColor = vbYellow
            Next
           
        End If
        RsEmail.MoveNext
        
    Loop
'End If
 .Redraw = flexRDBuffered
End With
Exit Sub
Adaerror:
End Sub
Private Sub tmrResize_Timer()
    On Error Resume Next
    GetCursorPos cursor_pos
    picLeft.Width = (Me.Width - ((cursor_pos.x * Screen.TwipsPerPixelX) - Me.Left)) - 90
 
End Sub
Private Sub MDIForm_Terminate()
On Error Resume Next
      CN.Execute "Update tbldata_user Set Statuslogin = 0 WHere NIP = '" & StrNIPUser & "'"
If CN.State = adStateOpen Then CN.Close
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
 CN.Execute "Update tbldata_user Set Statuslogin = 0 WHere NIP = '" & StrNIPUser & "'"
 Set MDIMENU = Nothing
End Sub

Private Sub MnuAbout_Click()
frmAbout.show vbModal
End Sub

Private Sub MnuChangePasswor_Click()
 With FrmPassword
        .RowFlex = r
        .NamaUser = StrNIPUser
        .UserPassword = StrPassword
        .IDUser = StrIDUser
        .Add = False
        .IDGroup = strGroup
        .show vbModal
    End With
End Sub

Private Sub MnuDataUser_Click()
FrmDataUser.show
End Sub

Private Sub MnuEditActualTimesheet_Click()
LoadForm FrmTimesheetEdit
End Sub

Private Sub MnuExit_Click()
CN.Execute "Update tbldata_user Set Statuslogin = 0 WHere NIP = '" & StrNIPUser & "'"

Unload Me

End Sub

Private Sub MnuGaji_Click()
FrmGaji.show
End Sub

Private Sub MnuGroupUser_Click()
FrmGroupUser.show
End Sub

Private Sub MnuLapTotalTs_Click()
LoadForm FrmLapTotalTms
End Sub

Private Sub MnuListPlan_Click()
LoadForm FrmListPlan
End Sub

Private Sub Mnulogin_Click()
LoadForm FrmUser
End Sub

Private Sub MnuLogout_Click()
    CN.Execute "Update tbldata_user Set Statuslogin = 0 WHere NIP = '" & StrNIPUser & "'"
    Unload Me
    If CN.State = adStateOpen Then CN.Close
    FrmLogin.show

End Sub

Private Sub mnuPlan_Click()
StatusForm = True
LoadForm FrmTimesheetPlan
End Sub

Private Sub MnuProject_Click()
LoadForm FrmProjectList
End Sub

Private Sub MnuRekapAbsen_Click()
LoadForm FrmRekapAbsen
End Sub

Private Sub MnuRekapBiayaLembur_Click()
LoadForm FrmLapBiayaLembur
End Sub

Private Sub MnuRekapBiayaPerproject_Click()
LoadForm frmLapBiayaProject
End Sub

Private Sub Mnurepair_Click()
LoadForm FrmReTimesheet
End Sub

Private Sub MnuRetingkat_Click()
LoadForm FrmRepairKaryawan
End Sub

Private Sub MnuRkpKaryTimesheet_Click()
LoadForm frmRekapTimesheet
End Sub

Private Sub Mnusetkoneksi_Click()
FrmSetKoneksi.show vbModal
End Sub

Private Sub MnuSettingLembur_Click()
LoadForm FrmSetLembur
End Sub

Private Sub Mnutimesheet_Click()
If KodeDivisi = 46 Or StrUser = 3578 Then
   LoadForm FrmTimesheet
Else
    LoadForm FrmTimesheet2
End If
End Sub

Private Sub MnuTotalhari_Click()
FrmTotal.show
End Sub

Private Sub MnuTotaljam_Click()
If KodeDivisi = 46 Or StrUser = "3578" Then
    LoadForm FrmRekapTotalPhase
Else
    LoadForm FrmRekapTotal
End If
End Sub

Private Sub mnuUC_Click()
    On Error Resume Next
    Shell "calc.exe", vbNormalFocus
End Sub

Private Sub mnuUN_Click()
    On Error Resume Next
    Shell "notepad.exe", vbNormalFocus
End Sub

Private Sub mnuUWE_Click()
    On Error Resume Next
    Shell "Explorer.exe", vbNormalFocus
End Sub

Private Sub MnuRSelect_Click()
On Error Resume Next
    ActiveForm.Perintah "Select"
End Sub

Private Sub mnuRAC_Click()
    On Error Resume Next
    ActiveForm.Perintah "Close"
End Sub

Private Sub mnuRNew_Click()
    On Error Resume Next
    ActiveForm.Perintah "New"
End Sub

Private Sub mnuRDelete_Click()
    On Error Resume Next
    ActiveForm.Perintah "Delete"
End Sub

Private Sub mnuRAES_Click()
    On Error Resume Next
    ActiveForm.Perintah "Edit"
End Sub

Private Sub mnuRPrint_Click()
    On Error Resume Next
    ActiveForm.Perintah "Print"
End Sub

Private Sub MnuRefresh_Click()
    On Error Resume Next
    ActiveForm.Perintah "Refresh"
End Sub
Private Sub MnuRView_Click()
    On Error Resume Next
    ActiveForm.Perintah "View"
End Sub
Private Sub MnuRSearch_Click()
    On Error Resume Next
    ActiveForm.Perintah "Search"
End Sub

Private Sub MnuVerifikasiTimesheetDivisi_Click()
LoadForm FrmAccDivisi2
End Sub

Private Sub MnuVerifikasitimesheetPM_Click()
LoadForm FrmAccPM
'LoadForm FrmAccDivisi
End Sub

Private Sub VSFlexGrid1_Click()
With FrmMessage
    .TxtFrom = StrNamaUser
    .TxtFrom.Tag = StrUser
    .CboKaryawan = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, 3)
    .TxtSubjek = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, 5)
    .TxtMessage = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, 6)
    .Command1.Enabled = False
    .show vbModal
End With
End Sub
