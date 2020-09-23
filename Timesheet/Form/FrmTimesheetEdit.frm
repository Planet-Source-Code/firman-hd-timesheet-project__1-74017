VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTimesheetEdit 
   Caption         =   "Edit Actual Timesheet"
   ClientHeight    =   6330
   ClientLeft      =   570
   ClientTop       =   2040
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   8970
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   8970
      TabIndex        =   9
      Top             =   5955
      Width           =   8970
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
   Begin ACTIVESKINLibCtl.Skin Skin2 
      Left            =   8400
      OleObjectBlob   =   "FrmTimesheetEdit.frx":0000
      Top             =   2760
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FrmTimesheetEdit.frx":0234
      Top             =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid VsFlex 
      Height          =   3975
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   6735
      _cx             =   11880
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
      SelectionMode   =   0
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
      FormatString    =   $"FrmTimesheetEdit.frx":0468
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
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3975
      Left            =   7680
      TabIndex        =   11
      ToolTipText     =   "Double Klik Kolom Tanggal Untuk Isi/Edit Timesheet, Double Klik Kolom Project Untuk Melihat PM Dan Mengirim Pesan"
      Top             =   1080
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
      FormatString    =   $"FrmTimesheetEdit.frx":0547
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
      Height          =   735
      Left            =   6840
      TabIndex        =   15
      Top             =   1680
      Width           =   3375
      _cx             =   5953
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
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   8940
      TabIndex        =   1
      Top             =   0
      Width           =   8970
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   4920
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VSFlex8Ctl.VSFlexGrid cboFlex 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Top             =   600
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
         ForeColorSel    =   4210752
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
         FormatString    =   $"FrmTimesheetEdit.frx":0626
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
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1080
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
         Format          =   52297731
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   6240
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52297730
         CurrentDate     =   39940
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   17
         Top             =   120
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   4200
         X2              =   4080
         Y1              =   120
         Y2              =   360
      End
      Begin VB.Label LKeluar 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "NIP  :"
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
         Left            =   480
         TabIndex        =   13
         Top             =   600
         Width           =   735
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
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
      Begin VB.Label LabelAbsen 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Masuk :"
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
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "List Timesheet Dan Lembur"
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
         Left            =   8640
         TabIndex        =   4
         Top             =   600
         Width           =   3495
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   3975
      Left            =   7560
      TabIndex        =   0
      Top             =   6000
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
      FormatString    =   $"FrmTimesheetEdit.frx":064F
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
Attribute VB_Name = "FrmTimesheetEdit"
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
           With VsFlex
                  .Row = .Rows - 1
                  .Col = 3
                  .EditCell
           End With
         Case "Search"
           With frmSearchs
                Set .srcForm = Me
                Set .srcColumnHeaders = VsFlex
                .srcNoOfCol = 5
                .show vbModal
            End With
         Case "Select"
            With VsFlex
                For Lrow = 1 To .Rows - 2
                    .TextMatrix(Lrow, 1) = "-1"
                Next
            End With
        Case "Delete"
          Call Hapus
        Case "Refresh"
           FillList
                
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

Private Sub Command1_Click()
If Trim(cboFlex) = "" Then
    MsgBox "Silahkan Pilih NIP Terlebih Dahulu", vbInformation
    cboFlex.SetFocus
    Exit Sub
End If
 FillList
End Sub
 
 

Private Sub fg_DblClick()
Dim J As String
With fg
If .Col = 2 Then
    DTPicker1.Value = Format(.TextMatrix(.Row, 2), "dd/MMM/yyyy")
    FillList
    Exit Sub
End If
If .Col >= 4 And .TextMatrix(.Row, .Col) <> "" And .TextMatrix(.Row, .Col) <> "Telat" Then
    J = FrmPM.Showdata(.TextMatrix(.Row, .Col), KodeDivisi)
    FrmPM.NIP = cboFlex
    FrmPM.show vbModal
End If
End With
End Sub

Private Sub Form_Load()
Dim RsTS As New ADODB.Recordset
    DTPicker1.Value = Date
    DTPicker1.CustomFormat = "dd/MMM/yyyy"
       AddKaryawan
     If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
      Skin1.ApplySkin hwnd
    End If
    Setgrid
    VSFg.Visible = False
    VSFlexGrid1.Visible = False
    If RsTS.State = adStateOpen Then RsTS.Close
     
    StrSQL = "Select IDtimesheet As Do,IDtimesheet,JamAwal As [Jam Awal],JamAkhir AS [Jam Akhir],Kd_divisi AS Divisi,NoProject As Project,Keterangan,Tanggal,Masuk,StatusPM From tbltimesheet Where Tanggal = '" & Format(DTPicker1, "yyyy/MM/dd") & "' And NIP = '" & cboFlex & "'  AND Status='Actual' Order By IDTimesheet Asc"
    RsTS.Open StrSQL, CN, adOpenStatic
    Set VsFlex.DataSource = RsTS
    With VsFlex
        
        .Rows = .Rows + 1
        .ColWidth(1) = 500
        .ColWidth(0) = 500
        .ColWidth(2) = 0
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 800
        .ColWidth(6) = 1000
        .ColWidth(7) = 1100
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColDataType(1) = flexDTBoolean
        .ColFormat(3) = "HH:mm"
        .ColFormat(4) = "HH:mm"
        .ColFormat(9) = "HH:mm"
    End With
End Sub
Private Sub AddKaryawan()

    Dim Cboid     As String
    Dim cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    cboid1 = vbNullString
    If StrUser = "3578" Or StrUser = "2976" Then
      StrSQL = "select * from Karyawan Where Status <> '14' AND Len(NIP) < 5 Order By NIP"
    Else
         StrSQL = "select * from Karyawan Where kd_divisi = '" & KodeDivisi & "' And Status <> '14'  AND Len(NIP) < 5  Order By NIP"
    End If
    Rscek.Open StrSQL, CN, adOpenStatic
    cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("NIP") & vbTab & Rscek("Nama")
      cboid1 = cboid1 + Cboid
      Rscek.MoveNext
    Loop
    cboFlex.ColComboList(0) = cboid1
    cboFlex.CellAlignment = flexAlignLeftCenter
End Sub
Public Sub FillList()

    Dim i, Split As Integer
    Dim CN2 As New ADODB.Connection
    Dim RsAbsensi As New ADODB.Recordset
    Dim Cboid, cboid1 As String
    Dim RsTS As New ADODB.Recordset
    Dim Hari As String, TglAbsen As Date
    Dim SQLAbsen As String
    On Error GoTo Adaerror
    Setgrid
    If Rscek.State = adStateOpen Then Rscek.Close
    Rscek.Open "Select * from absensi where nip = '" & cboFlex & "' And tgl= '" & Format(DTPicker1, "yyyy/MM/dd") & "'", CN, adOpenStatic
    If Not Rscek.EOF Then
        If Rscek!Kd_absensi = 4 Then
            LabelAbsen = "08:00"
            Label5 = "17:00"
        Else
            LabelAbsen = Format(Rscek!masuk, "HH:mm")
            Label5 = Format(Rscek!keluar, "HH:mm")
        End If
    Else
        If cboFlex = "287" Then
            LabelAbsen = "08:00"
            Label5 = "17:00"
        Else
            LabelAbsen = "00:00"
            LKeluar = "00:00"
            Label5 = "00:00"
        End If
    End If
      If LabelAbsen = "00:00" Then LKeluar = "00:00"
      Hari = Format(DTPicker1, "ddd")
   If Hari <> "Sat" And Hari <> "Sun" And Hari <> "Sabtu" And Hari <> "Minggu" Then
       StatusHari = "Kerja"
       If LabelAbsen <> "00:00" Then LKeluar = DateAdd("h", 9, LabelAbsen)
        If LKeluar <= "23:00" Then LKeluar = DateAdd("n", 30, LKeluar)
        
        StrSQL = "select tanggallibur from kalender " & _
            "where tanggallibur = '" & Format(DTPicker1, "MM/dd/yyyy") & "'"
        If Rscek.State = adStateOpen Then Rscek.Close
        Rscek.Open StrSQL, CN, adOpenStatic
        If Not Rscek.EOF Then
            StatusHari = "Libur"
            LKeluar = "00:00"
        End If
    Else
        StatusHari = "Libur"
        LKeluar = "00:00"
    End If
'    If LKeluar < "17:00" Then LKeluar = "17:30"
    LKeluar = Format(LKeluar, "HH:mm")
   With VsFlex
     
    Cboid = "08:00|08:30|09:00|09:30|10:00|10:30|11:00|11:30|12:00|13:00|13:30|14:00|14:30|15:00|15:30|16:00|16:30|17:00|17:30|18:00|18:30|19:00|19:30|20:00|20:30|21:00|21:30|22:00|22:30|23:00|23:30|00:00|00:30|01:00|01:30|02:00|02:30|03:00|03:30|04:00|04:30|05:00|05:30|06:00|06:30|07:00|07:30"
   
    If RsTS.State = adStateOpen Then RsTS.Close
    StrSQL = "Select IDtimesheet As Do,IDtimesheet,JamAwal As [Jam Awal],JamAkhir AS [Jam Akhir],Kd_divisi AS Divisi,NoProject As Project,Keterangan,Tanggal,Masuk,StatusPM From tbltimesheet Where Tanggal = '" & Format(DTPicker1, "yyyy/MM/dd") & "' And NIP = '" & cboFlex & "'  AND Status='Actual' Order By IDTimesheet Asc"
    RsTS.Open StrSQL, CN, adOpenStatic
    Set VsFlex.DataSource = RsTS
    VsFlex.ColDataType(1) = flexDTBoolean
         .Rows = .Rows + 1
        .ColWidth(1) = 500
        .ColWidth(0) = 500
        .ColWidth(2) = 0
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 800
        .ColWidth(6) = 1000
        .ColWidth(7) = 1100
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColFormat(3) = "HH:mm"
        .ColFormat(4) = "HH:mm"
        .ColFormat(9) = "HH:mm"
        .ColComboList(3) = Cboid
       .ColComboList(4) = Cboid
    If Rscek.State = adStateOpen Then Rscek.Close
    Rscek.Open "SELECT * from divisi where kd_bid >= 2 and kd_bid <= 20 order by kd_bid", CN, adOpenStatic
    'Rscek.Open StrSQL, CN, adOpenStatic
      cboid1 = " "
    Do Until Rscek.EOF
         Cboid = "|" & Rscek("Kd_div") & vbTab & Rscek("NM_DIV")
         cboid1 = cboid1 + Cboid
         Rscek.MoveNext
    Loop
        .ColComboList(5) = cboid1
    AddProject (KodeDivisi)
    For Lrow = 1 To .Rows - 2
        .TextMatrix(Lrow, 0) = Lrow
        .TextMatrix(Lrow, 1) = ""
        If Trim(.TextMatrix(Lrow, 7)) = "Lembur" Then
           For i = 3 To .Cols - 2
               .Row = Lrow
               .Col = i
               .CellBackColor = &HC0FFFF
           Next
        End If
       If .TextMatrix(Lrow, 10) = 1 Then
            For i = 1 To .Cols - 2
                .Row = Lrow
                .Col = i
                .CellBackColor = &HE0E0E0
            Next
        End If
        
    Next
    Showdata
        .Col = 3
        .Row = .Rows - 1
'        .SetFocus
    End With
Exit Sub
Adaerror:
  MsgBox err.Description

End Sub
Sub AddProject(Nama As String)
Dim Cboid, cboid1 As String
       If Rscek.State = adStateOpen Then Rscek.Close
    If Nama = 51 Then
       StrSQL = "select *  from project Where kd_divisi='" & Nama & "' and Status = 'Terpakai' order by project.kode"
    Else
       StrSQL = "select *  from project Where kd_divisi='" & Nama & "' and Tgl_Akhir >= '" & Format(DTPicker1, "MM/dd/yyyy") & "' order by project.kode"
    End If
    Rscek.Open StrSQL, CN, adOpenStatic
    cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("Kode") & vbTab & Rscek("Nama")
      cboid1 = cboid1 + Cboid
      Rscek.MoveNext
    Loop
        VsFlex.ColComboList(6) = cboid1
        
End Sub
Private Sub Form_Resize()
 On Error Resume Next
        CmdClose.Width = Me.Width - 100
        With VsFlex
            .Width = Me.ScaleWidth / 3 + 1200
            .Height = ScaleHeight - .Top - .Left - Picture3.Height
            fg.Width = Me.ScaleWidth / 1.7 - 150
            fg.Left = .Width + 120
            fg.Height = ScaleHeight - .Top - .Left - Picture3.Height

        End With
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
    Set FrmTimesheet2 = Nothing
End Sub

Private Sub VsFlex_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With VsFlex
On Error GoTo Adaerror
Dim i, Sjam, Smenit As Integer
Dim JmlLoop, x As Integer
Dim Jam1 As Date
Dim BatasJam As Boolean
Dim LastJam As Date
Dim Jam2 As String

If .TextMatrix(Row, 3) = "" Then Exit Sub
If .Col = 1 Then Exit Sub
If Col = 5 Then
   AddProject (.TextMatrix(Row, 5))
End If

If Col = 3 Or Col = 4 Then
    If Trim(.TextMatrix(Row, 5)) = "" Then .TextMatrix(Row, 5) = KodeDivisi
     If StatusHari = "Kerja" Then
      If .TextMatrix(.Row, .Col) > LKeluar Or .TextMatrix(.Row, Col) <= "07:00" Then
            .TextMatrix(.Row, 7) = "Lembur"
            For i = 3 To .Cols - 2
                .Row = .Row
                .Col = i
                .CellBackColor = &HC0FFFF
            Next
        Else
            .TextMatrix(.Row, 7) = "Timesheet"
            For i = 3 To .Cols - 2
                .Row = .Row
                .Col = i
                .CellBackColor = vbWhite
            Next
        End If
     Else
        .TextMatrix(.Row, 7) = "Lembur"
         For i = 3 To .Cols - 2
               .Row = .Row
               .Col = i
               .CellBackColor = &HC0FFFF
           Next
     End If
End If
 
If Row = 1 Then
    If Col = 3 Or Col = 4 Then
   'Untuk Ngcek Absen And .TextMatrix(Row, 7) = "Timesheet"
    If LabelAbsen >= "08:00" Then
        Jam1 = DateAdd("n", -30, LabelAbsen)
'        If Right(.TextMatrix(Row, 3), 2) <= 30 Then
            If Format(.TextMatrix(Row, 3), "HH:mm") < Format(Jam1, "HH:mm") Then
                  MsgBox "Invalid Time", vbCritical
                  FillList
               Exit Sub
             End If
      End If
        If Format(.TextMatrix(Row, 3), "HH:mm") = Format(.TextMatrix(Row, 4), "HH:mm") Then
               MsgBox "Invalid Time", vbCritical
               FillList
            Exit Sub
        End If
    If StatusHari = "Kerja" Then
         If Trim(.TextMatrix(.Row, 4)) = "" Then Exit Sub
       
            If .TextMatrix(Row, 4) >= "17:00" Then
                 .TextMatrix(.Row, 4) = "17:00"
                .TextMatrix(.Row, 7) = "Timesheet"
                For i = 3 To .Cols - 2
                    .Row = .Row
                    .Col = i
                    .CellBackColor = vbWhite
                Next
            End If
     
    End If
    End If
Else
    If Col = 3 Or Col = 4 Then
    
        LastJam = Format(.TextMatrix(.Row - 1, 4), "HH:mm")
       If StatusHari = "Kerja" Then If LastJam = "12:00" Then LastJam = "13:00"
        .TextMatrix(Row, 3) = LastJam
        If .TextMatrix(Row, Col) = "" Then Exit Sub
        
        If StatusHari = "Kerja" Then
            BatasJam = False
            For i = 1 To .Rows - 2
                If .TextMatrix(i, 4) = "17:00" Then
                    BatasJam = True
                    Exit For
                End If
            Next
            If BatasJam = False Then
               If .TextMatrix(Row, 4) > "17:00" Or .TextMatrix(Row, 4) < "08:00" Then .TextMatrix(Row, 4) = "17:00"
            End If
        End If
            
            If StatusHari = "Kerja" Then
                 Jam1 = CDate(.TextMatrix(Row, Col))
                 DTPicker3 = Format(LKeluar, "HH:mm")
                 Jam2 = DateDiff("n", DTPicker3, Jam1)
'               If (.TextMatrix(Row, Col) > LKeluar Or .TextMatrix(Row, Col) <= "07:00") And Abs(Jam2) > 30 Then
                 If (.TextMatrix(Row, Col) > LKeluar Or .TextMatrix(Row, Col) <= "07:00") Then
                  .TextMatrix(Row, 7) = "Lembur"
                   For i = 3 To .Cols - 2
                      .Row = .Row
                      .Col = i
                      .CellBackColor = &HC0FFFF
                  Next
               Else
                   .TextMatrix(.Row, 7) = "Timesheet"
                    For i = 3 To .Cols - 2
                        .Row = .Row
                        .Col = i
                        .CellBackColor = vbWhite
                    Next
               End If
            Else
               .TextMatrix(.Row, 7) = "Lembur"
                For i = 3 To .Cols - 2
                      .Row = .Row
                      .Col = i
                      .CellBackColor = &HC0FFFF
                  Next
            End If
          
            If .TextMatrix(.Row, 7) = "Lembur" And .TextMatrix(.Row, 4) <> "" Then
                  Jam1 = Format(.TextMatrix(.Row, 4), "HH:mm")
                   If Format(.TextMatrix(.Row, 4), "HH:mm") > "07:00" And Format(.TextMatrix(.Row, 4), "HH:mm") < Format(.TextMatrix(.Row, 3), "HH:mm") Then
                       MsgBox "Invalid Time", vbCritical
                       .TextMatrix(.Row, 4) = "07:00"
                       Exit Sub
                   End If
            End If

        For i = 1 To .Rows - 2
           If (Jam1 = Format(.TextMatrix(i, 3), "HH:mm") Or Jam1 < Format(.TextMatrix(i, 4), "HH:mm")) And i < Row And .TextMatrix(.Row, 7) = "Timesheet" Then
              MsgBox "Invalid Time", vbCritical
               FillList
              Exit Sub
           End If
        Next

    End If
End If
Select Case Col
    Case 3, 4, 5
        Col = Col + 1
        .Col = Col
        .EditCell
         
End Select
If Col >= 3 And VsFlex.TextMatrix(Row, 2) <> "" Then
    Call SimpanData(Row)
    Exit Sub
End If
If Col = 6 And .TextMatrix(Row, 2) = "" And Trim(.TextMatrix(Row, 6)) <> "" Then
    Call SimpanData(Row)
End If
End With
Exit Sub
Adaerror:
    MsgBox err.Description
End Sub

Private Sub VsFlex_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VsFlex
'    If .Col = Col And .CellBackColor = &HE0E0E0 Then .Editable = flexEDNone: Exit Sub

If LabelAbsen = "00:00" Then .Editable = flexEDNone: Exit Sub
   If .Col = 5 Or .Col = 7 Then
        .Editable = flexEDNone
        .EditMaxLength = 0
    Else
        .Editable = flexEDKbdMouse
    End If
End With
End Sub


Private Sub VsFlex_Click()
   VsFlex_GotFocus
End Sub

Private Sub VsFlex_GotFocus()
On Error Resume Next
 With VsFlex
 
 If LabelAbsen = "00:00" Then .Editable = flexEDNone: Exit Sub

'    If .Col = 7 Or .CellBackColor = &HE0E0E0 Then
'        .Editable = flexEDNone
'        .TextMatrix(.Row, 1) = ""
'        Exit Sub
'    Else
        .Editable = flexEDKbdMouse
'    End If
    If .Row = -1 Then .Row = 1
    If .Col = 1 And .TextMatrix(.Row, 3) = "" Then
        .TextMatrix(.Row, 1) = ""
    ElseIf .Col > 3 And .TextMatrix(.Row, 3) = "" Then
        .Col = 3
        .EditCell
    End If
    
End With
End Sub

Private Sub VsFlex_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96

End Sub

Private Sub VsFlex_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If Col = 7 Then KeyAscii = 0: Exit Sub
If KeyAscii = 39 Then KeyAscii = 96
If KeyAscii = 8 Or KeyAscii = 27 Then Exit Sub

End Sub

Private Sub VsFlex_KeyUp(KeyCode As Integer, Shift As Integer)
With VsFlex
If KeyCode = 13 Then
    Select Case .Col
        Case 3
            If .Text <> "" And .TextMatrix(.Row, 2) = "" Then
               
                .Col = 4
                .EditCell
                .TextMatrix(.Row, 8) = Format(DTPicker1, "MM/dd/yyyy")
             End If
                Exit Sub
        Case 4
            If .Text <> "" And .TextMatrix(.Row, 2) = "" Then
                .Col = 6
                .EditCell
                
            End If
            Exit Sub
        Case 6
            If .TextMatrix(.Row, 2) = "" And .TextMatrix(.Row, 3) <> "" Then
                Call SimpanData(.Row)
            End If
    End Select
    
End If
End With
End Sub

Private Sub VsFlex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    MDIMENU.MnuRView.Visible = False
    MDIMENU.MnuRBatal.Visible = False
    MDIMENU.mnuRSearch.Visible = False
    MDIMENU.mnuRPrint.Visible = False
    PopupMenu MDIMENU.mnuRecA
End If
End Sub
'Prosedur Simpan Data
Private Sub SimpanData(ByVal Row As Long)
Dim StrSQL As String
Dim i As Integer

On Error GoTo Adaerror

With VsFlex
If Trim(.TextMatrix(Row, 6)) = "" Then
    MsgBox "Project Tidak Boleh Kosong Atau Isikan Dengan Project Umum", vbCritical
    .Col = 6
    .EditCell
    Exit Sub
End If
If Trim(.TextMatrix(Row, 4)) = "" Then
    MsgBox "Jam Akhir Masih Kosong", vbCritical
    .Col = 4
    .EditCell
    Exit Sub
End If
If .Rows >= 3 Then
   If StatusHari = "Libur" Then
        If .TextMatrix(Row - 1, 6) = .TextMatrix(Row, 6) Then
            MsgBox "Untuk Project " & .TextMatrix(Row, 6) & "  Mohon Disatukan Jam Akhirnya Dengan Yang Atas", vbInformation
            .TextMatrix(Row - 1, 4) = .TextMatrix(Row, 4)
            .RemoveItem (Row)
            Row = Row - 1
        End If
   End If
End If
If .TextMatrix(Row, 2) = "" Then
      '-----Ambil ID -------------------
        GetNomorID ("tbltimesheet")
        StrKodeID = NewID
    '---------------------------------
    If StrUser = "3578" Then
        StrSQL = "Insert Into tbltimesheet(IDtimesheet,NIP,kd_divisi,Status,NoProject,Tanggal,JamAwal,JamAkhir,keterangan,Masuk,Keluar,last_update,last_user,StatusPM,StatusDivisi,Hari,TotalKerja,ProjectUmum) Values "
        StrSQL = StrSQL & "('" & StrKodeID & "','" & cboFlex & "','" & .TextMatrix(Row, 5) & "','Actual','" & .TextMatrix(Row, 6) & "','" & Format(DTPicker1, "MM/dd/yyyy") & "','" & .TextMatrix(Row, 3) & "','" & .TextMatrix(Row, 4) & "','" & .TextMatrix(Row, 7) & "','" & LabelAbsen & "','" & Label5 & "','" & ServerTime & "','" & StrUser & "',0,0,'" & StatusHari & "','" & TotalJamPUmum & "', '" & ProjectUmum & "')"
        PerintahExecute (StrSQL)
    Else
        StrSQL = "Insert Into tbltimesheet(IDtimesheet,NIP,kd_divisi,Status,NoProject,Tanggal,JamAwal,JamAkhir,keterangan,Masuk,Keluar,last_update,last_user,StatusPM,StatusDivisi,Hari,TotalKerja,ProjectUmum) Values "
        StrSQL = StrSQL & "('" & StrKodeID & "','" & cboFlex & "','" & .TextMatrix(Row, 5) & "','Actual','" & .TextMatrix(Row, 6) & "','" & Format(DTPicker1, "MM/dd/yyyy") & "','" & .TextMatrix(Row, 3) & "','" & .TextMatrix(Row, 4) & "','" & .TextMatrix(Row, 7) & "','" & LabelAbsen & "','" & Label5 & "','" & ServerTime & "','" & StrUser & "',0,0,'" & StatusHari & "','" & TotalJamPUmum & "', '" & ProjectUmum & "')"
        PerintahExecute (StrSQL)
        
        StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(ServerTime, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Tambah Timesheet, " & StrNIPUser & " &  " & .TextMatrix(Row, 3) & "','Edit Timesheet')"
        PerintahExecute (StrSQL)
    End If
    .TextMatrix(Row, 0) = Val(.TextMatrix(.Row - 1, 0)) + 1
    .TextMatrix(Row, 2) = StrKodeID
    
    If .TextMatrix(.Rows - 1, 3) <> "" Then
        .Rows = .Rows + 1
        .Row = .Rows - 1
    End If
    .Col = 3
    .EditCell
    .Refresh
     
Else
   If StrUser = "3578" Then
        StrSQL = "Update tbltimesheet Set Hari = '" & StatusHari & "',kd_divisi='" & .TextMatrix(Row, 5) & "',Status='Actual' ,Tanggal='" & Format(DTPicker1, "MM/dd/yyyy") & "',JamAwal='" & .TextMatrix(Row, 3) & "',JamAkhir='" & .TextMatrix(Row, 4) & "',NoProject = '" & .TextMatrix(Row, 6) & "',keterangan='" & .TextMatrix(Row, 7) & "',Masuk='" & LabelAbsen & "',last_update='" & ServerTime & "',last_user='" & StrNIPUser & "',TotalKerja = '" & TotalJamPUmum & "' Where IDtimesheet='" & .TextMatrix(Row, 2) & "' "
   Else
        StrSQL = "Update tbltimesheet Set Hari = '" & StatusHari & "',kd_divisi='" & .TextMatrix(Row, 5) & "',Status='Actual' ,Tanggal='" & Format(DTPicker1, "MM/dd/yyyy") & "',JamAwal='" & .TextMatrix(Row, 3) & "',JamAkhir='" & .TextMatrix(Row, 4) & "',NoProject = '" & .TextMatrix(Row, 6) & "',keterangan='" & .TextMatrix(Row, 7) & "',Masuk='" & LabelAbsen & "',last_update='" & ServerTime & "',last_user='" & StrNIPUser & "',TotalKerja = '" & TotalJamPUmum & "', ProjectUmum= '" & ProjectUmum & "' Where IDtimesheet='" & .TextMatrix(Row, 2) & "' "
   End If
    PerintahExecute (StrSQL)
 
    StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(ServerTime, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Rubah Timesheet, " & StrNIPUser & " &  " & .TextMatrix(Row, 3) & "','Edit Timesheet')"
    PerintahExecute (StrSQL)
 
End If
    If StatusHari = "Kerja" Then
        If .TextMatrix(Row, 7) = "Lembur" Then
           For i = 3 To .Cols - 2
               .Row = Row
               .Col = i
               .CellBackColor = &HC0FFFF
           Next
        End If
    Else
         For i = 3 To .Cols - 2
               .Row = Row
               .Col = i
               .CellBackColor = &HC0FFFF
           Next
    End If
'    FillList
Exit Sub
Adaerror:
MsgBox err.Description
End With
End Sub
Private Sub Hapus()
Dim StrSQL As String
Dim Lrow As Integer
Dim Tanya As String
Dim ErrConn As Long
 If CekCurek("hapus", VsFlex) = False Then FillList: Exit Sub
            With VsFlex
            If MsgBox("Apakah Anda yakin ingin menghapus Data ?", vbQuestion + vbYesNo, "Konfirmasi hapus") = vbNo Then
               Exit Sub
            Else
             Do Until Lrow = .Rows - 1
             
                If .TextMatrix(Lrow, 1) = "-1" Then
                         StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(ServerTime, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Hapus Timesheet, " & StrNIPUser & " &  " & .TextMatrix(Lrow, 3) & "','Edit Timesheet')"
                         PerintahExecute (StrSQL)
                                
'                        StrSQL = "Delete From Timesheet where IDtimesheet = '" & .TextMatrix(Lrow, 2) & "'"
'                        PerintahExecute (StrSQL)
'
                        StrSQL = "Delete From tbltimesheet"
                        StrSQL = StrSQL & " where IDtimesheet = '" & .TextMatrix(Lrow, 2) & "'"
                        CN.Execute StrSQL
                        .RemoveItem (Lrow)
                        Lrow = Lrow - 1
                End If
                Lrow = Lrow + 1
            Loop
                    FillList
                End If
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
Dim IDWaktu, Hari As String
Dim Jam1, Jam2 As Date
Dim TglAwal, TglAkhir As Date
Dim TotalHari, Selisih As Double
Dim RsTS As New ADODB.Recordset
 
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
    VSFg.Cols = 11
With VSFlexGrid1
        .Rows = 1
    If RsTS.State = adStateOpen Then RsTS.Close
    StrSQL = "SELECT IDtimesheet,Tanggal,JamAwal As [Jam Awal],JamAkhir AS [Jam Akhir],Status,NoProject As Project,Keterangan,Tanggal,Masuk,StatusPM,StatusDivisi,Keluar From tbltimesheet Where Tanggal Between '" & Format(TglAwal, "yyyy/MM/dd") & "' And '" & Format(TglAkhir, "yyyy/MM/dd") & "' And NIP = '" & cboFlex & "' And Status='Actual' Order By Tanggal DESC,IDTimesheet ASC"
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
            Rscek.Open "Select * from absensi where nip = '" & cboFlex & "' And tgl= '" & Format(TglAwal, "MM/dd/yyyy") & "'", CN, adOpenStatic
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
'                    VSFg.TextMatrix(VSFg.Rows - 1, 4) = StrNIPUser
                    VSFg.TextMatrix(VSFg.Rows - 1, 5) = .TextMatrix(Lrow, 5)
                    VSFg.TextMatrix(VSFg.Rows - 1, 6) = Format(.TextMatrix(Lrow, 9), "HH:mm")
                     VSFg.TextMatrix(VSFg.Rows - 1, 7) = .TextMatrix(Lrow, 7)
                     VSFg.TextMatrix(VSFg.Rows - 1, 8) = .TextMatrix(Lrow, 10)
                     VSFg.TextMatrix(VSFg.Rows - 1, 9) = .TextMatrix(Lrow, 11)
                     VSFg.TextMatrix(VSFg.Rows - 1, 10) = Format(.TextMatrix(Lrow, 12), "HH:mm")
                    JmlLoop = JmlLoop + 1
                    DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                    If Format(DTPicker3, "HH:mm") = "12:30" Then DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                Loop
    Next
End With
    
'Disiplit Perhari
With fg
'.Rows = 1
'If VSFg.Rows >= 2 Then .Rows = 2
'.Cols = 6
For i = 1 To VSFg.Rows - 1
    Jam = VSFg.TextMatrix(i, 1)

'    If i <> VSFg.Rows - 1 Then
'        If VSFg.TextMatrix(i, 4) = VSFg.TextMatrix(i + 1, 4) Then
'
'            If VSFg.TextMatrix(i, 2) = VSFg.TextMatrix(i + 1, 2) And VSFg.TextMatrix(i, 5) = VSFg.TextMatrix(i + 1, 5) Then
'                J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
'            Else
                 J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
'                 .Rows = .Rows + 1
'           End If
'        Else
'                .Rows = .Rows + 1
'                J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
              
'        End If
'    Else
'        If VSFg.TextMatrix(i, 4) <> .TextMatrix(.Rows - 1, 4) Then
'             J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
'        Else
'             J = Tampil(Jam, i, VSFg.TextMatrix(i, 5), VSFg.TextMatrix(i, 8), VSFg.TextMatrix(i, 9))
'      End If
'    End If
   
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
                    Jam2 = DateAdd("n", 15, Jam2)
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
                                .Col = J
                                .Row = i
                               If Jam1 = "00:00" Then .TextMatrix(i, J) = "LIBUR"
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
End Sub

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
Sub Setgrid()
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
    .TextMatrix(0, 5) = "Keluar"
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




