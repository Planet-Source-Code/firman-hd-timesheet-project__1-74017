VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmListPlan 
   Caption         =   "List Plan Timesheet"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   10695
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10695
      TabIndex        =   7
      Top             =   6225
      Width           =   10695
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
         TabIndex        =   8
         Top             =   0
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
      ScaleWidth      =   10665
      TabIndex        =   1
      Top             =   0
      Width           =   10695
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   4800
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
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
         Format          =   46530563
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   7800
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   46530562
         CurrentDate     =   39940
      End
      Begin VSFlex8Ctl.VSFlexGrid VSfg 
         Height          =   735
         Left            =   7680
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   7815
         _cx             =   13785
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
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3120
         TabIndex        =   10
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
         Format          =   46530563
         CurrentDate     =   39931
      End
      Begin VSFlex8Ctl.VSFlexGrid cboFlex 
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Top             =   720
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
         FormatString    =   $"FrmListPlan.frx":0000
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
      Begin VSFlex8Ctl.VSFlexGrid Combo2 
         Height          =   315
         Left            =   1320
         TabIndex        =   14
         Top             =   1080
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
         FormatString    =   $"FrmListPlan.frx":0029
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
         TabIndex        =   16
         Top             =   1080
         Width           =   1215
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
         TabIndex        =   15
         Top             =   720
         Width           =   1695
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
         TabIndex        =   11
         Top             =   240
         Width           =   375
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
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "List Plan Timesheet "
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
         Left            =   5640
         TabIndex        =   5
         Top             =   1080
         Width           =   3495
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   3975
      Left            =   7440
      TabIndex        =   0
      Top             =   4800
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
      FormatString    =   $"FrmListPlan.frx":0052
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
   Begin ACTIVESKINLibCtl.Skin Skin2 
      Left            =   8400
      OleObjectBlob   =   "FrmListPlan.frx":0131
      Top             =   2760
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FrmListPlan.frx":0365
      Top             =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3975
      Left            =   0
      TabIndex        =   9
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
      FormatString    =   $"FrmListPlan.frx":0599
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
Attribute VB_Name = "FrmListPlan"
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

Private Sub Command1_Click()
Showdata
End Sub

Private Sub fg_Click()
fg_GotFocus
End Sub

Private Sub fg_GotFocus()
With fg
    If .Col = 1 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With

End Sub

Private Sub fg_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
With fg
    If .Col = 1 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub

Private Sub Form_Load()
    AddKaryawan
    AddProject
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
    Showdata

End Sub
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
    .ColWidth(1) = 500
    .ColWidth(2) = 1200
    .ColWidth(3) = 0
     .ColWidth(4) = 800
     .ColWidth(5) = 0
'    .ColWidth(15) = 0
    For i = 6 To 21
       .ColWidth(i) = 1200
'       .ColComboList(I) = cboid1
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
Private Sub AddProject()

    Dim Cboid     As String
    Dim Cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    Cboid1 = vbNullString
    StrSQL = "SELECT * FROM Project Where KD_DIVisi = '" & KodeDivisi & "' And Project.Status ='Terpakai' Order By Kode"

    Rscek.Open StrSQL, CN, adOpenStatic
    Cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("Kode") & vbTab & Rscek("Nama")
      Cboid1 = Cboid1 + Cboid
      Rscek.MoveNext
    Loop
    Combo2.ColComboList(0) = Cboid1
End Sub
Private Sub AddKaryawan()

    Dim Cboid     As String
    Dim Cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    Cboid1 = vbNullString
    StrSQL = "select * from Karyawan Where kd_divisi = '" & KodeDivisi & "' And Status <> '14' Order By Nama"
    Rscek.Open StrSQL, CN, adOpenStatic
    Cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("NIP") & vbTab & Rscek("Nama")
      Cboid1 = Cboid1 + Cboid
      Rscek.MoveNext
    Loop
    CboFlex.ColComboList(0) = Cboid1
     CboFlex.CellAlignment = flexAlignLeftCenter
End Sub
Private Sub Form_Resize()
    On Error Resume Next
        CmdClose.Width = Me.Width - 100
        With fg
             .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture3.Height

        End With
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
    Set FrmListPlan = Nothing
End Sub

Private Sub fg_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    MDIMENU.MnuRView.Visible = False
    MDIMENU.MnuRBatal.Visible = False
    MDIMENU.mnuRSearch.Visible = False
    MDIMENU.mnuRPrint.Visible = False
    PopupMenu MDIMENU.mnuRecA
End If
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
                         StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Hapus Timesheet, " & StrNIPUser & " &  " & .TextMatrix(Lrow, 3) & "','Plan Timesheet')"
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
AdaError:
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
   
    
With VSFlexGrid1
        .Rows = 1
        VSFg.Rows = 1
        fg.Rows = 1
    If RsTS.State = adStateOpen Then RsTS.Close
    StrSQL = "SELECT IDtimesheet,Tanggal,JamAwal As [Jam Awal],JamAkhir AS [Jam Akhir],Status,NoProject As Project,Keterangan,Tanggal,Masuk,NIP From tbltimesheet Where Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And '" & Format(DTPicker2, "MM/dd/yyyy") & "' And Kd_Divisi = '" & KodeDivisi & "' And Status ='Plan'"
    
    If CboFlex <> "" Then StrSQL = StrSQL & " AND NIP = '" & CboFlex & "'"
    If Combo2 <> "" Then StrSQL = StrSQL & " AND NoProject = '" & Combo2 & "'"
    
    StrSQL = StrSQL & " Order By Tanggal DESC,Status,IDTimesheet ASC"
    
    RsTS.Open StrSQL, CN, adOpenStatic
    Set .DataSource = RsTS
    .ColDataType(2) = flexDTDate
    .ColFormat(2) = "dd/MM/yyyy"
    For Lrow = 1 To .Rows - 1
        If .TextMatrix(Lrow, 3) = "" Then .TextMatrix(Lrow, 3) = "00:00"
        If .TextMatrix(Lrow, 4) = "" Then .TextMatrix(Lrow, 4) = "00:00"
        Jam1 = CDate(.TextMatrix(Lrow, 3))
        Jam2 = CDate(.TextMatrix(Lrow, 4))
        DTPicker3.Value = Jam1
        i = DateDiff("n", Jam2, Jam1)
        i = Abs(i)
        Split = Abs(i / 30)
        JmlLoop = 0
        If Split > 1 Then
            Do Until JmlLoop = Split
                VSFg.Rows = VSFg.Rows + 1
                
                VSFg.TextMatrix(VSFg.Rows - 1, 0) = VSFg.Rows - 1
                IDWaktu = Format(DTPicker3.Value, "HH:mm")
                If Trim(.TextMatrix(Lrow, 7)) = "Timesheet" And IDWaktu = "17:00" Then IDWaktu = "16:30"
                VSFg.TextMatrix(VSFg.Rows - 1, 1) = IDWaktu
                VSFg.TextMatrix(VSFg.Rows - 1, 2) = Format(.TextMatrix(Lrow, 2), "dd/MM/yyyy")
                VSFg.TextMatrix(VSFg.Rows - 1, 3) = .TextMatrix(Lrow, 6)
                VSFg.TextMatrix(VSFg.Rows - 1, 4) = .TextMatrix(Lrow, 10)
                VSFg.TextMatrix(VSFg.Rows - 1, 5) = .TextMatrix(Lrow, 5)
                VSFg.TextMatrix(VSFg.Rows - 1, 6) = Format(.TextMatrix(Lrow, 9), "HH:mm")
                 VSFg.TextMatrix(VSFg.Rows - 1, 7) = .TextMatrix(Lrow, 7)
                JmlLoop = JmlLoop + 1
                DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                If Format(DTPicker3, "HH:mm") = "12:30" Then DTPicker3.Value = DateAdd("n", 30, DTPicker3)
            Loop
        Else
              VSFg.Rows = VSFg.Rows + 1
                
                VSFg.TextMatrix(VSFg.Rows - 1, 0) = VSFg.Rows - 1
                IDWaktu = Format(DTPicker3.Value, "HH:mm")
                If Trim(.TextMatrix(Lrow, 7)) = "Timesheet" And IDWaktu = "17:00" Then IDWaktu = "16:30"
                VSFg.TextMatrix(VSFg.Rows - 1, 1) = IDWaktu
                VSFg.TextMatrix(VSFg.Rows - 1, 2) = Format(.TextMatrix(Lrow, 2), "dd/MM/yyyy")
                VSFg.TextMatrix(VSFg.Rows - 1, 3) = .TextMatrix(Lrow, 6)
                VSFg.TextMatrix(VSFg.Rows - 1, 4) = .TextMatrix(Lrow, 10)
                VSFg.TextMatrix(VSFg.Rows - 1, 5) = .TextMatrix(Lrow, 5)
                VSFg.TextMatrix(VSFg.Rows - 1, 6) = Format(.TextMatrix(Lrow, 9), "HH:mm")
                 VSFg.TextMatrix(VSFg.Rows - 1, 7) = .TextMatrix(Lrow, 7)
                JmlLoop = JmlLoop + 1
                DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                If Format(DTPicker3, "HH:mm") = "12:30" Then DTPicker3.Value = DateAdd("n", 30, DTPicker3)
        End If
    Next
End With
    If VSFlexGrid1.Rows = 1 Then MsgBox "Data Tidak Ditemukan", vbCritical: Exit Sub
'Disiplit Perhari
With fg
.Rows = 1
If VSFlexGrid1.Rows >= 2 Then .Rows = 2
'.Cols = 6
For i = 1 To VSFg.Rows - 1
    Jam = VSFg.TextMatrix(i, 1)

    If i <> VSFg.Rows - 1 Then
        If VSFg.TextMatrix(i, 4) = VSFg.TextMatrix(i + 1, 4) Then
           
            If VSFg.TextMatrix(i, 2) = VSFg.TextMatrix(i + 1, 2) And VSFg.TextMatrix(i, 5) = VSFg.TextMatrix(i + 1, 5) Then
                J = Tampil(Jam, i, VSFg.TextMatrix(i, 5))
            Else
                 J = Tampil(Jam, i, VSFg.TextMatrix(i, 5))
                 .Rows = .Rows + 1
           End If
        Else
                .Rows = .Rows + 1
                J = Tampil(Jam, i, VSFg.TextMatrix(i, 5))
              
        End If
    Else
        If VSFg.TextMatrix(i, 4) <> .TextMatrix(.Rows - 1, 4) Then
             J = Tampil(Jam, i, VSFg.TextMatrix(i, 5))
        Else
             J = Tampil(Jam, i, VSFg.TextMatrix(i, 5))
      End If
    End If
   
Next
     
    
    For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = i
    Next
'            For J = 6 To .Cols - 1
'                .ColWidth(J) = 1300
'                Jam1 = Format(.TextMatrix(I, 3), "HH:mm")
'                Jam2 = Left(.TextMatrix(0, J), 5)
'                If J <= 22 Then If Jam2 < Jam1 Then .TextMatrix(I, J) = "Telat"
'
'            Next
'
'
'    Next
    .ColDataType(2) = flexDTDate
    .ColFormat(2) = "dd/MM/yyyy"
'    .ColWidth(3) = 800
   If .Cols >= 14 Then .ColWidth(14) = 0
End With
End Sub

Function Tampil(ByVal Jam As String, ByVal i As Integer, status As String)
Dim Project As String, AdaKolom As Boolean
Dim JmlCol, PosisiKol As Integer
Dim PosisiJam As String
With fg
 
    .TextMatrix(.Rows - 1, 2) = Format(VSFg.TextMatrix(i, 2), "dd/MM/yyyy")
    .TextMatrix(.Rows - 1, 3) = VSFg.TextMatrix(i, 6)
    .TextMatrix(.Rows - 1, 4) = VSFg.TextMatrix(i, 4)
    Project = VSFg.TextMatrix(i, 3)
    If status <> "" Then .TextMatrix(.Rows - 1, 5) = Trim(status)
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
        
    
         .TextMatrix(.Rows - 1, PosisiKol) = Project
    End If
  End With
End Function



