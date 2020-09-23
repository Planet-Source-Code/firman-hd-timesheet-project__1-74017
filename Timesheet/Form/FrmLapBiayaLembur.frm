VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmLapBiayaLembur 
   Caption         =   "Rekap Biaya Lembur"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   9180
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid VSFlex 
      Height          =   4695
      Left            =   0
      TabIndex        =   15
      Top             =   2160
      Width           =   8535
      _cx             =   15055
      _cy             =   8281
      Appearance      =   1
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
      SelectionMode   =   3
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
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9180
      TabIndex        =   19
      Top             =   6000
      Width           =   9180
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
         Left            =   840
         TabIndex        =   20
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2145
      ScaleWidth      =   9150
      TabIndex        =   0
      Top             =   0
      Width           =   9180
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmLapBiayaLembur.frx":0000
         Left            =   1680
         List            =   "FrmLapBiayaLembur.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Print out"
         Height          =   375
         Left            =   7320
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   5880
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Export To Sheet"
         Height          =   375
         Left            =   7320
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Before Sheet"
         Height          =   375
         Left            =   10560
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Open XLS"
         Height          =   375
         Left            =   5880
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Next Sheet"
         Height          =   375
         Left            =   12000
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog dlg 
         Left            =   5520
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   720
         TabIndex        =   8
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
         Left            =   3000
         TabIndex        =   9
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
         Left            =   1680
         TabIndex        =   21
         Top             =   1320
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
         FormatString    =   $"FrmLapBiayaLembur.frx":0021
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
         Left            =   1680
         TabIndex        =   22
         Top             =   1680
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
         FormatString    =   $"FrmLapBiayaLembur.frx":004A
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
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label5 
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
         TabIndex        =   18
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         TabIndex        =   13
         Top             =   1320
         Width           =   1695
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
         TabIndex        =   12
         Top             =   120
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
         Left            =   2520
         TabIndex        =   11
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Divisi"
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
         Top             =   960
         Width           =   1695
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1320
      OleObjectBlob   =   "FrmLapBiayaLembur.frx":0073
      Top             =   6000
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4695
      Left            =   0
      TabIndex        =   14
      Top             =   2160
      Width           =   8535
      _cx             =   15055
      _cy             =   8281
      Appearance      =   1
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
      SelectionMode   =   3
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
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   53870594
         CurrentDate     =   39940
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   960
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   53870594
         CurrentDate     =   39940
      End
   End
End
Attribute VB_Name = "FrmLapBiayaLembur"
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

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
AddProject
getKaryawan
End Sub
Private Sub AddProject()

    Dim Cboid     As String
    Dim cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    cboid1 = " "
    If Trim(Combo1) <> "" Then
       StrSQL = "SELECT Project.Kode, Project.Nama, Divisi.NM_DIV AS Divisi, Project.Status FROM Project INNER JOIN Divisi ON Project.Kd_Divisi = Divisi.KD_DIV Where NM_DIV = '" & Combo1 & "' Order By Kode"
    Else
        StrSQL = "SELECT Project.Kode, Project.Nama, Divisi.NM_DIV AS Divisi, Project.Status FROM Project INNER JOIN Divisi ON Project.Kd_Divisi = Divisi.KD_DIV Order By Kode"
    End If
    Rscek.Open StrSQL, CN, adOpenStatic
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("Kode") & vbTab & Rscek("Nama")
      cboid1 = cboid1 + Cboid
      Rscek.MoveNext
    Loop
    Combo2.ColComboList(0) = cboid1
End Sub
Private Sub Combo1_GotFocus()
Combo1.BackColor = &HC0FFFF
End Sub

Private Sub Combo1_LostFocus()
Combo1.BackColor = vbWhite
End Sub
Sub SimpanGaji()
Dim Row As Integer
If Rscek.State = adStateOpen Then Rscek.Close
Rscek.Open "SELECT * From tblgaji", CN, adOpenStatic
If Rscek.RecordCount = 0 Then
    With fg
        For Row = 2 To .Rows - 1
            If .TextMatrix(Row, 1) = "" Then Exit For
            StrSQL = "Insert Into TblGaji(IDGaji,tingkat,Nip,Gaji,Periode,Keterangan,last_update,last_user,statusaktif) Values ('" & .TextMatrix(Row, 0) & "','" & .TextMatrix(Row, 1) & "','" & .TextMatrix(Row, 2) & "','" & .TextMatrix(Row, 4) & "','" & Format(.TextMatrix(Row, 5), "yyyy-MM-dd") & "','" & .TextMatrix(Row, 6) & "','" & Now & "','" & StrUser & "','1')"
            CN.Execute StrSQL
        Next
    End With
End If

End Sub
Private Sub Command1_Click()
Dim RsBiaya As New ADODB.Recordset
Dim RsGaji As New ADODB.Recordset
Dim Jam1, Jam2 As Date
Dim J As String, Tanggal As String
Dim TotalLembur As Double
Dim JamBruto, MenitBruto As Double
Dim JamNetto, MenitNetto As Double
Dim TotalBruto, TotalNetto As Double
Dim TotalUpah, TotalUM As Double
Dim TotalUpah1 As Double

Dim JamAbsen, MenitAbsen As String
Dim JmlLoop  As Integer
On Error GoTo Adaerror
With VsFlex
    .Rows = 1
If fg.Rows <= 2 Then
   MsgBox "Data Gaji Karyawan Masih Kosong", vbCritical
   Exit Sub
End If
Tanggal = Format(DTPicker1, "MM")
If Mid(FileTitle, 6, 2) <> Tanggal Then
    If MsgBox("File Data Gaji Tidak Sama Dengan Tanggal Pencarian, Apakah Anda Akan Melanjutkan Proses ini ?", vbQuestion + vbYesNo, "Konfirmasi hapus") = vbNo Then
           Exit Sub
    End If
End If
Command1.Enabled = False
'SimpanGaji
fg.Visible = False
.Cols = 19
If RsBiaya.State = adStateOpen Then RsBiaya.Close
Set RsBiaya = Nothing

StrSQL = "SELECT * From vLembur WHERE TANGGAL BETWEEN '" & Format(DTPicker1, "mm/dd/yyyy") & "' And '" & Format(DTPicker2, "mm/dd/yyyy") & "'"
If Combo3.Text = "Timesheet" Then
   StrSQL = StrSQL & " And Keterangan ='Timesheet'"
ElseIf Combo3.Text = "Lembur" Then
    StrSQL = StrSQL & " And Keterangan ='Lembur'"
End If
If Trim(Combo1.Text) <> "" Then
    StrSQL = StrSQL & " And Divisi = '" & Combo1 & "'"
End If
If Trim(Combo2.Text) <> "" Then
    StrSQL = StrSQL & " And NoProject = '" & Combo2 & "'"
End If

If Trim(cboFlex.Text) <> "" Then
    StrSQL = StrSQL & " And NIP = '" & cboFlex.Text & "'"
End If
StrSQL = StrSQL & " Order by Tanggal,NIP,idTimesheet"
RsBiaya.Open StrSQL, CN, adOpenStatic
Set .DataSource = RsBiaya
.ColWidth(0) = 300
.ColWidth(1) = 1200
'.ColWidth(4) = 3000
.ColWidth(8) = 800
.ColWidth(9) = 800
.ColDataType(8) = flexDTDate
.ColDataType(9) = flexDTDate
.ColFormat(8) = "HH:mm"
.ColFormat(9) = "HH:mm"
.ColFormat(12) = "HH:mm"
.ColFormat(11) = "HH:mm"
.ColFormat(14) = "HH:mm"

.ColFormat(13) = "#,###"
.ColFormat(17) = "#,###"
.ColFormat(18) = "#,###"
.ColWidth(13) = 0
.ColWidth(19) = 0
.ColWidth(20) = 0
.TextMatrix(0, 16) = "Total Kerja"

If Combo3 = "Timesheet" Then
    .TextMatrix(0, 11) = "Total Netto"
    .ColWidth(12) = 0
Else
    .TextMatrix(0, 11) = "Lembur Bruto"
    .ColWidth(12) = 1000
End If
.TextMatrix(0, 17) = "Upah TS"
.TextMatrix(0, 18) = "Uang Makan"

For Lrow = 1 To .Rows - 1
    .TextMatrix(Lrow, 0) = Lrow
    .TextMatrix(Lrow, 10) = 0
     
     .TextMatrix(Lrow, 18) = 0
        JmlLoop = 0
        Jam1 = CDate(.TextMatrix(Lrow, 8))
        Jam2 = CDate(.TextMatrix(Lrow, 9))
        DTPicker3.Value = Jam1
        Do Until JmlLoop = 50
            If Format(DTPicker3.Value, "hh") = Format(Jam2, "hh") Then Exit Do
            JmlLoop = JmlLoop + 1
            DTPicker3.Value = DateAdd("n", 60, DTPicker3)
'            If Format(DTPicker3, "HH:mm") = "12:00" Then DTPicker3.Value = DateAdd("n", 60, DTPicker3)
        Loop
    
     DTPicker3 = .TextMatrix(Lrow, 8)
     DTPicker4 = .TextMatrix(Lrow, 9)
     
     If DTPicker3.Minute > DTPicker4.Minute Then
         JamAbsen = JmlLoop - 1
     Else
        JamAbsen = JmlLoop
     End If

'    .TextMatrix(Lrow, 10) = Format(DTPicker4 - DTPicker3, "HH:mm")
     MenitAbsen = Format(DTPicker4 - DTPicker3, "HH:mm")
     JamAbsen = JamAbsen & ":" & Right(MenitAbsen, 2)
     JamAbsen = CDate(JamAbsen)
    .TextMatrix(Lrow, 10) = Format(JamAbsen, "HH:mm")
    .TextMatrix(Lrow, 13) = 0
    .TextMatrix(Lrow, 11) = .TextMatrix(Lrow, 10)
    .TextMatrix(Lrow, 12) = 0
    If Trim(.TextMatrix(Lrow, 17)) = "Lembur" Then
       If Format(.TextMatrix(Lrow, 8), "HH:mm") <> "00:00" And Format(.TextMatrix(Lrow, 9), "HH:mm") <> "00:00" Then
             If .TextMatrix(Lrow, 2) = "Kerja" Then
                DTPicker3 = "09:00"
                DTPicker4 = .TextMatrix(Lrow, 10)
                .TextMatrix(Lrow, 11) = Format(DTPicker4 - DTPicker3, "HH:mm")
             Else
                .TextMatrix(Lrow, 11) = .TextMatrix(Lrow, 10)
             End If
             
        Else
          .TextMatrix(Lrow, 11) = 0
        End If
    Else
        DTPicker3 = "01:00"
        DTPicker4 = .TextMatrix(Lrow, 10)
        .TextMatrix(Lrow, 11) = Format(DTPicker4 - DTPicker3, "HH:mm")
    End If
    If Format(.TextMatrix(Lrow, 8), "HH:mm") = "00:00" Then .Col = 8: .Row = Lrow: .CellBackColor = vbRed
    If Format(.TextMatrix(Lrow, 9), "HH:mm") = "00:00" Then .Col = 9: .Row = Lrow: .CellBackColor = vbRed
         JmlLoop = 0
        Jam1 = CDate(.TextMatrix(Lrow, 14))
        Jam2 = CDate(.TextMatrix(Lrow, 15))
        DTPicker3.Value = Jam1
        .TextMatrix(Lrow, 16) = 0
        Do Until JmlLoop = 50
           If Format(DTPicker3.Value, "hh:mm") = Format(Jam2, "hh:mm") Then Exit Do
              If Format(DTPicker3, "HH:mm") = "12:00" Then
                 DTPicker3.Value = DateAdd("n", 60, DTPicker3)
                 
              Else
                DTPicker3.Value = DateAdd("n", 30, DTPicker3)
                .TextMatrix(Lrow, 16) = (.TextMatrix(Lrow, 16) + 0.5)
              End If
               JmlLoop = JmlLoop + 1
        Loop
'            Jam1 = CDate(.TextMatrix(Lrow, 14))
'            Jam2 = CDate(.TextMatrix(Lrow, 15))
'            J = Format(Jam2 - Jam1, "HH:mm")
'            If  .TextMatrix(Lrow,17) ="Timesheet
    NilaiGaji = 0
    Tingkat = 0
    temptotaljam = 0
    tempbiayalembur1 = 0
    TotalUM = 0
   
    J = GetSet(.TextMatrix(Lrow, 3), .TextMatrix(Lrow, 3))
    If Tingkat <> 0 Then
         If Trim(.TextMatrix(Lrow, 17)) = "Timesheet" Then
            J = Replace(.TextMatrix(Lrow, 16), ":", ".")
            .TextMatrix(Lrow, 17) = Round(CCur(((NilaiGaji) / .TextMatrix(Lrow, 19))) * CDbl(J), 2)
        
        ElseIf Trim(.TextMatrix(Lrow, 17)) = "Lembur" Then
             J = GetLembur(.TextMatrix(Lrow, 3), Lrow, .TextMatrix(Lrow, 2))
        End If
            
    End If
        .Col = 17
        .Row = Lrow
        .CellAlignment = flexAlignRightCenter
 
    If Trim(.TextMatrix(Lrow, 17)) = "Timesheet" Or Trim(.TextMatrix(Lrow, 17)) = "Lembur" Then .TextMatrix(Lrow, 17) = 0
     If .TextMatrix(Lrow, 18) = "" Then .TextMatrix(Lrow, 18) = 0
   
    
Next
'       Lrow = 0
'       Do Until Lrow = .Rows - 1
'
'              If .TextMatrix(Lrow, 10) < "09:00" Then
'                    .RemoveItem (Lrow)
'                    Lrow = Lrow - 10
'              End If
'           Lrow = Lrow + 1
'       Loop
    
    For Lrow = 1 To .Rows - 1
        
        TotalUpah = TotalUpah + CDbl(.TextMatrix(Lrow, 17))
        TotalUpah1 = TotalUpah1 + CDbl(.TextMatrix(Lrow, 13))
        TotalUM = TotalUM + CDbl(.TextMatrix(Lrow, 18))
        If Len(.TextMatrix(Lrow, 11)) = 5 Then
            MenitBruto = CDbl(MenitBruto) + CDbl(Mid(.TextMatrix(Lrow, 11), 4, 2))
            JamBruto = CDbl(JamBruto) + CDbl(Left(.TextMatrix(Lrow, 11), 2))
        End If
'        If Len(.TextMatrix(Lrow, 16)) = 5 Then
'            MenitNetto = CDbl(MenitNetto) + CDbl(Mid(.TextMatrix(Lrow, 16), 4, 2))
            JamNetto = CDbl(JamNetto) + CDbl(.TextMatrix(Lrow, 16))
'        End If
    Next
    
    If .Rows > 1 Then
        
         Dim ConvertMenit, ConvertJam As Double
         Dim HasilConvert As String
         .Rows = .Rows + 1
         HasilConvert = Round(MenitBruto / 60, 2)
         ConvertMenit = HasilConvert * 60
         If MenitBruto >= 60 Then
               ConvertJam = MenitBruto \ 60
               ConvertMenit = ConvertMenit - (60 * ConvertJam) '75
               ConvertMenit = ConvertJam & ":" & CInt(ConvertMenit)
              HasilConvert = Format(ConvertMenit, "HH:mm")
         Else
            ConvertMenit = CInt(ConvertMenit)
             HasilConvert = "00:" & ConvertMenit
         End If
         JamBruto = CDbl(Left(HasilConvert, 2)) + JamBruto
'         .TextMatrix(.Rows - 1, 16) = JamBruto & ":" & Mid(HasilConvert, 4, 2)
         '--------------TOTAL NETTO--------
         HasilConvert = Round(MenitNetto / 60, 2)
         ConvertMenit = HasilConvert * 60
         If MenitNetto >= 60 Then
               ConvertJam = MenitNetto \ 60
               ConvertMenit = ConvertMenit - (60 * ConvertJam) '75
               ConvertMenit = ConvertJam & ":" & CInt(ConvertMenit)
              HasilConvert = Format(ConvertMenit, "HH:mm")
              JamNetto = JamNetto + ConvertJam
         Else
            ConvertMenit = CInt(ConvertMenit)
             HasilConvert = "00:" & ConvertMenit
         End If
         .TextMatrix(.Rows - 1, 16) = JamNetto
         .TextMatrix(.Rows - 1, 13) = TotalUpah1
         .TextMatrix(.Rows - 1, 17) = TotalUpah
         .TextMatrix(.Rows - 1, 18) = TotalUM
         .Col = 17
         .Row = .Rows - 1
         .CellAlignment = flexAlignRightCenter
        For lCol = 1 To .Cols - 1
            .Row = .Rows - 1
            .Col = lCol
            .CellBackColor = &H8000000F
            .Col = 18
            .CellBackColor = &H8000000F
        Next
   Else
        MsgBox "Data Tidak Ditemukan", vbInformation
   End If
End With
Command1.Enabled = True
Exit Sub
Adaerror:
MsgBox err.Description
Command1.Enabled = True
End Sub
Function GetSet(Oldnip As String, NIP As String) As String
Dim i As Integer
GetSet = False
'If Oldnip = NIP Then Exit Function
With fg
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
Function GetLembur(ByVal NIP As String, ByVal Lrow As Integer, Project As String)
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
Dim NilaiUpah, NilaiUpah1 As Currency, Menitlembur As Currency
Dim TotalTs1, TotalTs2 As String
Dim TotalTs3, TotalTs4 As String
Dim RsLembur As New ADODB.Recordset
With VsFlex
         TotalLembur = 0
         Istirahat = 0
         TotalLemburBruto = 0
         UM = 0
      
         If RsSetting.State = adStateOpen Then RsSetting.Close
         RsSetting.Open "SELECT * FROM tblSETTING WHERE TINGKAT = '" & Tingkat & "' AND HARI = '" & .TextMatrix(Lrow, 2) & "' AND StatusAktif =1 And Berlaku_SD >= '" & Format(Date, "mm/dd/yyyy") & "'", CN, adOpenStatic
         If Not RsSetting.EOF Then
             Jam1 = RsSetting!jamlembur1
             Jam2 = RsSetting!jamlembur2
             Jam3 = RsSetting!jamlembur3
             Jam4 = RsSetting!jamlembur4
             UM = RsSetting!upahmakan
             JamMakan = RsSetting!jammakan1
             If Format(.TextMatrix(Lrow, 8), "HH:mm") = "00:00" Then
                TotalLembur = 0
                TotalTs3 = 0
             Else
                Dim Jam As Integer
                Dim Menit As Integer
                       TotalJam = .TextMatrix(Lrow, 10)
                       TotalTs1 = .TextMatrix(Lrow, 16)
                  Select Case .TextMatrix(Lrow, 2)
                      Case "Kerja"
                            TotalLemburBruto = Left(TotalJam, 2) - 9 & Mid(TotalJam, 3, 6)
                            TotalLembur = CDbl(Left(TotalJam, 2) - 9)
                            
                            TotalTs2 = TotalTs1 'Left(TotalTs1, 2) - 9 & Mid(TotalTs1, 3, 6)
                            TotalTs3 = CDbl(Left(TotalTs1, 2))
                      Case "Libur"
                            TotalLemburBruto = TotalJam
                            TotalLembur = CDbl(Left(TotalJam, 2))
                            
                            TotalTs2 = TotalTs1
                            TotalTs3 = CDbl(Left(TotalTs1, 2))
                  End Select
                  
                  If Len(TotalLemburBruto) = 4 Then TotalLemburBruto = "0" & TotalLemburBruto
                
                  If Len(TotalTs2) = 4 Then TotalTs2 = "0" & TotalTs2
                   
                  If TotalLembur >= RsSetting!ist1 And TotalLembur <= RsSetting!ist2 Then
                      Istirahat = CDbl(RsSetting!jamist_1) * 60
                      If Len(TotalLembur) = 1 Then
                          TotalLembur = "0" & TotalLembur & Mid(TotalJam, 3, 6)
                      ElseIf Len(TotalLembur) = 2 Then
                          TotalLembur = TotalLembur & Mid(TotalJam, 3, 6)
                      End If
                      If Istirahat >= 60 Then
                            Jam = Istirahat \ 60
                            Menit = Istirahat - (60 * Jam)
                            Istirahat = Jam & ":" & Menit
                      Else
                          Istirahat = "00:" & Istirahat
                      End If
                      If TotalLembur >= "00:30" Then
                        TotalLembur = Format(CDate(TotalLembur) - CDate(Istirahat), "HH:mm")

                     End If
                   ElseIf TotalLembur >= RsSetting!ist3 And TotalLembur <= RsSetting!ist4 Then
                      Istirahat = CDbl(RsSetting!jamist_2) * 60
                      If Len(TotalLembur) = 1 Then
                          TotalLembur = "0" & TotalLembur & Mid(TotalJam, 3, 6)
                      ElseIf Len(TotalLembur) = 2 Then
                          TotalLembur = TotalLembur & Mid(TotalJam, 3, 6)
                      End If
                      If Istirahat >= 60 Then
                            Jam = Istirahat \ 60
                            Menit = Istirahat - (60 * Jam)
                            Istirahat = Jam & ":" & Menit
                            Istirahat = Format(Istirahat, "HH:mm")
                      Else
                          Istirahat = "00:" & Istirahat
                      End If
                       TotalLembur = Format(CDate(TotalLembur) - CDate(Istirahat), "HH:mm")
                    ElseIf TotalLembur >= RsSetting!ist5 And TotalLembur <= RsSetting!ist6 Then
                      Istirahat = CDbl(RsSetting!jamist_3) * 60
                      If Len(TotalLembur) = 1 Then
                          TotalLembur = "0" & TotalLembur & Mid(TotalJam, 3, 6)
                      ElseIf Len(TotalLembur) = 2 Then
                          TotalLembur = TotalLembur & Mid(TotalJam, 3, 6)
                      End If
                      If Istirahat >= 60 Then
                            Jam = Istirahat \ 60
                            Menit = Istirahat - (60 * Jam)
                            Istirahat = Jam & ":" & Menit
                            Istirahat = Format(Istirahat, "HH:mm")
                      Else
                          Istirahat = "00:" & Istirahat
                      End If
                
                      TotalLembur = Format(CDate(TotalLembur) - CDate(Istirahat), "HH:mm")
                    ElseIf TotalLembur >= RsSetting!ist7 And TotalLembur <= RsSetting!ist8 Then
                      Istirahat = CDbl(RsSetting!jamist_4) * 60
                      If Len(TotalLembur) = 1 Then
                          TotalLembur = "0" & TotalLembur & Mid(TotalJam, 3, 6)
                      ElseIf Len(TotalLembur) = 2 Then
                          TotalLembur = TotalLembur & Mid(TotalJam, 3, 6)
                      End If
                      If Istirahat > 60 Then
                            Jam = Istirahat \ 60
                            Menit = Istirahat - (60 * Jam)
                            Istirahat = Jam & ":" & Menit
                            Istirahat = Format(Istirahat, "HH:mm")
                      Else
                          Istirahat = "00:" & Istirahat
                      End If
          
                      TotalLembur = Format(CDate(TotalLembur) - CDate(Istirahat), "HH:mm")
                    
                    End If
                    
                  '----------------TOTAL TS----------------
                  If TotalTs3 >= RsSetting!ist1 And TotalTs3 <= RsSetting!ist2 Then
                      Istirahat = RsSetting!jamist_1
                      TotalTs3 = TotalTs3 - Istirahat
 
                   ElseIf TotalTs3 >= RsSetting!ist3 And TotalTs3 <= RsSetting!ist4 Then
                      Istirahat = RsSetting!jamist_2
                      TotalTs3 = TotalTs3 - Istirahat
 
                    ElseIf TotalTs3 >= RsSetting!ist5 And TotalTs3 <= RsSetting!ist6 Then
                      Istirahat = RsSetting!jamist_3
                      TotalTs3 = TotalTs3 - Istirahat
 
                    ElseIf TotalTs3 >= RsSetting!ist7 And TotalTs3 <= RsSetting!ist8 Then
                      Istirahat = RsSetting!jamist_4
                      TotalTs3 = TotalTs3 - Istirahat
                    
                    End If
            End If
         
         
         'pembulatan menit
         If TotalLembur <> 0 Then
            Dim Pembulatan As Integer
            Pembulatan = Right(TotalLembur, 2)
            If Pembulatan >= 0 And Pembulatan <= 7 Then
               TotalLembur = Left(TotalLembur, 3) & "00"
            ElseIf Pembulatan >= 8 And Pembulatan <= 17 Then
               TotalLembur = Left(TotalLembur, 3) & 15
            ElseIf Pembulatan >= 18 And Pembulatan <= 22 Then
               TotalLembur = Left(TotalLembur, 3) & 15
            ElseIf Pembulatan >= 23 And Pembulatan <= 37 Then
               TotalLembur = Left(TotalLembur, 3) & 30
            ElseIf Pembulatan >= 38 And Pembulatan <= 60 Then
               TotalLembur = Left(TotalLembur, 3) & 45
            End If
            
'            Pembulatan = Right(TotalTs3, 2)
'            If Pembulatan >= 0 And Pembulatan <= 7 Then
'               TotalTs3 = Left(TotalTs3, 3) & "00"
'            ElseIf Pembulatan >= 8 And Pembulatan <= 17 Then
'               TotalTs3 = Left(TotalTs3, 3) & 15
'            ElseIf Pembulatan >= 18 And Pembulatan <= 22 Then
'               TotalTs3 = Left(TotalTs3, 3) & 15
'            ElseIf Pembulatan >= 23 And Pembulatan <= 37 Then
'               TotalTs3 = Left(TotalTs3, 3) & 30
'            ElseIf Pembulatan >= 38 And Pembulatan <= 60 Then
'               TotalTs3 = Left(TotalTs3, 3) & 45
'            End If
            
         If Left(TotalLembur, 2) < JamMakan Then UM = 0
            If Format(.TextMatrix(Lrow, 8), "HH:mm") = "00:00" Then
               If Format(.TextMatrix(Lrow, 8), "HH:mm") >= "00:00" And Format(.TextMatrix(Lrow, 8), "HH:mm") <= "05:00" Then
                  If .TextMatrix(Lrow, 2) = "Kerja" Then
                     Transport = "Ya"
                  Else
                     Transport = "Ya"
                  End If
               Else
                  Transport = 0
               End If
            Else
               Transport = 0
            End If
            
         'upah
           
         Else
            UM = 0
            Transport = 0
            NilaiUpah = 0
            TotalLembur = 0
         End If
         JamLembur = Replace(TotalLembur, ":", ".")
            Upah1 = RsSetting!Upah1
            Upah2 = RsSetting!Upah2
            NilaiUpah = 0
            If JamLembur >= Jam1 And JamLembur <= Jam2 Then
                NilaiUpah = (JamLembur * NilaiGaji / 173) * Upah1
            ElseIf JamLembur >= Jam3 And JamLembur <= Jam4 Then
                NilaiUpah = (Jam2 * NilaiGaji / 173) * Upah1
                SplitLembur = JamLembur - Jam2
                NilaiUpah = CDbl(NilaiUpah) + CDbl(SplitLembur * NilaiGaji / 173) * Upah2
            End If
            
            JamLembur = Replace(TotalTs3, ":", ".")
            Upah1 = RsSetting!Upah1
            Upah2 = RsSetting!Upah2
            NilaiUpah1 = 0
            If JamLembur >= Jam1 And JamLembur <= Jam2 Then
                NilaiUpah1 = (JamLembur * NilaiGaji / 173) * Upah1
            ElseIf JamLembur >= Jam3 And JamLembur <= Jam4 Then
                NilaiUpah1 = (Jam2 * NilaiGaji / 173) * Upah1
                SplitLembur = JamLembur - Jam2
                NilaiUpah1 = CDbl(NilaiUpah1) + CDbl(SplitLembur * NilaiGaji / 173) * Upah2
            End If
        End If
         If TotalLembur > 0 Then
            VsFlex.TextMatrix(Lrow, 12) = TotalLembur
            VsFlex.TextMatrix(Lrow, 13) = NilaiUpah
            VsFlex.TextMatrix(Lrow, 16) = TotalTs3
             VsFlex.TextMatrix(Lrow, 17) = NilaiUpah1
            VsFlex.TextMatrix(Lrow, 18) = UM
              
        End If
End With
End Function

Private Sub Command2_Click()
Dim x As String
If VsFlex.Rows > 1 Then
   VsFlex.SaveGrid "C:\BiayaLembur.xls", flexFileExcel, True
'    Shell PathOffice & "C:\BiayaLembur.csv", vbNormalFocus
    x = ShellExecute(Me.hwnd, "open", "C:\BiayaLembur.xls", vbNullString, "C:\BiayaLembur.xls", 1)
End If
End Sub

Private Sub Command3_Click()
On Error GoTo Adaerror
If VsFlex.Rows > 1 Then VsFlex.PrintGrid "Biaya Lembur Timesheet - Periode " & DTPicker1.Value & " S/D " & DTPicker2.Value, 2, 2, 900, 500
Exit Sub
Adaerror:
MsgBox err.Description
End Sub
Sub Showgaji()
On Error GoTo Adaerror
    MsgBox "Silahkan Pilih File Data Gaji Karyawan", vbInformation
'    CN.Execute "delete From tblgaji"
    dlg.FileName = ""
    dlg.Filter = "File Gaji (*.Csv)|*.Csv"
    dlg.DialogTitle = "Data Gaji Karyawan "
    dlg.ShowOpen
    If Len(dlg.FileName) = 0 Then Exit Sub
    FileTitle = dlg.FileTitle
    MousePointer = MousePointerConstants.vbHourglass
    fg.LoadGrid dlg.FileName, flexFileCommaText
    MousePointer = MousePointerConstants.vbDefault
    
    Caption = APPNAME + " " + dlg.FileName

    sheet = 0
    Command5.Enabled = True
    With fg
        Dim Lrow As Long
        For Lrow = 1 To .Rows - 1
          .TextMatrix(Lrow, 0) = Lrow
        Next
    End With
    fg.Visible = False
Exit Sub
Adaerror:
MousePointer = MousePointerConstants.vbDefault
MsgBox err.Description & " Or File NO Match "
End Sub

Private Sub Command4_Click()
Showgaji

fg.Visible = True
End Sub

Private Sub Command5_Click()
sheet = sheet - 1
    
    MousePointer = MousePointerConstants.vbHourglass
    
    On Error Resume Next
    fg.LoadGrid dlg.FileName, flexFileCommaText, sheet
     With fg
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
    fg.LoadGrid dlg.FileName, flexFileCommaText, sheet
     With fg
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
If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path & "\Skins\" & skinsFileName
      Skin1.ApplySkin hwnd
End If
If Rscek.State = adStateOpen Then Rscek.Close
Rscek.Open "SELECT * from divisi where kd_bid >= 0 and kd_bid <= 20 order by kd_bid", CN, adOpenStatic
   Combo1.AddItem ""
Do Until Rscek.EOF
   Combo1.AddItem Rscek!NM_DIV
   Rscek.MoveNext
Loop
AddProject
getKaryawan
Combo3.ListIndex = 0
If fg.Rows = 1 Then Showgaji
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker1.Value = DateSerial(Year(Now), Month(Now), 1)
    DTPicker1.Value = DateAdd("M", 0, DTPicker1.Value)
    DTPicker2.Value = DateSerial(Year(Now), Month(Now), 1)
    DTPicker2.Value = DateAdd("M", 1, DTPicker2.Value)
    DTPicker2.Value = DateAdd("d", -1, DTPicker2.Value)
    DTPicker1.CustomFormat = "dd/MMM/yyyy"
    DTPicker2.CustomFormat = "dd/MMM/yyyy"
End Sub

Private Sub Form_Resize()
On Error Resume Next
 With VsFlex
    .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture3.Height
    End With
CmdClose.Width = Picture3.Width
With fg
    .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture3.Height
End With
End Sub
Private Function getKaryawan()
    Dim Cboid, cboid1 As String
     If Combo1 = "" Then
        StrSQL = "Select * From Karyawan Where Status <> '14' Order By Nama "
     Else
        StrSQL = " SELECT Karyawan.NIP,Karyawan.NIP, Karyawan.Nama, Divisi.NM_DIV AS DIVISI FROM Divisi INNER JOIN Karyawan ON Divisi.KD_DIV = Karyawan.Kd_Divisi Where NM_DIV = '" & Combo1 & "' And Karyawan.Status <> '14' Order By Karyawan.NIP "
     End If
     If Rscek.State = adStateOpen Then Rscek.Close
     Rscek.Open StrSQL, CN, adOpenStatic
     cboFlex.Text = ""
     cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("NIP") & vbTab & Rscek("Nama")
      cboid1 = cboid1 + Cboid
      Rscek.MoveNext
    Loop
    cboFlex.ColComboList(0) = cboid1
End Function
Private Sub cboFlex_LostFocus()
cboFlex.BackColor = vbWhite
If Trim(cboFlex.Text) = vbNullString Then cboFlex.Text = ""
End Sub

