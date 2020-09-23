VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmRekapLembur 
   Caption         =   "Rekap Lembur"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   9090
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1785
      ScaleWidth      =   9060
      TabIndex        =   2
      Top             =   0
      Width           =   9090
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Terverifikasi"
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
         Left            =   5040
         TabIndex        =   21
         Top             =   720
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Belum Terverifikasi"
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
         Left            =   5040
         TabIndex        =   20
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Export To Sheet"
         Height          =   375
         Left            =   9360
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   7800
         TabIndex        =   5
         Top             =   720
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
         Left            =   1320
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Print out"
         Height          =   375
         Left            =   9360
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   720
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
         Format          =   52953091
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3000
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
         Format          =   52953091
         CurrentDate     =   39931
      End
      Begin VSFlex8Ctl.VSFlexGrid cboFlex 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   960
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
         FormatString    =   $"FrmRekapLembur.frx":0000
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
         TabIndex        =   10
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
         FormatString    =   $"FrmRekapLembur.frx":0029
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
         TabIndex        =   15
         Top             =   600
         Width           =   1695
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   120
         Width           =   495
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
         TabIndex        =   12
         Top             =   960
         Width           =   1695
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
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9090
      TabIndex        =   0
      Top             =   6570
      Width           =   9090
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
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1320
      OleObjectBlob   =   "FrmRekapLembur.frx":0052
      Top             =   6000
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlex 
      Height          =   4695
      Left            =   0
      TabIndex        =   19
      Top             =   1800
      Width           =   8535
      _cx             =   15055
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
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4695
      Left            =   0
      TabIndex        =   16
      Top             =   1800
      Width           =   8535
      _cx             =   15055
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
         Format          =   52953090
         CurrentDate     =   39940
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   960
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52953090
         CurrentDate     =   39940
      End
   End
End
Attribute VB_Name = "FrmRekapLembur"
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
Dim I As Integer, FileTitle As String

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
AddProject
getKaryawan
End Sub
Private Sub AddProject()

    Dim Cboid     As String
    Dim Cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    Cboid1 = vbNullString
    StrSQL = "SELECT Project.Kode, Project.Nama, Divisi.NM_DIV AS Divisi, Project.Status FROM Project INNER JOIN Divisi ON Project.Kd_Divisi = Divisi.KD_DIV Where NM_DIV = '" & Combo1 & "' And Project.Status ='Terpakai' Order By Kode"

    Rscek.Open StrSQL, CN, adOpenStatic
    Cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("Kode") & vbTab & Rscek("Nama")
      Cboid1 = Cboid1 + Cboid
      Rscek.MoveNext
    Loop
    Combo2.ColComboList(0) = Cboid1
End Sub
Private Sub Combo1_GotFocus()
Combo1.BackColor = &HC0FFFF
End Sub

Private Sub Combo1_LostFocus()
Combo1.BackColor = vbWhite
End Sub

Private Sub Command1_Click()
Dim RsBiaya As New ADODB.Recordset
Dim RsGaji As New ADODB.Recordset
Dim Jam1 As DTPicker, Jam2 As DTPicker
Dim J As String, Tanggal As String
Dim TotalLembur As Double
Dim JamBruto, MenitBruto As Double
Dim JamNetto, MenitNetto As Double
Dim TotalBruto, TotalNetto As Double
Dim TotalUpah As Currency, TotalUM As Currency
With VSFlex
fg.Visible = False

If RsBiaya.State = adStateOpen Then RsBiaya.Close

StrSQL = "SELECT * From vLemburAll WHERE TANGGAL BETWEEN '" & Format(DTPicker1, "mm/dd/yyyy") & "' And '" & Format(DTPicker2, "mm/dd/yyyy") & "'"

If Combo1.Text <> "" Then
    StrSQL = StrSQL & " And Divisi = '" & Combo1 & "'"
End If
If Trim(Combo2.Text) <> "" Then
    StrSQL = StrSQL & " And NoProject Like '" & Combo2 & "%'"
End If

If Trim(cboFlex.Text) <> "" Then
    StrSQL = StrSQL & " And NIP = '" & cboFlex.Text & "'"
End If
If Option1.Value = True Then
     StrSQL = StrSQL & " And StatusPM = 1"
End If
If Option2.Value = True Then
     StrSQL = StrSQL & " And  StatusPM = 0"
End If
StrSQL = StrSQL & " Order by Tanggal,NIP"
RsBiaya.Open StrSQL, CN, adOpenStatic
Set VSFlex.DataSource = RsBiaya
'.Cols = .Cols + 2
.ColWidth(1) = 1200
.ColWidth(0) = 300
.ColWidth(3) = 700
.ColWidth(4) = 3000
.ColWidth(5) = 2000
.ColWidth(7) = 800
.ColWidth(8) = 800
.ColWidth(11) = 0
.ColWidth(15) = 0
.ColWidth(16) = 0
.ColDataType(7) = flexDTDate
.ColDataType(8) = flexDTDate
.ColFormat(7) = "HH:mm"
.ColFormat(8) = "HH:mm"
.ColFormat(11) = "HH:mm"
.ColFormat(1) = "dd/MM/yyyy"
'.ColFormat(17) = "#,###"
'.TextMatrix(0, 16) = "Upah Lembur"
'.TextMatrix(0, 17) = "Uang Makan"

For Lrow = 1 To .Rows - 1
    .TextMatrix(Lrow, 0) = Lrow
    .TextMatrix(Lrow, 9) = 0
    DTPicker3 = .TextMatrix(Lrow, 7)
    DTPicker4 = .TextMatrix(Lrow, 8)
     
    .TextMatrix(Lrow, 9) = Format(DTPicker4 - DTPicker3, "HH:mm")
    .TextMatrix(Lrow, 10) = .TextMatrix(Lrow, 9)
'    .TextMatrix(Lrow, 12) = 0
    If .TextMatrix(Lrow, 2) = "Kerja" Then
      If Format(.TextMatrix(Lrow, 7), "HH:mm") <> "00:00" And Format(.TextMatrix(Lrow, 8), "HH:mm") <> "00:00" Then
        DTPicker3 = "09:00"
        DTPicker4 = .TextMatrix(Lrow, 9)
        .TextMatrix(Lrow, 10) = Format(DTPicker4 - DTPicker3, "HH:mm")
      Else
        .TextMatrix(Lrow, 10) = 0
      End If
    Else
        
    End If
    If Format(.TextMatrix(Lrow, 7), "HH:mm") = "00:00" Then .Col = 7: .Row = Lrow: .CellBackColor = vbRed
    If Format(.TextMatrix(Lrow, 8), "HH:mm") = "00:00" Then .Col = 8: .Row = Lrow: .CellBackColor = vbRed


    TotalLembur = TotalLembur + .TextMatrix(Lrow, 14)
    If Len(.TextMatrix(Lrow, 10)) = 5 Then
        MenitBruto = CDbl(MenitBruto) + CDbl(Mid(.TextMatrix(Lrow, 10), 4, 2))
        JamBruto = CDbl(JamBruto) + CDbl(Left(.TextMatrix(Lrow, 10), 2))
    End If
    If Len(.TextMatrix(Lrow, 10)) = 5 Then
        MenitNetto = CDbl(MenitNetto) + CDbl(Mid(.TextMatrix(Lrow, 10), 4, 2))
        JamNetto = CDbl(JamNetto) + CDbl(Left(.TextMatrix(Lrow, 10), 2))
    End If
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
         .TextMatrix(.Rows - 1, 10) = JamBruto & ":" & Mid(HasilConvert, 4, 2)
         '--------------TOTAL NETTO--------
         HasilConvert = Round(MenitNetto / 60, 2)
         ConvertMenit = HasilConvert * 60
         If MenitNetto >= 60 Then
               ConvertJam = MenitNetto \ 60
               ConvertMenit = ConvertMenit - (60 * ConvertJam) '75
               ConvertMenit = ConvertJam & ":" & CInt(ConvertMenit)
              HasilConvert = Format(ConvertMenit, "HH:mm")
         Else
            ConvertMenit = CInt(ConvertMenit)
             HasilConvert = "00:" & ConvertMenit
         End If
         .TextMatrix(.Rows - 1, 10) = JamNetto & ":" & Mid(HasilConvert, 4, 2)
         .TextMatrix(.Rows - 1, 14) = TotalLembur
        For Lcol = 1 To .Cols - 1
         .Row = .Rows - 1
         .Col = Lcol
         .CellBackColor = &H8000000F
        Next
        VSFlex.FrozenCols = 4
   Else
        MsgBox "Data Tidak Ditemukan", vbInformation
   End If
End With
    
End Sub

Private Sub Command2_Click()
If VSFlex.Rows > 1 Then
   VSFlex.SaveGrid "C:\BiayaLembur.csv", flexFileCommaText, True
    Shell PathOffice & "C:\BiayaLembur.csv", vbNormalFocus
End If
End Sub

Private Sub Command3_Click()
On Error GoTo Adaerror
If VSFlex.Rows > 1 Then VSFlex.PrintGrid "Rekap Lembur Timesheet - Periode " & DTPicker1.Value & " S/D " & DTPicker2.Value, , 2, 900, 500
Exit Sub
Adaerror:
MsgBox err.Description
End Sub
Private Sub Form_Load()
 If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path & "\Skins\" & skinsFileName
      Skin1.ApplySkin hwnd
     End If
If Rscek.State = adStateOpen Then Rscek.Close
Rscek.Open "SELECT * from divisi where kd_bid >= 0 and kd_bid <= 20 order by kd_bid", CN, adOpenStatic
   Combo1.Text = ""
   Combo1.AddItem ""
Do Until Rscek.EOF
   Combo1.AddItem Rscek!NM_DIV
   Rscek.MoveNext
Loop
Select Case UCase(strGroup)
    Case "ADMIN", "PM"
        Combo1.Text = NamaDivisi
        Combo1.Enabled = False
    Case "USER"
        Combo1.Text = NamaDivisi
        Combo1.Enabled = False
        cboFlex.Text = StrNIPUser
        cboFlex.Enabled = False
End Select
AddProject
getKaryawan


DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker1.Value = DateSerial(Year(Now), Month(Now), 1)
DTPicker1.Value = DateAdd("M", 0, DTPicker1.Value)
DTPicker2.Value = DateSerial(Year(Now), Month(Now), 1)
DTPicker2.Value = DateAdd("M", 1, DTPicker2.Value) - 1
DTPicker1.CustomFormat = "dd/MMM/yyyy"
DTPicker2.CustomFormat = "dd/MMM/yyyy"
End Sub

Private Sub Form_Resize()
 With VSFlex
    .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture3.Height
    End With
CmdClose.Width = Picture3.Width
With fg
    .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture3.Height
End With
End Sub
Private Function getKaryawan()
    Dim Cboid, Cboid1 As String
     If Combo1 = "" Then
        StrSQL = "Select * From Karyawan Where Status <> '14'Order By NIP "
     Else
        StrSQL = " SELECT Karyawan.NIP,Karyawan.NIP, Karyawan.Nama, Divisi.NM_DIV AS DIVISI FROM Divisi INNER JOIN Karyawan ON Divisi.KD_DIV = Karyawan.Kd_Divisi Where NM_DIV = '" & Combo1 & "' And Karyawan.Status <> '14' Order By Karyawan.NIP "
     End If
     If Rscek.State = adStateOpen Then Rscek.Close
     Rscek.Open StrSQL, CN, adOpenStatic
    Cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("NIP") & vbTab & Rscek("Nama")
      Cboid1 = Cboid1 + Cboid
      Rscek.MoveNext
    Loop
    cboFlex.ColComboList(0) = Cboid1
End Function
Private Sub cboFlex_LostFocus()
cboFlex.BackColor = vbWhite
If Trim(cboFlex.Text) = vbNullString Then cboFlex.Text = ""
End Sub


Private Sub VsFlex_Click()
Dim J As String
With VSFlex
If .Col = 6 And .TextMatrix(.Row, .Col) <> "" And .TextMatrix(.Row, .Col) <> "Telat" Then
    J = FrmPM.Showdata(.TextMatrix(.Row, .Col), KodeDivisi)
    FrmPM.Show vbModal
End If
End With
End Sub
