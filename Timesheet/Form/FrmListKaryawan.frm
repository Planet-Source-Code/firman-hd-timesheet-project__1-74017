VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmListKaryawan 
   Caption         =   "Daftar Karyawan"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   14175
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   14175
      TabIndex        =   3
      Top             =   7950
      Width           =   14175
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
         Left            =   2520
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   6435
         TabIndex        =   5
         Top             =   0
         Width           =   6435
         Begin VB.CommandButton btnFirst 
            Height          =   375
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "First 250"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton btnPrev 
            Height          =   375
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Previous 250"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton btnLast 
            Height          =   375
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Last 250"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton btnNext 
            Height          =   375
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Next 250"
            Top             =   0
            Width           =   375
         End
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2280
            TabIndex        =   10
            Top             =   60
            Width           =   2655
         End
      End
      Begin VB.Label lblCurrentRecord 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Record: 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   60
         Width           =   1365
      End
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   14175
      TabIndex        =   2
      Top             =   7935
      Width           =   14175
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   14175
      TabIndex        =   1
      Top             =   7920
      Width           =   14175
   End
   Begin VB.PictureBox shpBar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   14175
      TabIndex        =   0
      Top             =   0
      Width           =   14175
      Begin VB.CommandButton Command2 
         Caption         =   "&Export To Sheet"
         Height          =   615
         Left            =   8280
         TabIndex        =   34
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4080
         TabIndex        =   29
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   255
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3120
         TabIndex        =   25
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy"
         Format          =   52756483
         CurrentDate     =   40032
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1680
         TabIndex        =   24
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy"
         Format          =   52756483
         CurrentDate     =   40032
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   10800
         ScaleHeight     =   1545
         ScaleWidth      =   1665
         TabIndex        =   22
         Top             =   25
         Width           =   1695
         Begin VB.Image Image1 
            Height          =   1335
            Left            =   120
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         Height          =   615
         Left            =   8280
         TabIndex        =   20
         Top             =   120
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "FrmListKaryawan.frx":0000
         Left            =   4440
         List            =   "FrmListKaryawan.frx":0013
         TabIndex        =   19
         Text            =   "NIP"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox CmbStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "FrmListKaryawan.frx":003B
         Left            =   1440
         List            =   "FrmListKaryawan.frx":004E
         TabIndex        =   16
         Text            =   "Karyawan Aktif"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox TxtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         Top             =   120
         Width           =   3135
      End
      Begin VB.ComboBox CboSearch 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "FrmListKaryawan.frx":0098
         Left            =   120
         List            =   "FrmListKaryawan.frx":00A8
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   120
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   7080
         TabIndex        =   30
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy"
         Format          =   52756483
         CurrentDate     =   40032
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   5640
         TabIndex        =   31
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy"
         Format          =   52756483
         CurrentDate     =   40032
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   33
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun Keluar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   32
         Top             =   675
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun Masuk Dari"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   675
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Photo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   10200
         TabIndex        =   23
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Like"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   21
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Order By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   735
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8040
      OleObjectBlob   =   "FrmListKaryawan.frx":00C6
      Top             =   2280
   End
   Begin VSFlex8Ctl.VSFlexGrid VsFlex 
      Height          =   3975
      Left            =   0
      TabIndex        =   12
      ToolTipText     =   "Klik Kanan Mouse / Double Klik..."
      Top             =   1800
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
      BackColorAlternate=   16707036
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
      FormatString    =   $"FrmListKaryawan.frx":02FA
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
   Begin VB.Image imgBtnDn 
      Height          =   480
      Left            =   7920
      Picture         =   "FrmListKaryawan.frx":03D9
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBtnUp 
      Height          =   480
      Left            =   7560
      Picture         =   "FrmListKaryawan.frx":0CA3
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "FrmListKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CURR_COL As Integer
Dim RsKaryawan As New Recordset
Dim RecordPage As New clsPaging
Dim StrSQLParser As New clsSQLSelectParser


'Procedure used to filter records
Public Sub FilterRecord(ByVal srcCondition As String)
    StrSQLParser.RestoreStatement
    If Left(srcCondition, 3) = "Kode" Then srcCondition = "Karyawan." & srcCondition
    StrSQLParser.wCondition = srcCondition
    ReloadRecords StrSQLParser.StrSQLStatement
End Sub
'hanee0507@yahoo.com
Public Sub Perintah(ByVal What As String)
Dim Lrow As Long
Dim x As String
Dim lCol As Long
    On Error GoTo err
    Select Case What
        Case "New"
'           With VsFlex
'                  FrmAddKaryawan.show vbModal
'                  RefreshRecords
'           End With
         Case "Search"
'           With frmSearchs
'                Set .srcForm = Me
'                Set .srcColumnHeaders = VsFlex
'                .srcNoOfCol = 4
'                .show vbModal
'            End With
         Case "Select"
            With VsFlex
                For Lrow = 1 To .Rows - 2
                    .TextMatrix(Lrow, 1) = "-1"
                Next
            End With
        Case "Delete"
          Call Hapus
        Case "Refresh"
            RefreshRecords
        Case "Print"
            VsFlex.SaveGrid "C:\DataKaryawan.xls", flexFileExcel, True
'            Shell PathOffice & "C:\DataKaryawan.CSV", vbNormalFocus
           x = ShellExecute(Me.hwnd, "open", "C:\DataKaryawan.xls", vbNullString, "C:\DataKaryawan.xls", 1)
 
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

Public Sub RefreshRecords()
    StrSQLParser.RestoreStatement
    ReloadRecords StrSQLParser.StrSQLStatement
End Sub

'Procedure for reloadingrecords
Public Sub ReloadRecords(ByVal srcStrSQL As String)
    '-In this case I used StrSQL because it is faster than Filter function of VB
    '-when hundling millions of records.
    On Error GoTo err
    If CN.State = adStateClosed Then CN.Open
    With RsKaryawan
        If .State = adStateOpen Then .Close
        .Open srcStrSQL, CN, adOpenStatic
    End With
    RecordPage.Refresh
    FillList 1
    Exit Sub
err:
        If err.Number = -2147217913 Then
            srcStrSQL = Replace(srcStrSQL, "'", "", , , vbTextCompare)
            Resume
        ElseIf err.Number = -2147217900 Then
            MsgBox "Invalid search operation.", vbExclamation
            StrSQLParser.RestoreStatement
            srcStrSQL = StrSQLParser.StrSQLStatement
            Resume
       
        End If
End Sub


Private Sub btnFirst_Click()
    If RecordPage.PAGE_CURRENT <> 1 Then FillList 1
End Sub

Private Sub btnLast_Click()
    If RecordPage.PAGE_CURRENT <> RecordPage.PAGE_TOTAL Then FillList RecordPage.PAGE_TOTAL
End Sub

Private Sub btnNext_Click()
    If RecordPage.PAGE_CURRENT <> RecordPage.PAGE_TOTAL Then FillList RecordPage.PAGE_NEXT
End Sub

Private Sub btnPrev_Click()
    If RecordPage.PAGE_CURRENT <> 1 Then FillList RecordPage.PAGE_PREVIOUS
End Sub
 
Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub Command1_Click()
On Error GoTo AdaError
Dim SQL As String
With StrSQLParser
        .Fields = "*"
        .Tables = "karyawan"
'        If TxtSearch <> "" Then
            If CboSearch = "" And TxtSearch = "" Then
                MsgBox "Kriteria Pencarian Belum Lengkap", vbCritical
                Exit Sub
            End If
            If TxtSearch <> "" Then
                Select Case CboSearch
                    Case "NIP"
                        SQL = "NIP = " & TxtSearch & " And"
                    Case "Nama"
                        SQL = " Nama LIKE '%" & TxtSearch & "%' And"
                    Case "Divisi"
                        SQL = SQL & " DIVISI LIKe '" & TxtSearch & "%' AND "
                End Select
            End If
                If Check1.Value = 1 Then SQL = SQL & " Datepart(Year,Tgl_masuk) Between '" & Format(DTPicker1, "yyyy") & "' And '" & Format(DTPicker2, "yyyy") & "' And"
                If Check2.Value = 1 Then SQL = SQL & " Datepart(Year,Tgl_keluar) Between '" & Format(DTPicker4, "yyyy") & "' And '" & Format(DTPicker3, "yyyy") & "' And"
    
        If CmbStatus = "Karyawan Aktif" Then
           .wCondition = SQL & " status <> 14 And Len(NIP) < 5"
        ElseIf CmbStatus = "Karyawan Berhenti" Then
           .wCondition = SQL & "status = 14 And Len(NIP) < 5"
        ElseIf CmbStatus = "Tetap Aktif" Then
           .wCondition = SQL & " (status = 1 Or status = 11) And Len(NIP) < 5"
        ElseIf CmbStatus = "Kontrak Aktif" Then
           .wCondition = SQL & " (status = 2 Or status = 12) And  Len(NIP) < 5"
        Else
           .wCondition = SQL & " status < 100 And Len(NIP) < 5"
    
        End If
        .SortOrder = "karyawan." & Combo1 & ""
        .SaveStatement

    End With
    If RsKaryawan.State = adStateOpen Then RsKaryawan.Close
    RsKaryawan.CursorLocation = adUseClient
    RsKaryawan.Open StrSQLParser.StrSQLStatement, CN, adOpenStatic, adLockReadOnly

    With RecordPage
        .Start RsKaryawan, 100
        FillList 1
    End With
 Exit Sub
AdaError:
    MsgBox err.Description
End Sub

Private Sub Command2_Click()
Dim x As String
With VsFlex
If .Rows > 1 Then
     
     
        .AddItem "", 1
        .AddItem "", 2
        
        .Redraw = flexRDNone
        .Redraw = flexRDBuffered
        .TextMatrix(1, 1) = "Rekap Karyawan "
    For lCol = 1 To .Cols - 1
        
       .TextMatrix(2, lCol) = .TextMatrix(0, lCol)
        .Row = 2
        .Col = lCol
        .CellBackColor = vbGreen '&HE0E0E0
    Next
        .SaveGrid "C:\RekapKarywn.xls", flexFileExcel, False
'        Shell PathOffice & "C:\RekapKarywn.xls", vbNormalFocus
x = ShellExecute(Me.hwnd, "open", "C:\RekapKarywn.xls", vbNullString, "C:\Rekapkarywn.xls", 1)
          .RemoveItem (1)
          .RemoveItem (1)
End If
End With
End Sub

Private Sub Form_Activate()
'RefreshRecords
End Sub

Private Sub Form_Load()
     If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
      Skin1.ApplySkin hwnd
    End If
    'Set the graphics for the controls
    With MDIMENU
  
        btnFirst.Picture = .i16x16.ListImages(3).Picture
        btnPrev.Picture = .i16x16.ListImages(4).Picture
        btnNext.Picture = .i16x16.ListImages(5).Picture
        btnLast.Picture = .i16x16.ListImages(6).Picture
        
        btnFirst.DisabledPicture = .i16x16g.ListImages(3).Picture
        btnPrev.DisabledPicture = .i16x16g.ListImages(4).Picture
        btnNext.DisabledPicture = .i16x16g.ListImages(5).Picture
        btnLast.DisabledPicture = .i16x16g.ListImages(6).Picture
    End With
    With StrSQLParser
        .Fields = "*"
        .Tables = "karyawan"
        If CmbStatus = "Karyawan Aktif" Then
           .wCondition = "status <> 14"
        ElseIf CmbStatus = "Karyawan Berhenti" Then
            .wCondition = "status = 14 "
        Else
            .wCondition = "status < 100"
        End If
        .SortOrder = "" & Combo1 & ""
        .SaveStatement
       
    End With
    If RsKaryawan.State = adStateOpen Then RsKaryawan.Close
    RsKaryawan.CursorLocation = adUseClient
    RsKaryawan.Open StrSQLParser.StrSQLStatement, CN, adOpenStatic, adLockOptimistic
    
    With RecordPage
        .Start RsKaryawan, 50
        FillList 1
    End With
    CboSearch.ListIndex = 0
    DTPicker1.Value = Date
    DTPicker2.Value = Date
End Sub

Sub Reminder()
Dim TahunCuti As Date
Dim XRow, Xcol As Integer
Dim Selisih As Integer
Dim XNama As String
On Error Resume Next
With VsFlex
    For XRow = 1 To .Rows - 1
        If Trim(.TextMatrix(XRow, 9)) = "kontrak" Or Trim(.TextMatrix(XRow, 9)) = "kontrak (P)" Then
            TahunCuti = .TextMatrix(XRow, 19)
            Selisih = DateDiff("m", Date, TahunCuti)
           XNama = .TextMatrix(XRow, 3)
            If Selisih = 0 Or Selisih = 1 Then
                For lCol = 1 To .Cols - 1
                    .Col = lCol
                    .Row = XRow
                    .CellBackColor = vbRed
                Next
            End If
        End If
    Next
End With
End Sub
Public Sub FillList(ByVal whichPage As Long)
    Dim i As Integer
    Dim Cboid, Cboid1 As String
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call IsiGrid(VsFlex, RsKaryawan, RecordPage.PageStart, RecordPage.PageEnd, 8, 2, False, True, , , , "PK")
    With VsFlex
'        .ColWidth(3) = 700
'        .ColWidth(4) = 3000
'        .ColWidth(8) = 2500
''        .ColWidth(6) = 2500
'        .ColFormat(6) = "dd/MMM/yyyy"
'        .ColFormat(7) = "dd/MMM/yyyy"
'     .Rows = .Rows - 1
'     .Cols = .Cols + 10
'     .TextMatrix(0, 10) = "SK"
'     .TextMatrix(0, 11) = "Pengalaman"
'     .TextMatrix(0, 12) = "Pendidikan"
'     .TextMatrix(0, 13) = "Pelatihan"
'     .TextMatrix(0, 14) = "SKA"
'
'     .TextMatrix(0, 15) = "Keluarga"
'     .TextMatrix(0, 16) = "Bahasa"
'     .TextMatrix(0, 17) = "Organisasi"
'     .TextMatrix(0, 18) = "Jamsostek"
'     .TextMatrix(0, 19) = "Hbs.Kontrak"
'     .TextMatrix(0, 20) = "Upload File"
'     .ColWidth(10) = 700
'     .ColWidth(11) = 1200
'     .ColWidth(12) = 1000
'     .ColWidth(13) = 900
'     .ColWidth(14) = 900
'     .ColWidth(15) = 750
'     .ColWidth(16) = 1000
'    For Lrow = 1 To .Rows - 1
''        .TextMatrix(Lrow, 13) = .TextMatrix(Lrow, 8)
''        .TextMatrix(Lrow, 10) = ""
'       If UCase(.TextMatrix(Lrow, 9)) <> "BERHENTI" Then .TextMatrix(Lrow, 7) = "-"
'
'        .Cell(flexcpPicture, Lrow, 10) = imgBtnUp
'        .Cell(flexcpPictureAlignment, Lrow, 10) = flexAlignCenterCenter
'        .Cell(flexcpPicture, Lrow, 11) = imgBtnUp
'        .Cell(flexcpPictureAlignment, Lrow, 11) = flexAlignCenterCenter
'        .Cell(flexcpPicture, Lrow, 12) = imgBtnUp
'        .Cell(flexcpPictureAlignment, Lrow, 12) = flexAlignCenterCenter
'        .Cell(flexcpPicture, Lrow, 13) = imgBtnUp
'        .Cell(flexcpPictureAlignment, Lrow, 13) = flexAlignCenterCenter
'        .Cell(flexcpPicture, Lrow, 14) = imgBtnUp
'        .Cell(flexcpPictureAlignment, Lrow, 14) = flexAlignCenterCenter
'         .Cell(flexcpPicture, Lrow, 15) = imgBtnUp
'        .Cell(flexcpPictureAlignment, Lrow, 15) = flexAlignCenterCenter
'         .Cell(flexcpPicture, Lrow, 16) = imgBtnUp
'        .Cell(flexcpPictureAlignment, Lrow, 16) = flexAlignCenterCenter
'         .Cell(flexcpPicture, Lrow, 17) = imgBtnUp
'        .Cell(flexcpPictureAlignment, Lrow, 17) = flexAlignCenterCenter
'         .Cell(flexcpPicture, Lrow, 18) = imgBtnUp
'        .Cell(flexcpPictureAlignment, Lrow, 18) = flexAlignCenterCenter
'        If Rscek.State = adStateOpen Then Rscek.Close
'        Rscek.Open "Select * From Kar_Status Where NIP = '" & .TextMatrix(Lrow, 3) & "' Order By Id DEsc", CN, adOpenStatic
'        If Not Rscek.EOF Then
'            .TextMatrix(Lrow, 19) = IIf(IsNull(Rscek!Tgl_Akhir_K), "", Rscek!Tgl_Akhir_K)
'        End If
'         .Cell(flexcpPicture, Lrow, 20) = imgBtnUp
'        .Cell(flexcpPictureAlignment, Lrow, 20) = flexAlignCenterCenter
'    Next
    End With
    Reminder
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    SetNavigation
    'Display the page information
    lblPageInfo.Caption = "Record " & RecordPage.PageInfo
    'Display the selected record
   
    VsFlex.Sort = flexSortCustom
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        
        shpBar.Width = ScaleWidth
        CmdClose.Width = shpBar.Width - 5000
        VsFlex.Width = Me.ScaleWidth - 100
        VsFlex.Height = (Me.ScaleHeight - Picture1.Height) - VsFlex.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmListKaryawan = Nothing
End Sub

Private Sub SetNavigation()
    With RecordPage
        If .PAGE_TOTAL = 1 Then
            btnFirst.Enabled = False
            btnPrev.Enabled = False
            btnNext.Enabled = False
            btnLast.Enabled = False
        ElseIf .PAGE_CURRENT = 1 Then
            btnFirst.Enabled = False
            btnPrev.Enabled = False
            btnNext.Enabled = True
            btnLast.Enabled = True
        ElseIf .PAGE_CURRENT = .PAGE_TOTAL And .PAGE_CURRENT > 1 Then
            btnFirst.Enabled = True
            btnPrev.Enabled = True
            btnNext.Enabled = False
            btnLast.Enabled = False
        Else
            btnFirst.Enabled = True
            btnPrev.Enabled = True
            btnNext.Enabled = True
            btnLast.Enabled = True
        End If
    End With
End Sub


Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
End Sub

Private Sub VSFlex_Click()
On Error Resume Next
Dim rsStream As New ADODB.Stream
Dim RsGambar As New ADODB.Recordset
If rsStream.State = adStateOpen Then rsStream.Close
If RsGambar.State = adStateOpen Then RsGambar.Close
RsGambar.Open "Select * From Karyawan Where NIP = '" & VsFlex.TextMatrix(VsFlex.Row, 3) & "'", CN, adOpenKeyset, adLockOptimistic
Image1.Picture = Nothing
    If Not RsGambar.EOF Then
        If Len(RsGambar!Photo) <> 0 Then
            rsStream.Type = adTypeBinary
            rsStream.Open
            rsStream.Write RsGambar.Fields("Photo").Value
            rsStream.SaveToFile App.Path & "\Foto.jpg", adSaveCreateOverWrite
            Image1.Picture = LoadPicture(App.Path & "\Foto.jpg")
            Kill App.Path & "\Foto.jpg"
            rsStream.Close
        End If
    
        
    End If
    If VsFlex.Col = 1 Then
       VsFlex.Editable = flexEDKbdMouse
    Else
       VsFlex.Editable = flexEDNone
             
    End If
'    Image1.Picture = LoadPicture(VsFlex.TextMatrix(VsFlex.Row, 13))
    If VsFlex.Text <> "" Then lblCurrentRecord.Caption = "Selected Record: " & VsFlex.Row
    
Exit Sub
err:
        lblCurrentRecord.Caption = "Selected Record: NONE"
AdaError:
'MsgBox err.Description
End Sub


Private Sub VsFlex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    MDIMENU.MnuRView.Visible = False
    MDIMENU.MnuRBatal.Visible = False
    MDIMENU.mnuRSearch.Visible = True
'    MDIMENU.mnuRPrint.Visible = False
    PopupMenu MDIMENU.mnuRecA
End If
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth
End Sub

Private Sub Hapus()
Dim Lrow As Long
Dim StrSQL As String
Dim Karyawan As String
Dim Tanya As String
Dim ErrConn As Long
 VsFlex.Rows = VsFlex.Rows + 1
 If CekCurek("hapus", VsFlex) = False Then Exit Sub
 If CN.State = adStateClosed Then CN.Open
            With VsFlex
            If MsgBox("Apakah Anda yakin ingin menghapus Data ?", vbQuestion + vbYesNo, "Konfirmasi hapus") = vbNo Then
               Exit Sub
            Else
             Do Until Lrow = .Rows - 1
                If .TextMatrix(Lrow, 1) = "-1" Then
                            
                                
                            StrSQL = "Delete From Karyawan Where NIP = '" & .TextMatrix(Lrow, 3) & "'"
                            PerintahExecute (StrSQL)
                           
                            .RemoveItem (Lrow)
                            Lrow = Lrow - 1
                    End If
                Lrow = Lrow + 1
            Loop
                     RefreshRecords
                End If
'                VsFlex.Rows = VsFlex.Rows - 1
            End With
Exit Sub

AdaError:
If ErrConn > 0 Then CN.RollbackTrans
MsgBox err.Description
End Sub





