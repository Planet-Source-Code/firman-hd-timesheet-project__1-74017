VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Begin VB.Form FrmGroupPermission 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group User"
   ClientHeight    =   8040
   ClientLeft      =   225
   ClientTop       =   1455
   ClientWidth     =   8310
   Icon            =   "FrmGroupPermission.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   8310
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "F9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      Picture         =   "FrmGroupPermission.frx":74F2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton CmdSimpan 
      BackColor       =   &H00FFFFFF&
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      Picture         =   "FrmGroupPermission.frx":C852
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   975
   End
   Begin VSFlex8Ctl.VSFlexGrid VsFlex 
      Height          =   7095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8175
      _cx             =   14420
      _cy             =   12515
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
      BackColor       =   -2147483643
      ForeColor       =   4194304
      BackColorFixed  =   15648682
      ForeColorFixed  =   0
      BackColorSel    =   16761024
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
      AllowUserResizing=   4
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
      FormatString    =   $"FrmGroupPermission.frx":11AB8
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
      ExplorerBar     =   7
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
      AllowUserFreezing=   2
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   6360
         OleObjectBlob   =   "FrmGroupPermission.frx":11B97
         Top             =   2280
      End
   End
   Begin VB.Image Bottom 
      Height          =   60
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   4050
   End
   Begin VB.Image ImgRight 
      Height          =   3870
      Left            =   9960
      Stretch         =   -1  'True
      Top             =   0
      Width           =   60
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   1
      Left            =   10920
      Top             =   1320
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   2
      Left            =   10920
      ToolTipText     =   "close"
      Top             =   960
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image3 
      Height          =   450
      Index           =   0
      Left            =   10800
      Top             =   120
      Width           =   150
   End
   Begin VB.Image Image1 
      Height          =   210
      Index           =   0
      Left            =   10440
      Stretch         =   -1  'True
      Top             =   120
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   165
      Index           =   0
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   570
      Width           =   60
   End
   Begin VB.Image Image5 
      Height          =   150
      Index           =   0
      Left            =   10920
      Stretch         =   -1  'True
      Top             =   570
      Width           =   60
   End
   Begin VB.Image Image6 
      Height          =   60
      Index           =   0
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   720
      Width           =   660
   End
   Begin VB.Image Image2 
      Height          =   450
      Index           =   0
      Left            =   10320
      Top             =   120
      Width           =   150
   End
   Begin VB.Image Image6 
      Height          =   60
      Index           =   1
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image Image1 
      Height          =   210
      Index           =   1
      Left            =   10560
      Stretch         =   -1  'True
      Top             =   -360
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Image2 
      Height          =   450
      Index           =   1
      Left            =   10320
      Top             =   -360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Image3 
      Height          =   450
      Index           =   1
      Left            =   11040
      Top             =   -360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Image4 
      Height          =   165
      Index           =   1
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   -240
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Image5 
      Height          =   165
      Index           =   1
      Left            =   11160
      Stretch         =   -1  'True
      Top             =   -240
      Visible         =   0   'False
      Width           =   60
   End
End
Attribute VB_Name = "FrmGroupPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myNamaMenu As String
Dim MyNamaGroup As String
Public MyIDGroup As String
Dim MHidden1 As Integer, MHidden2 As Integer, MKategori As Integer
Dim PHidden1 As Integer, PHidden2 As Integer, PKategori As Integer
Dim LHidden1 As Integer, LHidden2 As Integer, LKategori As Integer
Dim MaHidden1 As Integer, MaHidden2 As Integer, MaKategori As Integer
Dim LaHidden1 As Integer, LaHidden2 As Integer, LaKategori As Integer
Dim THidden1 As Integer, THidden2 As Integer, TKategori As Integer
Dim PaHidden1 As Integer, PaHidden2 As Integer, PaKategori As Integer
Dim KHidden1 As Integer, KHidden2 As Integer, KKategori As Integer
Sub Setgrid()
Dim x As Long
With VSFlex
    .Cols = 11
    .FixedRows = 1
    .Rows = 1
    .Rows = 2
    .FixedCols = 0
    .TextMatrix(0, 0) = ""
    .TextMatrix(0, 1) = "Nama Menu"
    .TextMatrix(0, 2) = "ID"
    .TextMatrix(0, 3) = "Parents"
    .TextMatrix(0, 4) = "Hak User"
    .TextMatrix(0, 5) = "Tambah"
    .TextMatrix(0, 6) = "Ubah"
    .TextMatrix(0, 7) = "Hapus"
    .TextMatrix(0, 8) = "Print"
    .TextMatrix(0, 9) = "Nama Menu"
    .TextMatrix(0, 10) = "No"
    .ColWidth(0) = 0
    .ColWidth(1) = 4500
    .ColWidth(2) = 0
    .ColWidth(3) = 0
    .ColWidth(4) = 1000
    .ColWidth(5) = 0  'insert
    .ColWidth(6) = 0 'Edit
    .ColWidth(7) = 0 'delete
    .ColWidth(8) = 0 'print
    .ColWidth(9) = 0
    .ColWidth(10) = 700
'    .RowHeight(0) = 500
    .ColDataType(0) = flexDTBoolean
'    .ColDataType(3) = flexDTUnknown
    .ColDataType(4) = flexDTBoolean
    .ColDataType(5) = flexDTBoolean
    .ColDataType(6) = flexDTBoolean
    .ColDataType(7) = flexDTBoolean
    .ColDataType(8) = flexDTBoolean
    .AllowUserResizing = flexResizeColumns
    For x = 0 To .Cols - 1
        .FixedAlignment(x) = flexAlignCenterCenter
    Next
'    File 1 - 13
'
'Maste 14 - 31
'
'Penjualan 32 - 54
'
'Pembelian 55 - 60
'
'inventory 61 - 67
'
'Akuntansi 68 - 77
'
'Utility 78 - 83
'
'Help 84 - 87

    MKategori = 1 'Kategori File
    MHidden1 = 2 'File hidden
    MHidden2 = 13
    '-----------------------------
    PKategori = 14 'Kategori Master
    PHidden1 = 15 'Master hidden
    PHidden2 = 31
     '-----------------------------
    LKategori = 32 'Kategori PENJUALAN
    LHidden1 = 33 'PENJUALAN hidden
    LHidden2 = 54
      '-----------------------------
    MaKategori = 55 'Kategori Pembelian
    MaHidden1 = 56 'Pembelian hidden
    MaHidden2 = 60
'      '-----------------------------
    LaKategori = 61 'Laporan
    LaHidden1 = 62 'Laporan hidden
    LaHidden2 = 67
'      '-----------------------------
    TKategori = 68 'Tools Kategori
    THidden1 = 67 'Tools hidden
    THidden2 = 77
'
    PaKategori = 78 'Panduan Kategori
    PaHidden1 = 79 'Panduan hidden
    PaHidden2 = 83
'        '-----------------------------
    KKategori = 84 'Keluar Kategori
    KHidden1 = 85 'Keluar hidden
    KHidden2 = 87
'
End With
End Sub
Private Sub Showdata()
Dim Rscek As New ADODB.Recordset
Dim RsShow As New ADODB.Recordset
Dim StrSQL As String
Dim i As Integer, J As Integer
Dim NamaGroup As String
Dim GP As String
Dim strStartUp As String
Dim strGroupPermission As String

On Error GoTo AdaError

If RsShow.State = adStateOpen Then RsShow.Close

StrSQL = "select IDMenu,Nama_Menu,Parents,Menu_Caption from TblGroup_Menu order by Urutan"

RsShow.Open StrSQL, CN, adOpenStatic

With VSFlex
'    .Rows = 1
'    .Rows = 2
'    .Row = 1
'    .Redraw = flexRDNone
Do Until RsShow.EOF
         .Row = .Rows - 1
        .TextMatrix(.Row, 3) = RsShow!Parents
        .TextMatrix(.Row, 9) = RsShow!nama_menu
        .TextMatrix(.Row, 10) = .Row
        If .TextMatrix(.Row, 3) = "0" Then
            .TextMatrix(.Row, 1) = RsShow!Menu_Caption
        ElseIf .TextMatrix(.Row, 3) = "1" Then
            .TextMatrix(.Row, 1) = Space(10) & "-- " & RsShow!Menu_Caption
        ElseIf .TextMatrix(.Row, 3) = "2" Then
            .TextMatrix(.Row, 1) = Space(20) & "--- " & RsShow!Menu_Caption
        End If
'        StrSQL = "Select * from TblGroup_User_Detail Where Nama_Menu ='" & .TextMatrix(.Row, 9) & "' And Nama_User = '" & NamaGroups & "' "
         StrSQL = " select tblgroup_user_detail.Group_Permission AS Group_Permission,tblgroup_menu.Nama_Menu AS Nama_Menu,tblgroup_user_detail.IDGroup AS IDGroup from tblgroup_user_detail INNER JOIN tblgroup_menu ON tblgroup_user_detail.nama_menu = tblgroup_menu.nama_menu Where tblgroup_user_detail.IDGroup ='" & MyIDGroup & "' And tblgroup_menu.Nama_Menu='" & Trim(Replace(.TextMatrix(.Row, 9), "--", "")) & "'"

        Rscek.Open StrSQL, CN, adOpenStatic
        If Not Rscek.EOF Then
            If Left(Rscek!Group_Permission, 1) <> 0 Then .TextMatrix(.Row, 4) = "-1"
        End If
        Rscek.Close
        .TextMatrix(.Row, 2) = RsShow!IDMenu
        RsShow.MoveNext
        .Rows = .Rows + 1
       
Loop
     .Rows = .Rows - 1
'     .Row = .Row - 1
'     .Redraw = flexRDBuffered


'    For I = 0 To .Cols - 1 'KategoriFile
'        .Row = MKategori
'        .Col = I
'        .CellBackColor = &H8000000F
'    Next
'    For I = MHidden1 To MHidden2 'File hidden
'        .RowHidden(I) = True
'    Next
'
'    For I = 0 To .Cols - 1 'KategoriMaster
'        .Row = PKategori
'        .Col = I
'        .CellBackColor = &H8000000F
'    Next
'     For I = PHidden1 To PHidden2 'Master hidden
'        .RowHidden(I) = True
'    Next
'
'    For I = 0 To .Cols - 1 'Kategori
'        .Row = LKategori
'        .Col = I
'        .CellBackColor = &H8000000F
'    Next
'    For I = LHidden1 To LHidden2 'PENJUALAN
'        .RowHidden(I) = True
'    Next
'
'    For I = 0 To .Cols - 1 'Kategori Pembelian
'        .Row = MaKategori
'        .Col = I
'        .CellBackColor = &H8000000F
'    Next
'    For I = MaHidden1 To MaHidden2 'Pembelian
'        .RowHidden(I) = True
'    Next
'
'    For I = 0 To .Cols - 1 'Kategori laporan
'        .Row = LaKategori
'        .Col = I
'        .CellBackColor = &H8000000F
'    Next
'    For I = LaHidden1 To LaHidden2 'laporan
'        .RowHidden(I) = True
'    Next
'
'    For I = 0 To .Cols - 1 'Kategori Tools
'        .Row = TKategori
'        .Col = I
'        .CellBackColor = &H8000000F
'    Next
'     For I = THidden1 To THidden2 'tools
'        .RowHidden(I) = True
'    Next
'
'    For I = 0 To .Cols - 1 'Kategori Panduan
'        .Row = PaKategori
'        .Col = I
'        .CellBackColor = &H8000000F
'    Next
'    For I = PaHidden1 To PaHidden2 'Panduan
'        .RowHidden(I) = True
'    Next
'
'    For I = 0 To .Cols - 1 'Kategori Keluar
'        .Row = KKategori
'        .Col = I
'        .CellBackColor = &H8000000F
'    Next
'
'    For I = KHidden1 To KHidden2 'keluar
'        .RowHidden(I) = True
'    Next

End With
If RsShow.State = adStateOpen Then RsShow.Close
Set RsShow.ActiveConnection = Nothing
Exit Sub
AdaError:
MsgBox err.Number & Chr(13) & err.Description, vbInformation, "Pesan Error : ShowData - FrmGroupPermission"
End Sub

Private Sub CmdExit_Click()
Unload Me
Set FrmGroupPermission = Nothing
End Sub

Private Sub CmdSimpan_Click()
Dim CmSimpan As New ADODB.Command
Dim RsSimpan As New ADODB.Recordset
Dim i As Long, J As Long
Dim NamaGroup As String
Dim GP As String
Dim ErrConn As Long
'NamaGroup = Newlogin.Encrypt(Right(NamaGroups, Len(NamaGroups) - 22), 1)
On Error GoTo AdaError

'If CN.State = adStateClosed Then CN.Open

'ErrConn = CN.BeginTrans

With VSFlex
For i = 1 To .Rows - 1
    GP = IIf(.TextMatrix(i, 4) = "-1", 1, 0) & IIf(.TextMatrix(i, 5) = "-1", 1, 0) & IIf(.TextMatrix(i, 6) = "-1", 1, 0) & IIf(.TextMatrix(i, 7) = "-1", 1, 0) & IIf(.TextMatrix(i, 8) = "-1", 1, 0)
 
    If RsSimpan.State = adStateOpen Then RsSimpan.Close
    RsSimpan.Open "Select Group_Permission,IDMenu from TblGroup_User_Detail Where IDGroup='" & MyIDGroup & "' and Nama_menu='" & .TextMatrix(i, 9) & "'", CN, adOpenStatic
    If Not RsSimpan.EOF Then
             
            CmSimpan.ActiveConnection = CN
            CmSimpan.CommandText = "Update TblGroup_User_Detail set Group_Permission='" & GP & "' Where IDGroup='" & MyIDGroup & "' and Nama_Menu='" & .TextMatrix(i, 9) & "'"
            CmSimpan.Execute
     Else
        
        CmSimpan.ActiveConnection = CN
        CmSimpan.CommandText = "Insert into TblGroup_User_Detail (IDGroup,IDMenu,Group_Permission,Nama_Menu) VALUES ('" & MyIDGroup & "','" & .TextMatrix(i, 2) & "','" & GP & "','" & .TextMatrix(i, 9) & "')"
        CmSimpan.Execute
    End If
Next

'CN.CommitTrans
End With

If RsSimpan.State = adStateOpen Then RsSimpan.Close
Set RsSimpan.ActiveConnection = Nothing

'If CN.State = adStateOpen Then CN.Close

MsgBox "Data berhasil disimpan !", vbInformation
Exit Sub

AdaError:
If ErrConn > 0 Then CN.RollbackTrans
MsgBox err.Number & Chr(13) & err.Description, vbInformation, "Pesan Error : SimpanData - FrmGroupPermission"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
      If CmdSimpan.Enabled = True Then CmdSimpan_Click
ElseIf KeyCode = vbKeyF9 Then
    Unload Me
    Set FrmGroupUser = Nothing
End If
End Sub

Private Sub Form_Load()
VSFlex.Editable = True
Me.Caption = MyNamaGroup
Setgrid
Showdata
If Len(skinsFileName) <> 0 Then
     Skin1.LoadSkin App.Path & "\Skins\" & skinsFileName
     Skin1.ApplySkin hwnd
End If
End Sub
Public Property Get NamaGroups() As String
    NamaGroups = MyNamaGroup
End Property

Public Property Let NamaGroups(ByVal vNewValue As String)
    MyNamaGroup = vNewValue
End Property

Public Property Let NamaMenu(ByVal vNewValue As String)
    myNamaMenu = vNewValue
End Property

Private Sub VSFlex_Click()
'Dim I As Long
'Dim j As Long
'
With VSFlex
If .Col = 4 Then
   .Editable = flexEDKbdMouse
Else
   .Editable = flexEDNone
End If
'    'Cek Parents
'    If .Col = 0 Then
'       If .TextMatrix(.Row, 0) = "-1" And UCase(.TextMatrix(.Row, 1)) = "File" Then
'            For I = MHidden1 To MHidden2 'File
'                .RowHidden(I) = False
'            Next
'            For I = MHidden1 To MHidden2 'File
'                .TextMatrix(I, 0) = "-1"
'            Next
'
'       ElseIf .TextMatrix(.Row, 0) = "0" And UCase(.TextMatrix(.Row, 1)) = "File" Then
'            For I = MHidden1 To MHidden2 'File
'               .RowHidden(I) = True
'            Next
'       End If
'
'       If .TextMatrix(.Row, 0) = "-1" And UCase(.TextMatrix(.Row, 1)) = "Master" Then
'           For I = PHidden1 To PHidden2 'Master
'            .RowHidden(I) = False
'           Next
'            For I = PHidden1 To PHidden2 'Master
'                .TextMatrix(I, 0) = "-1"
'            Next
'       ElseIf .TextMatrix(.Row, 0) = "0" And UCase(.TextMatrix(.Row, 1)) = "Master" Then
'            For I = PHidden1 To PHidden2 'Master
'            .RowHidden(I) = True
'            Next
'       End If
'
'       If .TextMatrix(.Row, 0) = "-1" And UCase(.TextMatrix(.Row, 1)) = "PENJUALAN" Then
'            For I = LHidden1 To LHidden2 'PENJUALAN
'                .RowHidden(I) = False
'            Next
'            For I = LHidden1 To LHidden2 'PENJUALAN
'                .TextMatrix(I, 0) = "-1"
'            Next
'       ElseIf .TextMatrix(.Row, 0) = "0" And UCase(.TextMatrix(.Row, 1)) = "PENJUALAN" Then
'            For I = LHidden1 To LHidden2 'PENJUALAN
'                .RowHidden(I) = True
'            Next
'       End If
'
'       If .TextMatrix(.Row, 0) = "-1" And UCase(.TextMatrix(.Row, 1)) = "Pembelian" Then
'            For I = MaHidden1 To MaHidden2 'Pembelian
'                .RowHidden(I) = False
'            Next
'            For I = MaHidden1 To MaHidden2 'Pembelian
'                .TextMatrix(I, 0) = "-1"
'            Next
'       ElseIf .TextMatrix(.Row, 0) = "0" And UCase(.TextMatrix(.Row, 1)) = "Pembelian" Then
'            For I = MaHidden1 To MaHidden2 'Pembelian
'                .RowHidden(I) = True
'            Next
'       End If
'
'       If .TextMatrix(.Row, 0) = "-1" And UCase(.TextMatrix(.Row, 1)) = "LAPORAN" Then
'            For I = LaHidden1 To LaHidden2 'laporan
'                .RowHidden(I) = False
'            Next
'            For I = LaHidden1 To LaHidden2 'laporan
'                .TextMatrix(I, 0) = "-1"
'            Next
'       ElseIf .TextMatrix(.Row, 0) = "0" And UCase(.TextMatrix(.Row, 1)) = "LAPORAN" Then
'            For I = LaHidden1 To LaHidden2 'laporan
'                .RowHidden(I) = True
'            Next
'       End If
'
'       If .TextMatrix(.Row, 0) = "-1" And UCase(.TextMatrix(.Row, 1)) = "TOOLS" Then
'            For I = THidden1 To THidden2 'TOOLS
'                .RowHidden(I) = False
'            Next
'            For I = THidden1 To THidden2 'TOOLS
'                .TextMatrix(I, 0) = "-1"
'            Next
'       ElseIf .TextMatrix(.Row, 0) = "0" And UCase(.TextMatrix(.Row, 1)) = "TOOLS" Then
'            For I = THidden1 To THidden2 'TOOLS
'                .RowHidden(I) = True
'            Next
'       End If
'
'       If .TextMatrix(.Row, 0) = "-1" And UCase(.TextMatrix(.Row, 1)) = "PANDUAN" Then
'            For I = PaHidden1 To PaHidden2 'PANDUAN
'                .RowHidden(I) = False
'            Next
'            For I = PaHidden1 To PaHidden2 'PANDUAN
'                .TextMatrix(I, 0) = "-1"
'            Next
'       ElseIf .TextMatrix(.Row, 0) = "0" And UCase(.TextMatrix(.Row, 1)) = "PANDUAN" Then
'            For I = PaHidden1 To PaHidden2 'PANDUAN
'                .RowHidden(I) = True
'            Next
'       End If
'
'        If .TextMatrix(.Row, 0) = "-1" And UCase(.TextMatrix(.Row, 1)) = "KELUAR" Then
'            For I = KHidden1 To KHidden2 'KELUAR
'                .RowHidden(I) = False
'            Next
'            For I = KHidden1 To KHidden2 'keluar
'                .TextMatrix(I, 0) = "-1"
'            Next
'       ElseIf .TextMatrix(.Row, 0) = "0" And UCase(.TextMatrix(.Row, 1)) = "KELUAR" Then
'           For I = KHidden1 To KHidden2 'KELUAR
'                .RowHidden(I) = True
'           Next
'       End If
'       .Col = 4
'
'    ElseIf .Col = 4 Then
'        'Cek Header
'        If Mid(.TextMatrix(.Row, 1), 11, 2) <> "--" And Mid(.TextMatrix(.Row, 1), 21, 2) <> "--" Then
'            If .TextMatrix(.Row, 4) = "-1" Then 'untuk ceklist menu all hak akses
'                For I = 5 To 8
'                    .TextMatrix(.Row, I) = "-1"
'                Next
'            Else
'                For I = 5 To 8
'                    .TextMatrix(.Row, I) = ""
'                Next
'            End If
'
'        ElseIf Mid(.TextMatrix(.Row, 1), 11, 2) = "--" Or Mid(.TextMatrix(.Row, 1), 21, 2) = "--" Then
'            If .Row >= MHidden1 And .CellChecked = flexChecked And .Row <= MHidden2 And .CellChecked = flexChecked Then
'                .TextMatrix(MHidden1 + 1, 4) = "-1" 'File
'            End If
'            If .Row >= PHidden1 And .CellChecked = flexChecked And .Row <= PHidden2 And .CellChecked = flexChecked Then
'                .TextMatrix(PHidden1 - 1, 4) = "-1" 'Master
'            End If
'
'            If .Row >= LHidden1 And .CellChecked = flexChecked And .Row <= LHidden2 And .CellChecked = flexChecked Then
'                .TextMatrix(LHidden1 - 1, 4) = "-1" 'PENJUALAN
'            End If
'
'            If .Row >= MaHidden1 And .CellChecked = flexChecked And .Row <= MaHidden2 And .CellChecked = flexChecked Then
'                .TextMatrix(MaHidden1 - 1, 4) = "-1" 'Pembelian
'            End If
'
'            If .Row >= LaHidden1 And .CellChecked = flexChecked And .Row <= LaHidden2 And .CellChecked = flexChecked Then
'                .TextMatrix(LaHidden1 - 1, 4) = "-1"
'            End If
'            If .Row >= THidden1 And .CellChecked = flexChecked And .Row <= THidden2 And .CellChecked = flexChecked Then
'                .TextMatrix(THidden1 - 1, 4) = "-1"
'            End If
'
'            If .Row >= PaHidden1 And .CellChecked = flexChecked And .Row <= PaHidden2 And .CellChecked = flexChecked Then
'                .TextMatrix(PaHidden1 - 1, 4) = "-1"
'            End If
'            If .Row >= KHidden1 And .CellChecked = flexChecked And .Row <= KHidden2 And .CellChecked = flexChecked Then
'                .TextMatrix(KHidden1 - 1, 4) = "-1"
'            End If
'           Exit Sub
'        End If
'         .Refresh
'     End If
End With
End Sub

Private Sub VsFlex_GotFocus()
VSFlex_Click
End Sub

Private Sub VsFlex_KeyPress(KeyAscii As Integer)
VSFlex_Click
End Sub
