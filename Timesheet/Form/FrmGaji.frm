VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Begin VB.Form FrmGaji 
   Caption         =   "Setting Gaji"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   7305
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   7305
      TabIndex        =   4
      Top             =   4935
      Width           =   7305
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
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
         Left            =   1920
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   4755
         TabIndex        =   6
         Top             =   0
         Width           =   4755
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "First 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3795
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Previous 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Last 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   4110
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Next 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1440
            TabIndex        =   11
            Top             =   60
            Width           =   1935
         End
      End
      Begin VB.Label lblCurrentRecord 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Record: 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   12
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
      ScaleWidth      =   7305
      TabIndex        =   3
      Top             =   4920
      Width           =   7305
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   7305
      TabIndex        =   2
      Top             =   4905
      Width           =   7305
   End
   Begin VB.PictureBox shpBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   6735
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Setting Gaji"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   4815
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8040
      OleObjectBlob   =   "FrmGaji.frx":0000
      Top             =   2280
   End
   Begin VSFlex8Ctl.VSFlexGrid VsFlex 
      Height          =   3975
      Left            =   0
      TabIndex        =   13
      Top             =   240
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
      FormatString    =   $"FrmGaji.frx":0234
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
Attribute VB_Name = "FrmGaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CURR_COL As Integer

Dim RsGaji As New Recordset
Dim RecordPage As New clsPaging
Dim StrSQLParser As New clsSQLSelectParser


'Procedure used to filter records
Public Sub FilterRecord(ByVal srcCondition As String)
    StrSQLParser.RestoreStatement
    If Left(srcCondition, 3) = "NIP" Then srcCondition = "tblgaji." & srcCondition
    StrSQLParser.wCondition = srcCondition
    ReloadRecords StrSQLParser.StrSQLStatement
End Sub

Public Sub Perintah(ByVal What As String)
Dim Lrow As Long
Dim lCol As Long
    On Error GoTo err
    Select Case What
        Case "New"
           With VSFlex
                  .Row = .Rows - 1
                  .Col = 3
                  .EditCell
           End With
         Case "Search"
           With frmSearchs
                Set .srcForm = Me
                Set .srcColumnHeaders = VSFlex
                .srcNoOfCol = 5
                .show vbModal
            End With
         Case "Select"
            With VSFlex
                For Lrow = 1 To .Rows - 2
                    .TextMatrix(Lrow, 1) = "-1"
                Next
            End With
        Case "Delete"
          Call Hapus
        Case "Refresh"
            RefreshRecords
                
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
    With RsGaji
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

Private Sub Form_Activate()
RefreshRecords
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
        .Fields = "tblgaji.Idgaji,tblgaji.Tingkat, tblgaji.NIP, karyawan.Nama, tblgaji.Gaji, tblgaji.Keterangan"
        .Tables = "karyawan INNER JOIN tblgaji ON karyawan.NIP = tblgaji.NIP"
        .wCondition = "statusaktif = 1"
        .SortOrder = "tblgaji.tingkat,Karyawan.Nama ASC"
        .SaveStatement
    End With
    If RsGaji.State = adStateOpen Then RsGaji.Close
    RsGaji.CursorLocation = adUseClient
    RsGaji.Open StrSQLParser.StrSQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPage
        .Start RsGaji, 50
        FillList 1
    End With

End Sub

Public Sub FillList(ByVal whichPage As Long)
    Dim i As Integer
    Dim Cboid, Cboid1 As String
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call IsiGrid(VSFlex, RsGaji, RecordPage.PageStart, RecordPage.PageEnd, 8, 2, False, True, , , , "PK")
    With VSFlex
        .ColWidth(3) = 800
        .ColWidth(4) = 800
        .ColWidth(5) = 2000
        .ColWidth(7) = 4500
        .ColFormat(6) = "#,###"
        For i = 3 To 15
           Cboid = "|" & i
           Cboid1 = Cboid1 + Cboid

         Next
       .ColComboList(3) = Cboid1
    End With
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    SetNavigation
    'Display the page information
    lblPageInfo.Caption = "Record " & RecordPage.PageInfo
    'Display the selected record
    VSFlex_Click
    VSFlex.Sort = flexSortCustom
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        
        shpBar.Width = ScaleWidth
        CmdClose.Width = shpBar.Width - 5000
        VSFlex.Width = Me.ScaleWidth - 100
        VSFlex.Height = (Me.ScaleHeight - Picture1.Height) - VSFlex.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
    Set FrmGaji = Nothing
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


Private Sub VsFlex_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlex
  Select Case Col
        Case 3
          .EditMaxLength = 3
          
        Case 6
            .EditMaxLength = 254
  End Select
End With
End Sub


Private Sub VSFlex_Click()
    Dim x As Long
    On Error GoTo err
    With VSFlex
    If .Col = 1 And .TextMatrix(.Row, 3) = "" Then
        .TextMatrix(.Row, 1) = ""
    ElseIf .Col > 3 And .TextMatrix(.Row, 3) = "" Then
        .Col = 3
        .EditCell
    End If
End With
    If VSFlex.Text <> "" Then lblCurrentRecord.Caption = "Selected Record: " & VSFlex.Row
    
Exit Sub
err:
        lblCurrentRecord.Caption = "Selected Record: NONE"
End Sub
Private Sub VsFlex_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96

End Sub

Private Sub VsFlex_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96
If KeyAscii = 8 Or KeyAscii = 27 Then Exit Sub
If KeyAscii = 13 Then
    If Col >= 3 And VSFlex.TextMatrix(Row, 2) <> "" Then
          Call SimpanData(Row)
    End If
    Exit Sub
End If
If VSFlex.Col = 4 Or VSFlex.Col = 6 Then
   If KeyAscii < 46 Or KeyAscii > 57 Then KeyAscii = 0
   If KeyAscii < 46 Or KeyAscii > 57 Then KeyAscii = 0
End If
End Sub

Private Sub VSFlex_KeyUp(KeyCode As Integer, Shift As Integer)
With VSFlex

    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then
      VSFlex_Click
      Exit Sub
    ElseIf KeyCode = 27 Then
       VSFlex.Refresh
       Exit Sub
    End If
    If KeyCode = 13 Then
      Select Case .Col
        Case 3, 6
           If .Text = "" Then
              .EditCell
              Exit Sub
           Else
                If .TextMatrix(.Row, 2) = "" Then
                    .Col = .Col + 1
                    .EditCell
                    .Refresh
                    Exit Sub
               End If
           End If
         Case 4, 5
         If getNama(.Text, .Col) = True Then
             
                  If cekGaji(.TextMatrix(.Row, 4)) = True And .TextMatrix(.Row, 2) = "" Then
                      MsgBox "Data Sudah Ada", vbCritical
                      .EditCell
                      .Refresh
                      Exit Sub
                   Else
                    If .TextMatrix(.Row, 2) = "" Then
                      .Col = .Col + 1
                      .EditCell
                      .Refresh
                      Exit Sub
                    Else
                        .Col = .Col + 1
                        .Refresh
                        Exit Sub
                    End If
                 End If
           Else
           .Col = .Col
'           .EditCell
           Exit Sub
        End If
             
         Case 7
             If .TextMatrix(.Row, 2) = "" Then SimpanData (.Row)
      End Select
End If

End With
End Sub
Private Function getNama(ByVal Search As String, Col As Integer) As Boolean
getNama = False
With frmSearch

    If Len(Search) < 1 Then MsgBox "Masukan Minimal 1 Digit", vbInformation: Exit Function
        
    Select Case Col
        Case 4
            .AsalStrSQL = "Select NIP,NIP,Nama From Karyawan Where NIP Like '%" & Search & "%'"
        Case 5
            .AsalStrSQL = "Select NIP,NIP,Nama From Karyawan Where Nama Like '%" & Trim(Search) & "%'"
    End Select
    .Caption = "Data Karyawan"
    .JmlRec = 2
    .IsiList
    
    If .VsSearch.Rows > 1 Then
        .show 1
        
        If .VsSearch.Row > 0 Then
             VSFlex.TextMatrix(VSFlex.Row, 4) = .VsSearch.TextMatrix(.VsSearch.Row, 2)
             VSFlex.TextMatrix(VSFlex.Row, 5) = .VsSearch.TextMatrix(.VsSearch.Row, 3)
               getNama = True
        Else
            VSFlex.Text = ""
            getNama = False
           Unload frmSearch
        End If
    Else
       VSFlex.Text = ""
       Unload frmSearch
End If
End With
    
Set frmSearch = Nothing
End Function
Private Sub VsFlex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    MDIMENU.MnuRView.Visible = False
    MDIMENU.MnuRBatal.Visible = False
    MDIMENU.mnuRSearch.Visible = True
    MDIMENU.mnuRPrint.Visible = False
    PopupMenu MDIMENU.mnuRecA
End If
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth
End Sub

Private Sub Hapus()
Dim Lrow As Long
Dim StrSQL As String
Dim Gaji As String
Dim Tanya As String
Dim ErrConn As Long
 If CekCurek("hapus", VSFlex) = False Then Exit Sub
 If CN.State = adStateClosed Then CN.Open
            With VSFlex
            If MsgBox("Apakah Anda yakin ingin menghapus Data ?", vbQuestion + vbYesNo, "Konfirmasi hapus") = vbNo Then
               Exit Sub
            Else
             Do Until Lrow = .Rows - 1
                If .TextMatrix(Lrow, 1) = "-1" Then
                            StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Hapus Data Gaji, " & .TextMatrix(Lrow, 4) & "','Data Gaji')"
                            PerintahExecute (StrSQL)
                                
                            StrSQL = "Update TblGaji set statusaktif = 0,last_update ='" & Now & "',last_user = '" & StrUser & "' Where IDGaji = '" & .TextMatrix(Lrow, 2) & "'"
                            PerintahExecute (StrSQL)
                            .RemoveItem (Lrow)
                            Lrow = Lrow - 1
                    End If
                Lrow = Lrow + 1
            Loop
                     RefreshRecords
                End If
            End With
Exit Sub

AdaError:
If ErrConn > 0 Then CN.RollbackTrans
MsgBox err.Description
End Sub


'Prosedur Simpan Data
Private Sub SimpanData(ByVal Row As Long)
Dim StrSQL As String
Dim Gaji As String

On Error GoTo AdaError

With VSFlex

If .TextMatrix(Row, 2) = "" Then
     '-----Ambil ID -------------------
        GetNomorID ("TblGaji")
        StrKodeID = NewID
    '---------------------------------
    
    StrSQL = "Insert Into TblGaji(IDGaji,tingkat,Nip,Gaji,keterangan,periode,last_update,last_user,statusaktif) Values ('" & StrKodeID & "','" & .TextMatrix(Row, 3) & "','" & .TextMatrix(Row, 4) & "','" & .TextMatrix(Row, 6) & "','" & .TextMatrix(Row, 7) & "','" & Now & "','" & Now & "','" & StrUser & "','1')"
    PerintahExecute (StrSQL)
    
    StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Tambah Data Gaji, " & .TextMatrix(Row, 4) & "','Data Gaji')"
    PerintahExecute (StrSQL)
    TotalRecord = TotalRecord + 1
    lblPageInfo.Caption = "Record " & ARecord & " - " & TotalRecord & " of " & TotalRecord
    .TextMatrix(.Row, 0) = Val(.TextMatrix(.Row - 1, 0)) + 1
    .TextMatrix(Row, 2) = StrKodeID
    lblCurrentRecord.Caption = "Selected Record: " & VSFlex.TextMatrix(VSFlex.Row, 0)

    .Rows = .Rows + 1
    .Row = .Rows - 1
    .Col = 3
    .EditCell
    .Refresh
Else
   
    StrSQL = "Update TblGaji set statusaktif = 0,last_update ='" & Now & "',last_user = '" & StrUser & "' Where IDGaji = '" & .TextMatrix(Row, 2) & "'"
    PerintahExecute (StrSQL)
    '-----Ambil ID -------------------
        GetNomorID ("TblGaji")
        StrKodeID = NewID
    '---------------------------------
    StrSQL = "Insert Into TblGaji(IDGaji,tingkat,Nip,Gaji,keterangan,periode,last_update,last_user,statusaktif) Values ('" & StrKodeID & "','" & .TextMatrix(Row, 3) & "','" & .TextMatrix(Row, 4) & "','" & .TextMatrix(Row, 6) & "','" & .TextMatrix(Row, 7) & "','" & Now & "','" & Now & "','" & StrUser & "','1')"
    PerintahExecute (StrSQL)
    .TextMatrix(Row, 2) = StrKodeID
    StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Tambah Data Gaji, " & .TextMatrix(Row, 4) & "','Data Gaji')"
    PerintahExecute (StrSQL)
    .TextMatrix(.Row, 0) = .Row
End If

Exit Sub
AdaError:
MsgBox err.Description
End With
End Sub

Private Function cekGaji(Nama As String) As Boolean
Dim Rscek As New ADODB.Recordset
Dim StrSQL As String
On Error GoTo AdaError

cekGaji = False
If CN.State = adStateClosed Then CN.Open
If Rscek.State = adStateOpen Then Rscek.Close

With VSFlex
    
    StrSQL = "Select * from TblGaji Where Nip='" & Nama & "'"
    Rscek.Open StrSQL, CN, adOpenStatic
    If Not Rscek.EOF Then
        cekGaji = True
    End If
End With

If Rscek.State = adStateOpen Then Rscek.Close
Exit Function
AdaError:
MsgBox err.Description
End Function


