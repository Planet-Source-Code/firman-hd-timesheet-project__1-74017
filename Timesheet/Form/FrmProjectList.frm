VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Begin VB.Form FrmProjectList 
   Caption         =   "List Project"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   9285
   WindowState     =   2  'Maximized
   Begin VB.PictureBox shpBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   6735
      TabIndex        =   11
      Top             =   0
      Width           =   6735
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "List Project"
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
         TabIndex        =   12
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   9285
      TabIndex        =   10
      Top             =   5955
      Width           =   9285
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   9285
      TabIndex        =   9
      Top             =   5970
      Width           =   9285
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9285
      TabIndex        =   0
      Top             =   5985
      Width           =   9285
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
         Left            =   1920
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   0
         Width           =   4755
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   4110
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Next 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Last 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3795
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Previous 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "First 250"
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
            TabIndex        =   7
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
         TabIndex        =   8
         Top             =   60
         Width           =   1365
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8040
      OleObjectBlob   =   "FrmProjectList.frx":0000
      Top             =   2280
   End
   Begin VSFlex8Ctl.VSFlexGrid VsFlex 
      Height          =   3975
      Left            =   0
      TabIndex        =   13
      ToolTipText     =   "Klik Kanan Mouse / Double Klik..."
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
      FormatString    =   $"FrmProjectList.frx":0234
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
End
Attribute VB_Name = "FrmProjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CURR_COL As Integer

Dim RsProject As New Recordset
Dim RecordPage As New clsPaging
Dim StrSQLParser As New clsSQLSelectParser


'Procedure used to filter records
Public Sub FilterRecord(ByVal srcCondition As String)
    StrSQLParser.RestoreStatement
    If Left(srcCondition, 3) = "Kode" Then srcCondition = "Project." & srcCondition
    StrSQLParser.wCondition = srcCondition
    ReloadRecords StrSQLParser.StrSQLStatement
End Sub

Public Sub Perintah(ByVal What As String)
Dim Lrow As Long
Dim lCol As Long
    On Error GoTo err
    Select Case What
        Case "New"
           With VsFlex
                  FrmProject.DeAktif
                  FrmProject.show vbModal
'                  RefreshRecords
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
    With RsProject
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
        .Fields = "Project.ID,Project.Kode, Project.Nama, Project.Status,Project.Nip_PM, karyawan.Nama AS [Nama PM], Project.Keterangan, Project.Tgl_Update As [Tgl Awal],Project.Tgl_Akhir, Project.Kd_Divisi"
        .Tables = "karyawan INNER JOIN Project ON karyawan.NIP = Project.Nip_PM "
        
        If UCase(strGroup) = "IT" Or UCase(strGroup) = "PTW" Then
            .wCondition = "Project.kd_divisi <> ''"
        Else
            .wCondition = "Project.kd_divisi = '" & KodeDivisi & "'"
        End If
         .SortOrder = "Project.Kode,Project.kd_divisi"
        .SaveStatement
       
    End With
    If RsProject.State = adStateOpen Then RsProject.Close
    RsProject.CursorLocation = adUseClient
    RsProject.Open StrSQLParser.StrSQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPage
        .Start RsProject, 50
        FillList 1
    End With

End Sub

Public Sub FillList(ByVal whichPage As Long)
    Dim i As Integer
    Dim Cboid, cboid1 As String
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call IsiGrid(VsFlex, RsProject, RecordPage.PageStart, RecordPage.PageEnd, 8, 2, False, True, , , , "PK")
    With VsFlex
        .ColWidth(3) = 1000
        .ColWidth(4) = 3000
        .ColWidth(5) = 1200
        .ColWidth(7) = 3500
        .ColWidth(6) = 1000
        .ColWidth(11) = 0
'        For I = 3 To 15
'           Cboid = "|" & I
'           cboid1 = cboid1 + Cboid
'
'         Next
'       .ColComboList(3) = cboid1
    .Rows = .Rows - 1
    End With
    
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
 
    Set FrmProject = Nothing
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


Private Sub VsFlex_Click()
    If VsFlex.Col = 1 Then
       VsFlex.Editable = flexEDKbdMouse
    Else
       VsFlex.Editable = flexEDNone
             
    End If
    If VsFlex.Text <> "" Then lblCurrentRecord.Caption = "Selected Record: " & VsFlex.Row
    
Exit Sub
err:
        lblCurrentRecord.Caption = "Selected Record: NONE"
End Sub

Private Sub VsFlex_DblClick()
With VsFlex
    If .Col >= 3 Then
       
        FrmProject.Showdata (.TextMatrix(.Row, 2))
        FrmProject.show vbModal
'        RefreshRecords
    End If
End With
End Sub

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
Dim Project As String
Dim Tanya As String
Dim ErrConn As Long
   VsFlex.Rows = VsFlex.Rows + 1
 If CekCurek("hapus", VsFlex) = False Then Exit Sub
 If CN.State = adStateClosed Then CN.Open
            With VsFlex
            If MsgBox("Apakah Anda yakin ingin menghapus Data ?", vbQuestion + vbYesNo, "Konfirmasi hapus") = vbNo Then
               Exit Sub
            Else
                Lrow = 1
                
             Do Until Lrow = .Rows - 1
                If .TextMatrix(Lrow, 1) = "-1" Then
                            StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Hapus Data Project, " & .TextMatrix(Lrow, 4) & "','Project')"
                            PerintahExecute (StrSQL)
                                
                            StrSQL = "Delete From Project Where ID = '" & .TextMatrix(Lrow, 2) & "'"
                            PerintahExecute (StrSQL)
                            .RemoveItem (Lrow)
                            Lrow = Lrow - 1
                    End If
                Lrow = Lrow + 1
            Loop
                     RefreshRecords
                End If
            End With
    VsFlex.Rows = VsFlex.Rows - 1
Exit Sub

Adaerror:
If ErrConn > 0 Then CN.RollbackTrans
MsgBox err.Description
End Sub



