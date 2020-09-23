VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSetLembur 
   Caption         =   "Setting Lembur"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   8235
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   8235
      TabIndex        =   12
      Top             =   7485
      Width           =   8235
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   8235
      TabIndex        =   11
      Top             =   7500
      Width           =   8235
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   8235
      TabIndex        =   2
      Top             =   7515
      Width           =   8235
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
         Left            =   1680
         TabIndex        =   3
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
         TabIndex        =   4
         Top             =   0
         Width           =   4755
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   4110
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Next 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Last 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3795
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Previous 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   5
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
            TabIndex        =   9
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
         TabIndex        =   10
         Top             =   60
         Width           =   1365
      End
   End
   Begin VB.PictureBox shpBar 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   6735
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Setting Lembur"
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
      Left            =   7200
      OleObjectBlob   =   "FrmSetLembur.frx":0000
      Top             =   720
   End
   Begin VSFlex8Ctl.VSFlexGrid VsFlex 
      Height          =   3975
      Left            =   0
      TabIndex        =   13
      Top             =   360
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
      BackColor       =   -2147483643
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
      FormatString    =   $"FrmSetLembur.frx":0234
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
      AllowUserFreezing=   2
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin MSComCtl2.DTPicker DTTanggal 
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   57802755
         CurrentDate     =   39551
      End
   End
End
Attribute VB_Name = "FrmSetLembur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CURR_COL As Integer

Dim RsSetting As New Recordset
Dim RecordPage As New clsPaging
Dim StrSQLParser As New clsSQLSelectParser
'Procedure used to filter records
Public Sub FilterRecord(ByVal srcCondition As String)
    StrSQLParser.RestoreStatement
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
                  .Row = .Rows - 1
                  .Col = 4
                  .EditCell
                  Exit Sub
           End With
        Case "Search"
            With frmSearchs
                Set .srcForm = Me
                Set .srcColumnHeaders = VsFlex
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
    With RsSetting
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




Private Sub DTTanggal_Change()
      VsFlex.Text = Format(DTTanggal.Value, "dd/MMM/yyyy")

End Sub

Private Sub DTTanggal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
   If VsFlex.TextMatrix(VsFlex.Row, 2) = "" Then VsFlex.Text = ""
    DTTanggal.Visible = False
End If
If KeyCode = 13 Then
      
    ' update grid value whenever the data changes
    With VsFlex
         .Col = .Col
         .Row = .Row
         If .Col = 3 Then
            .TextMatrix(.Row, 3) = Format(DTTanggal.Value, "dd/MMM/yyyy")
'            If .TextMatrix(.Row, Col + 1) <> "" Then
'                .Editable = flexEDKbdMouse
                .Col = 4
                .EditCell
                .Refresh
'            End If
            .SetFocus
        End If
        DTTanggal.Visible = False
             
   End With
   
End If
End Sub

Private Sub DTTanggal_LostFocus()
DTTanggal.Visible = False
VsFlex.TextMatrix(VsFlex.Row, 3) = Format(DTTanggal.Value, "dd/MMM/yyyy")
VsFlex.Col = 4
VsFlex.EditCell
VsFlex.Refresh
  
End Sub

'Private Sub VSFlex_DblClick()
'MsgBox VSFlex.Col
'End Sub

Private Sub VsFlex_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With VsFlex
    ' if this is a date column, edit it with the date picker control
    If VsFlex.ColDataType(Col) = flexDTDate Then
        Cancel = True

        DTTanggal.Move VsFlex.CellLeft, VsFlex.CellTop, VsFlex.CellWidth, VsFlex.CellHeight

        If VsFlex <> "" Then
            DTTanggal.Value = VsFlex
            DTTanggal.Tag = VsFlex
        End If
        DTTanggal.Visible = True
        DTTanggal.SetFocus
    Else
    End If
End With
End Sub

Private Sub VsFlex_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    ' don't scroll while editing dates
    If DTTanggal.Visible Then Cancel = True

End Sub

Private Sub VsFlex_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If DTTanggal.Visible Then Cancel = True

End Sub
Private Sub Form_Activate()
RefreshRecords
VsFlex.FrozenCols = 5
End Sub
Private Sub Form_Load()
 
    DTTanggal.CustomFormat = "dd/MMM/yyyy"
       DTTanggal.Value = Date
'    dttanggal.Value = DateSerial(Year(Now), Month(Now), 1)
    DTTanggal.Value = DateAdd("M", 0, DTTanggal.Value)
  
     If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path & "\Skins\" & skinsFileName
      Skin1.ApplySkin hwnd
     End If
        Me.Top = 0
        Me.Left = 0
   
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
        .Tables = "TblSetting"
        .SortOrder = "Tingkat,hari ASC"
        .wCondition = " StatusAktif = 1"
        .SaveStatement
    End With
    If RsSetting.State = adStateOpen Then RsSetting.Close
    RsSetting.CursorLocation = adUseClient
    RsSetting.Open StrSQLParser.StrSQLStatement, CN, adOpenStatic, adLockReadOnly
  

    With RecordPage
        .Start RsSetting, 50
        FillList 1
    End With

End Sub

Public Sub FillList(ByVal whichPage As Long)
    Dim Cboid     As String
    Dim cboid1    As String
    Dim cboID2    As String
    Dim i As Integer
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    
    Call IsiGrid(VsFlex, RsSetting, RecordPage.PageStart, RecordPage.PageEnd, 8, 2, False, True, , , , "PK")
    With VsFlex
        .FixedRows = 2
        .AddItem "", 1
        .ColDataType(3) = flexDTDate
        .ColFormat(3) = "dd/MMM/yyyy"
        .TextMatrix(0, 6) = "Upah"
        .TextMatrix(1, 6) = "Lembur"
        .TextMatrix(0, 7) = "Jam"
        .TextMatrix(1, 7) = "Lembur Ke"
        .TextMatrix(0, 8) = "S/D"
        .TextMatrix(0, 9) = "Upah"
        .TextMatrix(0, 10) = "Jam Lembur Ke"
        .TextMatrix(1, 10) = "Lembur Ke"
        .TextMatrix(0, 11) = "S/D"
        .TextMatrix(0, 12) = "Upah"
        .TextMatrix(0, 13) = "Jam Lembur Ke"
        .TextMatrix(1, 13) = "Lembur Ke"
        .TextMatrix(0, 14) = "S/D"
        .TextMatrix(0, 15) = "Upah"
        .TextMatrix(0, 16) = "Istirahat 1"
        .TextMatrix(0, 17) = "S/D"
        .TextMatrix(0, 18) = "Waktu"
        .TextMatrix(0, 19) = "Istirahat 2"
        .TextMatrix(0, 20) = "S/D"
        .TextMatrix(0, 21) = "Waktu"
        .TextMatrix(0, 22) = "Istirahat 3"
        .TextMatrix(0, 23) = "S/D"
        .TextMatrix(0, 24) = "Waktu"
        .TextMatrix(0, 25) = "Istirahat 4"
        .TextMatrix(0, 26) = "S/D"
        .TextMatrix(0, 27) = "Waktu"
        .TextMatrix(1, 28) = "Dari"
        .TextMatrix(1, 29) = "S/D"
        .TextMatrix(1, 30) = "Rp"
        .TextMatrix(0, 28) = "Jam Makan"
        .TextMatrix(0, 29) = "Jam Makan"
        .TextMatrix(0, 30) = "Jam Makan"
        
        .ColWidth(2) = 0
        .ColWidth(6) = 800
        .ColWidth(7) = 1000
        .ColWidth(10) = 1000
        .ColWidth(13) = 1000
        .ColWidth(31) = 0
        .ColWidth(32) = 0
        .ColWidth(33) = 0
        .ColWidth(4) = 800
        .ColWidth(5) = 800
        .ColWidth(8) = 800
        .ColWidth(9) = 800
        .ColWidth(11) = 800
        .ColWidth(12) = 800
        .ColWidth(13) = 1000
        .ColWidth(14) = 700
        .ColWidth(15) = 800
        .ColWidth(16) = 1000
        .ColWidth(17) = 700
        .ColWidth(18) = 800
        .ColWidth(19) = 1000
        .ColWidth(20) = 700
        .ColWidth(21) = 800
        .ColWidth(22) = 1000
        .ColWidth(23) = 700
        .ColWidth(24) = 800
        .ColWidth(25) = 1000
        .ColWidth(26) = 700
        .ColWidth(27) = 800
        .ColWidth(28) = 800
        .ColWidth(29) = 700
        .ColWidth(30) = 1000
         .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        For i = 3 To 15
           Cboid = "|" & i
           cboid1 = cboid1 + Cboid

         Next
         .ColComboList(4) = cboid1
         Cboid = vbNullString
         Cboid = "|Kerja|Libur"
         .ColComboList(5) = Cboid
        For lCol = 6 To 27
           .ColDataType(lCol) = flexDTCurrency
        Next
         Cboid = vbNullString
         cboid1 = vbNullString
         cboID2 = vbNullString
        For i = 0 To 24
            Cboid = "|" & i
            cboID2 = "|" & i & ".5"
            If i < 24 Then
                cboid1 = cboid1 + Cboid + cboID2
            Else
                cboid1 = cboid1 + Cboid
            End If
         Next i
          
        
        For lCol = 13 To 27
            If lCol >= 13 And lCol <= 15 Then .ColComboList(lCol) = cboid1
            If lCol = 16 Or lCol = 17 Then .ColComboList(lCol) = cboid1
            If lCol = 19 Or lCol = 20 Then .ColComboList(lCol) = cboid1
             If lCol = 22 Or lCol = 23 Then .ColComboList(lCol) = cboid1
             If lCol = 25 Or lCol = 26 Then .ColComboList(lCol) = cboid1
            .ColDataType(lCol) = flexDTCurrency
            
        Next
'        For i = 0 To 200
'            Cboid = "|" & i
'                Cboid1 = Cboid1 + Cboid
'
'         Next i
            
'            .ColComboList(15) = ""
'            .ColComboList(15) = Cboid1
            
            .ColFormat(30) = "#,###"
    End With
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    SetNavigation
    lblPageInfo.Caption = "Record " & RecordPage.PageInfo
    VsFlex_Click
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

    MDIMENU.mnuRPrint.Visible = False
    MDIMENU.mnuRSearch.Visible = False
    
    Set FrmSetLembur = Nothing
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
With VsFlex
    If Col = 3 Then
       .EditMaxLength = 200
    ElseIf Col = 4 Then
       .EditMaxLength = 15
    ElseIf Col >= 5 And .Col < 26 Then
      .EditMaxLength = 5
    End If
End With
End Sub

Private Sub VsFlex_Click()
    Dim x As Long
    On Error GoTo err
    If VsFlex.Text <> "" Then lblCurrentRecord.Caption = "Selected Record: " & VsFlex.TextMatrix(VsFlex.Row, 0)
Exit Sub
err:
        lblCurrentRecord.Caption = "Selected Record: NONE"
End Sub

Private Sub VsFlex_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96
If KeyAscii = 8 Then Exit Sub
If KeyAscii = 13 Then Exit Sub
If KeyAscii = vbKeyEscape Then Exit Sub
With VsFlex
    If .Col >= 5 Then
        If KeyAscii = 44 Then KeyAscii = 46: Exit Sub
        If KeyAscii < 44 Or KeyAscii > 57 Then KeyAscii = 0
    End If
    
End With
End Sub

Private Sub VsFlex_KeyUp(KeyCode As Integer, Shift As Integer)
With VsFlex

    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then
      VsFlex_Click
      Exit Sub
    ElseIf KeyCode = 27 Or KeyCode = vbKeyEscape Then
       VsFlex.Refresh
       Exit Sub
    End If
    If KeyCode = 13 Then
     
        If .Col >= 3 And .TextMatrix(.Row, 2) <> "" Then
              Call SimpanData(.Row)
              Exit Sub
           End If

        If .Col = 4 Then
            If .Text = "" Then
             .EditCell
             Exit Sub
          ElseIf .TextMatrix(.Row, 5) = "" Then
            .Col = 5
            .EditCell
             Exit Sub
          End If
       ElseIf .Col = 5 Then
             If .Text = "" Then
                .EditCell
                Exit Sub
             Else
               If cekSetting(.TextMatrix(.Row, 4), .TextMatrix(.Row, 5)) = True Then
                  MsgBox "Data Setting Sudah Ada", vbCritical
'                  .Text = ""
                  .EditCell
                  Exit Sub
               Else
                   .Col = 6
                   .EditCell
                    Exit Sub
               End If
            End If
        ElseIf .Col >= 6 And .Col <= 26 And .TextMatrix(.Row, 3) <> "" Then
            If .Text = "" Then
             .EditCell
             Exit Sub
          Else
            .Col = .Col + 1
            .EditCell
             Exit Sub
          End If
        ElseIf .Col = 27 And .TextMatrix(.Row, 3) <> "" Then
            If .Text = "" Then
               .EditCell
               Exit Sub
            Else
              SimpanData (.Row)
              Exit Sub
            End If
        End If
         .Refresh
    End If
 
End With
End Sub


Private Sub VsFlex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    MDIMENU.mnuRPrint.Visible = False
    MDIMENU.mnuRSearch.Visible = False
     MDIMENU.MnuRView.Visible = False
    MDIMENU.MnuRBatal.Visible = False
    PopupMenu MDIMENU.mnuRecA
End If
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth
End Sub

Private Sub Hapus()
Dim StrSQL As String
Dim Tanya As String
Dim ErrConn As Long
 If CekCurek("hapus", VsFlex) = False Then Exit Sub
            With VsFlex
            If MsgBox("Apakah Anda yakin ingin menghapus Data ?", vbQuestion + vbYesNo, "Konfirmasi hapus") = vbNo Then
               Exit Sub
            Else
             Do Until Lrow = .Rows - 1
                If .TextMatrix(Lrow, 1) = "-1" Then
                        StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(ServerTime, "yyyy-MM-dd HH:mm:ss") & "','" & StrUser & "','Hapus User  " & .TextMatrix(Lrow, 3) & "','Setting Lembur')"
                        CN.Execute StrSQL
                                
                        StrSQL = "Delete From TblSetting Where IDSetting = '" & .TextMatrix(Lrow, 2) & "'"
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
Adaerror:
If ErrConn > 0 Then CN.RollbackTrans
MsgBox err.Description
End Sub
'Prosedur Simpan Data
Private Sub SimpanData(ByVal Row As Long)
Dim StrSQL As String
Dim Setting As String

On Error GoTo Adaerror
With VsFlex
 
If .TextMatrix(Row, 2) = "" Then
     '-----Ambil ID -------------------
        GetNomorID ("TblSetting")
        StrKodeID = NewID
    '---------------------------------
    StrSQL = "Insert Into tblsetting (IDSetting,Berlaku_SD,Tingkat,Hari,UpahPerJam,JamLembur1,JamLembur2,"
    StrSQL = StrSQL & "Upah1,JamLembur3,JamLembur4,Upah2,JamLembur5,JamLembur6,Upah3,Ist1,Ist2,JamIst_1,Ist3,"
    StrSQL = StrSQL & "Ist4,JamIst_2,Ist5,Ist6,JamIst_3,Ist7,Ist8,JamIst_4,"
    StrSQL = StrSQL & "JamMakan1,JamMakan2,UpahMakan,statusaktif,last_user,last_update)Values"
    StrSQL = StrSQL & "('" & StrKodeID & "','" & Format(.TextMatrix(Row, 3), "MM-dd-yyyy") & "','" & .TextMatrix(Row, 4) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 5) & "','" & .TextMatrix(Row, 6) & "','" & .TextMatrix(Row, 7) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 8) & "','" & .TextMatrix(Row, 9) & "','" & .TextMatrix(Row, 10) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 11) & "','" & .TextMatrix(Row, 12) & "','" & .TextMatrix(Row, 13) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 14) & "','" & .TextMatrix(Row, 15) & "','" & .TextMatrix(Row, 16) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 17) & "','" & .TextMatrix(Row, 18) & "','" & .TextMatrix(Row, 19) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 20) & "','" & .TextMatrix(Row, 21) & "','" & .TextMatrix(Row, 22) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 23) & "','" & .TextMatrix(Row, 24) & "','" & .TextMatrix(Row, 25) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 26) & "','" & .TextMatrix(Row, 27) & "','" & .TextMatrix(Row, 28) & "','" & .TextMatrix(Row, 29) & "','" & .TextMatrix(Row, 30) & "','1','" & StrUser & "','" & ServerTime & "')"
    PerintahExecute (StrSQL)
    
    StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(ServerTime, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Tambah Setting Lembur, " & DTTanggal & "','Setting Lembur')"
    PerintahExecute (StrSQL)
       
    TotalRecord = TotalRecord + 1
    lblPageInfo.Caption = "Record " & ARecord & " - " & TotalRecord & " of " & TotalRecord
    .TextMatrix(.Row, 0) = Val(.TextMatrix(.Row - 1, 0)) + 1
    .TextMatrix(Row, 2) = StrKodeID
    lblCurrentRecord.Caption = "Selected Record: " & VsFlex.TextMatrix(VsFlex.Row, 0)
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .Col = 3
    .EditCell
    .Refresh
Else
    

'    StrSQL = "Update tblsetting Set last_user='" & StrUser & "',last_update='" & Now & "',StatusAktif ='0' Where IDSetting = '" & .TextMatrix(Row, 2) & "'"
'    PerintahExecute (StrSQL)
'    StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Rubah Setting Lembur, " & DTTanggal & "','Setting Lembur')"
'    PerintahExecute (StrSQL)
     StrSQL = "Delete tblsetting where IDSetting = '" & .TextMatrix(Row, 2) & "'"
     CN.Execute StrSQL
       '-----Ambil ID -------------------
        GetNomorID ("TblSetting")
        StrKodeID = NewID
    '---------------------------------
    StrSQL = "Insert Into tblsetting (IDSetting,Berlaku_SD,Tingkat,Hari,UpahPerJam,JamLembur1,JamLembur2,"
    StrSQL = StrSQL & "Upah1,JamLembur3,JamLembur4,Upah2,JamLembur5,JamLembur6,Upah3,Ist1,Ist2,JamIst_1,Ist3,"
    StrSQL = StrSQL & "Ist4,JamIst_2,Ist5,Ist6,JamIst_3,Ist7,Ist8,JamIst_4,"
    StrSQL = StrSQL & "JamMakan1,JamMakan2,UpahMakan,statusaktif,last_user,last_update)Values"
    StrSQL = StrSQL & "('" & StrKodeID & "','" & Format(.TextMatrix(Row, 3), "MM-dd-yyyy") & "','" & .TextMatrix(Row, 4) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 5) & "','" & .TextMatrix(Row, 6) & "','" & .TextMatrix(Row, 7) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 8) & "','" & .TextMatrix(Row, 9) & "','" & .TextMatrix(Row, 10) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 11) & "','" & .TextMatrix(Row, 12) & "','" & .TextMatrix(Row, 13) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 14) & "','" & .TextMatrix(Row, 15) & "','" & .TextMatrix(Row, 16) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 17) & "','" & .TextMatrix(Row, 18) & "','" & .TextMatrix(Row, 19) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 20) & "','" & .TextMatrix(Row, 21) & "','" & .TextMatrix(Row, 22) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 23) & "','" & .TextMatrix(Row, 24) & "','" & .TextMatrix(Row, 25) & "',"
    StrSQL = StrSQL & "'" & .TextMatrix(Row, 26) & "','" & .TextMatrix(Row, 27) & "','" & .TextMatrix(Row, 28) & "','" & .TextMatrix(Row, 29) & "','" & .TextMatrix(Row, 30) & "','1','" & StrUser & "','" & Now & "')"
    PerintahExecute (StrSQL)
    .TextMatrix(Row, 2) = StrKodeID

    StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & StrUser & "','Tambah Setting Lembur, " & DTTanggal & "','Setting Lembur')"
    PerintahExecute (StrSQL)
     .TextMatrix(.Row, 0) = .Row
End If

Exit Sub
Adaerror:
MsgBox err.Description
End With
End Sub

Private Function cekSetting(Tingkat As String, Hari As String) As Boolean
Dim Rscek As New ADODB.Recordset
Dim StrSQL As String
Dim TglAwal, TglAkhir As Date
On Error GoTo Adaerror

cekSetting = False

If Rscek.State = adStateOpen Then Rscek.Close

With VsFlex
    StrSQL = "select * from tblsetting " & _
    "where tingkat = '" & Tingkat & "' " & _
    "and hari = '" & Hari & "' And StatusAktif =1"
   Rscek.Open StrSQL, CN, adOpenStatic
    If Not Rscek.EOF Then
        cekSetting = True
    End If
End With

If Rscek.State = adStateOpen Then Rscek.Close
Exit Function
Adaerror:
MsgBox err.Description
End Function
