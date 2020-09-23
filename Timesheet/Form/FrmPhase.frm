VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Begin VB.Form FrmPhase 
   Caption         =   "Phase"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   9840
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9840
      TabIndex        =   3
      Top             =   7275
      Width           =   9840
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
      ScaleWidth      =   9840
      TabIndex        =   2
      Top             =   7260
      Width           =   9840
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   9840
      TabIndex        =   1
      Top             =   7245
      Width           =   9840
   End
   Begin VB.PictureBox shpBar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   9840
      TabIndex        =   0
      Top             =   0
      Width           =   9840
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8040
      OleObjectBlob   =   "FrmPhase.frx":0000
      Top             =   2280
   End
   Begin VSFlex8Ctl.VSFlexGrid VsFlex 
      Height          =   3975
      Left            =   0
      TabIndex        =   12
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
      FormatString    =   $"FrmPhase.frx":0234
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
Attribute VB_Name = "FrmPhase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CURR_COL As Integer

Dim RsPhase As New Recordset
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
                  .EditCell
           End With
         Case "Search"
           With frmSearchs
                Set .srcForm = Me
                Set .srcColumnHeaders = VsFlex
                .srcNoOfCol = 4
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
    With RsPhase
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
        .Fields = "*"
        .Tables = "TblPhase"
        If StrUser <> 3578 Then .wCondition = "kd_Divisi = '" & KodeDivisi & "'"
        .SaveStatement
    End With
    If RsPhase.State = adStateOpen Then RsPhase.Close
    RsPhase.CursorLocation = adUseClient
    RsPhase.Open StrSQLParser.StrSQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPage
        .Start RsPhase, 50
        FillList 1
    End With
    
End Sub
Function SimpanData(Row As Integer)

With VsFlex
    If .TextMatrix(Row, 4) = "" Then
        MsgBox "Kode Phase Masih Kosong", vbCritical
        .Col = 4
        .EditCell
    End If
    If .TextMatrix(Row, 5) = "" Then
        MsgBox "Keterangan Masih Kosong", vbCritical
        .Col = 5
        .EditCell
    End If
    If .TextMatrix(Row, 2) = "" Then
        StrSQL = "Insert Into TblPhase(kd_divisi,kode_phase,keterangan)Values('" & KodeDivisi & "','" & .TextMatrix(Row, 4) & "','" & .TextMatrix(Row, 5) & "')"
        CN.Execute StrSQL
        
        If Rscek.State = adStateOpen Then Rscek.Close
        Rscek.Open "Select * From TblPhase Where kd_divisi = '" & KodeDivisi & "' And Kode_Phase = '" & .TextMatrix(Row, 4) & "' ", CN, adOpenStatic
        If Not Rscek.EOF Then .TextMatrix(Row, 2) = Rscek!ID
         
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .TextMatrix(.Row, 0) = .Rows - 1
        .Col = 4
        .EditCell
        
    Else
        StrSQL = "Update TblPhase Set kode_phase='" & .TextMatrix(Row, 4) & "',keterangan='" & .TextMatrix(Row, 5) & "' Where Id = '" & .TextMatrix(Row, 2) & "'"
        CN.Execute StrSQL
    End If
End With
End Function
Public Sub FillList(ByVal whichPage As Long)
    Dim i As Integer
    Dim Cboid, cboid1 As String
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call IsiGrid(VsFlex, RsPhase, RecordPage.PageStart, RecordPage.PageEnd, 8, 2, False, True, , , , "PK")
     With VsFlex
        .ColWidth(3) = 0
        .ColWidth(4) = 1000
        .ColWidth(5) = 3000
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
    Set FrmPhase = Nothing
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
 

Private Sub VsFlex_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If VsFlex.TextMatrix(Row, 2) <> "" Then SimpanData (Row)
End Sub

Private Sub VsFlex_Click()
     
       VsFlex.Editable = flexEDKbdMouse
     

    If VsFlex.Text <> "" Then lblCurrentRecord.Caption = "Selected Record: " & VsFlex.Row
    
Exit Sub
err:
        lblCurrentRecord.Caption = "Selected Record: NONE"
End Sub
 

Private Sub VsFlex_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96
If KeyAscii = 13 Or KeyAscii = 8 Then Exit Sub

End Sub

Private Sub VsFlex_KeyUp(KeyCode As Integer, Shift As Integer)
With VsFlex

        Select Case .Col
            Case 4
                If KeyCode = 13 Then
                    .Col = .Col + 1
                    .EditCell
                End If
            
           Case 5
                If KeyCode = 13 Then
                    If .TextMatrix(.Row, 2) = "" Then SimpanData (.Row)
                End If
        End Select
    
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
Dim Divisi As String
Dim Tanya As String
Dim ErrConn As Long
  
 If CekCurek("hapus", VsFlex) = False Then Exit Sub
 If CN.State = adStateClosed Then CN.Open
            With VsFlex
            If MsgBox("Apakah Anda yakin ingin menghapus Data ?", vbQuestion + vbYesNo, "Konfirmasi hapus") = vbNo Then
               Exit Sub
            Else
             Do Until Lrow = .Rows - 1
                If .TextMatrix(Lrow, 1) = "-1" Then
                     StrSQL = "Delete From TblPhase Where ID= '" & .TextMatrix(Lrow, 2) & "'"
                     CN.Execute StrSQL
                            
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









