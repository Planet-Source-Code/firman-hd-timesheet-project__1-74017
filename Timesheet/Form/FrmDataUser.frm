VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Begin VB.Form FrmDataUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data User"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5565
   Icon            =   "FrmDataUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   5565
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Menu"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "FrmDataUser.frx":0E42
      Top             =   7200
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4545
      Picture         =   "FrmDataUser.frx":1076
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7005
      Width           =   735
   End
   Begin VB.CommandButton CmdHapus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Picture         =   "FrmDataUser.frx":63D6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Delete "
      Top             =   7005
      Width           =   735
   End
   Begin VSFlex8Ctl.VSFlexGrid VsFlex 
      Height          =   6855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      _cx             =   9763
      _cy             =   12091
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
      FormatString    =   $"FrmDataUser.frx":B2BA
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
   End
   Begin VB.Image imgBtnDn 
      Height          =   480
      Left            =   480
      Picture         =   "FrmDataUser.frx":B399
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBtnUp 
      Height          =   480
      Left            =   120
      Picture         =   "FrmDataUser.frx":BC63
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Bottom 
      Height          =   60
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   4050
   End
End
Attribute VB_Name = "FrmDataUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Setgrid()
With VsFlex
    .Rows = 1
    .Rows = 2
    .Cols = 9
    .TextMatrix(0, 0) = "No"
    .TextMatrix(0, 1) = "Do"
    .TextMatrix(0, 2) = "ID"
    .TextMatrix(0, 3) = "NIP"
    .TextMatrix(0, 4) = "Group User"
    .TextMatrix(0, 5) = "Password"
    .TextMatrix(0, 6) = "Password"
    .TextMatrix(0, 7) = "OldNIP"
    .ColWidth(0) = 300
    .ColWidth(1) = 400
    .ColWidth(2) = 0
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 950
    .ColWidth(6) = 0
    .ColWidth(7) = 0
     .ColWidth(8) = 500
    .ColDataType(1) = flexDTBoolean
    For Lrow = 1 To .Rows - 1
        .Cell(flexcpPicture, Lrow, 5) = imgBtnUp
        .Cell(flexcpPictureAlignment, Lrow, 5) = flexAlignCenterCenter
    Next
    For lCol = 1 To .Cols - 1
        .FixedAlignment(lCol) = flexAlignCenterCenter
    Next
End With
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdHapus_Click()
Dim StrSQL As String
Dim Lrow As Integer
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
                       
                        StrSQL = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(ServerTime, "yyyy-MM-dd HH:mm:ss") & "','" & StrUser & "','Hapus User  " & .TextMatrix(Lrow, 3) & "','Data User')"
                        CN.Execute StrSQL
                        
                        StrSQL = "Delete from tbldata_user Where ID='" & .TextMatrix(Lrow, 2) & "'"
                        PerintahExecute (StrSQL)
                        
                        .RemoveItem (Lrow)
                        Lrow = Lrow - 1
                End If
                Lrow = Lrow + 1
            Loop
                     
                     
             End If
                      
         End With

Exit Sub
Adaerror:
If err.Number = "-2147217873" Then
    MsgBox "Data tidak dapat dihapus karena masih dipakai oleh tabel " & Mid(err.Description, InStr(1, err.Description, "table") + 5, InStr(1, err.Description, "column") - 2), vbInformation, "Pesan Error"
Else
    MsgBox err.Number & Chr(13) & err.Description
End If
End Sub

Private Sub Command1_Click()
FrmGroupMenu.show
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path & "\Skins\" & skinsFileName
      Skin1.ApplySkin hwnd
    End If
 
Setgrid
Showdata
LoadGroup
If StrUser <> "3578" Then Command1.Visible = False
End Sub
Sub LoadGroup()
    Dim Cboid     As String
    Dim cboid1    As String

    Cboid = vbNullString
    cboid1 = vbNullString
    If Rscek.State = adStateOpen Then Rscek.Close
    Rscek.Open "SELECT * FROM TblGroup_User ORDER BY IDGroup ASC", CN, adOpenStatic
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("NamaGroup")
      cboid1 = cboid1 + Cboid
      Rscek.MoveNext
    Loop
    VsFlex.ColComboList(4) = cboid1
    If Rscek.State = adStateOpen Then Rscek.Close
End Sub
Sub Showdata()
Dim cboid1 As String
cboid1 = "0|1"
With VsFlex
  If Rscek.State = adStateOpen Then Rscek.Close
  Rscek.Open "Select * From tbldata_user Order By NIP ASC", CN, adOpenStatic
  
  Do Until Rscek.EOF
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = .Row
    .TextMatrix(.Row, 2) = Rscek!ID
    .TextMatrix(.Row, 3) = Rscek!NIP
    .TextMatrix(.Row, 4) = Rscek!Type
    .TextMatrix(.Row, 6) = Enc.DecryptString(Rscek!Password)
'    .TextMatrix(.Row, 6) = Rscek!Password
    .TextMatrix(.Row, 7) = Rscek!NIP
    .TextMatrix(.Row, 8) = Rscek!StatusLogin
   
    .Rows = .Rows + 1
    Rscek.MoveNext
  Loop
    .ColComboList(8) = cboid1
   For Lrow = 1 To .Rows - 1
        .Cell(flexcpPicture, Lrow, 5) = imgBtnUp
        .Cell(flexcpPictureAlignment, Lrow, 5) = flexAlignCenterCenter
    Next
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmDataUser = Nothing
End Sub

Private Sub VsFlex_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, Cancel As Boolean)
    
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    VsFlex.Row = VsFlex.Row
    ' get cell that was clicked
    Dim r&, c&
    r = VsFlex.MouseRow
    c = VsFlex.MouseCol
    
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    If VsFlex.TextMatrix(r, 3) = "" Then Exit Sub
 
    ' make sure the click was on a cell with a button
    If Not (VsFlex.Cell(flexcpPicture, r, c) Is imgBtnUp) Then Exit Sub
    
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = VsFlex.Cell(flexcpLeft, r, c) + VsFlex.Cell(flexcpWidth, r, c) - x
'    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    VsFlex.Cell(flexcpPicture, r, c) = imgBtnDn
If c = 5 Then
   With FrmPassword
        .RowFlex = r
        .NamaUser = VsFlex.TextMatrix(r, 3)
        .UserPassword = VsFlex.TextMatrix(r, 6)
        .Add = True
        .IDUser = VsFlex.TextMatrix(r, 2)
        .IDGroup = VsFlex.TextMatrix(r, 4)
        .show vbModal
    End With
    With VsFlex
        If .TextMatrix(.Rows - 1, 3) <> "" Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 3
            .EditCell
        End If
    End With
     VsFlex.Row = VsFlex.Rows - 1
     VsFlex.Cell(flexcpPicture, VsFlex.Row, 5) = imgBtnUp
     VsFlex.Cell(flexcpPictureAlignment, VsFlex.Row, 5) = flexAlignCenterCenter
  End If
    VsFlex.Cell(flexcpPicture, r, c) = imgBtnUp
    Cancel = True

End Sub

Private Sub VsFlex_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96
End Sub

Private Sub VsFlex_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96
If KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = vbKeyEscape Then Exit Sub
With VsFlex
    If .Col = 5 Then If KeyAscii > 1 Then KeyAscii = 0
    If .Col = 3 Then
               
       If KeyAscii = 44 Then KeyAscii = 46
       If KeyAscii < 44 Or KeyAscii > 57 Then KeyAscii = 0
    End If
End With
End Sub

Private Sub VsFlex_KeyUp(KeyCode As Integer, Shift As Integer)
With VsFlex
   If KeyCode = 27 Then
       VsFlex.Refresh
       Exit Sub
    End If
    
    If KeyCode = 13 Then
        If .Col = 3 Then
          If .Text = "" Then
             .EditCell
             Exit Sub
          Else
               If CekUser(.Text) = True Then
                    MsgBox "NIP User Sudah Ada", vbCritical
                    .EditCell
                    .Refresh
                    Exit Sub
                ElseIf .TextMatrix(.Row, 4) = "" Then
                    .TextMatrix(.Row, 0) = .Row
                    .Col = 4
                    .EditCell
                    .Refresh
                    Exit Sub
               End If
          End If
        
       ElseIf .Col = 4 And .TextMatrix(.Row, 2) = "" Then
             With FrmPassword
                    .RowFlex = VsFlex.Row
                    .NamaUser = VsFlex.TextMatrix(VsFlex.Row, 3)
                    .UserPassword = VsFlex.TextMatrix(VsFlex.Row, 6)
                    .Add = True
                    .IDUser = VsFlex.TextMatrix(VsFlex.Row, 2)
                    .IDGroup = VsFlex.TextMatrix(VsFlex.Row, 4)
                    .show vbModal
             End With
            With VsFlex
               If .TextMatrix(.Rows - 1, 3) <> "" Then
                         .Rows = .Rows + 1
                        .Row = .Rows - 1
                        .Col = 3
                        .EditCell
                    End If
             End With
                 VsFlex.Row = VsFlex.Rows - 1
                 VsFlex.Cell(flexcpPicture, VsFlex.Row, 5) = imgBtnUp
                 VsFlex.Cell(flexcpPictureAlignment, VsFlex.Row, 5) = flexAlignCenterCenter
       Else
            SimpanData (.Row)
       End If
    End If
End With
End Sub
Private Sub SimpanData(ByVal Row As Long, Optional Col As Long)
Dim SQL As String
Dim CmSimpan As New ADODB.Command

On Error GoTo Adaerror

If CN.State = adStateClosed Then CN.Open
With VsFlex
    
      
    CmSimpan.ActiveConnection = CN
    CmSimpan.CommandText = "Update tbldata_user set NIP='" & VsFlex.TextMatrix(Row, 3) & "',Type = '" & VsFlex.TextMatrix(Row, 4) & "',StatusLogin ='" & .TextMatrix(Row, 8) & "' Where ID='" & VsFlex.TextMatrix(Row, 2) & "'"
    CmSimpan.Execute
    

    CmSimpan.ActiveConnection = CN
    CmSimpan.CommandText = "Insert into TblLog_User (Tanggal,Nama_User,Log_User,Modul) VALUES ('" & Format(ServerTime, "yyyy-MM-dd HH:mm:ss") & "','" & StrUser & "','Rubah User & " & .TextMatrix(Row, 7) & "','Data User')"
    CmSimpan.Execute
End With
    
Set CmSimpan.ActiveConnection = Nothing
Exit Sub
Adaerror:
MsgBox err.Number & Chr(13) & err.Description
End Sub
Private Function CekUser(Nama As String) As Boolean
On Error GoTo Adaerror

CekUser = False
If Rscek.State = adStateOpen Then Rscek.Close

With VsFlex
    StrSQL = "select * From tbldata_user Where NIP ='" & Nama & "'"
    Rscek.Open StrSQL, CN, adOpenStatic
    If Not Rscek.EOF Then
        CekUser = True
    End If
End With

If Rscek.State = adStateOpen Then Rscek.Close
Exit Function
Adaerror:
MsgBox err.Description
End Function
