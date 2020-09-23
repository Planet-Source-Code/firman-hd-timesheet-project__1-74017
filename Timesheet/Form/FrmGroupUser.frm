VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmGroupUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group User"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   Icon            =   "FrmGroupUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   4335
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
      Left            =   2760
      Picture         =   "FrmGroupUser.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Delete "
      Top             =   5085
      Width           =   735
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
      Left            =   3480
      Picture         =   "FrmGroupUser.frx":5D26
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5085
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "FrmGroupUser.frx":B086
      Top             =   5280
   End
   Begin VSFlex8Ctl.VSFlexGrid VsFlex 
      Height          =   4935
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4335
      _cx             =   7646
      _cy             =   8705
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
      FormatString    =   $"FrmGroupUser.frx":B2BA
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
   Begin VB.Image Bottom 
      Height          =   60
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   4050
   End
   Begin VB.Image imgBtnUp 
      Height          =   480
      Left            =   120
      Picture         =   "FrmGroupUser.frx":B399
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBtnDn 
      Height          =   480
      Left            =   480
      Picture         =   "FrmGroupUser.frx":BC63
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "FrmGroupUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdHapus_Click()
Dim ErrConn As Long

On Error GoTo AdaError
If CekCurek("Hapus", VsFlex) = False Then Exit Sub
If MsgBox("Data akan dihapus ?", vbCritical + vbYesNo, "Pesan") = vbYes Then
    If CN.State = adStateClosed Then CN.Open
    I = 1
    With VsFlex
    Do Until I = .Rows - 1
     If .TextMatrix(I, 1) = "-1" Then
     If Rscek.State = adStateOpen Then Rscek.Close
'        StrSQL = "Select IDGroup From Login Where IDGroup='" & .TextMatrix(i, 1) & "'"
'        Rscek.Open StrSQL, CN, adOpenStatic
'        If Rscek.EOF Then
            StrSQL = "Delete from TblGroup_User_Detail Where IDGroup='" & .TextMatrix(I, 2) & "'"
            CN.Execute StrSQL
            StrSQL = "Delete from TblGroup_User Where IDGroup='" & .TextMatrix(I, 2) & "'"
            CN.Execute StrSQL
            .RemoveItem (I)
            I = I - 1

'        Else
'           GoTo AdaError
        End If
'    End If
    I = I + 1
    Loop

    End With
 End If

Exit Sub
AdaError:
    MsgBox "Nama Group Tidak Dapat Dihapus Karena Sudah Dipakai", vbCritical

End Sub

Private Sub SimpanData(ByVal Row As Long)
On Error GoTo AdaError
If CN.State = adStateClosed Then CN.Open
With VsFlex
If VsFlex.TextMatrix(Row, 2) = "" Then
    StrSQL = "Insert into TblGroup_User (NamaGroup) VALUES ('" & VsFlex.TextMatrix(Row, 3) & "')"
    CN.Execute StrSQL

    If Rscek.State = adStateOpen Then Rscek.Close
    Rscek.Open "Select IDGroup from TblGroup_User Where NamaGroup='" & VsFlex.TextMatrix(Row, 3) & "'", CN, adOpenStatic
    If Not Rscek.EOF Then
        VsFlex.TextMatrix(Row, 2) = Rscek!IDGroup
    End If
    If Rscek.State = adStateOpen Then Rscek.Close
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .Cell(flexcpPicture, .Row, 4) = imgBtnUp
    .Cell(flexcpPictureAlignment, .Row, 4) = flexAlignCenterCenter

    .Col = 3
    .EditCell
    .Refresh
Else
    StrSQL = "Update TblGroup_User set NamaGroup='" & VsFlex.TextMatrix(Row, 3) & "' Where IDGroup='" & VsFlex.TextMatrix(Row, 2) & "'"
    CN.Execute StrSQL
End If
End With

Exit Sub
AdaError:
MsgBox err.Number & Chr(13) & err.Description
End Sub

Private Sub Showdata()
Dim SQL As String

On Error GoTo AdaError

If Rscek.State = adStateOpen Then Rscek.Close
SQL = "select idgroup,namagroup from Tblgroup_user order by IDGroup"

Rscek.Open SQL, CN, adOpenStatic

With VsFlex
    .Rows = 1
    .Rows = 2
Do Until Rscek.EOF
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = .Row
    .TextMatrix(.Row, 2) = Rscek!IDGroup
    .TextMatrix(.Row, 3) = Rscek!NamaGroup

    Rscek.MoveNext
    .Rows = .Rows + 1
Loop
For Lrow = 1 To .Rows - 1
        .Cell(flexcpPicture, Lrow, 4) = imgBtnUp
        .Cell(flexcpPictureAlignment, Lrow, 4) = flexAlignCenterCenter
    Next
End With
If Rscek.State = adStateOpen Then Rscek.Close

Exit Sub
AdaError:
MsgBox err.Number & Chr(13) & err.Description
End Sub
Sub Setgrid()
With VsFlex
    .Rows = 1
    .Rows = 2
    .Cols = 5
    .TextMatrix(0, 0) = "No"
    .TextMatrix(0, 1) = "Do"
    .TextMatrix(0, 2) = "ID"
    .TextMatrix(0, 3) = "Nama Group"
    .TextMatrix(0, 4) = "Menu"
    .ColWidth(0) = 300
    .ColWidth(1) = 400
    .ColWidth(2) = 0
    .ColWidth(3) = 2200
    .ColWidth(4) = 1000
    .ColDataType(1) = flexDTBoolean
    For Lrow = 1 To .Rows - 1
        .Cell(flexcpPicture, Lrow, 4) = imgBtnUp
        .Cell(flexcpPictureAlignment, Lrow, 4) = flexAlignCenterCenter
    Next
    For Lcol = 1 To .Cols - 1
        .FixedAlignment(Lcol) = flexAlignCenterCenter
    Next
End With
End Sub

Private Sub CmdExit_Click()
Unload Me
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
End Sub
Private Sub VSFlex_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col >= 3 And VsFlex.TextMatrix(Row, 2) <> "" Then
    Call SimpanData(Row)
End If

End Sub

Private Sub VsFlex_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VsFlex
   ' -----------Cek Len Karakter
    If .Col = 3 Then
       .EditMaxLength = 100
    End If
End With
End Sub

Private Sub VSFlex_Click()
With VsFlex
    If .Col = 1 And .TextMatrix(.Row, 3) = "" Then
        .TextMatrix(.Row, 1) = ""
    ElseIf .Col > 3 And .TextMatrix(.Row, 3) = "" Then
        .Col = 3
        .EditCell
    End If
End With
End Sub
Private Sub VsFlex_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single, Cancel As Boolean)
    
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
    d = VsFlex.Cell(flexcpLeft, r, c) + VsFlex.Cell(flexcpWidth, r, c) - X
'    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    VsFlex.Cell(flexcpPicture, r, c) = imgBtnDn
If c = 4 Then
    FrmGroupPermission.MyIDGroup = VsFlex.TextMatrix(r, 2)
     FrmGroupPermission.Caption = VsFlex.TextMatrix(r, 3)
    FrmGroupPermission.Show vbModal
End If
    VsFlex.Cell(flexcpPicture, r, c) = imgBtnUp
    Cancel = True

End Sub

Private Sub VsFlex_GotFocus()
VSFlex_Click
End Sub

Private Sub VsFlex_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96
End Sub

Private Sub VsFlex_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96
End Sub

Private Sub VSFlex_KeyUp(KeyCode As Integer, Shift As Integer)
With VsFlex
    If KeyCode = 13 Then
        If .Col = 3 Then
          If .Text = "" Then
             .EditCell
             Exit Sub
          Else
               If cekGroup(.Text) = True And .TextMatrix(.Row, 2) = "" Then
                    MsgBox "Nama Group Sudah Ada", vbCritical
                    .EditCell
                    .Refresh
                    Exit Sub
                ElseIf .TextMatrix(.Row, 4) = "" Then
'                    .TextMatrix(.Row, 0) = .Row
'                    .Col = 4
'                    .EditCell
'                    .Refresh
'                    Exit Sub
                    SimpanData (VsFlex.Row)
               End If
            End If
       ElseIf .Col = 4 And .TextMatrix(.Row, 3) <> "" Then
'                    With frmBarangDetail
'                         .RowBarang = VsFlex.Row
'                         SimpanData (VsFlex.Row)
'                        .StrIDBarang = VsFlex.TextMatrix(VsFlex.Row, 2)
'                        .StrKategori = CmbKategori.Text
'                        ShowQty = True
'                        .Show 1
'                         ShowQty = False
'                    End With
'                        VsFlex.Row = VsFlex.Rows - 1
'                        VsFlex.Cell(flexcpPicture, VsFlex.Row, 13) = imgBtnUp
'                        VsFlex.Cell(flexcpPictureAlignment, VsFlex.Row, 13) = flexAlignRightCenter
'                        VsFlex.Cell(flexcpPicture, VsFlex.Row, 14) = imgBtnUp
'                        VsFlex.Cell(flexcpPictureAlignment, VsFlex.Row, 14) = flexAlignRightCenter
'                        VsFlex.Col = 3
'                        VsFlex.EditCell
'                        VsFlex.SetFocus
'                        Exit Sub
            End If
         .Refresh
       End If

 
End With
End Sub
Private Function cekGroup(Nama As String) As Boolean
On Error GoTo AdaError

cekGroup = False
If Rscek.State = adStateOpen Then Rscek.Close

With VsFlex
    StrSQL = "select * From Tblgroup_user Where namaGroup ='" & Nama & "'"
    Rscek.Open StrSQL, CN, adOpenStatic
    If Not Rscek.EOF Then
        cekGroup = True
    End If
End With

If Rscek.State = adStateOpen Then Rscek.Close
Exit Function
AdaError:
MsgBox err.Description
End Function

