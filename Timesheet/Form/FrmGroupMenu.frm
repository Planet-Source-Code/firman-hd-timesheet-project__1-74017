VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmGroupMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Menu"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "FrmGroupMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6735
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "FrmGroupMenu.frx":0E42
      Top             =   4920
   End
   Begin VB.CommandButton CmdHapus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "F8"
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
      Left            =   5040
      Picture         =   "FrmGroupMenu.frx":1076
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Delete "
      Top             =   5160
      Width           =   735
   End
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
      Height          =   615
      Left            =   5760
      Picture         =   "FrmGroupMenu.frx":5F5A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton CmdAuto 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Auto Fill"
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
      Left            =   4080
      Picture         =   "FrmGroupMenu.frx":B2BA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   855
   End
   Begin VSFlex8Ctl.VSFlexGrid VsFlex 
      Height          =   4935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6735
      _cx             =   11880
      _cy             =   8705
      Appearance      =   2
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
      FormatString    =   $"FrmGroupMenu.frx":10431
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
End
Attribute VB_Name = "FrmGroupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim myNamaMenu As String

Private Sub CmdAuto_Click()
Dim CmSimpan As New ADODB.Command
Dim Rscek As New ADODB.Recordset
Dim SQL As String
Dim MA As CONTROL

On Error Resume Next
If CN.State = adStateClosed Then CN.Open
CmSimpan.ActiveConnection = CN
SQL = "Delete From TblGroup_Menu"
CmSimpan.CommandText = SQL
CmSimpan.Execute

If Rscek.State = adStateOpen Then Rscek.Close

With VsFlex
    .Rows = 1
    .Rows = 2
For Each MA In MDIMENU.Controls
    Debug.Print MA.Name
    .Row = .Rows - 1
     If InStr(1, MA.Name, "Skin1") = 0 And InStr(1, MA.Name, "mnuRecA") = 0 And InStr(1, MA.Name, "mnuRNew") = 0 And InStr(1, MA.Name, "MnuRSelect") = 0 And InStr(1, MA.Name, "mnuRSearch") = 0 And InStr(1, MA.Name, "mnuRDelete") = 0 And InStr(1, MA.Name, "mnuRefresh") = 0 And InStr(1, MA.Name, "mnuRPrint") = 0 And InStr(1, MA.Name, "MnuRView") = 0 And InStr(1, MA.Name, "MnuRBatal") = 0 And InStr(1, MA.Name, "mnuRAC") = 0 And InStr(1, MA.Name, "PicSeparatorKanan") = 0 And InStr(1, MA.Name, "ImgOpen") = 0 And InStr(1, MA.Name, "ImgClose") = 0 And InStr(1, MA.Name, "LvMenu") = 0 And InStr(1, MA.Name, "PicRight") = 0 And MA.Name <> "" And Not (TypeOf MA Is Image) And InStr(1, MA.Name, "StatusBar") = 0 And InStr(1, MA.Name, "Timer") = 0 And Not (TypeOf MA Is Label) And Not (TypeOf MA Is ImageList) And Not (TypeOf MA Is Frame) And InStr(1, MA.Name, "Pic") = 0 Then
        .TextMatrix(.Row, 0) = .Row
        .TextMatrix(.Row, 3) = MA.Caption
        If UCase(Mid(MA.Name, 4, 3)) = "KEP" Then
           .TextMatrix(.Row, 4) = 0
        Else
            .TextMatrix(.Row, 4) = 1
        End If
'        MsgBox MA.Name
        .TextMatrix(.Row, 5) = MA.Name
        
        SimpanData (.Row)
        .Rows = .Rows + 1
    End If
Next
End With
If Rscek.State = adStateOpen Then Rscek.Close
Set Rscek.ActiveConnection = Nothing
Set CmSimpan.ActiveConnection = Nothing
End Sub

Private Sub CmdExit_Click()
Unload Me
Set FrmGroupMenu = Nothing
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
                        
                        StrSQL = "Delete from TblGroup_Menu Where IDMenu='" & .TextMatrix(Lrow, 2) & "'"
                        PerintahExecute (StrSQL)
                        .RemoveItem (Lrow)
                        Lrow = Lrow - 1
                End If
                Lrow = Lrow + 1
            Loop
                     
                     
             End If
            For Lrow = 1 To .Rows - 2
                .TextMatrix(Lrow, 0) = Lrow
                SimpanData (Lrow)
            Next
            
         End With

Exit Sub

AdaError:
If ErrConn > 0 Then CN.RollbackTrans
MsgBox err.Description
End Sub

'Private Sub VsFlex_ButtonClick(ByVal RowIndex As Integer)
'    If RowIndex <> VsFlex.Rows - 1 Then
'        FrmGroupPermission.IDGroup = VsFlex.TextMatrix(RowIndex, 1)
'        FrmGroupPermission.NamaGroups = "Permission on Group : " & VsFlex.TextMatrix(RowIndex, 4)
'        FrmGroupPermission.Show 1
'    End If
'End Sub
Private Sub CekHak()

On Error GoTo AdaError
'1 = View, 2 = insert 3 = edit, 4 = delete, 5 = print
If InStr(1, strGroup, "Admin") > 0 Then
    Exit Sub
End If

If Left(StrHakUser, 1) = 1 Then
    Me.Show
Else
    MsgBox "Untuk menggunakan menu ini hubungi Administrator", vbInformation
    Unload Me
    Set FrmGroupMenu = Nothing
End If

If Mid(StrHakUser, 2, 1) = "1" Then 'insert
    VsFlex.EditEnable = True
    CmdAuto.Enabled = True
Else
    VsFlex.EditEnable = False
    CmdAuto.Enabled = False
End If

If Mid(StrHakUser, 4, 1) = "1" Then 'delete
    CmdHapus.Enabled = True
Else
    CmdHapus.Enabled = False
End If
Exit Sub
AdaError:
MsgBox err.Number & Chr(13) & err.Description
End Sub

Private Sub VsFlex_ReachColsEnd(Row As Long)
    Call SimpanData(Row)
End Sub

Private Sub VsFlex_TextKeyPress(KeyAscii As Integer, RowIndex As Long, ColIndex As Long)
'    If VsFlex.TextMatrix(0, ColIndex) = "Group User" Then
        If VsFlex.TextMatrix(RowIndex, 1) <> "" Then Call SimpanData(RowIndex)
'    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
        CmdHapus_Click
ElseIf KeyCode = vbKeyF9 Then
        Unload Me
        Set FrmGroupMenu = Nothing

ElseIf KeyCode = vbKeyEscape Then
    CmdHapus.Enabled = True
    CmdExit.Enabled = True
End If
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

Private Sub SimpanData(ByVal Row As Long)
Dim CmSimpan As New ADODB.Command
Dim RsSimpan As New ADODB.Recordset
Dim SQL As String

On Error GoTo AdaError

If CN.State = adStateClosed Then CN.Open
If Rscek.State = adStateOpen Then Rscek.Close
    Rscek.Open "Select * From TblGroup_Menu Where Nama_Menu = '" & VsFlex.TextMatrix(Row, 5) & "'", CN, adOpenStatic
Debug.Print VsFlex.TextMatrix(Row, 5)
If VsFlex.TextMatrix(Row, 5) = "picSeparator" Or VsFlex.TextMatrix(Row, 5) = "picLeft" Or VsFlex.TextMatrix(Row, 5) = "lvWin" Then Exit Sub
If VsFlex.TextMatrix(Row, 5) = "picLine" Or VsFlex.TextMatrix(Row, 5) = "tmrResize" Or VsFlex.TextMatrix(Row, 5) = "tmrMemStatus" Then Exit Sub
If VsFlex.TextMatrix(Row, 5) = "picContainer" Or VsFlex.TextMatrix(Row, 5) = "picFreeMem" Or VsFlex.TextMatrix(Row, 5) = "Line1" Then Exit Sub
If VsFlex.TextMatrix(Row, 5) = "Line2" Or VsFlex.TextMatrix(Row, 5) = "Skin1" Or VsFlex.TextMatrix(Row, 5) = "Line1" Then Exit Sub

If Rscek.EOF Then
        '---------------------------------
        CmSimpan.ActiveConnection = CN
        SQL = "Insert into TblGroup_Menu (IDMenu,Nama_Menu,Parents,urutan,Menu_Caption) VALUES ('" & Row & "','" & VsFlex.TextMatrix(Row, 5) & "'," & VsFlex.TextMatrix(Row, 4) & "," & Row & ",'" & VsFlex.TextMatrix(Row, 3) & "')"
        CmSimpan.CommandText = SQL
        CmSimpan.Execute

            VsFlex.TextMatrix(Row, 2) = Row
Else
    CmSimpan.ActiveConnection = CN
    CmSimpan.CommandText = "Update TblGroup_Menu set Nama_Menu='" & VsFlex.TextMatrix(Row, 5) & "',Parents=" & VsFlex.TextMatrix(Row, 4) & ",Urutan=" & Row & ",Menu_Caption='" & VsFlex.TextMatrix(Row, 3) & "' Where IDMenu='" & VsFlex.TextMatrix(Row, 2) & "'"
    CmSimpan.Execute


End If



If RsSimpan.State = adStateOpen Then RsSimpan.Close
'If CN.State = adStateOpen Then CN.Close

Exit Sub
AdaError:
MsgBox err.Number & Chr(13) & err.Description
End Sub

Private Sub Showdata()
Dim RsShow As New ADODB.Recordset
Dim SQL As String

On Error GoTo AdaError
If CN.State = adStateClosed Then CN.Open
If RsShow.State = adStateOpen Then RsShow.Close
SQL = "select * from TblGroup_Menu order by Urutan"
RsShow.Open SQL, CN, adOpenStatic

With VsFlex
    .Rows = 1
    .Rows = 2
Do Until RsShow.EOF
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = .Row
    .TextMatrix(.Row, 2) = RsShow!IDMenu
    .TextMatrix(.Row, 3) = RsShow!Menu_Caption
    .TextMatrix(.Row, 4) = RsShow!Parents
    .TextMatrix(.Row, 5) = RsShow!nama_menu
    RsShow.MoveNext
    .Rows = .Rows + 1
Loop
End With
If RsShow.State = adStateOpen Then RsShow.Close
Set RsShow.ActiveConnection = Nothing
'If CN.State = adStateOpen Then CN.Close

Exit Sub
AdaError:
MsgBox err.Number & Chr(13) & err.Description

End Sub
Sub Setgrid()
With VsFlex
    .Rows = 1
    .Rows = 2
    .Cols = 6
    .TextMatrix(0, 0) = "No"
    .TextMatrix(0, 1) = "Do"
    .TextMatrix(0, 2) = "ID"
    .TextMatrix(0, 3) = "Caption Menu"
    .TextMatrix(0, 4) = "Parents"
    .TextMatrix(0, 5) = "Nama Menu"
    .ColWidth(0) = 300
    .ColWidth(1) = 400
    .ColWidth(2) = 0
    .ColWidth(3) = 2500
    .ColWidth(4) = 1000
    .ColWidth(5) = 2050
    .ColDataType(1) = flexDTBoolean
    For Lrow = 1 To .Cols - 1
        .FixedAlignment(Lrow) = flexAlignCenterCenter
    Next

End With
End Sub

Private Sub VSFlex_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SimpanData (VsFlex.Row)
End Sub
