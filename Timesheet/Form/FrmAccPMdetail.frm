VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmAccPMdetail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detail Timesheet"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11445
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   11445
      TabIndex        =   0
      Top             =   6540
      Width           =   11445
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
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlex 
      Height          =   4815
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8535
      _cx             =   15055
      _cy             =   8493
      Appearance      =   1
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
      AutoSearch      =   0
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
         Left            =   720
         TabIndex        =   3
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52690946
         CurrentDate     =   39940
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52690946
         CurrentDate     =   39940
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   4935
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8535
      _cx             =   15055
      _cy             =   8705
      Appearance      =   1
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
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14737632
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
      AutoSearch      =   0
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
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FrmAccPMdetail.frx":0000
      Top             =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4815
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   8535
      _cx             =   15055
      _cy             =   8493
      Appearance      =   1
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
      AutoSearch      =   0
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
End
Attribute VB_Name = "FrmAccPMdetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdClose_Click()
Unload Me
End Sub

Public Function ShowDetail1(tgl1 As Date, tgl2 As Date, NIP As String, Slot As String)

Dim Rscek As New ADODB.Recordset
Dim Rs As New ADODB.Recordset
Dim X As Integer
VsFlex.Rows = 1
fg.Rows = 1
CN.Execute "delete from timesheet2 where id = '9' or id ='10'"
If Rscek.State = adStateOpen Then Rscek.Close

StrSQL = "SELECT Timesheet2.Tanggal,Timesheet2.ID AS Jam,Timesheet2.Slot AS Kode,Timesheet2.NIP, Karyawan.Nama,Timesheet2.Status  FROM Karyawan INNER JOIN Timesheet2 ON Karyawan.NIP = Timesheet2.NIP"
StrSQL = StrSQL & " WHERE TANGGAL BETWEEN '" & Format(tgl1, "yyyy-MM-dd") & "' AND '" & Format(tgl2, "yyyy-MM-dd") & "'"
StrSQL = StrSQL & " And Timesheet2.NIP = '" & NIP & "'"
StrSQL = StrSQL & " And Timesheet2.kd_divisi = '" & KodeDivisi & "' Order By Timesheet2.Tanggal,Timesheet2.NIP,Timesheet2.status,Timesheet2.ID Asc"
Rscek.Open StrSQL, CN, adOpenStatic
Set fg.DataSource = Rscek

With fg
.TextMatrix(0, 0) = "No"
fg.ColWidth(1) = 1300
fg.ColWidth(2) = 1200
fg.ColWidth(5) = 3000
fg.ColWidth(0) = 500
fg.ColDataType(1) = flexDTDate
.ColFormat(1) = "dd/MM/yyyy"
For X = 1 To .Rows - 1
    Jam = .TextMatrix(X, 2)
    Select Case Jam
        Case 1
             .TextMatrix(X, 2) = "08:00-08:30"
        Case 2
            .TextMatrix(X, 2) = "08:30-09:00"
        Case 3
            .TextMatrix(X, 2) = "09:00-09:30"
        Case 4
            .TextMatrix(X, 2) = "09:30-10:00"
        Case 5
            .TextMatrix(X, 2) = "10:00-10:30"
        Case 6
            .TextMatrix(X, 2) = "10:30-11:00"
        Case 7
            .TextMatrix(X, 2) = "11:00-11:30"
        Case 8
            .TextMatrix(X, 2) = "11:30-12:00"
        Case 9
              .TextMatrix(X, 2) = "ISTIRAHAT"
        Case 10
             .TextMatrix(X, 2) = "ISTIRAHAT"
        Case 11
             .TextMatrix(X, 2) = "13:00-13:30"
        Case 12
            .TextMatrix(X, 2) = "13:30-14:00"
        Case 13
            .TextMatrix(X, 2) = "14:00-14:30"
        Case 14
            .TextMatrix(X, 2) = "14:30-15:00"
        Case 15
            .TextMatrix(X, 2) = "15:00-15:30"
        Case 16
            .TextMatrix(X, 2) = "15:30-16:00"
        Case 17
            .TextMatrix(X, 2) = "16:00-16:30"
        Case 18
            .TextMatrix(X, 2) = "16:30-17:00"
    End Select
    fg.TextMatrix(X, 0) = X
Next
  
End With

If Rs.State = adStateOpen Then Rs.Close
StrSQL = "SELECT Timesheet2.ID AS Jam,Timesheet2.Tanggal,Timesheet2.Slot AS Kode,Timesheet2.NIP, Karyawan.Nama,timesheet2.id,timesheet2.Status FROM Karyawan INNER JOIN Timesheet2 ON Karyawan.NIP = Timesheet2.NIP"
StrSQL = StrSQL & " WHERE TANGGAL BETWEEN '" & Format(tgl1, "yyyy-MM-dd") & "' AND '" & Format(tgl2, "yyyy-MM-dd") & "'"
StrSQL = StrSQL & "And Timesheet2.NIP = '" & NIP & "'"
StrSQL = StrSQL & " And Timesheet2.kd_divisi = '" & KodeDivisi & "' Order By Timesheet2.Tanggal,Timesheet2.NIP,Timesheet2.status,Timesheet2.ID Asc"

'StrSQL = "SELECT Timesheet2.Tanggal,Timesheet2.ID AS Jam,Timesheet2.Slot AS Kode,Timesheet2.NIP, Karyawan.Nama  FROM Karyawan INNER JOIN Timesheet2 ON Karyawan.NIP = Timesheet2.NIP"
'StrSQL = StrSQL & " WHERE TANGGAL BETWEEN '" & Format(tgl1, "yyyy-MM-dd") & "' AND '" & Format(tgl2, "yyyy-MM-dd") & "' And Timesheet2.Status = 'Actual'"
'StrSQL = StrSQL & "And Timesheet2.Slot NOT LIKE '%" & Slot & "%' And Timesheet2.NIP = '" & NIP & "'"
'StrSQL = StrSQL & " And Timesheet2.kd_divisi = '" & KodeDivisi & "' Order By Timesheet2.Tanggal,Timesheet2.NIP,Timesheet2.ID Asc"

Rs.Open StrSQL, CN, adOpenStatic
Set VSFlexGrid1.DataSource = Rs
VSFlexGrid1.Refresh
With VSFlexGrid1
.TextMatrix(0, 0) = "No"
.ColWidth(1) = 1000
.ColWidth(2) = 800
.ColWidth(0) = 500

For X = 1 To .Rows - 1
    Jam = .TextMatrix(X, 1)
    Select Case Jam
        Case 1
             .TextMatrix(X, 1) = "08:00-08:30"
        Case 2
            .TextMatrix(X, 1) = "08:30-09:00"
        Case 3
            .TextMatrix(X, 1) = "09:00-09:30"
        Case 4
            .TextMatrix(X, 1) = "09:30-10:00"
        Case 5
            .TextMatrix(X, 1) = "10:00-10:30"
        Case 6
            .TextMatrix(X, 1) = "10:30-11:00"
        Case 7
            .TextMatrix(X, 1) = "11:00-11:30"
        Case 8
            .TextMatrix(X, 1) = "11:30-12:00"
        Case 9
              .TextMatrix(X, 1) = "ISTIRAHAT"
        Case 10
             .TextMatrix(X, 1) = "ISTIRAHAT"
        Case 11
             .TextMatrix(X, 1) = "13:00-13:30"
        Case 12
            .TextMatrix(X, 1) = "13:30-14:00"
        Case 13
            .TextMatrix(X, 1) = "14:00-14:30"
        Case 14
            .TextMatrix(X, 1) = "14:30-15:00"
        Case 15
            .TextMatrix(X, 1) = "15:00-15:30"
        Case 16
            .TextMatrix(X, 1) = "15:30-16:00"
        Case 17
            .TextMatrix(X, 1) = "16:00-16:30"
        Case 18
            .TextMatrix(X, 1) = "16:30-17:00"
    End Select
    .TextMatrix(X, 0) = X
Next
End With
Showdata
End Function
Private Sub Command2_Click()
Dim X As String
On Error GoTo AdaError
If Option3.Value = True Then
    fg.SaveGrid "C:\Rekaptimesheet.xls", flexFileExcel, True
'    Shell PathOffice & "C:\Rekaptimesheet.csv", vbNormalFocus
    X = ShellExecute(Me.hwnd, "open", "C:\Rekaptimesheet.xls", vbNullString, "C:\RekapTimesheet.xls", 1)
ElseIf Option4.Value = True Then
    VsFlex.ColWidth(12) = 0
    VsFlex.SaveGrid "C:\Rekaptimesheet.xls", flexFileExcel, True
'    Shell PathOffice & "C:\Rekaptimesheet.xls", vbNormalFocus
X = ShellExecute(Me.hwnd, "open", "C:\RekapTimesheet.xls", vbNullString, "C:\RekapTimesheet.xls", 1)
End If

Exit Sub
AdaError:
MsgBox err.Description
End Sub

Private Sub fg_Click()
Dim J As String
With fg
If .TextMatrix(.Row, 3) <> "" And .Col = 3 Then
    J = FrmPM.Showdata(.TextMatrix(.Row, 3), KodeDivisi)
    FrmPM.show vbModal
End If
End With
End Sub

Private Sub Form_Load()
    Setgrid
'    AddKaryawan
    If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
      Skin1.ApplySkin hwnd
    End If
    VsFlex.FrozenCols = 2
End Sub
Private Sub Form_Resize()
 On Error Resume Next
    With fg
    .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture3.Height
    End With
    With VsFlex
    .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture3.Height
    End With
    With VSPlan
    .Move .Left, VsFlex.Height, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Me.Height / 2
    End With
    CmdClose.Width = Picture3.Width
End Sub
Sub Setgrid()
Dim i As Integer
With VsFlex

    .Rows = 1
'    .Rows = 2
    .Cols = 22
    .TextMatrix(0, 0) = "No"
    .TextMatrix(0, 1) = "Tanggal"
    .TextMatrix(0, 2) = "NIP"
    .TextMatrix(0, 3) = "Nama"
    .TextMatrix(0, 4) = "08:00-08:30"
    .TextMatrix(0, 5) = "08:30-09:00"
    .TextMatrix(0, 6) = "09:00-09:30"
    .TextMatrix(0, 7) = "09:30-10:00"
    .TextMatrix(0, 8) = "10:00-10:30"
    .TextMatrix(0, 9) = "10:30-11:00"
    .TextMatrix(0, 10) = "11:00-11:30"
    .TextMatrix(0, 11) = "11:30-12:00"
    .TextMatrix(0, 12) = "12:00-13:00"
    .TextMatrix(0, 13) = "13:00-13:30"
    .TextMatrix(0, 14) = "13:30-14:00"
    .TextMatrix(0, 15) = "14:00-14:30"
    .TextMatrix(0, 16) = "14:30-15:00"
    .TextMatrix(0, 17) = "15:00-15:30"
    .TextMatrix(0, 18) = "15:30-16:00"
    .TextMatrix(0, 19) = "16:00-16:30"
    .TextMatrix(0, 20) = "16:30-17:00"
    .TextMatrix(0, 21) = "Status"
'    .TextMatrix(0, 22) = "Lembur"
'    .TextMatrix(0, 23) = "T.Lembur"
    
    .ColWidth(0) = 500
    For i = 4 To 21
       .ColWidth(i) = 1000
    Next
    .ColDataType(1) = flexDTDate
    .ColFormat(1) = "dd/MM/yyyy"
    .ColWidth(12) = 0
End With
End Sub
Sub Showdata()
On Error GoTo AdaError
Dim i, J As Integer
Dim Jam As Integer
With VsFlex
.Rows = 1
.Rows = 2
.ColWidth(1) = 1200
.ColWidth(2) = 800
For i = 1 To VSFlexGrid1.Rows - 1
    Jam = VSFlexGrid1.TextMatrix(i, 6)

    If i <> VSFlexGrid1.Rows - 1 Then
        If VSFlexGrid1.TextMatrix(i, 4) = VSFlexGrid1.TextMatrix(i + 1, 4) Then
           
            If VSFlexGrid1.TextMatrix(i, 2) = VSFlexGrid1.TextMatrix(i + 1, 2) Then
                J = Tampil(Jam, i, VSFlexGrid1.TextMatrix(i, 7))
                If Jam = 18 Then .Rows = .Rows + 1
                
            Else
                 J = Tampil(Jam, i, VSFlexGrid1.TextMatrix(i, 7))
                .Rows = .Rows + 1
           End If
        Else
                J = Tampil(Jam, i, VSFlexGrid1.TextMatrix(i, 7))
                .Rows = .Rows + 1
        End If
    Else
        If VSFlexGrid1.TextMatrix(i, 4) <> .TextMatrix(.Rows - 1, 2) Then
             J = Tampil(Jam, i, VSFlexGrid1.TextMatrix(i, 7))
        Else
             J = Tampil(Jam, i, VSFlexGrid1.TextMatrix(i, 7))
      End If
    End If
   
Next
''    .Cols = .Cols + 2
'     Dim RsLembur As New ADODB.Recordset
'     Dim JamAwal, JamAkhir As String
    For i = 1 To .Rows - 1
          .TextMatrix(i, 0) = i
'        If RsLembur.State = adStateOpen Then RsLembur.Close
'
'            StrSQL = "Select * From Lembur1 Where Tanggal='" & Format(.TextMatrix(I, 1), "yyyy-MM-dd") & "' And NIP = '" & .TextMatrix(I, 2) & "' And noproject Not LIKE '%*%'"
'         RsLembur.Open StrSQL, CN, adOpenStatic
'        If Not RsLembur.EOF Then
'
'            .TextMatrix(I, 22) = RsLembur!NoProject
'            JamAwal = Replace(RsLembur!Jam_Awal, ".", ":")
'            JamAkhir = Replace(RsLembur!Jam_Akhir, ".", ":")
'            DTPicker3 = CDate(JamAwal)
'            DTPicker4 = CDate(JamAkhir)
'            .TextMatrix(I, 23) = Format(DTPicker4 - DTPicker3, "HH:mm")
'        End If
    Next
End With
Exit Sub
AdaError:
MsgBox err.Description & "-" & i
End Sub

Function Tampil(ByVal Jam As Integer, ByVal i As Integer, status As String)
Dim Project As String
With VsFlex
    .TextMatrix(.Rows - 1, 1) = VSFlexGrid1.TextMatrix(i, 2)
    .TextMatrix(.Rows - 1, 2) = VSFlexGrid1.TextMatrix(i, 4)
    .TextMatrix(.Rows - 1, 3) = VSFlexGrid1.TextMatrix(i, 5)
    Project = VSFlexGrid1.TextMatrix(i, 3)
    If Project = "Telat*" Then Project = "Telat"
  Select Case Jam
      Case 1
          .TextMatrix(.Rows - 1, 4) = Project
      Case 2
          .TextMatrix(.Rows - 1, 5) = Project
      Case 3
          .TextMatrix(.Rows - 1, 6) = Project
      Case 4
          .TextMatrix(.Rows - 1, 7) = Project
      Case 5
          .TextMatrix(.Rows - 1, 8) = Project
      Case 6
          .TextMatrix(.Rows - 1, 9) = Project
      Case 7
          .TextMatrix(.Rows - 1, 10) = Project
      Case 8
          .TextMatrix(.Rows - 1, 11) = Project
      Case 9, 10
          .TextMatrix(.Rows - 1, 12) = "ISTIRAHAT"
      Case 11
          .TextMatrix(.Rows - 1, 13) = Project
      Case 12
          .TextMatrix(.Rows - 1, 14) = Project
      Case 13
          .TextMatrix(.Rows - 1, 15) = Project
      Case 14
          .TextMatrix(.Rows - 1, 16) = Project
      Case 15
          .TextMatrix(.Rows - 1, 17) = Project
      Case 16
          .TextMatrix(.Rows - 1, 18) = Project
      Case 17
          .TextMatrix(.Rows - 1, 19) = Project
      Case 18
          .TextMatrix(.Rows - 1, 20) = Project
  End Select
        .TextMatrix(.Rows - 1, 21) = status
End With
End Function

Private Sub Form_Unload(Cancel As Integer)
Set FrmAccPMdetail = Nothing
End Sub


Private Sub VSFlex_Click()
Dim J As String
With VsFlex
If .Col >= 4 And .Col <= 21 And .TextMatrix(.Row, .Col) <> "" And .TextMatrix(.Row, .Col) <> "Telat" Then
    J = FrmPM.Showdata(.TextMatrix(.Row, .Col), KodeDivisi)
    FrmPM.show vbModal
End If
End With
End Sub


