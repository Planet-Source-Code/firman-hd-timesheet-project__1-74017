VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTotal 
   Caption         =   "Total Hari"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   11745
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fix"
      Height          =   495
      Left            =   7080
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20643842
      CurrentDate     =   39940
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20643843
      CurrentDate     =   39931
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20643843
      CurrentDate     =   39931
   End
   Begin VSFlex8Ctl.VSFlexGrid VSProUmum 
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   6735
      _cx             =   11880
      _cy             =   2990
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
      SelectionMode   =   1
      GridLines       =   2
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
      FormatString    =   $"FrmTotal.frx":0000
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
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20643842
      CurrentDate     =   39940
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   5895
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   11535
      _cx             =   20346
      _cy             =   10398
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
      SelectionMode   =   1
      GridLines       =   2
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
      FormatString    =   $"FrmTotal.frx":00DF
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
   Begin VB.Label LKeluar 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dari "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "S.D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "FrmTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalHari As Double

Private Sub Command1_Click()
Dim Hari As String
Dim RsTS As New ADODB.Recordset
Dim x, Selisih As Integer
Dim SaldoUmum As Double
Command1.Enabled = False
If RsTS.State = adStateOpen Then RsTS.Close
RsTS.Open "select * From TblTimesheet Where Hari = 'Kerja' And Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And '" & Format(DTPicker2, "MM/dd/yyyy") & "'  And kd_divisi = '" & Text1 & "' Order By Tanggal ASC", CN, adOpenStatic
Set VSFlexGrid1.DataSource = RsTS
 
Selisih = DateDiff("d", DTPicker1, DTPicker2)
TotalHari = 0

For x = 0 To Selisih
   If x > 0 Then DTPicker1 = DateAdd("d", 1, DTPicker1)
   Hari = Format(DTPicker1, "ddd")
   If Hari <> "Sat" And Hari <> "Sun" And Hari <> "Sabtu" And Hari <> "Minggu" Then
       StatusHari = "Kerja"
        StrSQL = "select tanggallibur from kalender " & _
            "where tanggallibur = '" & Format(DTPicker1, "MM/dd/yyyy") & "'"
        If Rscek.State = adStateOpen Then Rscek.Close
        Rscek.Open StrSQL, CN, adOpenStatic
        If Rscek.EOF Then
           TotalHari = TotalHari + 1
        End If
    End If
Next
Getumum
TotalHari = TotalHari * 8
For Selisih = 1 To VSFlexGrid1.Rows - 1
        VSFlexGrid1.TextMatrix(Selisih, 0) = Selisih
        VSFlexGrid1.TextMatrix(Selisih, 17) = 0
    With VSProUmum
        For x = 1 To .Rows - 1
            If .TextMatrix(x, 1) = VSFlexGrid1.TextMatrix(Selisih, 3) Then
               VSFlexGrid1.TextMatrix(Selisih, 19) = TotalHari
               VSFlexGrid1.TextMatrix(Selisih, 20) = .TextMatrix(x, 2)
               DTPicker3 = Format(VSFlexGrid1.TextMatrix(Selisih, 10), "HH:mm")
               DTPicker4 = Format(VSFlexGrid1.TextMatrix(Selisih, 11), "HH:mm")
'               DTPicker4 = "00:00"
              SaldoUmum = DateDiff("n", DTPicker3, DTPicker4)
              SaldoUmum = Abs(SaldoUmum) / 60
'              VSFlexGrid1.TextMatrix(Selisih, 21) = VSFlexGrid1.TextMatrix(Selisih, 19) - SaldoUmum
               Exit For
            End If
        Next
    End With
Next
'MsgBox TotalHari & "," & X

StrSQL = "Update Tbltimesheet Set TotalKerja = '" & TotalHari & "'"
' DTPicker1.Value = DateSerial(Year(Now), Month(Now), 1)
'    DTPicker1.Value = DateAdd("M", 0, DTPicker1.Value)
Command1.Enabled = True
End Sub
Sub Getumum()
Dim RsUmum As New ADODB.Recordset
If RsUmum.State = adStateOpen Then RsUmum.Close
RsUmum.Open "Select kd_divisi,Kode,Nama From Project Where Nama Like '%Umum%' Order By ID", CN, adOpenStatic
Set VSProUmum.DataSource = RsUmum
End Sub

Private Sub Command2_Click()
Dim x, xx As Integer
Dim status As Boolean
Dim Saldo, SaldoUmum As Double
Dim NilaiSaldo As String
With VSFlexGrid1
 If MsgBox("Apakah Anda yakin ingin menghapus Data ?", vbQuestion + vbYesNo, "Konfirmasi hapus") = vbNo Then
    Exit Sub
 End If
     StrSQL = "Update tbltimesheet Set TotalKerja = '" & .TextMatrix(1, 19) & "' Where Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And '" & Format(DTPicker2, "MM/dd/yyyy") & "'"
     CN.Execute StrSQL
    For x = 1 To .Rows - 1
        Command1.Caption = x
'        If Rscek.State = adStateOpen Then Rscek.Close
'        StrSQL = "Select * From Tbltimesheet Where NIP = '" & .TextMatrix(X, 2) & "' And Tanggal = '" & Format(.TextMatrix(X, 9), "MM/dd/yyyy") & "'"
'        Rscek.Open StrSQL, CN, adOpenStatic
'        status = False
'        For xx = 1 To VSProUmum.Rows - 1
'            If VSProUmum.TextMatrix(xx, 1) = .TextMatrix(X, 3) And .TextMatrix(X, 7) = VSProUmum.TextMatrix(xx, 2) Then
'
'               status = True
'               Exit For
'            End If
'        Next
'
'         If Trim(.TextMatrix(X, 15)) = "Timesheet" Then
'        Saldo = .TextMatrix(X, 19)
'        If status = True Then
'            DTPicker4 = Format(VSFlexGrid1.TextMatrix(X, 11), "HH:mm")
'        SaldoUmum = DateDiff("n", DTPicker3, DTPicker4)
'        SaldoUmum = Abs(SaldoUmum) / 60
'
'              If SaldoUmum = 9 Then SaldoUmum = 8
'              NilaiSaldo = IIf(IsNull(Rscek!SaldoUmum), "", Rscek!SaldoUmum)
'              If Trim(NilaiSaldo) = "" Then
'                 Saldo = Saldo - SaldoUmum
'              Else
'                Saldo = Rscek!SaldoUmum - SaldoUmum
''                Saldo = .TextMatrix(X, 19)
'              End If
'               DTPicker3 = Format(VSFlexGrid1.TextMatrix(X, 10), "HH:mm")
'
'        VSFlexGrid1.TextMatrix(X, 18) = CDbl(0)
'        VSFlexGrid1.TextMatrix(X, 18) = CDbl(SaldoUmum)
''        VSFlexGrid1.TextMatrix(X, 17) = VSFlexGrid1.TextMatrix(X, 17) + CDbl(SaldoUmum)
'              VSFlexGrid1.TextMatrix(X, 21) = CDbl(Saldo)
''       Debug.Print Saldo
        
            StrSQL = "Update tbltimesheet Set TotalKerja = '" & .TextMatrix(x, 19) & "',ProjectUmum = '" & .TextMatrix(x, 20) & "' Where NIP = '" & .TextMatrix(x, 2) & "' And Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And '" & Format(DTPicker2, "MM/dd/yyyy") & "'"
            
            CN.Execute StrSQL
'         End If
'        End If
    Next
End With
End Sub

Private Sub Form_Load()
 DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker1.Value = DateSerial(Year(Now), Month(Now), 1)
    DTPicker1.Value = DateAdd("M", 0, DTPicker1.Value)
    DTPicker2.Value = DateSerial(Year(Now), Month(Now), 1)
    DTPicker2.Value = DateAdd("M", 1, DTPicker2.Value) - 1
    DTPicker1.CustomFormat = "dd/MMM/yyyy"
    DTPicker2.CustomFormat = "dd/MMM/yyyy"
End Sub

Private Sub Form_Resize()
With VSFlexGrid1
     .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left

End With
End Sub
