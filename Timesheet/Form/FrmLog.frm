VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmLog 
   Caption         =   "Log User"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   12360
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   12360
      TabIndex        =   10
      Top             =   5535
      Width           =   12360
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
         TabIndex        =   11
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   12330
      TabIndex        =   0
      Top             =   0
      Width           =   12360
      Begin VB.TextBox Txtnip 
         Height          =   285
         Left            =   960
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Last Login"
         Height          =   375
         Left            =   4920
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Print out"
         Height          =   375
         Left            =   6960
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   4920
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Export To Sheet"
         Height          =   375
         Left            =   6960
         TabIndex        =   1
         Top             =   120
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   720
         TabIndex        =   4
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
         Format          =   53542915
         CurrentDate     =   39931
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2880
         TabIndex        =   5
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
         Format          =   53542915
         CurrentDate     =   39931
      End
      Begin VSFlex8Ctl.VSFlexGrid CboFlex 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   600
         Width           =   3465
         _cx             =   6112
         _cy             =   556
         Appearance      =   0
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
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   16777215
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmLog.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   60
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
         Editable        =   2
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "NIP"
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
         TabIndex        =   15
         Top             =   960
         Width           =   1215
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
         Left            =   240
         TabIndex        =   9
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
         Left            =   2400
         TabIndex        =   8
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Modul"
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
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "FrmLog.frx":0029
      Top             =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4095
      Left            =   0
      TabIndex        =   12
      Top             =   1320
      Width           =   8295
      _cx             =   14631
      _cy             =   7223
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
      BackColorSel    =   14737632
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
      FormatString    =   $"FrmLog.frx":025D
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
Attribute VB_Name = "FrmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim RsTS As New ADODB.Recordset
Dim LogSQL As String

If Trim(cboFlex) = "" Then
    LogSQL = "Select * From TblLog_user Where Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And  '" & Format(DTPicker2, "MM/dd/yyyy") & "' Order By ID Desc"
Else
   If cboFlex = "Timesheet" Then
      If Txtnip = "" Then
          LogSQL = "Select NIP,NoProject,Tanggal,Last_update As Tanggal_Isi,Jamawal,JamAkhir,Masuk From Tbltimesheet Where Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And  '" & Format(DTPicker2, "MM/dd/yyyy") & "'   Order By IDtimesheet Desc"
      Else
          LogSQL = "Select NIP,NoProject,Tanggal,Last_update As Tanggal_Isi,Jamawal,JamAkhir,Masuk From Tbltimesheet Where Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And  '" & Format(DTPicker2, "MM/dd/yyyy") & "' And NIP = '" & Txtnip & "' Order By IDtimesheet Desc"

      End If
   ElseIf cboFlex = "Delete User" Then
        LogSQL = "Select * From vuser where status = 14 order by nip"
'        For Lrow = 1 To fg.Rows - 1
'        Next
   Else
    LogSQL = "Select * From TblLog_user Where Tanggal Between '" & Format(DTPicker1, "MM/dd/yyyy") & "' And  '" & Format(DTPicker2, "MM/dd/yyyy") & "' And Modul = '" & cboFlex & "' Order By ID Desc"
   End If
End If
If RsTS.State = adStateOpen Then RsTS.Close
RsTS.Open LogSQL, CN, adOpenDynamic, adLockOptimistic

Set fg.DataSource = RsTS
fg.DataRefresh
fg.Refresh
With fg
    For Lrow = 1 To .Rows - 1
        .TextMatrix(Lrow, 0) = Lrow
        
       If cboFlex = "Delete User" Then CN.Execute "Delete From tbldata_user Where NIp = '" & fg.TextMatrix(Lrow, 1) & "'"

    Next
End With
End Sub

Private Sub Command3_Click()
On Error GoTo Adaerror
If fg.Rows > 1 Then fg.PrintGrid "Log Timesheet " & DTPicker1.Value & " S/D " & DTPicker2.Value, 2, 2, 900, 500
Exit Sub
Adaerror:
MsgBox err.Description

End Sub

Private Sub Command4_Click()
If Rscek.State = adStateOpen Then Rscek.Close
Rscek.Open "SELECT NIP,NamaUser,NamaDivisi,last_login,StatusLogin,ComputerName,IPAddress,UserComp,VersiApp From tbldata_user Order By Last_login DESC", CN, adOpenStatic
Set fg.DataSource = Rscek
With fg
    For Lrow = 1 To .Rows - 1
        .TextMatrix(Lrow, 0) = Lrow
    Next
End With
End Sub

Private Sub Form_Load()
    AddModul
 DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker1.Value = Date
    DTPicker1.Value = DateSerial(Year(Now), Month(Now), 1)
    DTPicker1.Value = DateAdd("M", 0, DTPicker1.Value)
    DTPicker2.Value = DateSerial(Year(Now), Month(Now), 1)
    DTPicker2.Value = DateAdd("M", 1, DTPicker2.Value) - 1
    DTPicker1.CustomFormat = "dd/MMM/yyyy"
    DTPicker2.CustomFormat = "dd/MMM/yyyy"
    
     If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
      Skin1.ApplySkin hwnd
    End If
End Sub
Private Sub AddModul()

    Dim Cboid     As String
    Dim cboid1    As String
If Rscek.State = adStateOpen Then Rscek.Close
    Cboid = vbNullString
    cboid1 = vbNullString
     StrSQL = "Select Distinct Modul From tbllog_user Order By Modul"
    Rscek.Open StrSQL, CN, adOpenStatic
    cboid1 = " "
    Do Until Rscek.EOF
      Cboid = "|" & Rscek("Modul")
      cboid1 = cboid1 + Cboid
      Rscek.MoveNext
    Loop
    cboid1 = cboid1 + "|Delete User"
    cboFlex.ColComboList(0) = cboid1
    cboFlex.CellAlignment = flexAlignLeftCenter
End Sub
Private Sub Form_Resize()
    On Error Resume Next
        CmdClose.Width = Me.Width - 100
        With fg
             .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left - Picture2.Height

        End With
       
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmLapBiayaProject = Nothing
End Sub
