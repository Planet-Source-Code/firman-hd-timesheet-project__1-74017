VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10125
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   10065
      TabIndex        =   0
      Top             =   0
      Width           =   10125
      Begin VB.CommandButton Command1 
         Caption         =   "Load Login"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Create Login"
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   4080
         TabIndex        =   3
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
         Format          =   48889859
         CurrentDate     =   39931
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlex 
      Height          =   9975
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   7215
      _cx             =   12726
      _cy             =   17595
      Appearance      =   1
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
      SelectionMode   =   1
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
Attribute VB_Name = "FrmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Rscek.State = adStateOpen Then Rscek.Close
Rscek.Open "Select Distinct NIP,Type From Login Order By NIP", CN, adOpenStatic
Set VSFlex.DataSource = Rscek
VSFlex.TextMatrix(0, 0) = VSFlex.Rows
If Rscek.State = adStateOpen Then Rscek.Close
Rscek.Open "Select NIP From karyawan where Len(NIP) < 5  Order By NIP", CN, adOpenStatic
Set VSFlex.DataSource = Rscek
VSFlex.TextMatrix(0, 0) = VSFlex.Rows
For Lrow = 1 To VSFlex.Rows - 1
     VSFlex.TextMatrix(Lrow, 0) = Lrow
Next
End Sub

Private Sub Command2_Click()
Dim UserPassword, NamaUser As String
Dim IDGroup As String
On Error GoTo AdaError
With VSFlex
'Untuk Input User Baru
For Lrow = 1 To .Rows - 1
    Command1.Caption = Lrow & " - " & .TextMatrix(Lrow, 1)
'    IDGroup = .TextMatrix(Lrow, 2)
    
    If Rscek.State = adStateOpen Then Rscek.Close
    Rscek.Open "Select * From tbldata_user Where NIP ='" & .TextMatrix(Lrow, 1) & "'", CN, adOpenStatic
    If .TextMatrix(Lrow, 1) = "" Then Exit For
    If Rscek.EOF Then
'        StrSQL = "Insert into tbldata_user (NIP,Password,Type,StatusLogin) VALUES ('" & .TextMatrix(Lrow, 1) & "','" & Enc.EncryptString(.TextMatrix(Lrow, 1)) & "','" & IDGroup & "',0)"
          StrSQL = "Insert into tbldata_user (NIP,Password,Type,StatusLogin) VALUES ('" & .TextMatrix(Lrow, 1) & "','" & Enc.EncryptString(.TextMatrix(Lrow, 1)) & "','User',0)"

        CN.Execute StrSQL
    End If
Next
End With
Exit Sub
AdaError:
Call Error
End Sub
Sub Error()
Dim UserPassword, NamaUser As String
Dim IDGroup As String
On Error GoTo AdaError
With VSFlex
For Lrow = 1 To .Rows - 1
    Command1.Caption = Lrow & " - " & .TextMatrix(Lrow, 1)
'    IDGroup = User
    If Rscek.State = adStateOpen Then Rscek.Close
    Rscek.Open "Select * From tbldata_user Where NIP ='" & .TextMatrix(Lrow, 1) & "'", CN, adOpenStatic
    If Rscek.EOF Then
             StrSQL = "Insert into tbldata_user (NIP,Password,Type,StatusLogin) VALUES ('" & .TextMatrix(Lrow, 1) & "','" & Enc.EncryptString(.TextMatrix(Lrow, 1)) & "','User',0)"

        CN.Execute StrSQL
    End If
Next
End With
Exit Sub
AdaError:
Command2_Click
End Sub
Private Sub Form_Resize()
 On Error Resume Next
    VSFlex.Height = Me.Height - 1500
End Sub
