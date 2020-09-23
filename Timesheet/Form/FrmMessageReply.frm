VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Begin VB.Form FrmMessageReply 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5385
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4635
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin ACTIVESKINLibCtl.Skin Skin2 
         Left            =   4080
         OleObjectBlob   =   "FrmMessageReply.frx":0000
         Top             =   4200
      End
      Begin VB.TextBox TxtDate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   " "
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox TxtSubjek 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Text            =   " "
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox TxtFrom 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   " "
         Top             =   120
         Width           =   4335
      End
      Begin VB.TextBox TxtMessage 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "FrmMessageReply.frx":0234
         Top             =   1920
         Width           =   5175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Reply"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Close"
         Height          =   495
         Left            =   2040
         TabIndex        =   1
         Top             =   3960
         Width           =   1455
      End
      Begin VSFlex8Ctl.VSFlexGrid TxtTo 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   480
         Width           =   4305
         _cx             =   7594
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
         BackColorSel    =   16777215
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
         FormatString    =   $"FrmMessageReply.frx":0236
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
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3000
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         TabIndex        =   11
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         TabIndex        =   10
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
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
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
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
         TabIndex        =   8
         Top             =   1560
         Width           =   855
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2760
      OleObjectBlob   =   "FrmMessageReply.frx":025F
      Top             =   1920
   End
End
Attribute VB_Name = "FrmMessageReply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function Showdata(NIP As String, Tanggal As Date)
StrSQL = "SELECT tblEmail.NIP, tblEmail.Nama, tblEmail.NIPTo,"
StrSQL = StrSQL & " Karyawan.Nama AS NamaTo, tblEmail.Tanggal,"
StrSQL = StrSQL & " tblEmail.Subjeck, tblEmail.Email, tblEmail.StatusBaca,tblEmail.StatusBalas"
StrSQL = StrSQL & " FROM tblEmail INNER JOIN Karyawan ON tblEmail.NIPTo = Karyawan.NIP Where Tblemail.NIP = '" & NIP & "' And Tanggal ='" & Tanggal & "' And nipto = '" & StrUser & "'"
If Rscek.State = adStateOpen Then Rscek.Close
Rscek.Open StrSQL, CN, adOpenStatic
If Not Rscek.EOF Then
    TxtFrom.Tag = Rscek!NIP
    TxtFrom = Rscek!Nama
    TxtSubjek = Rscek!Subjeck
    TxtTo.Text = Rscek!Namato
    TxtTo.Tag = Rscek!NIPto
    TxtMessage = Rscek!Email
    TxtDate = Rscek!Tanggal
End If
End Function

Private Sub Command1_Click()
Dim StrTO, StrNIpTO As String
On Error GoTo Adaerror
If Command1.Caption = "&Reply" Then
    StrFrom = TxtTo.Text
    StrNIpTO = TxtTo.Tag
    TxtTo = TxtFrom
    TxtTo.Tag = TxtFrom.Tag
    TxtFrom = StrFrom
    TxtFrom.Tag = StrNIpTO
    TxtSubjek = "Re: " & TxtSubjek
    TxtMessage.Text = ""
    TxtMessage.SetFocus
    TxtDate = Now
ElseIf Command1.Caption = "Send Message" Then
    If Trim(TxtMessage) = "" Then
        MsgBox "Pesan Masih Kosong", vbCritical
        TxtMessage.SetFocus
        Exit Sub
    End If
 
    StrSQL = "Insert Into tblEmail(NIP,Nama,NIPTo,Tanggal,Subjeck,Email,StatusBaca,StatusBalas)Values"
    StrSQL = StrSQL & "('" & TxtFrom.Tag & "','" & TxtFrom & "','" & TxtTo.Tag & "','" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & TxtSubjek & "','" & TxtMessage & "',0,0)"
    CN.Execute StrSQL
    StrSQL = "Insert Into tblEmailLog(NIP,Nama,NIPTo,Tanggal,Subjeck,Email,StatusBaca,StatusBalas)Values"
    StrSQL = StrSQL & "('" & TxtFrom.Tag & "','" & TxtFrom & "','" & TxtTo.Tag & "','" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & TxtSubjek & "','" & TxtMessage & "',0,0)"
    CN.Execute StrSQL
    
    MsgBox "Message Sent", vbInformation
    Unload Me
End If
Command1.Caption = "Send Message"
Exit Sub
Adaerror:
    MsgBox err.Description
End Sub
Sub AddKaryawan()
Dim RsPeg As New ADODB.Recordset
    Dim Cboid     As String
    Dim cboid1    As String
If RsPeg.State = adStateOpen Then RsPeg.Close
    Cboid = vbNullString
    cboid1 = vbNullString
    StrSQL = "SELECT karyawan.NIP, karyawan.Nama, Divisi.NM_DIV FROM karyawan INNER JOIN Divisi ON karyawan.Kd_Divisi = Divisi.KD_DIV Where Status <> '14' order by Nama,NM_DIV"
    RsPeg.Open StrSQL, CN, adOpenStatic
    Do Until RsPeg.EOF
      Cboid = "|" & RsPeg("Nama") & vbTab & RsPeg("NIP") & vbTab & RsPeg("NM_DIV")
      cboid1 = cboid1 + Cboid
       RsPeg.MoveNext
    Loop
    TxtTo.ColComboList(0) = cboid1
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
      Skin1.ApplySkin hwnd
    End If
AddKaryawan
End Sub

