VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Begin VB.Form FrmMessage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Message"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   0
      ScaleHeight     =   4035
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton Command2 
         Caption         =   "&Close"
         Height          =   495
         Left            =   2040
         TabIndex        =   10
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Send Message"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   3480
         Width           =   1455
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
         TabIndex        =   7
         Text            =   "FrmMessage.frx":0000
         Top             =   1440
         Width           =   5175
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
         TabIndex        =   6
         Text            =   " "
         Top             =   120
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
         TabIndex        =   4
         Text            =   " "
         Top             =   840
         Width           =   4335
      End
      Begin VSFlex8Ctl.VSFlexGrid CboKaryawan 
         Height          =   315
         Left            =   960
         TabIndex        =   1
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
         FormatString    =   $"FrmMessage.frx":0002
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
         TabIndex        =   11
         Top             =   120
         Width           =   1095
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
         Top             =   1200
         Width           =   855
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
         TabIndex        =   5
         Top             =   840
         Width           =   855
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
         TabIndex        =   3
         Top             =   120
         Width           =   495
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
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2760
      OleObjectBlob   =   "FrmMessage.frx":002B
      Top             =   1920
   End
End
Attribute VB_Name = "FrmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NIPKry, kdDivisi As String
Public NIPPM, NamaPM As String
Public Project, NamaKry As String

Private Sub Command1_Click()
Dim x As Integer
Dim RsSimpan As New ADODB.Recordset
On Error GoTo Adaerror
If Trim(CboKaryawan) = "" Then
    MsgBox "Isian 'To' Masih Kosong", vbCritical
    CboKaryawan.SetFocus
    Exit Sub
End If

If Trim(TxtSubjek) = "" Then
    MsgBox "Subjek Masih Kosong", vbCritical
    TxtSubjek.SetFocus
    Exit Sub
End If
If Trim(TxtMessage) = "" Then
    MsgBox "Pesan Masih Kosong", vbCritical
    TxtMessage.SetFocus
    Exit Sub
End If

RsSimpan.Open "Select * from Karyawan Where Nama = '" & CboKaryawan & "'", CN, adOpenStatic
If Not RsSimpan.EOF Then CboKaryawan.Text = RsSimpan!NIP
If Trim(CboKaryawan) = "All" Then
    For x = 0 To Combo1.ListCount - 1
        Combo1.ListIndex = x
        StrSQL = "Insert Into tblEmail(NIP,Nama,NIPTo,Tanggal,Subjeck,Email,StatusBaca,StatusBalas)Values"
        StrSQL = StrSQL & "('" & TxtFrom.Tag & "','" & TxtFrom & "','" & Combo1.Text & "','" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & TxtSubjek & "','" & TxtMessage & "',0,0)"
        CN.Execute StrSQL
        StrSQL = "Insert Into tblEmailLog(NIP,Nama,NIPTo,Tanggal,Subjeck,Email,StatusBaca,StatusBalas)Values"
        StrSQL = StrSQL & "('" & TxtFrom.Tag & "','" & TxtFrom & "','" & Combo1.Text & "','" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & TxtSubjek & "','" & TxtMessage & "',0,0)"
        CN.Execute StrSQL
    Next
Else
    StrSQL = "Insert Into tblEmail(NIP,Nama,NIPTo,Tanggal,Subjeck,Email,StatusBaca,StatusBalas)Values"
    StrSQL = StrSQL & "('" & TxtFrom.Tag & "','" & TxtFrom & "','" & CboKaryawan & "','" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & TxtSubjek & "','" & TxtMessage & "',0,0)"
    CN.Execute StrSQL
    
    StrSQL = "Insert Into tblEmailLog(NIP,Nama,NIPTo,Tanggal,Subjeck,Email,StatusBaca,StatusBalas)Values"
    StrSQL = StrSQL & "('" & TxtFrom.Tag & "','" & TxtFrom & "','" & CboKaryawan & "','" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "','" & TxtSubjek & "','" & TxtMessage & "',0,0)"
    CN.Execute StrSQL
End If
    CboKaryawan.Text = ""
    TxtSubjek.Text = ""
    TxtMessage.Text = ""
    RsSimpan.Close
    MsgBox "Message Sent", vbInformation

Exit Sub
Adaerror:
    MsgBox err.Description
End Sub
 
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Cboid, cboid1 As String
Dim x As Integer
If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
      Skin1.ApplySkin hwnd
    End If
'If Rscek.State = adStateOpen Then Rscek.Close
'    Combo1.Clear
'Select Case UCase(strGroup)
'    Case "USER", "IT"
'        Combo1.AddItem NIPPM
'        cboid1 = " "
'        cboid1 = cboid1 & "|All" & vbTab & "PM & DIVISI"
'        Cboid = cboid1 & "|" & NIPPM & vbTab & NamaPM & " (PM)"
'        StrSQL = "Select * from VUser Where Kd_divisi ='" & KodeDivisi & "' And type = 'admin'"
'        Rscek.Open StrSQL, CN, adOpenStatic
'        Do Until Rscek.EOF
'
'           Combo1.AddItem Rscek!NIP
'           Cboid = Cboid & "|" & Rscek!NIP & vbTab & Rscek!Nama & " (DIVISI)"
'           Rscek.MoveNext
'        Loop
'    Case "ADMIN"
'       StrSQL = "Select * from Karyawan Where NIP ='" & NIPKry & "'"
'       Rscek.Open StrSQL, CN, adOpenStatic
'        If Not Rscek.EOF Then NamaKry = Rscek!Nama
'            cboid1 = " "
'            Combo1.AddItem NIPKry
'            Combo1.AddItem NIPPM
'            cboid1 = cboid1 & "|All" & vbTab & "Karyawan & PM"
'            Cboid = cboid1 & "|" & NIPKry & vbTab & Trim(NamaKry) & " (Karyawan)"
'            Cboid = Cboid & "|" & NIPPM & vbTab & NamaPM & " (PM)"
'    Case "PM"
'         StrSQL = "Select * from Karyawan Where NIP ='" & NIPKry & "'"
'          Rscek.Open StrSQL, CN, adOpenStatic
'        If Not Rscek.EOF Then NamaKry = Rscek!Nama
'            Combo1.AddItem NIPKry
'            cboid1 = " "
'            cboid1 = cboid1 & "|All" & vbTab & "Karyawan & DIVISI"
'            Cboid = cboid1 & "|" & NIPKry & vbTab & Trim(NamaKry) & " (Karyawan)"
'            If Rscek.State = adStateOpen Then Rscek.Close
'            StrSQL = "Select * from VUser Where Kd_divisi ='" & KodeDivisi & "' And type = 'admin'"
'            Rscek.Open StrSQL, CN, adOpenStatic
'            Do Until Rscek.EOF
'               Combo1.AddItem NIP
'               Cboid = Cboid & "|" & Rscek!NIP & vbTab & Rscek!Nama & " (DIVISI)"
'               Rscek.MoveNext
'            Loop
'End Select
'    TxtSubjek = "Perihal Project " & Project
    AddKaryawan
'    CboKaryawan.ColComboList(0) = Cboid
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
    CboKaryawan.ColComboList(0) = cboid1
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set FrmMessage = Nothing
End Sub

Private Sub TxtMessage_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96

End Sub

Private Sub TxtSubjek_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 96
End Sub
