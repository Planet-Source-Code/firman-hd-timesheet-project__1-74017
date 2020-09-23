VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSearchs 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Records"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearchs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbFields 
      Height          =   315
      ItemData        =   "frmSearchs.frx":0A02
      Left            =   1800
      List            =   "frmSearchs.frx":0A04
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4995
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5520
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   4200
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   " Condition "
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6615
      Begin VB.ComboBox cmbOperation 
         Height          =   315
         Index           =   0
         ItemData        =   "frmSearchs.frx":0A06
         Left            =   120
         List            =   "frmSearchs.frx":0A10
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   360
         Width           =   2470
      End
      Begin VB.ComboBox cmbOperation 
         Height          =   315
         Index           =   1
         ItemData        =   "frmSearchs.frx":0A1D
         Left            =   120
         List            =   "frmSearchs.frx":0A27
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1080
         Width           =   2470
      End
      Begin VB.TextBox txtFilter 
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtFilter 
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   7
         Top             =   1080
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   20709379
         CurrentDate     =   38207
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "Or"
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "And"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Value           =   -1  'True
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   1
         Left            =   5040
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   20709379
         CurrentDate     =   38207
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   2
         Left            =   3120
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   20709379
         CurrentDate     =   38207
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   3
         Left            =   5040
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   20709379
         CurrentDate     =   38207
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "And"
         Height          =   255
         Left            =   4560
         TabIndex        =   14
         Top             =   1110
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "And"
         Height          =   255
         Left            =   4560
         TabIndex        =   13
         Top             =   390
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   360
         Width           =   240
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmSearchs.frx":0A34
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Records Where?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   165
      Width           =   1935
   End
End
Attribute VB_Name = "frmSearchs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public srcColumnHeaders As VSFlexGrid    'Source column headers
Public srcNoOfCol As Long
Public srcForm As Form 'Source form

Private Sub cmbOperation_Click(Index As Integer)
    If Index = 0 Then
        If cmbOperation(Index).ListIndex = 7 Then
            dtpDate(0).Visible = True
            dtpDate(1).Visible = True
            txtFilter(0).Visible = False
        Else
            txtFilter(0).Visible = True
            dtpDate(0).Visible = False
            dtpDate(1).Visible = False
        End If
    Else
        If cmbOperation(Index).ListIndex = 7 Then
            dtpDate(2).Visible = True
            dtpDate(3).Visible = True
            txtFilter(1).Visible = False
        Else
            txtFilter(1).Visible = True
            dtpDate(2).Visible = False
            dtpDate(3).Visible = False
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    'Verify
    If cmbOperation(0).ListIndex <> 7 Then If txtFilter(0).Text = "" Then txtFilter(0).SetFocus: Exit Sub
    
    On Error GoTo err
    Dim strFilter As String
    'Initialize the fields
    strFilter = Replace(cmbFields.Text, "/", "") 'ex. City/Town for tblCustomer
    strFilter = Replace(cmbFields.Text, " ", "")
    strFilter = "" & strFilter & ""
    'Initialize the operation used
    'First operation
    Select Case cmbOperation(0).ListIndex
        Case 0: strFilter = strFilter & " LIKE '%" & txtFilter(0).Text & "%'"
        Case 1: strFilter = strFilter & " = '" & txtFilter(0).Text & "'"
        Case 2: strFilter = strFilter & " <> '" & txtFilter(0).Text & "'"
        Case 3: strFilter = strFilter & " > '" & txtFilter(0).Text & "'"
        Case 4: strFilter = strFilter & " >= '" & txtFilter(0).Text & "'"
        Case 5: strFilter = strFilter & " < '" & txtFilter(0).Text & "'"
        Case 6: strFilter = strFilter & " <= '" & txtFilter(0).Text & "'"
        Case 7: strFilter = strFilter & " BETWEEN #" & dtpDate(0).Value & "# AND #" & dtpDate(1).Value & "#"
    End Select
    If cmbOperation(1).Text <> "" Then
        '-Second operation
        If Option1.Value = True Then
            strFilter = strFilter & " AND "
        Else
            strFilter = strFilter & " OR "
        End If
        
        Select Case cmbOperation(1).ListIndex
            Case 0: strFilter = strFilter & " LIKE '%" & txtFilter(1).Text & "%'"
            Case 1: strFilter = strFilter & " = '" & txtFilter(1).Text & "'"
            Case 2: strFilter = strFilter & " <> '" & txtFilter(1).Text & "'"
            Case 3: strFilter = strFilter & " > '" & txtFilter(1).Text & "'"
            Case 4: strFilter = strFilter & " >= '" & txtFilter(1).Text & "'"
            Case 5: strFilter = strFilter & " < '" & txtFilter(1).Text & "'"
            Case 6: strFilter = strFilter & " <= '" & txtFilter(1).Text & "'"
            Case 7: strFilter = strFilter & " BETWEEN #" & dtpDate(2).Value & "# AND #" & dtpDate(3).Value & "#"
        End Select
    End If
        
    'InputBox "", , strFilter
    'Pass the condition to filtered records
    srcForm.FilterRecord strFilter
    'Clear used variables
    strFilter = vbNullString
    
    Unload Me
    Exit Sub
err:
        If err.Number = -2147352571 Then
            MsgBox "Invalid search operation.", vbExclamation
            Unload Me
        ElseIf err.Number = 3001 Then
            Resume Next
        
        End If
End Sub

Private Sub Form_Load()
     If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path & "\Skins\" & skinsFileName
      Skin1.ApplySkin hwnd
    End If

    'Initialize values
    dtpDate(0).Value = Date
    dtpDate(1).Value = Date
    dtpDate(2).Value = Date
    dtpDate(3).Value = Date
    
    Dim I As Integer
    If srcNoOfCol = 0 Then srcNoOfCol = srcColumnHeaders.Cols - 3
    
    For I = 3 To srcNoOfCol
        If srcColumnHeaders.TextMatrix(0, I) <> "" Then cmbFields.AddItem srcColumnHeaders.TextMatrix(0, I)
    Next I
    I = 0
    
    cmbFields.ListIndex = 0
    cmbOperation(0).ListIndex = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSearchs = Nothing
End Sub

Private Sub txtFilter_GotFocus(Index As Integer)
    HLText txtFilter(Index)
End Sub
