VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "CybEr_ClonE.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDateChecker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Current Date and Time Checker"
   ClientHeight    =   5070
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   3495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDateChecker.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   5055
      Left            =   0
      ScaleHeight     =   4995
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.CheckBox Check3 
         Caption         =   "No,Let me adjust it!"
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   3360
         Width           =   3060
      End
      Begin VB.Timer tmrBlink 
         Interval        =   500
         Left            =   2925
         Top             =   825
      End
      Begin VB.CheckBox Check1 
         Caption         =   "No,Let me adjust it!"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   3060
      End
      Begin VB.Timer tmrCurrTime 
         Interval        =   1000
         Left            =   3120
         Top             =   2280
      End
      Begin VB.Timer tmrCurrDate 
         Interval        =   1000
         Left            =   3120
         Top             =   1800
      End
      Begin VB.CommandButton btnOK 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clos&e"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   4320
         Width           =   1530
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "User &Guide"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Picture         =   "frmDateChecker.frx":000C
         TabIndex        =   1
         ToolTipText     =   "Help"
         Top             =   4320
         Width           =   1530
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   0
         OleObjectBlob   =   "frmDateChecker.frx":070E
         Top             =   0
      End
      Begin MSComCtl2.DTPicker dpTime 
         Height          =   390
         Left            =   240
         TabIndex        =   8
         Top             =   2640
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53280770
         CurrentDate     =   38557
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   390
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   0
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   12632256
         CustomFormat    =   "dd MMMM, yyyy"
         Format          =   53280771
         CurrentDate     =   38207
      End
      Begin MSComCtl2.DTPicker Dttime 
         Height          =   390
         Left            =   240
         TabIndex        =   11
         Top             =   3675
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53280770
         CurrentDate     =   38557
      End
      Begin VB.CheckBox Check2 
         Caption         =   "No,Let me adjust it!"
         Height          =   315
         Left            =   225
         TabIndex        =   4
         Top             =   2325
         Width           =   3060
      End
      Begin VB.Label Labels 
         Caption         =   "Reminder Is this correct?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   3120
         Width           =   2235
      End
      Begin VB.Label Labels 
         Caption         =   "Is this correct?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   7
         Top             =   2100
         Width           =   1665
      End
      Begin VB.Label Labels 
         Caption         =   "Is this correct?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   6
         Top             =   900
         Width           =   1665
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Please make sure that the current time and date settings are correct!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   765
         Index           =   0
         Left            =   675
         TabIndex        =   5
         Top             =   150
         Width           =   2640
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   75
         Picture         =   "frmDateChecker.frx":0942
         Top             =   150
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmDateChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
 
Private Sub btnOK_Click()
    If Check1.Value = 1 Then Date = dtpDate.Value
    If Check2.Value = 1 Then Time = dpTime.Value
    If Check3.Value = 1 Then
        Call SaveSetting(App.EXEName, App.EXEName, "Remind", Format(Dttime, "hh:mm"))
        RemindTime = GetSetting(App.EXEName, App.EXEName, "Remind")
    End If
    Unload Me
End Sub

Private Sub Check1_Click()
    DisplayCap
    If Check1.Value = 1 Then
        dtpDate.Enabled = True
        tmrCurrDate.Enabled = False
    Else
        dtpDate.Enabled = False
    
        dtpDate.Value = Date
        tmrCurrDate.Enabled = True
    End If
End Sub

Private Sub Check2_Click()
    DisplayCap
    If Check2.Value = 1 Then
        dpTime.Enabled = True
        tmrCurrTime.Enabled = False
    Else
        dpTime.Enabled = False
        tmrCurrTime.Enabled = True
        dpTime.Value = Time
    End If
End Sub

Private Sub Check3_Click()
 DisplayCap
 If Check3.Value = 1 Then
        Dttime.Enabled = True
 End If
End Sub

Private Sub Command1_Click()
Dim strhfile As String
 strhfile = App.Path & "\timesheet.chm"
 ShellExecute Me.hwnd, "open", strhfile, "", "", vbMaximizedFocus
End Sub

Private Sub Form_Load()
    dtpDate.Value = Date
    dpTime.Value = Time
    Dttime.Value = RemindTime
    dtpDate.Enabled = False
    dpTime.Enabled = False
    Dttime.Enabled = False
If Len(skinsFileName) <> 0 Then
      Skin1.LoadSkin App.Path + "\Skins\" + skinsFileName
      Skin1.ApplySkin hwnd
    End If
End Sub

Private Sub tmrBlink_Timer()
    Labels(0).Visible = Not Labels(0).Visible
End Sub

Private Sub tmrCurrDate_Timer()
    If dtpDate.Value <> Date Then dtpDate.Value = Date
End Sub

Private Sub tmrCurrTime_Timer()
    dpTime.Value = Time
End Sub

Private Sub DisplayCap()
    If Check1.Value = 1 Or Check2.Value = 1 Or Check3.Value = 1 Then
        btnOK.Caption = "Adjust"
    Else
        btnOK.Caption = "Close"
    End If
End Sub
