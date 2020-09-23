VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmRkp 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Rekap Biaya Per Project"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   20774913
      CurrentDate     =   39867
   End
   Begin VB.Timer Timer2 
      Left            =   3960
      Top             =   4800
   End
   Begin VB.Timer Timer1 
      Left            =   3480
      Top             =   4800
   End
   Begin VB.TextBox txtpath 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4200
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtnm_divisi 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3840
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtkd_divisi 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3480
      TabIndex        =   8
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtnama 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3840
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtnip 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3480
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   840
      Top             =   5520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdprint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         TabIndex        =   5
         Top             =   960
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20709377
         CurrentDate     =   39850
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20709377
         CurrentDate     =   39850
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   840
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   840
      Top             =   5880
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   840
      Top             =   6240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   840
      Top             =   6600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   2040
      Top             =   5520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   2040
      Top             =   5880
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   2040
      Top             =   6240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   2040
      Top             =   6600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   3240
      Top             =   5520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   3240
      Top             =   5880
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   3240
      Top             =   6240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc12 
      Height          =   330
      Left            =   3240
      Top             =   6600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRkp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tglawal As String
Dim tglakhir As String
Dim fileName As String
Dim pathOpen As String
Dim tempbln1 As Integer
Dim tempbln2 As Integer
Dim tempbln3 As Integer
Dim tempkd_divisi As String
Dim tempnm_divisi As String
Dim tempkode As String
Dim tempnamakode As String
Dim tempnip As String
Dim tempnama As String
Dim posisi As Integer
Dim temptotalactualts As String
Dim amounttotalactualts As Double
Dim temptotallembur As Double
Dim gaji As Double
Dim tempgaji As Double
Dim tanggal As String
Dim temptotaljam As String
Dim temptotal As Double
Dim temptotalbiayats As String
Dim tempupahperjam As Double
Dim tempjamlembur1 As Double
Dim tempjamlembur2 As Double
Dim tempjamlembur3 As Double
Dim tempjamlembur4 As Double
Dim tempupah1 As Double
Dim tempupah2 As Double
Dim tempist1 As Double
Dim tempist2 As Double
Dim tempist3 As Double
Dim tempist4 As Double
Dim tempist5 As Double
Dim tempist6 As Double
Dim tempist7 As Double
Dim tempist8 As Double
Dim tempjamist_1 As Double
Dim tempjamist_2 As Double
Dim tempjamist_3 As Double
Dim tempjamist_4 As Double
Dim temphari As String
Dim temptingkat As Integer
Dim tempjammakan1 As Double
Dim tempjammakan2 As Double
Dim tempupahmakan As Double
Dim tempbiayalembur As Double
Dim tempbiaya1 As Double
Dim tempbiaya2 As Double
Dim tempupah As Double
Dim temp As Integer
Dim tempjammasuk As Integer
Dim tempjamkeluar As Integer
Dim tempmenitmasuk As Integer
Dim tempmenitkeluar As Integer
Dim tempjam As Integer
Dim tempmenit As Integer
Dim temptotalbiaya As Double
Dim temptotalbiayalembur As Double
Dim tempbiayalembur1 As String
Dim NilaiGaji As Currency
Private Sub cmdprint_Click()
tempbln1 = DTPicker1.Month
tempbln2 = DTPicker2.Month
tempbln3 = tempbln2 - tempbln1
If tempbln3 <> 0 Then
    MsgBox "Rekap perBulan", vbInformation, "PT.Wiratman"
Else
    tglawal = Format(DTPicker1.Value, "MM/DD/YYYY")
    tglakhir = Format(DTPicker2.Value, "MM/DD/YYYY")
    fileName = "C:\RekapBiayaProject.csv"
    pathOpen = PathOffice & " "
    Open fileName For Output As 1
    Print #1, "Periode : " & DTPicker1.Day & "-" & DTPicker2.Day & "/" & DTPicker2.Month & "/" & DTPicker2.Year
    Close #1
    querydivisi
    StrSQL = "delete from tmpnip"
    connserver3.Execute StrSQL
    Timer1.Enabled = True
    Timer1.Interval = 10
End If
End Sub
'1
Sub querydivisi()
Dim sql As String
sql = "select status.kd_divisi, divisi.nm_div " & _
    "from status, divisi " & _
    "where status.kd_divisi = divisi.kd_div " & _
    "and status.flag in ('Timesheet','Lembur') " & _
    "and status.periode >= '" & tglawal & "' " & _
    "and status.periode <= '" & tglakhir & "' " & _
    "and status.status = 'Valid' " & _
    "group by status.kd_divisi, divisi.nm_div " & _
    "order by status.kd_divisi"
Adodc1.ConnectionString = koneksiserver3
Adodc1.RecordSource = sql
Adodc1.Refresh
End Sub
'2
Sub querykode()
Dim sql As String
sql = "select kode,nama from project " & _
    "where kd_divisi = '" & tempkd_divisi & "' " & _
    "and status = 'Terpakai' " & _
    "order by kode"
Adodc2.ConnectionString = koneksiserver3
Adodc2.RecordSource = sql
Adodc2.Refresh
End Sub

Sub querynip()
Dim sql As String
'sql = "select timesheet2.nip " & _
'    "from timesheet2, karyawan, status " & _
'    "where timesheet2.nip = karyawan.nip " & _
'    "and karyawan.nip = status.nip " & _
'    "and timesheet2.slot = '" & tempkode & "' " & _
'    "and timesheet2.status = 'Actual' " & _
'    "and status.status = 'Valid' " & _
'    "and status.flag = 'Timesheet' " & _
'    "and status.periode >= '" & tglawal & "' " & _
'    "and status.periode <= '" & tglakhir & "' " & _
'    "group by timesheet2.nip " & _
'    "order by timesheet2.nip"
sql = "select karyawan.nip,karyawan.nama,karyawan.kd_divisi,divisi.nm_div " & _
    "from karyawan, divisi where karyawan.kd_divisi = divisi.kd_div " & _
    "and divisi.nm_div = '" & tempnm_divisi & "' " & _
    "and karyawan.status <> '14' " & _
    "order by nip"
Adodc3.ConnectionString = koneksiserver3
Adodc3.RecordSource = sql
Adodc3.Refresh


Do While Not Adodc3.Recordset.EOF
    tempnip = Adodc3.Recordset!NIP
    StrSQL = "insert into tmpnip (nip_divisi,nip,kode) values " & _
        "('" & txtnip.Text & "','" & tempnip & "','" & tempkode & "')"
    connserver3.Execute StrSQL
    Adodc3.Recordset.MoveNext
Loop
'sql = "select lembur1.nip " & _
'    "from lembur1, karyawan, status " & _
'    "where lembur1.nip = karyawan.nip " & _
'    "and karyawan.nip = status.nip " & _
'    "and lembur1.noproject = '" & tempkode & "' " & _
'    "and status.status = 'Valid' " & _
'    "and status.flag = 'Lembur' " & _
'    "and status.periode >= '" & tglawal & "' " & _
'    "and status.periode <= '" & tglakhir & "' " & _
'    "group by lembur1.nip " & _
'    "order by lembur1.nip"
'strsql = "Select * From Lembur1 Where NIP = '" & varNip & "' And Tanggal BETWEEN '" & Format(DTPicker1.Value, "YYYY-mm-DD") & "' AND '" & Format(DTPicker2.Value, "YYYY-mm-DD") & "' And NoProject NOT LIKE '%*%' Order By Tanggal"
'
'Adodc4.ConnectionString = koneksiserver3
'Adodc4.RecordSource = sql
'Adodc4.Refresh
'Do While Not Adodc4.Recordset.EOF
'    tempnip = Adodc4.Recordset!NIP
'    strsql = "insert into tmpnip (nip_divisi,nip,kode) values " & _
'        "('" & txtnip.Text & "','" & tempnip & "','" & tempkode & "')"
'    connserver3.Execute strsql
'    Adodc4.Recordset.MoveNext
'Loop
'sql = "select tmpnip.nip, karyawan.nama " & _
'    "from tmpnip, karyawan, status " & _
'    "where tmpnip.nip = karyawan.nip " & _
'    "and tmpnip.nip_divisi = '" & txtnip.Text & "' " & _
'    "and tmpnip.kode = '" & tempkode & "' " & _
'    "group by tmpnip.nip, karyawan.nama " & _
'    "order by tmpnip.nip"
'Adodc5.ConnectionString = koneksiserver3
'Adodc5.RecordSource = sql
'Adodc5.Refresh
End Sub

Sub queryactualts()
Dim sql As String
sql = "select count(slot) as slot " & _
    "from timesheet2 " & _
    "where nip = '" & tempnip & "' " & _
    "and tanggal >= '" & tglawal & "' " & _
    "and tanggal <= '" & tglakhir & "' " & _
    "and status = 'Actual' " & _
    "and slot = '" & tempkode & "'"
Adodc6.ConnectionString = koneksiserver3
Adodc6.RecordSource = sql
Adodc6.Refresh
End Sub

Sub querylembur()
Dim sql As String
sql = "select *from lembur1 " & _
    "where nip = '" & tempnip & "' " & _
    "and tanggal >= '" & tglawal & "' " & _
    "and tanggal <= '" & tglakhir & "' " & _
    "and noproject = '" & tempkode & "'"
Adodc7.ConnectionString = koneksiserver3
Adodc7.RecordSource = sql
Adodc7.Refresh
End Sub

Sub querygaji()
Dim sql As String
sql = "select *from gaji where nip = '" & tempnip & "'"
Adodc8.ConnectionString = koneksiserver3
Adodc8.RecordSource = sql
Adodc8.Refresh
End Sub

Sub querysetting()
Dim sql As String
sql = "Select *from setting " & _
    "where tingkat = '" & temptingkat & "' " & _
    "and hari = '" & temphari & "'"
Adodc9.ConnectionString = koneksiserver3
Adodc9.RecordSource = sql
Adodc9.Refresh
End Sub

Sub querytelat()
Dim sql As String
sql = "Select masuk,keluar from absensi " & _
    "where nip = '" & tempnip & "' " & _
    "and tgl >= '" & tanggal & "' " & _
    "and tgl <= '" & tanggal & "'"
Adodc10.ConnectionString = koneksiserver3
Adodc10.RecordSource = sql
Adodc10.Refresh
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker1.Value = DateSerial(Year(Now), Month(Now), 1)
    DTPicker1.Value = DateAdd("M", 0, DTPicker1.Value)
    DTPicker2.Value = DateSerial(Year(Now), Month(Now), 1)
    DTPicker2.Value = DateAdd("M", 1, DTPicker2.Value) - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
frmMenu.Show
frmMenu.Enabled = True
End Sub

Private Sub timer1_timer()
Dim RsActual As New ADODB.Recordset
Dim rsGaji As New ADODB.Recordset


If (Adodc1.Recordset.EOF = True) Or (Adodc1.Recordset.BOF = True) Then
    Timer1.Enabled = False
    Shell pathOpen & fileName, vbNormalFocus
    Label3.Caption = "Selesai"
Else
    tempnm_divisi = Adodc1.Recordset!nm_div
    tempkd_divisi = Adodc1.Recordset!kd_divisi
    Label3.Caption = tempnm_divisi
    querykode
    If (Adodc2.Recordset.EOF = True) And (Adodc2.Recordset.BOF = True) Then
    Else
            Open fileName For Append As 2
            Print #2, " "
            Print #2, "Divisi " & tempnm_divisi
            Close #2
      Do While Not Adodc2.Recordset.EOF
            tempkode = Adodc2.Recordset!Kode
            tempnamakode = Adodc2.Recordset!Nama
            Label3.Caption = tempnm_divisi & "-" & tempkode
            
'     If rs.State = adStateOpen Then rs.Close
'        strsql = " SELECT Timesheet2.NIP, Timesheet2.Tanggal,"
'        strsql = strsql & " Timesheet2.Kd_Divisi, Timesheet2.Slot, Project.Nama, Project.Status"
'        strsql = strsql & " FROM Timesheet2 INNER JOIN"
'        strsql = strsql & " Project ON Timesheet2.Kd_Divisi = Project.Kd_Divisi AND"
'        strsql = strsql & " Timesheet2.Slot = Project.kode where Project.Status ='Terpakai' And Timesheet2.Tanggal Between '" & DTPicker1 & "' And '" & DTPicker2 & "'"
'        strsql = strsql & " Order By Timesheet2.NIP,Timesheet2.Tanggal"
'        rs.Open strsql, connserver3, adOpenStatic
''        If Not rs.EOF Then
          
            
            If Rs.State = adStateOpen Then Rs.Close
                StrSQL = "select timesheet2.nip, timesheet2.Slot, karyawan.nama " & _
                  "from timesheet2, karyawan " & _
                  "where timesheet2.nip = karyawan.nip " & _
                  "and timesheet2.slot = '" & tempkode & "' And Timesheet2.Status ='Actual'" & _
                  "and timesheet2.tanggal >= '" & Format(DTPicker1, "yyyy-MM-dd") & "'" & _
                  "and timesheet2.tanggal <= '" & Format(DTPicker2, "yyyy-MM-dd") & "'" & _
                  "group by timesheet2.nip, karyawan.nama,timesheet2.slot " & _
                  "order by timesheet2.nip"
            Rs.Open StrSQL, connserver3, adOpenStatic
              
           If Not Rs.EOF Then
                
                Open fileName For Append As 3
                Print #3, " "
                Print #3, "No Project" & "," & "Nama Project"
                Print #3, tempkode & "," & tempnamakode
                'Print #3, " "
                Print #3, "NIP" & "," & "Nama" & "," & "Jam Kerja" & "," & "Biaya Gaji" & "," & "Jam Lembur" & "," & "Biaya Lembur"
                Close #3
              
             Do Until Rs.EOF
                    temptotalactualts = 0
                    temptotalbiayats = 0
                    NilaiGaji = 0
                    tempbiayalembur1 = 0
                    temptotaljam = 0
                    If RsActual.State = adStateOpen Then RsActual.Close
                   StrSQL = "select count(slot) as slot " & _
                    "from timesheet2 " & _
                    "where nip = '" & Rs!NIP & "' " & _
                    "and tanggal >= '" & tglawal & "' " & _
                    "and tanggal <= '" & tglakhir & "' " & _
                    "and status = 'Actual' " & _
                    "and slot = '" & tempkode & "' And Slot NOT Like 'Telat%'"
                    RsActual.Open StrSQL, connserver3, adOpenStatic
                    If Not RsActual.EOF Then
                       temptotalactualts = RsActual!slot / 2
                                    End If
                    If rsGaji.State = adStateOpen Then rsGaji.Close
                    rsGaji.Open "SELECT * FROM GAJI WHERE NIP = '" & Rs!NIP & "'", connserver3, adOpenStatic
                   
                    If Not rsGaji.EOF Then
                        NilaiGaji = rsGaji!gaji
                        temptingkat = rsGaji!tingkat
                        temptotalbiayats = Round(CCur(((NilaiGaji) / 173)) * CDbl(temptotalactualts), 2)
                        temptotalbiaya = temptotalbiaya + temptotalbiayats
                        temptotalbiayats = Replace(temptotalbiayats, ",", ".")
                    End If
                    GetLembur (Rs!NIP)
                    temptotalactualts = Replace(temptotalactualts, ",", ".")
                    temptotaljam = Replace(temptotaljam, ",", ".")
                    tempbiayalembur1 = Replace(tempbiayalembur1, ",", ".")
                    Open fileName For Append As 4
                    Print #4, Rs!NIP & "," & Replace(Rs!Nama, ",", " ") & "," & temptotalactualts & "," & temptotalbiayats & "," & temptotaljam & "," & tempbiayalembur1
                    Close #4
                    Rs.MoveNext
                Loop
                
            End If
         Adodc2.Recordset.MoveNext
        Loop
    End If
    Adodc1.Recordset.MoveNext
End If
End Sub

Sub GetLembur(ByVal NIP As String)
Dim jamMasuk, Hari As String
Dim JamKeluar As String
Dim TotalJam, TotalLemburBruto As String
Dim TotalLembur, Terlambat As String
Dim tingkat, Gapok As Double
Dim rsAbsen As New ADODB.Recordset
Dim rsGaji As New ADODB.Recordset
Dim RsSetting As New ADODB.Recordset
Dim Istirahat As String
Dim Jam1, Jam2 As Double
Dim Jam3, Jam4 As Double
Dim JamMakan As Integer
Dim Transport As String
Dim Upah1, Upah2, UM As Currency
Dim JamLembur, SplitLembur As Double
Dim NilaiUpah As Currency
Dim rsLembur As New ADODB.Recordset
If rsLembur.State = adStateOpen Then rsLembur.Close
StrSQL = "Select * From Lembur1 Where NIP = '" & NIP & "' And Tanggal BETWEEN '" & Format(DTPicker1.Value, "YYYY-mm-DD") & "' AND '" & Format(DTPicker2.Value, "YYYY-mm-DD") & "' And NoProject NOT LIKE '%*%' Order By Tanggal"
rsLembur.Open StrSQL, connserver3, adOpenStatic

If Not rsLembur.EOF Then
     Open fileName For Append As 2
     Do Until rsLembur.EOF
         TotalLembur = 0
         Istirahat = 0
         TotalLemburBruto = 0
         UM = 0
         Hari = rsLembur!Hari
         StrSQL = "Select * From Absensi Where NIP = '" & NIP & "' And Tgl = '" & Format(rsLembur!tanggal, "yyyy-MM-dd") & "'"
         If rsAbsen.State = adStateOpen Then rsAbsen.Close
         rsAbsen.Open StrSQL, connserver3, adOpenStatic
         If Not rsAbsen.EOF Then
'             Terlambat = "08:00"
'
'             If Format(rsAbsen!Masuk, "HH:mm") > CDate(Terlambat) Then
'                Terlambat = Format(rsAbsen!Masuk - CDate(Terlambat), "HH:mm")
'             Else
'                Terlambat = 0
'             End If
             If Format(rsAbsen!masuk, "yyyy-MM-dd") = "2000-01-01" Then
                jamMasuk = "TIDAK ABSEN"
             Else
                jamMasuk = Format(rsAbsen!masuk, "HH:mm")
             End If
             
             If Format(rsAbsen!keluar, "yyyy-MM-dd") = "2000-01-01" Then
                JamKeluar = "TIDAK ABSEN"
                TotalJam = 0
             Else
                JamKeluar = Format(rsAbsen!keluar, "HH:mm")
                TotalJam = Format(rsAbsen!keluar - rsAbsen!masuk, "HH:mm")
             End If
            
         End If
         
'         If rsGaji.State = adStateOpen Then rsGaji.Close
'         rsGaji.Open "SELECT * FROM GAJI WHERE NIP ='" & NIP & "'", connserver3, adOpenStatic
'
'         If Not rsGaji.EOF Then
'             tingkat = rsGaji!tingkat
'             Gapok = rsGaji!gaji
'         End If
                    
         If RsSetting.State = adStateOpen Then RsSetting.Close
         RsSetting.Open "SELECT * FROM SETTING WHERE TINGKAT = '" & temptingkat & "' AND HARI = '" & Hari & "' AND Berlaku = '" & Format(DTPicker1.Value, "MM/yyyy") & "'", connserver3, adOpenStatic
         If Not RsSetting.EOF Then
             Jam1 = RsSetting!jamlembur1
             Jam2 = RsSetting!jamlembur2
             Jam3 = RsSetting!jamlembur3
             Jam4 = RsSetting!jamlembur4
             UM = RsSetting!upahmakan
             JamMakan = RsSetting!jammakan1
            If JamKeluar = "TIDAK ABSEN" Then
                TotalLembur = 0
             Else
                Dim Jam As Integer
                Dim menit As Integer

                  Select Case Hari
                      Case "Kerja"
                            TotalLemburBruto = Left(TotalJam, 2) - 9 & Mid(TotalJam, 3, 6)
                            TotalLembur = CDbl(Left(TotalJam, 2) - 9)
                      Case "Libur"
                            TotalLembur = CDbl(Left(TotalJam, 2))
                            TotalLemburBruto = TotalJam
                  End Select
                  
                  If Len(TotalLemburBruto) = 4 Then TotalLemburBruto = "0" & TotalLemburBruto
         
                  If TotalLembur >= RsSetting!ist1 And TotalLembur <= RsSetting!ist2 Then
                      Istirahat = CDbl(RsSetting!jamist_1) * 60
                      If Len(TotalLembur) = 1 Then
                          TotalLembur = "0" & TotalLembur & Mid(TotalJam, 3, 6)
                      ElseIf Len(TotalLembur) = 2 Then
                          TotalLembur = TotalLembur & Mid(TotalJam, 3, 6)
                      End If
                      If Istirahat > 60 Then
                            Jam = Istirahat \ 60
                            menit = Istirahat - (60 * Jam)
                            Istirahat = Jam & ":" & menit
                      Else
                          Istirahat = "00:" & Istirahat
                      End If
                      TotalLembur = Format(CDate(TotalLembur) - CDate(Istirahat), "HH:mm")
                   ElseIf TotalLembur >= RsSetting!ist3 And TotalLembur <= RsSetting!ist4 Then
                      Istirahat = CDbl(RsSetting!jamist_2) * 60
                      If Len(TotalLembur) = 1 Then
                          TotalLembur = "0" & TotalLembur & Mid(TotalJam, 3, 6)
                      ElseIf Len(TotalLembur) = 2 Then
                          TotalLembur = TotalLembur & Mid(TotalJam, 3, 6)
                      End If
                      If Istirahat > 60 Then
                            Jam = Istirahat \ 60
                            menit = Istirahat - (60 * Jam)
                            Istirahat = Jam & ":" & menit
                            Istirahat = Format(Istirahat, "HH:mm")
                      Else
                          Istirahat = "00:" & Istirahat
                      End If

                      TotalLembur = Format(CDate(TotalLembur) - CDate(Istirahat), "HH:mm")
                    ElseIf TotalLembur >= RsSetting!ist5 And TotalLembur <= RsSetting!ist6 Then
                      Istirahat = CDbl(RsSetting!jamist_3) * 60
                      If Len(TotalLembur) = 1 Then
                          TotalLembur = "0" & TotalLembur & Mid(TotalJam, 3, 6)
                      ElseIf Len(TotalLembur) = 2 Then
                          TotalLembur = TotalLembur & Mid(TotalJam, 3, 6)
                      End If
                      If Istirahat > 60 Then
                            Jam = Istirahat \ 60
                            menit = Istirahat - (60 * Jam)
                            Istirahat = Jam & ":" & menit
                            Istirahat = Format(Istirahat, "HH:mm")
                      Else
                          Istirahat = "00:" & Istirahat
                      End If
                      TotalLembur = Format(CDate(TotalLembur) - CDate(Istirahat), "HH:mm")
                    ElseIf TotalLembur >= RsSetting!ist7 And TotalLembur <= RsSetting!ist8 Then
                      Istirahat = CDbl(RsSetting!jamist_4) * 60
                      If Len(TotalLembur) = 1 Then
                          TotalLembur = "0" & TotalLembur & Mid(TotalJam, 3, 6)
                      ElseIf Len(TotalLembur) = 2 Then
                          TotalLembur = TotalLembur & Mid(TotalJam, 3, 6)
                      End If
                      If Istirahat > 60 Then
                            Jam = Istirahat \ 60
                            menit = Istirahat - (60 * Jam)
                            Istirahat = Jam & ":" & menit
                            Istirahat = Format(Istirahat, "HH:mm")
                      Else
                          Istirahat = "00:" & Istirahat
                      End If
                      TotalLembur = Format(CDate(TotalLembur) - CDate(Istirahat), "HH:mm")
                    End If
                  
            End If
         End If
         
         'pembulatan menit
         If TotalLembur <> 0 Then
            Dim Pembulatan As Integer
            Pembulatan = Right(TotalLembur, 2)
            If Pembulatan >= 0 And Pembulatan <= 7 Then
               TotalLembur = Left(TotalLembur, 3) & "00"
            ElseIf Pembulatan >= 8 And Pembulatan <= 17 Then
               TotalLembur = Left(TotalLembur, 3) & 15
            ElseIf Pembulatan >= 18 And Pembulatan <= 22 Then
               TotalLembur = Left(TotalLembur, 3) & 15
            ElseIf Pembulatan >= 23 And Pembulatan <= 37 Then
               TotalLembur = Left(TotalLembur, 3) & 30
            ElseIf Pembulatan >= 38 And Pembulatan <= 60 Then
               TotalLembur = Left(TotalLembur, 3) & 45
            End If
         
         If Left(TotalLembur, 2) < JamMakan Then UM = 0
         If JamKeluar <> "TIDAK ABSEN" Then
            If JamKeluar >= "00:00" And JamKeluar <= "05:00" Then
               Transport = "Ya"
            Else
               Transport = "Tidak"
            End If
         Else
            Transport = "Tidak"
         End If
         'upah
           JamLembur = Replace(TotalLembur, ":", ".")
            Upah1 = RsSetting!Upah1
            Upah2 = RsSetting!Upah2
            NilaiUpah = 0
            If JamLembur >= Jam1 And JamLembur <= Jam2 Then
                NilaiUpah = (JamLembur * NilaiGaji / 173) * Upah1
            ElseIf JamLembur >= Jam3 And JamLembur <= Jam4 Then
                NilaiUpah = (Jam2 * NilaiGaji / 173) * Upah1
                SplitLembur = JamLembur - Jam2
                NilaiUpah = NilaiUpah + (SplitLembur * Gapok / 173) * Upah2
            End If
         Else
            UM = 0
            Transport = "Tidak"
            NilaiUpah = 0
         End If
            temptotaljam = CDbl(temptotaljam) + CDbl(JamLembur)
            tempbiayalembur1 = CCur(tempbiayalembur1) + NilaiUpah
         rsLembur.MoveNext
     Loop
     If RsSetting.State = adStateOpen Then RsSetting.Close
     If rsGaji.State = adStateOpen Then rsGaji.Close
     If rsAbsen.State = adStateOpen Then rsAbsen.Close
     If rsLembur.State = adStateOpen Then rsLembur.Close
    Close #2
End If
End Sub
