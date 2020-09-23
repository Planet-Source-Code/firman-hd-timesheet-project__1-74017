VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmRkpPerNip 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Rekap Biaya perNip"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   1680
      TabIndex        =   16
      Top             =   120
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Height          =   3135
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   5655
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2775
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   5655
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
         Left            =   1200
         TabIndex        =   7
         Top             =   960
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1200
         TabIndex        =   8
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
         Format          =   52953089
         CurrentDate     =   39850
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3840
         TabIndex        =   9
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
         Format          =   52953089
         CurrentDate     =   39850
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         Height          =   240
         Left            =   3120
         TabIndex        =   10
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.TextBox txtnip 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6480
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtnama 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6840
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtkd_divisi 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6480
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtnm_divisi 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6840
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtpath 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7200
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   255
      Left            =   7320
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   52953089
      CurrentDate     =   39867
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2760
      Top             =   7440
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2760
      Top             =   7800
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
      Left            =   2760
      Top             =   8160
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
      Left            =   2760
      Top             =   8520
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
      Left            =   3960
      Top             =   7440
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
      Left            =   3960
      Top             =   7800
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
      Left            =   3960
      Top             =   8160
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
      Left            =   3960
      Top             =   8520
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
      Left            =   5160
      Top             =   7440
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
      Left            =   5160
      Top             =   7800
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
      Left            =   5160
      Top             =   8160
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
      Left            =   5160
      Top             =   8520
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih Divisi"
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   4920
      Width           =   120
   End
End
Attribute VB_Name = "frmRkpPerNip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varID, varNip, varNama As String
Dim pathOpen As String
Dim fileName As String
Private Sub cmdprint_Click()
If Combo1.Text = "" Then
    MsgBox "Pilih Divisi Terlebih Dahulu", vbInformation, "PT.Wiratman"
Else
    If (varNip = "") And (varNama = "") Then
        MsgBox "Pilih NIP Yang Ingin Dicetak", vbInformation, "PT.Wiratman"
    Else
        tglawal = Format(DTPicker1.Value, "MM/DD/YYYY")
        tglakhir = Format(DTPicker2.Value, "MM/DD/YYYY")
        fileName = "C:\RekapPerNIP.csv"
        pathOpen = PathOffice & " "
        Open fileName For Output As 1
        Print #1, "REKAP LEMBUR"
        Print #1, "PERIODE : " & Format(DTPicker1.Value, "dd/MMM/yyyy") & " - " & Format(DTPicker2.Value, "dd/MMM/yyyy")
        Print #1, "NIP : " & varNip; "," & "NAMA : " & Replace(varNama, ",", " "); "," & "DIVISI : " & Combo1
        Print #1, " "
        Print #1, "TANGGAL  " & "," & "PROJECT " & "," & " MASUK " & "," & " KELUAR " & "," & " TOTAL KERJA " & "," & " LEMBUR BRUTO " & "," & " ISTIRAHAT " & "," & " LEMBUR NETTO " & "," & " BIAYA LEMBUR " & "," & " U.MAKAN " & "," & " U.TRANSPORT " & ""
        Close #1
        GetLembur (varNip)
        Shell pathOpen & fileName, vbNormalFocus
         
'         varNip = ""
'         varNama = ""
    End If
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

If Rs.State = adStateOpen Then Rs.Close
StrSQL = "Select * From Lembur1 Where NIP = '" & varNip & "' And Tanggal BETWEEN '" & Format(DTPicker1.Value, "YYYY-mm-DD") & "' AND '" & Format(DTPicker2.Value, "YYYY-mm-DD") & "' And NoProject NOT LIKE '%*%' Order By Tanggal"
'    strsql = "SELECT Status.NIP, Lembur1.Hari, Lembur1.NoProject,"
'    strsql = strsql & " Lembur1.Tanggal, Lembur1.Jam_Awal, Lembur1.Jam_Akhir,"
'    strsql = strsql & " Lembur1.Total, Status.Kd_Divisi, Status.Status,"
'    strsql = strsql & " Status.Flag"
'    strsql = strsql & " FROM Lembur1 INNER JOIN"
'    strsql = strsql & " Status ON Lembur1.NIP = Status.NIP Where Status.NIP = '" & varNip & "' And Lembur1.Tanggal BETWEEN '" & Format(DTPicker1.Value, "YYYY-mm-DD") & "' AND '" & Format(DTPicker2.Value, "YYYY-mm-DD") & "' Order By Lembur1.Tanggal"
Rs.Open StrSQL, connserver3, adOpenStatic

If Not Rs.EOF Then
     Open fileName For Append As 2
     Do Until Rs.EOF
         TotalLembur = 0
         Istirahat = 0
         TotalLemburBruto = 0
         UM = 0
         Hari = Rs!Hari
         StrSQL = "Select * From Absensi Where NIP = '" & varNip & "' And Tgl = '" & Format(Rs!tanggal, "yyyy-MM-dd") & "'"
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
             If Format(rsAbsen!masuk, "yyyy-MM-DD") = "2000-01-01" Then
                jamMasuk = "TIDAK ABSEN"
             Else
                jamMasuk = Format(rsAbsen!masuk, "HH:mm")
             End If
             
             If Format(rsAbsen!keluar, "yyyy-MM-DD") = "2000-01-01" Then
                JamKeluar = "TIDAK ABSEN"
                TotalJam = 0
             Else
                JamKeluar = Format(rsAbsen!keluar, "HH:mm")
                TotalJam = Format(rsAbsen!keluar - rsAbsen!masuk, "HH:mm")
             End If
            
         End If
         
         If rsGaji.State = adStateOpen Then rsGaji.Close
         rsGaji.Open "SELECT * FROM GAJI WHERE NIP ='" & varNip & "'", connserver3, adOpenStatic
       
         If Not rsGaji.EOF Then
             tingkat = rsGaji!tingkat
             Gapok = rsGaji!gaji
         End If
                    
         If RsSetting.State = adStateOpen Then RsSetting.Close
         RsSetting.Open "SELECT * FROM SETTING WHERE TINGKAT = '" & tingkat & "' AND HARI = '" & Hari & "' AND Berlaku = '" & Format(DTPicker1.Value, "MM/yyyy") & "'", connserver3, adOpenStatic
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
                NilaiUpah = (JamLembur * Gapok / 173) * Upah1
            ElseIf JamLembur >= Jam3 And JamLembur <= Jam4 Then
                NilaiUpah = (Jam2 * Gapok / 173) * Upah1
                SplitLembur = JamLembur - Jam2
                NilaiUpah = NilaiUpah + (SplitLembur * Gapok / 173) * Upah2
            End If
         Else
            UM = 0
            Transport = "Tidak"
            NilaiUpah = 0
         End If
         Print #2, Rs!tanggal & ", " & Rs!NoProject & ", " & "`" & jamMasuk & ", " & "`" & JamKeluar & ", " & "`" & TotalJam & ", " & "`" & TotalLemburBruto & ", " & "`" & Istirahat & "," & "`" & TotalLembur & "," & NilaiUpah & ", " & UM & ", " & Transport & ""
     
         Rs.MoveNext
     Loop
     If RsSetting.State = adStateOpen Then RsSetting.Close
     If rsGaji.State = adStateOpen Then rsGaji.Close
     If rsAbsen.State = adStateOpen Then rsAbsen.Close
     If Rs.State = adStateOpen Then Rs.Close
    Close #2
End If
End Sub
Sub querydivisi()
Dim sql As String
sql = "select kd_div, nm_div " & _
    "from divisi " & _
    "where kd_bid >= 0 and kd_bid <= 20 " & _
    "order by kd_bid"
Adodc1.ConnectionString = koneksiserver3
Adodc1.RecordSource = sql
Adodc1.Refresh
Do While Not Adodc1.Recordset.EOF
    Combo1.AddItem (Adodc1.Recordset!nm_div)
    Adodc1.Recordset.MoveNext
Loop
End Sub

Private Sub Combo1_Change()
querynip
End Sub

Private Sub Combo1_Click()
querynip
End Sub

Sub querynip()
Dim sql As String
sql = "select karyawan.nip,karyawan.nama,karyawan.kd_divisi,divisi.nm_div " & _
    "from karyawan, divisi where karyawan.kd_divisi = divisi.kd_div " & _
    "and divisi.nm_div = '" & Combo1.Text & "' " & _
    "and karyawan.status <> '14' " & _
    "order by nip"
Adodc2.ConnectionString = koneksiserver3
Adodc2.RecordSource = sql
Adodc2.Refresh
Set DataGrid1.DataSource = Adodc2
DataGrid1.Columns(0).Width = 750
DataGrid1.Columns(1).Width = 4000
DataGrid1.Columns(2).Visible = False
DataGrid1.Columns(3).Visible = False
End Sub

Private Sub DataGrid1_Click()
If (Adodc2.Recordset.EOF = True) And (Adodc2.Recordset.BOF = True) Then
    MsgBox "Pilih NIP Terlebih Dahulu", vbInformation, "PT.Wiratman"
Else

    varNip = Adodc2.Recordset!NIP
    varNama = Adodc2.Recordset!Nama
   
End If
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
