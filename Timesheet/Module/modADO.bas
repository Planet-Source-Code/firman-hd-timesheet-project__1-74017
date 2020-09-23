Attribute VB_Name = "modADO"

Option Explicit
Public StrUser As String
Public StrNamaUser As String
Public StrPassword As String
Public StrNIPUser As String
Public StrIDUser As String
Public strGroup As String
Public NamaDivisi As String
Public KodeDivisi As String
Public strIDGroup As String
Public NewID As String
Public StrHakUser As String
Public Pesan As String
Public StrKodeID As String
Public PosisiPage As Integer
Public TotalRecord As Long
Public ARecord As Integer
Public AkRecord As Integer
Public StrServer As String
Public StrDatabase As String
Public StrUserDB As String
Public strPasswordDB As String
Public Adaerror As String
Public CN   As New ADODB.Connection
Public Rscek As New ADODB.Recordset
Public StrSQL As String
Public skinsFileName  As String
Public MnChat As String
Public Lrow As Integer
Public lCol As Integer
Public Type MyMenu
    IDMenu As Single
    NamaMenu As String
    Parents As Integer
    Urutan As Long
    MenuCaption As String
End Type
Public Enc  As New ClsEncrypt
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public MasterFile As String
Public MasterUpdate As String
Public StrBackup As String
Public StrRestore As String
Public MyValue As String * 255
Public i As Integer
Public PathOffice As String
'----------regional Seeting
Private Const LOCALE_SSHORTDATE = &H1F
Private Const LOCALE_STIMEFORMAT = &H1003
Private Const WM_SETTINGCHANGE = &H1A
Private Const HWND_BROADCAST = &HFFFF&
Private Const LOCALE_SDECIMAL = &HE
Private Const LOCALE_STHOUSAND = &HF
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'For tracking mouse cursor position
Public Declare Function GetCursorPos Lib "user32" _
            (lpPoint As POINTAPI) As Long
            
Public Type POINTAPI
        x As Long
        y As Long
End Type
Public ProjectUmum As String
Public NMProjectUmum As String
Public TotalJamPUmum As Double
'The code bellow is an API(Abstract Window Programming Interface) fuction.
'To build your own API fuction just click Add-Ins Menu and select Add-in MAnager and in the list Select API Text viewer then explore and have fun
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

Public Const MAX_COMPUTERNAME_LENGTH As Long = 31
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Const NERR_SUCCESS = 0&
Public Const TIME_ZONE_ID_DAYLIGHT As Long = 2

Public Type TIME_OF_DAY_INFO
  tod_elapsedt    As Long
  tod_msecs       As Long
  tod_hours       As Long
  tod_mins        As Long
  tod_secs        As Long
  tod_hunds       As Long
  tod_timezone    As Long
  tod_tinterval   As Long
  tod_day         As Long
  tod_month       As Long
  tod_year        As Long
  tod_weekday     As Long
End Type

Public Type SYSTEMTIME
   wYear         As Integer
   wMonth        As Integer
   wDayOfWeek    As Integer
   wDay          As Integer
   wHour         As Integer
   wMinute       As Integer
   wSecond       As Integer
   wMilliseconds As Integer
End Type

Public Type TIME_ZONE_INFORMATION
   bias           As Long
   StandardName(0 To 63) As Byte  'unicode (0-based)
   StandardDate   As SYSTEMTIME
   StandardBias   As Long
   DaylightName(0 To 63) As Byte  'unicode (0-based)
   DaylightDate   As SYSTEMTIME
   DaylightBias   As Long
End Type

Public Declare Function NetRemoteTOD Lib "Netapi32" _
  (UncServerName As Byte, _
   BufferPtr As Long) As Long

Public Declare Function NetApiBufferFree Lib "Netapi32" _
  (ByVal lpBuffer As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (pTo As Any, uFrom As Any, _
   ByVal lSize As Long)

Public Declare Function GetTimeZoneInformation Lib "kernel32" _
  (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
   
Public Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" _
  (lpTimeZoneInformation As TIME_ZONE_INFORMATION, _
   lpUniversalTime As SYSTEMTIME, _
   lpLocalTime As SYSTEMTIME) As Long
Public RemindTime As String
Public StrTime As String
Public Sub Main()
On Error GoTo Adaerror
Dim strFileName As String
Dim lngCount As Long
Dim Retval As String
Dim strCommand As String

If App.PrevInstance = True Then
'  MsgBox ("This program is already running.")
  End
End If
skinsFileName = GetSetting(App.EXEName, App.EXEName, "Skins")
If skinsFileName = "" Then
    Call SaveSetting(App.EXEName, App.EXEName, "Skins", "DE.skn")
    skinsFileName = GetSetting(App.EXEName, App.EXEName, "Skins")
End If
RemindTime = GetSetting(App.EXEName, App.EXEName, "Remind")
If RemindTime = "" Then
    Call SaveSetting(App.EXEName, App.EXEName, "Remind", "15:00")
    RemindTime = GetSetting(App.EXEName, App.EXEName, "Remind")
End If

PathOffice = "C:\Program Files\OpenOffice.org 2.2\program\scalc "
strFileName = String(255, 0)
lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
strFileName = Left(strFileName, lngCount)
ReadKoneksi

ConvertSysDate 'rubah regional setting
 sServer = "\\" & StrServer
server_date = GetRemoteTOD(sServer)
DisplayData server_date
If ServerTime = "01/01/1970" Then ServerTime = Now
''-------UNTUK sementara ini dapat dihapus -----------------------------
'Dim HlpFile As String
'Dim Copyhelp As String
'Copyhelp = GetSetting(App.EXEName, App.EXEName, "update")
'If Copyhelp = "" Then
'    HlpFile = Replace(MasterUpdate, "Timesheet.exe", "Mod.NFO")
'    FileCopy HlpFile, App.Path & "\Mod.NFO"
'     Call SaveSetting(App.EXEName, App.EXEName, "update", "Mod.NFO")
'End If
''--------------------------------------

ReadKoneksi
If UCase(Right(strFileName, 7)) = "VB6.EXE" Then
   FrmLogin.show vbModal
Else
  If Cekupdate(App.Path & "\" & MasterFile) = True Then
        Shell App.Path & "\AutoUpdate.EXE " & MasterFile
      End
  Else
     FrmLogin.show vbModal
  End If
End If

'If CN.State = adStateOpen Then CN.Close
'CN = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID='" & StrUserDB & "';Password = '" & strPasswordDB & "';Initial Catalog= '" & StrDatabase & "';Data Source='" & StrServer & "'"
'If CN.State = adStateClosed Then CN.Open


Exit Sub
Adaerror:
If err.Number = -2147467259 Or err.Number = -2147217843 Then
    MsgBox "Silahkan Setting Koneksi Database", vbInformation
    FrmSetKoneksi.show vbModal
Else
    MsgBox err.Description
    FrmLogin.show vbModal
End If
End Sub
Function Cekupdate(filespec) As Boolean
Dim strFileName As String
Dim lngCount As Long
Dim fso, NewFile, FileOld
Dim Neww, Old, SizeFile
Cekupdate = False
On Error Resume Next
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set NewFile = fso.GetFile(MasterUpdate)
  Set FileOld = fso.GetFile(filespec)



strFileName = String(255, 0)
lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
strFileName = Left(strFileName, lngCount)
Neww = Format(NewFile.DateLastModified, "MM/dd/yyyy")
SizeFile = NewFile.Size
Old = Format(FileOld.DateLastModified, "MM/dd/yyyy")
If Trim(Neww) <> "" Then
    If Neww <> Old Then
       Open App.Path & "\lastfile.log" For Output As #1
        Print #1, """" & Neww & """;""" & Old & """"
        Close #1
       Cekupdate = True
       
    End If
End If
End Function
Sub MapDriveServer()
Dim NetR As NETRESOURCE
Dim ErrInfo As Long
Dim MyPass As String, MyUSer As String

NetR.dwScope = RESOURCE_GLOBALNET
NetR.dwType = RESOURCETYPE_DISK
NetR.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
NetR.dwUsage = RESOURCEUSAGE_CONNECTABLE
NetR.lpLocalName = "R:" ' If undefined, Connect with no device
NetR.lpRemoteName = "\\Server2\Share"   ' Your valid share
'NetR.lpComment = "Optional Comment"
'NetR.lpProvider =    ' Leave this undefined
MyPass = "sdm"
MyUSer = "Server2\sdm"
' If the UserName and Password arguments are NULL, the user context
' for the process provides the default user name.
'ErrInfo = WNetAddConnection2(NetR, MyPass, MyUSer, _
'CONNECT_UPDATE_PROFILE)
 ErrInfo = WNetCancelConnection2(NetR.lpLocalName, _
CONNECT_UPDATE_PROFILE, False)
'If ErrInfo = NO_ERROR Then
'    ErrInfo = WNetCancelConnection2(NetR.lpLocalName, _
'    CONNECT_UPDATE_PROFILE, False)
'Else
' ErrInfo = WNetCancelConnection2(NetR.lpLocalName, _
'CONNECT_UPDATE_PROFILE, False)
'    ErrInfo = WNetAddConnection2(NetR, MyPass, MyUSer, _
'    CONNECT_UPDATE_PROFILE)
'ErrInfo = WNetCancelConnection2(NetR.lpLocalName, _
'CONNECT_UPDATE_PROFILE, False)
'End If
End Sub
Public Sub ConvertSysDate()
  Dim dwLCID As Long
  
  dwLCID = GetSystemDefaultLCID()
  
  If SetLocaleInfo(dwLCID, LOCALE_SSHORTDATE, "MM/dd/yyyy") = False Then
      Exit Sub
  End If
  If SetLocaleInfo(dwLCID, LOCALE_STIMEFORMAT, "HH:mm:ss") = False Then
      Exit Sub
  End If
  If SetLocaleInfo(dwLCID, LOCALE_SDECIMAL, ".") = False Then
       MsgBox "Error Set Decimal Separator"
       Exit Sub
    End If
    
    If SetLocaleInfo(dwLCID, LOCALE_STHOUSAND, ",") = False Then
       MsgBox "Error Set Thousand Separator"
       Exit Sub
    End If
         

  PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0

End Sub
Public Sub ReadKoneksi()
On Error GoTo ErrReadKoneksi
Dim Tmpr As String
Dim IsiText As String, IsiKoneksi() As String
Open App.Path & "\Mod.NFO" For Input As #1
Do Until EOF(1)
    Line Input #1, Tmpr
    IsiText = IsiText & Tmpr
Loop
Close #1
IsiKoneksi = Split(IsiText, ";")
StrServer = Mid(IsiKoneksi(0), 2, Len(IsiKoneksi(0)) - 2)
StrDatabase = Mid(IsiKoneksi(1), 2, Len(IsiKoneksi(1)) - 2)
MasterUpdate = Mid(IsiKoneksi(2), 2, Len(IsiKoneksi(2)) - 2)
MasterFile = Mid(IsiKoneksi(3), 2, Len(IsiKoneksi(3)) - 2)
StrUserDB = Mid(IsiKoneksi(4), 2, Len(IsiKoneksi(4)) - 2)
strPasswordDB = Mid(IsiKoneksi(5), 2, Len(IsiKoneksi(5)) - 2)

Exit Sub
ErrReadKoneksi:
Close #1
Open App.Path & "\Mod.NFO" For Append As #1
Print #1, ""; ""; ""; ";"
Close #1
On Error GoTo 0
End Sub

Public Function PerintahExecute(Ssql As String)
On Error GoTo Adaerror

Dim cmd As New ADODB.Command
If CN.State = adStateClosed Then CN.Open
    cmd.ActiveConnection = CN
    cmd.CommandText = Ssql
    cmd.Execute
Exit Function
Adaerror:
MsgBox err.Description
End Function

Public Function GetNomorID(NamaTable As String, Optional Update As Boolean = False) As String
Dim ID(0 To 9) As String
Dim LastID As String
Dim Rscek As New ADODB.Recordset
Dim cmCek As New ADODB.Command
Dim i As Integer
Dim StrSQL As String
On Error GoTo Adaerror
Const Katalog = "0123456789abcdefghijklmnopqrstuvwxyz"
If CN.State = adStateClosed Then CN.Open
If Rscek.State = adStateOpen Then Rscek.Close
StrSQL = "Select " & NamaTable & " from TblKodeID"
Rscek.Open StrSQL, CN, adOpenStatic

If Rscek.EOF Then
    LastID = ""
Else
   If IsNull(Rscek.Fields(NamaTable).Value) = True Then
        LastID = ""
    Else
       LastID = Rscek.Fields(NamaTable).Value
    End If
End If


'------------ Hitung
If LastID <> "" Then
    For i = 0 To 9
        ID(i) = Mid(LastID, i + 1, 1)
    Next

    For i = 9 To 1 Step -1
        If InStr(1, Katalog, ID(i)) = 36 Then
            ID(i) = "0"
        Else
            ID(i) = Mid(Katalog, InStr(1, Katalog, ID(i)) + 1, 1)
            Exit For
        End If
    Next
    
    NewID = ""
    For i = 0 To 9
        NewID = NewID + ID(i)
    Next
Else
    NewID = "0000000000"
End If

'--------- Update ID
If Update = False Then
    If LastID <> "" Then
        cmCek.ActiveConnection = CN
'        StrSQL = "Update TblKodeID set " & NamaTable & "='" & NewID & "' Where " & NamaTable & "='" & LastID & "'"
        StrSQL = "Update TblKodeID set " & NamaTable & "='" & NewID & "'"
        cmCek.CommandText = StrSQL
        cmCek.Execute
    Else
        cmCek.ActiveConnection = CN
'        StrSQL = "Update TblKodeID set " & NamaTable & "='" & NewID & "' Where " & NamaTable & "='" & LastID & "'"
        StrSQL = "Update TblKodeID set " & NamaTable & "='" & NewID & "'"
        cmCek.CommandText = StrSQL
        cmCek.Execute
    End If
End If

If Rscek.State = adStateOpen Then Rscek.Close

Set cmCek.ActiveConnection = Nothing
Exit Function
Adaerror:
MsgBox "Belum Didaftarkan Di TblKodeID", vbCritical
End Function

Function getAutoNo(nKode As String, Optional Update As Boolean = False) As String
On Error GoTo Adaerror
Dim rc As New ADODB.Recordset
If CN.State = adStateClosed Then CN.Open
If rc.State = adStateOpen Then rc.Close
StrSQL = "SELECT * FROM TblNomorUrut where IDNomor='" & nKode & "'"
rc.Open StrSQL, CN, adOpenStatic
'If hErr = "" Then
   If Not rc.EOF Then
      Dim hFormat As String
      Dim ChangeNo As String
      Dim LastYear As String
      Dim LastMonth As String
      Dim LenNo As String
      Dim LastNo As String
      Dim hState As Boolean

      hFormat = IIf(IsNull(rc!FormatNomor), "", rc!FormatNomor)
      ChangeNo = IIf(IsNull(rc!UbahNomorJika), "", rc!UbahNomorJika)
      LastMonth = IIf(IsNull(rc!BulanTerakhir), "", rc!BulanTerakhir)
      LastYear = IIf(IsNull(rc!TahunTerakhir), "", rc!TahunTerakhir)
      LastNo = IIf(IsNull(rc!NomorTerakhir), "", rc!NomorTerakhir)
      LenNo = IIf(IsNull(rc!SizeNomor), "", rc!SizeNomor)
      If ChangeNo <> "" Then
             Select Case LCase(ChangeNo)
                    Case "{bln}"
                        If LastMonth <> Format(Date, "mm") Then hState = True
                        If Update Then StrSQL = "UPDATE TblNomorUrut SET BulanTerakhir='" & Format(Date, "mm") & "' Where IDNomor='" & nKode & "'"
                    
                    Case "{thn}"
                        If LastYear <> Format(Date, "yy") Then hState = True
                        If Update Then StrSQL = "UPDATE TblNomorUrut SET TahunTerakhir='" & Format(Date, "yy") & "' Where IDNomor='" & nKode & "'"
                    
                    Case "{tahun}"
                        If LastYear <> Format(Date, "yyyy") Then hState = True
                        If Update Then StrSQL = "UPDATE TblNomorUrut SET TahunTerakhir='" & Format(Date, "yyyy") & "' Where IDNomor='" & nKode & "'"
                    
                    Case "{bln}{thn}"
                        If LastMonth & LastYear <> Format(Date, "mmyy") Then hState = True
                        If Update Then StrSQL = "UPDATE TblNomorUrut SET BulanTerakhir='" & Format(Date, "mm") & "', TahunTerakhir='" & Format(Date, "yy") & "' Where IDNomor='" & nKode & "'"
                    
                    Case "{thn}{bln}"
                        If LastYear & LastMonth <> Format(Date, "yymm") Then hState = True
                        If Update Then StrSQL = "UPDATE TblNomorUrut SET TahunTerakhir='" & Format(Date, "yy") & "', BulanTerakhir='" & Format(Date, "mm") & "' Where IDNomor='" & nKode & "'"
                    
                    Case "{tahun}{bln}"
                        If LastYear & LastMonth <> Format(Date, "yyyymm") Then hState = True
                        If Update Then StrSQL = "UPDATE TblNomorUrut SET LastYear='" & Format(Date, "yyyy") & "', LastMonth='" & Format(Date, "mm") & "' Where IDNomor='" & nKode & "'"
                    
                    Case "{bln}{tahun}"
                        If LastMonth & LastYear <> Format(Date, "mmyyyy") Then hState = True
                        If Update Then StrSQL = "UPDATE TblNomorUrut SET LastMonth='" & Format(Date, "mm") & "', LastYear='" & Format(Date, "yyyy") & "' Where IDNomor='" & nKode & "'"
            End Select
                   If Update Then PerintahExecute (StrSQL)
            
            If hState = False Then
               LastNo = Val(LastNo) + 1
            Else
               LastNo = 1
            End If
      Else
         LastNo = Val(LastNo) + 1
      End If
        LastNo = String(LenNo - Len(CStr(LastNo)), "0") & LastNo
        hFormat = Replace(hFormat, "{thn}", Format(Date, "yy"))
        hFormat = Replace(hFormat, "{tahun}", Format(Date, "yyyy"))
        hFormat = Replace(hFormat, "{bln}", Format(Date, "mm"))
        hFormat = Replace(hFormat, "{nourut}", LastNo)
        
        If Update Then
          StrSQL = "UPDATE TblNomorUrut SET NomorTerakhir='" & LastNo & "' Where IDNomor='" & nKode & "'"
          PerintahExecute (StrSQL)
        End If
        getAutoNo = hFormat
       
   End If
   rc.Close
'   If CN.State = adStateOpen Then CN.Close
Exit Function
Adaerror:
MsgBox err.Description
End Function

Public Function CekHakUser(StrGroupUSer As String, NamaMenu As String) As Long
Dim Rscek As New ADODB.Recordset

Dim StrSQL As String

If CN.State = adStateClosed Then CN.Open
If Rscek.State = adStateOpen Then Rscek.Close
StrSQL = "select `tblgroup_user_detail`.`IDMenu` AS `IDMenu`,`tblgroup_user_detail`.`IDGroup` AS `IDGroup`,`tblgroup_user_detail`.`Group_Permissions` AS `Group_Permissions`,`tblgroup_menu`.`Nama_Menu` AS `Nama_Menu`,`tbldata_user`.`Nama_User` AS `Nama_User`,`tbldata_user`.`IDUser` AS `IDUser` from ((`tblgroup_user_detail` join `tblgroup_menu` on((`tblgroup_user_detail`.`IDMenu` = `tblgroup_menu`.`IDMenu`))) join `tbldata_user` on((`tblgroup_user_detail`.`IDGroup` = `tbldata_user`.`IDGroup`)))"
StrSQL = StrSQL & "where IDUser = '" & StrNIPUser & "' and IDGroup = '" & strIDGroup & "' And Nama_Menu = '" & NamaMenu & "'"
Rscek.Open StrSQL, CN, adOpenStatic
If Not Rscek.EOF Then
   StrHakUser = Rscek!Group_Permissions
End If
If Rscek.State = adStateOpen Then Rscek.Close
Set Rscek.ActiveConnection = Nothing
'If CN.State = adStateOpen Then CN.Close
End Function
Public Sub CloseDB()
    'Close the connection
    CN.Close
    Set CN = Nothing
End Sub

Public Sub ShowReport(rpt As Object, strlink As String, StrSQL As String)
   
    rpt.Tag = strlink
    
    rpt.DataControl1.Source = StrSQL
    
    rpt.show vbModal
End Sub

Public Function CekCurek(status As String, Ctr As VSFlexGrid) As Boolean
Dim i As Long, J As Long
Dim Keterangan As String

CekCurek = False
With Ctr
    J = 0
    For i = 1 To .Rows - 2
        If .TextMatrix(i, 1) = "-1" Then
            CekCurek = True
            J = J + 1
        End If
    Next
    
    If J > 0 Then
        Exit Function
    Else
        MsgBox "Check List Di Kolom 'Do' pada baris data yang akan di" & status & " ...", vbCritical
       CekCurek = False
    End If
        
End With

End Function

Public Sub CekMenu()
'--- This sub check Permission from vTblGroupUserDetail
'--- Note : Administrator get Full permission !!!

Dim Rscek As New ADODB.Recordset
Dim MA As CONTROL
Dim StrMenu As String

On Error Resume Next

For Each MA In MDIMENU.Controls
  If InStr(1, MA.name, "LvWIn") = 0 And InStr(1, MA.name, "Skin1") = 0 And InStr(1, MA.name, "PicSeparatorKanan") = 0 And InStr(1, MA.name, "ImgOpen") = 0 And InStr(1, MA.name, "ImgClose") = 0 And InStr(1, MA.name, "LvMenu") = 0 And InStr(1, MA.name, "PicRight") = 0 And MA.name <> "" And Not (TypeOf MA Is Image) And InStr(1, MA.name, "StatusBar") = 0 And InStr(1, MA.name, "Timer") = 0 And Not (TypeOf MA Is Label) And Not (TypeOf MA Is ImageList) And Not (TypeOf MA Is Frame) And InStr(1, MA.name, "Pic") = 0 Then
'    If InStr(1, LCase(StrUser), "admin") = 0 Then
        If Rscek.State = adStateOpen Then Rscek.Close
'        Debug.Print MA.Name
        StrMenu = MA.name
        StrSQL = " SELECT tbldata_user.NIP, Tblgroup_user.namagroup,"
        StrSQL = StrSQL & " tblgroup_user_detail.idgroup, tblgroup_user_detail.idmenu,"
        StrSQL = StrSQL & " tblgroup_user_detail.group_permission,tblgroup_user_detail.nama_menu"
        StrSQL = StrSQL & " FROM tbldata_user INNER JOIN Tblgroup_user ON"
        StrSQL = StrSQL & " tbldata_user.Type = Tblgroup_user.namagroup INNER JOIN"
        StrSQL = StrSQL & " tblgroup_user_detail ON Tblgroup_user.IDGroup = tblgroup_user_detail.IDGroup"
        StrSQL = StrSQL & " Where tbldata_user.NIP = '" & StrUser & "' And tblgroup_user_detail.Nama_Menu = '" & StrMenu & "'"
 
       Rscek.Open StrSQL, CN, adOpenStatic
        
        If Not Rscek.EOF Then
    '       StrMenu = Rscek!Group_Permissions
                If Left(Rscek!Group_Permission, 1) = "1" Then
                    MA.Enabled = True
                Else
                    MA.Enabled = False
                End If
        Else
            If MA.name = "PicAdvisor" Or MA.name = "picSeparator" Or MA.name = "picLeft" Or MA.name = "lvWin" Or MA.name = "PicLine" Or MA.name = "picContainer" Then
                MA.Enabled = True
            Else
                MA.Enabled = False
            End If
        End If
'    Else
'        MA.Enabled = True
'
'    End If
        
        If StrMenu = "MnuGroupMenu" Or StrMenu = "mnuSO" Or StrMenu = "mnuRecA" Or StrMenu = "InjectData" Or StrMenu = "MnuRBatal" Or StrMenu = "MnuRView" Then MA.Enabled = False
        If StrMenu = "mnuRNew" Or StrMenu = "MnuRPrint" Or StrMenu = "mnuRefresh" Or StrMenu = "mnuRPrint" Or StrMenu = "mnuRSearch" Or StrMenu = "mnuRDelete" Or StrMenu = "MnuRSelect" Or StrMenu = "mnuRAC" Then MA.Enabled = True
 
 End If
Next
If Rscek.State = adStateOpen Then Rscek.Close

'MDIMAIN.StatusBar1.Visible = True
End Sub

Public Sub IsiGrid(ByRef Ctr As VSFlexGrid, ByRef Rs As Recordset, ByVal pos_start As Long, ByVal pos_end As Long, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional match_field As String, Optional match_str As String, Optional match_ico As Byte, Optional srcHiddenField As String)
On Error Resume Next
Dim lRows As Long
Dim lCol As Long
Dim lCols As Long
Dim Lrow As Long
Dim sColor As String
Dim SelRecord As Long
    lCols = Rs.Fields.Count
        With Ctr
            .Redraw = flexRDNone
            .Rows = 1
            .Rows = 2
            .Cols = 2 + lCols
            .Row = 0
            For lCol = 0 To lCols - 1
                .Col = 2 + lCol
                .Text = Rs(lCol).name
                .ColWidth(lCol) = 1500
                .FixedAlignment(lCol) = flexAlignCenterCenter
            Next
    If Rs.RecordCount > 0 Then
        Rs.AbsolutePosition = pos_start
            Lrow = 1
     For SelRecord = pos_start To pos_end
                For lCol = 0 To lCols - 1
                    .Col = 2 + lCol
                    .Row = Lrow
                    .Text = Rs.Fields(lCol).Value
                Next
                 .TextMatrix(Lrow, 0) = SelRecord
                 .TextMatrix(Lrow, 1) = ""
                 Rs.MoveNext
                 Lrow = Lrow + 1
                .Rows = Lrow + 1
    Next
Else
           Ctr.Rows = 1
           Ctr.Rows = 2
    End If
           .ColDataType(1) = flexDTBoolean
           .ColWidth(1) = 350
           .TextMatrix(0, 1) = "Do"
           .ColWidth(0) = 500
           .ColWidth(2) = 0
           .TextMatrix(0, 0) = "No"
           .Redraw = flexRDBuffered
           .SetFocus
End With

End Sub
