Attribute VB_Name = "ModCompName"
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
'Public Const SOCKET_ERROR = -1
Public Const WSADESCRIPTION_LEN = 256
Public Const WSASYS_STATUS_LEN = 128
Public IPPc As String
Public ServerTime As Date
'Public Type HOSTENT
'    hName As Long
'    hAliases As Long
'    hAddrType As Integer
'    hLength As Integer
'    hAddrList As Long
'End Type


Public Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADESCRIPTION_LEN) As Byte
    szSystemStatus(0 To WSASYS_STATUS_LEN) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type


Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long


Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal _
    wVersionRequired&, lpWSADATA As WSAData) As Long


Public Declare Function WSACleanup Lib "wsock32.dll" () As Long


Public Declare Function gethostname Lib "wsock32.dll" (ByVal hostname$, _
    ByVal HostLen As Long) As Long


Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal _
    hostname$) As Long


Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal _
    hpvSource&, ByVal cbCopy&)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public server_date As TIME_OF_DAY_INFO
 Public sServer As String
Public Function UserName() As String

    Dim llReturn As Long
    Dim lsUserName As String
    Dim lsBuffer As String
    
    lsUserName = ""
    lsBuffer = Space$(255)
    llReturn = GetUserName(lsBuffer, 255)
    
    
    If llReturn Then
       lsUserName = Left$(lsBuffer, InStr(lsBuffer, Chr(0)) - 1)
    End If
    
    UserName = lsUserName
End Function
Public Function ComputerName() As String
  Dim lsBuffer As String
  Dim llReturn As Long
  Dim lsName As String
 
  lsName = ""
  lsBuffer = Space$(255)
  llReturn = GetComputerName(lsBuffer, 255)
  
  If llReturn Then
        lsName = Left$(lsBuffer, InStr(lsBuffer, Chr(0)) - 1)
  End If
  
  ComputerName = lsName
End Function


Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function

Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function

Sub SocketsInitialize()
    Dim WSAD As WSAData
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String
    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)


    If iReturn <> 0 Then
        MsgBox "Winsock.dll is Not responding."
        End
    End If
    If lobyte(WSAD.wVersion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wVersion) = _
    WS_VERSION_MAJOR And hibyte(WSAD.wVersion) < WS_VERSION_MINOR) Then
    sHighByte = Trim$(Str$(hibyte(WSAD.wVersion)))
    sLowByte = Trim$(Str$(lobyte(WSAD.wVersion)))
    sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
    sMsg = sMsg & " is Not supported by winsock.dll "
    MsgBox sMsg
    End
End If
'iMaxSockets is not used in winsock 2. S
'     o the following check is only
'necessary for winsock 1. If winsock 2 i
'     s requested,
'the following check can be skipped.


If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
    sMsg = "This application requires a minimum of "
    sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
    MsgBox sMsg
    End
End If
End Sub

Sub SocketsCleanup()
    Dim lReturn As Long
    lReturn = WSACleanup()


    If lReturn <> 0 Then
        MsgBox "Socket Error " & Trim$(Str$(lReturn)) & " occurred In Cleanup "
        End
    End If
End Sub

Public Function GetTheIP()
Call SocketsInitialize
    
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim Host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String


    If gethostname(hostname, 256) = SOCKET_ERROR Then
        MsgBox "Windows Sockets Error " & Str(WSAGetLastError())
        Exit Function
    Else
        hostname = Trim$(hostname)
    End If
    hostent_addr = gethostbyname(hostname)


    If hostent_addr = 0 Then
        MsgBox "Winsock.dll is Not responding."
        Exit Function
    End If
    RtlMoveMemory Host, hostent_addr, LenB(Host)
    RtlMoveMemory hostip_addr, Host.hAddrList, 4

    Do
        ReDim temp_ip_address(1 To Host.hLength)
        RtlMoveMemory temp_ip_address(1), hostip_addr, Host.hLength


        For i = 1 To Host.hLength
            ip_address = ip_address & temp_ip_address(i) & "."
        Next
        ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
        IPPc = ip_address
        Host.hAddrList = Host.hAddrList + LenB(Host.hAddrList)
        RtlMoveMemory hostip_addr, Host.hAddrList, 4
    Loop While (hostip_addr <> 0)
End Function
Public Function GetRemoteTOD(ByVal sServer As String) As TIME_OF_DAY_INFO

   Dim success       As Long
   Dim bServer()     As Byte
   Dim tod           As TIME_OF_DAY_INFO
   Dim systime_utc   As SYSTEMTIME
   Dim systime_local As SYSTEMTIME
   Dim tzi           As TIME_ZONE_INFORMATION
   Dim bufptr        As Long

   If sServer <> vbNullChar Then
      If Left$(sServer, 2) <> "\\" Then
            bServer = "\\" & sServer & vbNullChar
      Else: bServer = sServer & vbNullChar
      End If
      
   Else
   
      bServer = sServer & vbNullChar
   
   End If
   
   If NetRemoteTOD(bServer(0), bufptr) = NERR_SUCCESS Then

      CopyMemory tod, ByVal bufptr, LenB(tod)
      Call GetTimeZoneInformation(tzi)

      With systime_utc
         .wDay = tod.tod_day
         .wDayOfWeek = tod.tod_weekday
         .wMonth = tod.tod_month
         .wYear = tod.tod_year
         .wHour = tod.tod_hours
         .wMinute = tod.tod_mins
         .wSecond = tod.tod_secs
      End With
      
      Call SystemTimeToTzSpecificLocalTime(tzi, systime_utc, systime_local)
 
      With tod
         .tod_mins = systime_local.wMinute
         .tod_hours = systime_local.wHour
         .tod_secs = systime_local.wSecond
         .tod_day = systime_local.wDay
         .tod_month = systime_local.wMonth
         .tod_year = systime_local.wYear
         .tod_weekday = systime_local.wDayOfWeek
      End With
      
   End If
   
   Call NetApiBufferFree(bufptr)
   GetRemoteTOD = tod

End Function
Public Sub DisplayData(server_date As TIME_OF_DAY_INFO)

   Dim newtime  As Date
   
        newtime = DateAdd("s", server_date.tod_elapsedt, #1/1/1970#)
        newtime = DateAdd("n", -server_date.tod_timezone, newtime)
   
       ServerTime = newtime
End Sub

