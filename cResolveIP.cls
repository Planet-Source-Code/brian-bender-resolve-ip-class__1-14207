VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cResolveIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const WSADescription_Len = 256
Const WSASYS_Status_Len = 128

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Private Declare Function WSACleanup Lib "wsock32" () As Long
Private Declare Function WSAStartup Lib "wsock32" (ByVal VersionReq As Long, WSADataReturn As WSADATA) As Long
Private Declare Function WSAGetLastError Lib "wsock32" () As Long
Private Declare Function gethostbyaddr Lib "wsock32" (addr As Long, addrLen As Long, addrType As Long) As Long
Private Declare Function GetHostByName Lib "wsock32" Alias "gethostbyname" (ByVal hostname As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
    
Public Function IsIP(ByVal strIP As String) As Boolean
    '-- Checks if string is valid IP address
    On Error Resume Next
    Dim t As String
    Dim s As String
    Dim i As Integer
    s = strIP
    While InStr(s, ".") <> 0
        t = Left(s, InStr(s, ".") - 1)
        If IsNumeric(t) And Val(t) >= 0 And Val(t) <= 255 Then s = Mid(s, InStr(s, ".") + 1) _
    Else Exit Function
        i = i + 1
    Wend
    t = s
    If IsNumeric(t) And InStr(t, ".") = 0 And Len(t) = Len(Trim(Str(Val(t)))) And _
    Val(t) >= 0 And Val(t) <= 255 And strIP <> "255.255.255.255" And i = 3 Then IsIP = True
    If Err.Number > 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
        Err.Clear
    End If
End Function

Public Function MakeIP(strIP As String) As Long
    On Error Resume Next
    '-- Converts IP address from string to sin addr
    Dim lIP As Long
    lIP = Left(strIP, InStr(strIP, ".") - 1)
    strIP = Mid(strIP, InStr(strIP, ".") + 1)
    lIP = lIP + Left(strIP, InStr(strIP, ".") - 1) * 256
    strIP = Mid(strIP, InStr(strIP, ".") + 1)
    lIP = lIP + Left(strIP, InStr(strIP, ".") - 1) * 256 * 256
    strIP = Mid(strIP, InStr(strIP, ".") + 1)
    If strIP < 128 Then
        lIP = lIP + strIP * 256 * 256 * 256
    Else
        lIP = lIP + (strIP - 256) * 256 * 256 * 256
    End If
    MakeIP = lIP
    If Err.Number > 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
        Err.Clear
    End If
End Function

Function GetHostNameByIPAddress(strAddr As String) As String
    On Error Resume Next
    '-- Resolves IP address to host name
    Dim nRet As Long
    Dim lIP As Long
    Dim strHost As String * 255
    Dim strTemp As String
    Dim hst As HOSTENT
    If IsIP(strAddr) Then
        lIP = MakeIP(strAddr)
        nRet = gethostbyaddr(lIP, 4, 2)
        If nRet <> 0 Then
            RtlMoveMemory hst, nRet, Len(hst)
            RtlMoveMemory ByVal strHost, hst.hName, 255
            strTemp = strHost
            If InStr(strTemp, Chr(10)) <> 0 Then strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
            strTemp = Trim(strTemp)
            GetHostNameByIPAddress = strTemp
        Else
            Err.Raise vbObjectError + 501, "Get Host Name By IP Address", "Host name not found."
            Exit Function
        End If
    Else
        Err.Raise vbObjectError + 502, "Get Host Name By IP Address", "Invalid IP address."
        Exit Function
    End If
    If Err.Number > 0 Then
        Err.Raise Err.Number, "Get Host Name By IP Address", Err.Description
    End If
End Function

Function GetIPAddressbyHostName(ByVal strHost As String)
    '--Resolves host name to IP address
    On Error Resume Next
    Dim hostent_addr As Long
    Dim hst As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String
    If IsIP(strHost) Then
        GetIPAddressbyHostName = strHost
        Exit Function
    End If
    hostent_addr = GetHostByName(strHost)
    If hostent_addr = 0 Then
        Err.Raise vbObjectError + 500, "Get IP Address by HostName", "Cannot resolve host"
        Exit Function
    End If
    RtlMoveMemory hst, hostent_addr, LenB(hst)
    RtlMoveMemory hostip_addr, hst.hAddrList, 4
    ReDim temp_ip_address(1 To hst.hLength)
    RtlMoveMemory temp_ip_address(1), hostip_addr, hst.hLength
    For i = 1 To hst.hLength
        ip_address = ip_address & temp_ip_address(i) & "."
    Next
    ip_address = Mid(ip_address, 1, Len(ip_address) - 1)
    GetIPAddressbyHostName = ip_address
    If Err.Number > 0 Then
        Err.Raise Err.Number, "Get IP Address by Host Name", Err.Description
    End If
End Function

Private Sub Class_Initialize()
    Dim udtWSAData As WSADATA
    If WSAStartup(257, udtWSAData) Then
        Err.Raise Err.Number, "Initialize Winsock", Err.Description
    End If
End Sub

Private Sub Class_Terminate()
    WSACleanup
End Sub
