Attribute VB_Name = "modInternetFile"
'----------------------------------------------------------------------------------------
'
'   Module          : modInternetFile
'   Description     : Reading Internet Files using High-level Internet API
'   Author          : Eswar Santhosh
'   Last Updated    : 25th August, 2001
'
'   Copyright Info  :
'
'   This Class module is provided AS-IS. This Class module can be used as a part of a compiled
'   executable whether freeware or not. This Class module may not be posted to any web site
'   or BBS or any redistributable media like CD-ROM without the consent of the author.
'
'   Web Site : http://www.quickvb.com/
'
'   e-mail   : esanthosh@quickvb.com
'
'
'   Revision History :
'
'---------------------------------------------------------------------------------------------

Option Explicit

'
' API Constants
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' InternetOpen
Private Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
Private Const INTERNET_OPEN_TYPE_PROXY As Long = 3

' InternetConnect
Private Const INTERNET_INVALID_PORT_NUMBER As Long = 0

' HTTPRequest or FTPOpenFile
Private Const INTERNET_FLAG_KEEP_CONNECTION As Long = &H400000
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

' Internet Schemes
Private Const INTERNET_SCHEME_PARTIAL As Long = -2
Private Const INTERNET_SCHEME_UNKNOWN As Long = -1
Private Const INTERNET_SCHEME_DEFAULT As Long = 0
Private Const INTERNET_SCHEME_FTP  As Long = 1
Private Const INTERNET_SCHEME_GOPHER  As Long = 2
Private Const INTERNET_SCHEME_HTTP  As Long = 3
Private Const INTERNET_SCHEME_HTTPS As Long = 4
Private Const INTERNET_SCHEME_FILE As Long = 5
Private Const INTERNET_SCHEME_NEWS As Long = 6
Private Const INTERNET_SCHEME_MAILTO As Long = 7
Private Const INTERNET_SCHEME_SOCKS As Long = 8
Private Const INTERNET_SCHEME_JAVASCRIPT As Long = 9
Private Const INTERNET_SCHEME_VBSCRIPT As Long = 10
Private Const INTERNET_SCHEME_FIRST  As Long = INTERNET_SCHEME_FTP
Private Const INTERNET_SCHEME_LAST As Long = INTERNET_SCHEME_VBSCRIPT
    
' Service Types
Private Const INTERNET_SERVICE_FTP As Long = 1
Private Const INTERNET_SERVICE_GOPHER  As Long = 2
Private Const INTERNET_SERVICE_HTTP As Long = 3

' Query Info
Private Const HTTP_QUERY_CONTENT_LENGTH As Long = 5
Private Const HTTP_QUERY_FLAG_NUMBER As Long = &H20000000
Private Const HTTP_QUERY_CONTENT_TYPE As Long = 1

'
' API Types
'
Private Type URL_COMPONENTS
    dwStructSize As Long
    lpszScheme As String
    dwSchemeLength As Long
    nScheme As Long
    lpszHostName As String
    dwHostNameLength As Long
    INTERNET_PORT As Long
    lpszUserName As String
    dwUserNameLength As Long
    lpszPassword As String
    dwPasswordLength As Long
    lpszUrlPath As String
    dwUrlPathLength As Long
    lpszExtraInfo As String
    dwExtraInfoLength As Long
End Type

'
' API Declarations
'
Private Declare Function InternetCrackUrl Lib "wininet.dll" Alias "InternetCrackUrlA" _
(ByVal lpszUrl As String, ByVal dwUrlLength As Long, ByVal dwFlags As Long, _
ByRef lpUrlComponents As URL_COMPONENTS) As Long

Private Declare Function InternetOpen Lib "wininet.dll" Alias _
   "InternetOpenA" (ByVal sAgent As String, _
   ByVal lAccessType As Long, ByVal sProxyName As String, _
   ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
   
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" _
(ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, _
ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
   
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
      (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
     ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, _
    ByVal lFlags As Long, ByVal lContext As Long) As Long

Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
    (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
    ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
   
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" _
  (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, _
   ByVal sOptional As String, ByVal lOptionalLength As Long) As Integer

Private Declare Function InternetReadFile Lib "wininet.dll" _
    (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
    lNumberOfBytesRead As Long) As Integer

Private Declare Function InternetCloseHandle Lib "wininet.dll" _
   (ByVal hInet As Long) As Integer

Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
(ByVal hRequest As Long, ByVal dwInfoLevel As Long, _
lpBuffer As Any, ByRef lpdwBufferLength As Long, ByRef lpdwIndex As Long) _
As Long

' Support Function
Private Declare Sub PathStripPath Lib "shlwapi" Alias "PathStripPathA" _
(ByVal sPath As String)

'
' Public Enums
'
Public Enum eRequestMethod
    eInetRequestGet
    eInetRequestPost
End Enum

'
' Local Constants
'
Private Const sUserAgent As String = "Mozilla/4.0 (compatible; MSIE 5.5; Windows NT 5.0)"
Private Const sPostHeader As String = "Content-Type: application/x-www-form-urlencoded"
Private Const sReferrer As String = ""
Private Const sHTTPVersion As String = "HTTP/1.1"
Private Const sNonFileChars As String = "/,\,:," & """" & "*,?,<,>,|,"

Public Function ReadURL(ByVal sURL As String, _
ByRef sContentType As String, _
Optional ByVal sRequestData As String, _
Optional ByVal nChunkSize As Long = 1024, _
Optional ByVal eMethod As eRequestMethod = eInetRequestPost) As String
'
' Returns the Response to the Posted Request
'

' Path Components
Dim tURL As URL_COMPONENTS
Dim sHost As String
Dim sPath As String
Dim sUserName As String
Dim sPassword As String
Dim sExtraInfo As String
Dim nPort As Long
Dim nScheme As Long

' Connection and File Objects
Dim nService As Long
Dim sVerb As String
Dim sHeader As String
Dim nHeaderLength As Long

Dim hSession As Long
Dim hConnection As Long
Dim hRequest As Long ' File Handle : HTTP Request Handle, FTP File Handle, Gopher File Handle
Dim hSentRequest As Long

Dim sFileChunk As String
Dim nBytesRead As Long

With tURL
    .dwStructSize = Len(tURL)
        
    .lpszScheme = Space$(255)
    .dwSchemeLength = 255
    
    .lpszHostName = Space$(255)
    .dwHostNameLength = 255
    
    .lpszUrlPath = Space$(1024)
    .dwUrlPathLength = 1024
    
    .lpszUserName = Space$(255)
    .dwUserNameLength = 255
    
    .lpszPassword = Space$(255)
    .dwPasswordLength = 255
    
    .lpszExtraInfo = Space$(1024)
    .dwExtraInfoLength = 1024
End With

If InternetCrackUrl(sURL, Len(sURL), 0&, tURL) <> 0 Then
    sHost = Left$(tURL.lpszHostName, tURL.dwHostNameLength)
    sPath = Left$(tURL.lpszUrlPath, tURL.dwUrlPathLength)
    sUserName = Left$(tURL.lpszUserName, tURL.dwUserNameLength)
    sPassword = Left$(tURL.lpszPassword, tURL.dwPasswordLength)
    sExtraInfo = Left$(tURL.lpszExtraInfo, tURL.dwExtraInfoLength)
    nPort = tURL.INTERNET_PORT
    
    nScheme = tURL.nScheme
    
    Select Case nScheme
    Case INTERNET_SCHEME_FTP
        nService = INTERNET_SERVICE_FTP
        
    Case INTERNET_SCHEME_GOPHER
        nService = INTERNET_SERVICE_GOPHER
        
    Case Else
        nService = INTERNET_SERVICE_HTTP
    End Select
Else
    Exit Function
End If

hSession = InternetOpen(sUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, "", "", 0&)

If hSession = 0 Then Exit Function

hConnection = InternetConnect(hSession, sHost, nPort, sUserName, sPassword, nService, 0&, 0&)

If hConnection = 0 Then
    Call InternetCloseHandle(hSession)
    Exit Function
End If

If (nService = INTERNET_SERVICE_HTTP) And _
   (Trim$(sRequestData) <> "") Then
   
    If eMethod = eInetRequestPost Then
        sVerb = "POST"
        sHeader = sPostHeader
        nHeaderLength = Len(sPostHeader)
    Else
        sVerb = "GET"
        sPath = sPath & "?" & sRequestData
        sRequestData = ""
    End If
    
    hRequest = HttpOpenRequest(hConnection, sVerb, sPath, sHTTPVersion, sReferrer, 0&, INTERNET_FLAG_RELOAD, 0&)
    
    If hRequest = 0 Then
        Call InternetCloseHandle(hConnection)
        Call InternetCloseHandle(hSession)
        
        Exit Function
    End If
    
    hSentRequest = HttpSendRequest(hRequest, sHeader, nHeaderLength, sRequestData, Len(sRequestData))
    
    If hSentRequest = 0 Then
        Call InternetCloseHandle(hRequest)
        Call InternetCloseHandle(hConnection)
        Call InternetCloseHandle(hSession)
        Exit Function
    End If
Else
    hRequest = InternetOpenUrl(hSession, sURL, vbNullString, 0&, INTERNET_FLAG_RELOAD, 0&)
End If

If hRequest <> 0 Then
    sContentType = GetContentType(hRequest)
    
    sFileChunk = Space$(nChunkSize)
    
    Do While InternetReadFile(hRequest, sFileChunk, nChunkSize, nBytesRead) <> 0
        If nBytesRead = 0 Then Exit Do
        
        sFileChunk = Left$(sFileChunk, nBytesRead)
        ReadURL = ReadURL & sFileChunk
        
        sFileChunk = Space$(nChunkSize)
        nBytesRead = nChunkSize
    Loop
End If

Call InternetCloseHandle(hRequest)
Call InternetCloseHandle(hConnection)
Call InternetCloseHandle(hSession)
End Function

Public Function GetLocalFileNameFromURL(ByVal sURL As String) As String
'
' Strips the Host, Protocol and Returns the Path of the File.
' Used here for generating Local File names
'
Dim tURL As URL_COMPONENTS
Dim sPath As String
Dim iPos As Long
Dim bInLoop As Boolean
Dim iLastPos As Long

With tURL
    .dwStructSize = Len(tURL)
    .lpszUrlPath = Space$(1024)
    .dwUrlPathLength = 1024
End With

If InternetCrackUrl(sURL, Len(sURL), 0&, tURL) <> 0 Then
    sPath = Left$(tURL.lpszUrlPath, tURL.dwUrlPathLength)
    
    Call PathStripPath(sPath)
    sPath = TrimNull(sPath)
    GetLocalFileNameFromURL = sPath

    iPos = 1
    iLastPos = 0

    Do While iPos <= Len(sPath)
        If Not bInLoop Then
            If InStr(1, sNonFileChars, Mid$(sPath, iPos, 1) & ",", vbTextCompare) > 0 Then
                bInLoop = True
                iLastPos = iPos
            End If
        Else
            ' Exit when one of chars does not match specified chars
            Do While InStr(1, sNonFileChars, Mid$(sPath, iPos, 1) & ",", vbTextCompare) > 0
                iPos = iPos + 1
                
                If iPos > Len(sPath) Then
                    Exit Do
                End If
            Loop
            
            bInLoop = False
            
            sPath = Left$(sPath, iLastPos - 1) & "_" & Mid$(sPath, iPos)
            
            iPos = iLastPos
        End If
        
        iPos = iPos + 1
    Loop
        
    GetLocalFileNameFromURL = sPath
End If

End Function

Private Function GetContentType(ByVal nHandle As Long) As String
'
' nHandle must be the one returned by InternetOpenURL or HttpOpenRequest Function.
' Returns type of the Requested File
'
Dim sType As String
Dim nLength As Long
Dim nIndex As Long

sType = Space$(255)
nLength = 255

If HttpQueryInfo(nHandle, HTTP_QUERY_CONTENT_TYPE, ByVal sType, nLength, nIndex) <> 0 Then
    GetContentType = TrimNull(sType)
End If
End Function

Private Function TrimNull(ByVal sStr As String) As String
'
' Strips upto the First Null Character in the String
'
Dim iPos As Long

iPos = InStr(1, sStr, vbNullChar)

If iPos > 0 Then
    TrimNull = Left$(sStr, iPos - 1)
Else
    TrimNull = sStr
End If
End Function

