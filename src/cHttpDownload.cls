VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' VbAsyncSocket Project (c) 2018-2019 by wqweto@gmail.com
'
' Simple and thin WinSock API wrappers for VB6
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cHttpDownload"
#Const ImplUseTls = True

'=========================================================================
' Events
'=========================================================================

Event DownloadProgress(ByVal BytesRead As Double, ByVal BytesTotal As Double)
Event DownloadComplete(ByVal LocalFileName As String)
Event UploadProgress(ByVal BytesWritten As Double, ByVal BytesTotal As Double)
Event UploadComplete(ByVal LocalFileName As String)
Event OperationStart()
Event OperationError(ByVal Number As Long, ByVal Description As String, ByVal Source As String)

'=========================================================================
' API
'=========================================================================

'--- for SHCreateStreamOnFile
Private Const STGM_READ                     As Long = 0
Private Const STGM_WRITE                    As Long = 1
Private Const STGM_CREATE                   As Long = &H1000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private Declare Function SHCreateStreamOnFileW Lib "shlwapi" (ByVal pszFile As Long, ByVal grfMode As Long, ppstm As IUnknown) As Long
Private Declare Function SHCreateStreamOnFileA Lib "shlwapi" (ByVal pszFile As String, ByVal grfMode As Long, ppstm As IUnknown) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const HDR_CONTENT_LENGTH    As String = "content-length:"
Private Const HDR_LOCATION          As String = "location:"
Private Const DEF_BUFFER_SIZE       As Long = 256& * 1024
'--- errors
Private Const ERR_INVALID_URL       As String = "Invalid URL"
Private Const ERR_HTTP_ONLY         As String = "Http protocol supported only"
Private Const ERR_INVALID_REDIRECT  As String = "Invalid redirect to %1"
Private Const ERR_INVALID_RESPONSE  As String = "Server returned %1"

#If ImplUseTls Then
Private WithEvents m_oSocket    As cTlsSocket
Attribute m_oSocket.VB_VarHelpID = -1
#Else
Private WithEvents m_oSocket    As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
#End If
Private m_uRemote               As UcsParsedUrl
Private m_sLocalFileName        As String
Private m_lStreamFlags          As Long
Private m_pFileStream           As IUnknown
Private m_dStartDate            As Date
Private m_lCallbackPtr          As Long
Private m_eState                As UcsStateEnum
Private m_baFileBuffer()        As Byte
Private m_dblBytesProgress      As Double
Private m_dblContentLength      As Double
Private m_sBoundaryData         As String
Private m_lBufferSize           As Long
Private m_sFileFormName         As String
Private m_sBody                 As String

Private Enum UcsStateEnum
    ucsIdle
    ucsWaitRecvHeaders
    ucsWaitRecvBody
    ucsWaitSendBody
    ucsWaitBoundary
End Enum

Private Type UcsParsedUrl
    Protocol        As String
    Host            As String
    Port            As Long
    Path            As String
    QueryString     As String
    Anchor          As String
    User            As String
    Pass            As String
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get Url() As String
    Url = m_uRemote.Protocol & "://" & m_uRemote.Host & m_uRemote.Path & m_uRemote.QueryString
End Property

Property Get LocalFileName() As String
    LocalFileName = m_sLocalFileName
End Property

Property Get StartDate() As Date
    StartDate = m_dStartDate
End Property

Property Get CallbackWeakRef() As Object
    Call vbaObjSetAddref(CallbackWeakRef, m_lCallbackPtr)
End Property

Property Set CallbackWeakRef(oValue As Object)
    m_lCallbackPtr = ObjPtr(oValue)
End Property

#If ImplUseTls Then
Property Get Socket() As cTlsSocket
#Else
Property Get Socket() As cAsyncSocket
#End If
    Set Socket = m_oSocket
End Property

Property Get BufferSize() As Long
    BufferSize = m_lBufferSize
End Property

Property Let BufferSize(ByVal lValue As Long)
    If lValue > 0 Then
        m_lBufferSize = lValue
    End If
End Property

Property Get Body() As String
    Body = m_sBody
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub DownloadFile(Url As String, LocalFileName As Variant)
    pvInit Url, LocalFileName, STGM_WRITE Or STGM_CREATE, vbNullString
End Sub

Public Sub UploadFile(Url As String, LocalFileName As Variant, Optional FileFormName As String)
    pvInit Url, LocalFileName, STGM_READ, IIf(LenB(FileFormName) <> 0, FileFormName, "file")
End Sub

Public Sub CancelOperation()
    Const FUNC_NAME     As String = "CancelOperation"
    
    On Error GoTo EH
    m_oSocket.Close_
    Set m_oSocket = Nothing
    Set m_pFileStream = Nothing
    m_lCallbackPtr = 0
    m_eState = ucsIdle
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'= private ===============================================================

Private Sub pvInit(Url As String, LocalFileName As Variant, ByVal StreamFlags As Long, FileFormName As String)
    Const FUNC_NAME     As String = "pvInit"
    Dim hResult     As Long
    
    On Error GoTo EH
    If Not pvParseUrl(Url, m_uRemote, "http") Then
        On Error GoTo 0
        Err.Raise vbObjectError, , ERR_INVALID_URL
    End If
    If Not pvIsProtocolSupported(m_uRemote.Protocol) Then
        On Error GoTo 0
        Err.Raise vbObjectError, , ERR_HTTP_ONLY
    End If
    If IsObject(LocalFileName) Then
        m_sLocalFileName = "[stream]"
        Set m_pFileStream = LocalFileName
    Else
        m_sLocalFileName = LocalFileName
        m_lStreamFlags = StreamFlags
        If IsNT Then
            hResult = SHCreateStreamOnFileW(StrPtr(m_sLocalFileName), m_lStreamFlags, m_pFileStream)
        Else
            hResult = SHCreateStreamOnFileA(m_sLocalFileName, m_lStreamFlags, m_pFileStream)
        End If
        If hResult < 0 Then
            On Error GoTo 0
            Err.Raise hResult
        End If
    End If
    m_sFileFormName = FileFormName
    m_dStartDate = Now
    m_eState = ucsIdle
    m_dblBytesProgress = 0
    If m_lStreamFlags = STGM_READ Then
        If IsObject(LocalFileName) Then
            m_dblContentLength = -1
        Else
            m_dblContentLength = FileLen(m_sLocalFileName)
        End If
    Else
        m_dblContentLength = 0
    End If
    m_baFileBuffer = vbNullString
    #If ImplUseTls Then
        Set m_oSocket = New cTlsSocket
    #Else
        Set m_oSocket = New cAsyncSocket
    #End If
    If Not m_oSocket.Create(SocketType:=ucsSckStream) Then
        On Error GoTo 0
        Err.Raise vbObjectError, , m_oSocket.GetErrorDescription(m_oSocket.LastError)
    End If
    #If ImplUseTls Then
    If Not m_oSocket.Connect(m_uRemote.Host, m_uRemote.Port, pvIsProtocolSecure(m_uRemote.Protocol), ucsTlsIgnoreServerCertificateErrors Or ucsTlsSupportAll) Then
    #Else
    If Not m_oSocket.Connect(m_uRemote.Host, m_uRemote.Port) Then
    #End If
        On Error GoTo 0
        Err.Raise vbObjectError, , m_oSocket.GetErrorDescription(m_oSocket.LastError)
    End If
    RaiseEvent OperationStart
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Function pvRecvHeaders(baBuffer() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvRecvHeaders"
    Const STR_DELIM     As String = vbCrLf & vbCrLf
    Dim lPos            As Long
    Dim sHeaders        As String
    Dim vSplit          As Variant
    Dim vElem           As Variant
    Dim uRedirect       As UcsParsedUrl
    
    On Error GoTo EH
    sHeaders = m_oSocket.FromTextArray(m_baFileBuffer, ucsScpAcp) & m_oSocket.FromTextArray(baBuffer, ucsScpAcp)
    lPos = InStr(sHeaders, STR_DELIM)
    If lPos > 0 Then
        vSplit = Split(Left$(sHeaders, lPos), vbCrLf)
        Select Case CLng(Val(Mid$(vSplit(0), 10, 3)))
        Case 200 To 299
TryDownload:
            m_dblContentLength = -1
            For Each vElem In vSplit
                If Left$(LCase$(vElem), Len(HDR_CONTENT_LENGTH)) = HDR_CONTENT_LENGTH Then
                    m_dblContentLength = Val(Mid$(vElem, Len(HDR_CONTENT_LENGTH) + 1))
                End If
            Next
            m_eState = ucsWaitRecvBody
            pvRecvBody m_oSocket.ToTextArray(Mid$(sHeaders, lPos + Len(STR_DELIM)), ucsScpAcp)
        Case 300 To 399
            For Each vElem In vSplit
                If Left$(LCase$(vElem), Len(HDR_LOCATION)) = HDR_LOCATION Then
                    vElem = Trim$(Mid$(vElem, Len(HDR_LOCATION) + 1))
                    If Not pvParseUrl(CStr(vElem), uRedirect, m_uRemote.Protocol) Then
                        pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "pvParseUrl", Replace(ERR_INVALID_REDIRECT, "%1", vElem)
                        GoTo QH
                    End If
                    uRedirect.User = m_uRemote.User
                    uRedirect.Pass = m_uRemote.Pass
                    Exit For
                End If
            Next
            If Not pvIsProtocolSupported(uRedirect.Protocol) Then
                #If UseTls Then
                    pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, ERR_HTTP_ONLY
                    GoTo QH
                #Else
                    GoTo TryDownload
                #End If
            End If
            m_uRemote = uRedirect
            m_oSocket.Close_
            If Not m_oSocket.Create(SocketType:=ucsSckStream) Then
                pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.Create"
                GoTo QH
            End If
            #If ImplUseTls Then
            If Not m_oSocket.Connect(m_uRemote.Host, m_uRemote.Port, UseTls:=pvIsProtocolSecure(m_uRemote.Protocol)) Then
            #Else
            If Not m_oSocket.Connect(m_uRemote.Host, m_uRemote.Port) Then
            #End If
                pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.Connect"
                GoTo QH
            End If
        Case Else
            pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Replace(ERR_INVALID_RESPONSE, "%1", Mid$(vSplit(0), 10))
            GoTo QH
        End Select
    End If
    '--- success
    pvRecvHeaders = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvRecvBody(baBuffer() As Byte, Optional ByVal Flush As Boolean) As Boolean
    Const FUNC_NAME     As String = "pvRecvBody"
    Dim lIdx            As Long
    Dim hResult         As Long
    
    On Error GoTo EH
    If UBound(baBuffer) >= 0 Then
        m_dblBytesProgress = m_dblBytesProgress + UBound(baBuffer) + 1
        RaiseEvent DownloadProgress(m_dblBytesProgress, m_dblContentLength)
        If m_lCallbackPtr <> 0 Then
            Call CallbackWeakRef.DownloadProgress(Me, m_dblBytesProgress, m_dblContentLength)
        End If
        lIdx = UBound(m_baFileBuffer) + 1
        ReDim Preserve m_baFileBuffer(0 To lIdx + UBound(baBuffer)) As Byte
        Call CopyMemory(m_baFileBuffer(lIdx), baBuffer(0), UBound(baBuffer) + 1)
    End If
    If UBound(m_baFileBuffer) >= m_lBufferSize Or Flush Then
        hResult = IStream_Write(m_pFileStream, m_baFileBuffer)
        m_baFileBuffer = vbNullString
        If hResult < 0 Then
            pvSetError hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "IStream_Write"
            GoTo QH
        End If
    End If
    '--- success
    pvRecvBody = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvSendBody() As Boolean
    Const FUNC_NAME     As String = "pvSendBody"
    Dim hResult         As Long
    Dim lSize           As Long
    
    On Error GoTo EH
    If UBound(m_baFileBuffer) < 0 And Not m_pFileStream Is Nothing Then
        ReDim m_baFileBuffer(0 To m_lBufferSize - 1) As Byte
        hResult = IStream_Read(m_pFileStream, m_baFileBuffer, lSize)
        If hResult < 0 Then
            pvSetError hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "IStream_Read"
            GoTo QH
        End If
    Else
        lSize = UBound(m_baFileBuffer) + 1
    End If
    If lSize > 0 Then
        If m_oSocket.Send(VarPtr(m_baFileBuffer(0)), lSize) < 0 Then
            If Not m_oSocket.HasPendingEvent Then
                pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.Send"
                GoTo QH
            End If
            If lSize < UBound(m_baFileBuffer) + 1 Then
                ReDim Preserve m_baFileBuffer(0 To lSize - 1) As Byte
            End If
        Else
            m_baFileBuffer = vbNullString
            RaiseEvent UploadProgress(m_dblBytesProgress, m_dblContentLength)
            If m_lCallbackPtr <> 0 Then
                Call CallbackWeakRef.UploadProgress(Me, m_dblBytesProgress, m_dblContentLength)
            End If
            m_dblBytesProgress = m_dblBytesProgress + lSize
            If Not m_oSocket.HasPendingEvent Then
                m_oSocket.PostEvent ucsSfdWrite
            End If
        End If
    Else
        Set m_pFileStream = Nothing
        m_baFileBuffer = vbNullString
        If Not m_oSocket.SendText(m_sBoundaryData, CodePage:=ucsScpAcp) Then
            If Not m_oSocket.HasPendingEvent Then
                pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.SendText"
                GoTo QH
            End If
        Else
            m_eState = ucsWaitBoundary
            RaiseEvent UploadProgress(m_dblBytesProgress, m_dblContentLength)
            If m_lCallbackPtr <> 0 Then
                Call CallbackWeakRef.UploadProgress(Me, m_dblBytesProgress, m_dblContentLength)
            End If
            If Not m_oSocket.HasPendingEvent Then
                m_oSocket.PostEvent ucsSfdWrite
            End If
        End If
    End If
    '--- success
    pvSendBody = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvSendComplete(baBuffer() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvSendComplete"
    Const STR_DELIM     As String = vbCrLf & vbCrLf
    Dim sHeaders        As String
    Dim lPos            As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    sHeaders = m_oSocket.FromTextArray(baBuffer, ucsScpAcp)
    lPos = InStr(sHeaders, STR_DELIM)
    If lPos > 0 Then
        vSplit = Split(Left$(sHeaders, lPos), vbCrLf)
        Select Case CLng(Val(Mid$(vSplit(0), 10, 3)))
        Case 200 To 299, 300 To 399
            m_sBody = Mid$(sHeaders, lPos + Len(STR_DELIM))
            RaiseEvent UploadComplete(m_sLocalFileName)
            If m_lCallbackPtr <> 0 Then
                Call CallbackWeakRef.UploadComplete(Me, m_sLocalFileName)
            End If
        Case Else
            pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Replace(ERR_INVALID_RESPONSE, "%1", Mid$(vSplit(0), 10))
            GoTo QH
        End Select
        m_oSocket.Close_
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Sub pvSetError(ByVal ErrNumber As Long, Optional ErrSource As String, Optional ErrDescription As String)
    Const FUNC_NAME     As String = "pvSetError"
    
    On Error GoTo EH
    If LenB(ErrDescription) = 0 Then
        If Not m_oSocket Is Nothing Then
            ErrDescription = m_oSocket.GetErrorDescription(ErrNumber)
        Else
            ErrDescription = "Error " & ErrNumber
        End If
    End If
    CancelOperation
    RaiseEvent OperationError(ErrNumber, ErrDescription, ErrSource)
    If m_lCallbackPtr <> 0 Then
        Call CallbackWeakRef.OperationError(Me, ErrNumber, ErrDescription, ErrSource)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Function pvParseUrl(sUrl As String, uParsed As UcsParsedUrl, Optional DefProtocol As String) As Boolean
    With CreateObject("VBScript.RegExp")
        .Pattern = "^(?:(?:(.+?):)?//)?(?:(?:([^:]*?):)?([^@]*)@)?([A-Za-z0-9\-\.]+)?(:[0-9]+)?(/[^?#]*?)?(\?[^#]*?)?(#.*)?$"
        With .Execute(sUrl)
            If .Count > 0 Then
                With .Item(0).SubMatches
                    uParsed.Protocol = IIf(LenB(.Item(0)) = 0, DefProtocol, .Item(0))
                    uParsed.User = .Item(1)
                    If LenB(uParsed.User) = 0 Then
                        uParsed.User = .Item(2)
                    Else
                        uParsed.Pass = .Item(2)
                    End If
                    uParsed.Host = .Item(3)
                    uParsed.Port = Val(Mid$(.Item(4), 2))
                    If uParsed.Port = 0 Then
                        Select Case LCase$(uParsed.Protocol)
                        Case "https"
                            uParsed.Port = 443
                        Case "socks5"
                            uParsed.Port = 1080
                        Case Else
                            uParsed.Port = 80
                        End Select
                    End If
                    uParsed.Path = .Item(5)
                    If LenB(uParsed.Path) = 0 Then
                        uParsed.Path = "/"
                    End If
                    uParsed.QueryString = .Item(6)
                    uParsed.Anchor = .Item(7)
                End With
                pvParseUrl = True
            End If
        End With
    End With
End Function

Private Function pvIsProtocolSupported(sProtocol As String) As Boolean
    #If ImplUseTls Then
        pvIsProtocolSupported = (LCase$(sProtocol) = "http" Or LCase$(sProtocol) = "https")
    #Else
        pvIsProtocolSupported = (LCase$(sProtocol) = "http")
    #End If
End Function

Private Function pvIsProtocolSecure(sProtocol As String) As Boolean
    pvIsProtocolSecure = (LCase$(sProtocol) = "https")
End Function

Private Function IStream_Read(pStm As IUnknown, baBuffer() As Byte, Optional BytesRead As Long) As Long
    If Not pStm Is Nothing And UBound(baBuffer) >= 0 Then
        IStream_Read = DispCallByVtbl(pStm, 3, VarPtr(baBuffer(0)), UBound(baBuffer) + 1, VarPtr(BytesRead))
    End If
End Function

Private Function IStream_Write(pStm As IUnknown, baBuffer() As Byte, Optional BytesWritten As Long) As Long
    If Not pStm Is Nothing And UBound(baBuffer) >= 0 Then
        IStream_Write = DispCallByVtbl(pStm, 4, VarPtr(baBuffer(0)), UBound(baBuffer) + 1, VarPtr(BytesWritten))
    End If
End Function

Private Function DispCallByVtbl(pUnk As IUnknown, ByVal lIndex As Long, ParamArray A() As Variant) As Variant
    Const CC_STDCALL    As Long = 4
    Dim lIdx            As Long
    Dim vParam()        As Variant
    Dim vType(0 To 63)  As Integer
    Dim vPtr(0 To 63)   As Long
    Dim hResult         As Long
    
    vParam = A
    For lIdx = 0 To UBound(vParam)
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
    Next
    hResult = DispCallFunc(ObjPtr(pUnk), lIndex * 4, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
    If hResult < 0 Then
        Err.Raise hResult, "DispCallFunc"
    End If
End Function

Private Function IsNT() As Boolean
    IsNT = (GetVersion() >= 0)
End Function

'=========================================================================
' Socket events
'=========================================================================

Private Sub m_oSocket_OnConnect()
    Const FUNC_NAME     As String = "m_oSocket_OnConnect"
    Const STR_BOUNDARY  As String = "3fbd04f5-b1ed-4060-99b9-fca7ff59c113"
    Dim sPostData       As String
    
    On Error GoTo EH
    If m_lStreamFlags <> STGM_READ Then
        m_eState = ucsWaitRecvHeaders
        If Not m_oSocket.SendText("GET " & m_uRemote.Path & m_uRemote.QueryString & " HTTP/1.0" & vbCrLf & _
                "Host: " & m_uRemote.Host & vbCrLf & _
                "Accept: */*" & vbCrLf & vbCrLf, CodePage:=ucsScpAcp) Then
            pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.SendText"
        End If
    Else
        m_eState = ucsWaitSendBody
        sPostData = "--" & STR_BOUNDARY & vbCrLf & _
            "Content-Disposition: form-data; name=""" & m_sFileFormName & """; filename=""" & Mid$(m_sLocalFileName, InStrRev(m_sLocalFileName, "\") + 1) & """" & vbCrLf & _
            "Content-Type: application/octet-stream" & vbCrLf & vbCrLf
        m_sBoundaryData = vbCrLf & "--" & STR_BOUNDARY & "--"
        If Not m_oSocket.SendText("POST " & m_uRemote.Path & m_uRemote.QueryString & " HTTP/1.0" & vbCrLf & _
                "Host: " & m_uRemote.Host & vbCrLf & _
                "Content-Length: " & Len(sPostData) + m_dblContentLength + Len(m_sBoundaryData) & vbCrLf & _
                "Content-Type: multipart/form-data; boundary=" & STR_BOUNDARY & vbCrLf & _
                "Accept: */*" & vbCrLf & vbCrLf & _
                sPostData, CodePage:=ucsScpAcp) Then
            pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.SendText"
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnReceive()
    Const FUNC_NAME     As String = "m_oSocket_OnReceive"
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    If m_eState = ucsWaitRecvHeaders Then
        If Not m_oSocket.ReceiveArray(baBuffer) Then
            pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.ReceiveArray"
            GoTo QH
        End If
        If UBound(baBuffer) < 0 Then
            GoTo QH
        End If
        If Not pvRecvHeaders(baBuffer) Then
            GoTo QH
        End If
    End If
    If m_eState = ucsWaitRecvBody Then
        If Not m_oSocket.ReceiveArray(baBuffer) Then
            If m_dblBytesProgress < m_dblContentLength Then
                pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.ReceiveArray"
            End If
            GoTo QH
        End If
        If UBound(baBuffer) < 0 Then
            GoTo QH
        End If
        If Not pvRecvBody(baBuffer) Then
            GoTo QH
        End If
    End If
    If m_eState >= ucsWaitSendBody Then
        If Not m_oSocket.ReceiveArray(baBuffer) Then
            pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.ReceiveArray"
            GoTo QH
        End If
        If UBound(baBuffer) < 0 Then
            GoTo QH
        End If
        If Not pvSendComplete(baBuffer) Then
            GoTo QH
        End If
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnSend()
    Const FUNC_NAME     As String = "m_oSocket_OnSend"
    
    On Error GoTo EH
    If m_eState = ucsWaitBoundary Then
        m_oSocket.ShutDown
    End If
    If m_eState = ucsWaitSendBody Then
        If Not pvSendBody() Then
            GoTo QH
        End If
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnClose()
    Const FUNC_NAME     As String = "m_oSocket_OnClose"
    Dim baBuffer()      As Byte
    Dim dblBytes        As Double
    
    On Error GoTo EH
    If m_eState = ucsWaitRecvBody Then
        Do
            dblBytes = m_dblBytesProgress
            m_oSocket_OnReceive
        Loop While m_dblBytesProgress > dblBytes
        baBuffer = vbNullString
        pvRecvBody baBuffer, Flush:=True
        RaiseEvent DownloadComplete(m_sLocalFileName)
        If m_lCallbackPtr <> 0 Then
            Call CallbackWeakRef.DownloadComplete(Me, m_sLocalFileName)
        End If
    End If
    CancelOperation
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Const FUNC_NAME     As String = "m_oSocket_OnError"
    
    On Error GoTo EH
    #If ImplUseTls Then
        With m_oSocket.LastError
            If .Number <> 0 Then
                pvSetError .Number, .Source, .Description
            End If
        End With
    #Else
        With m_oSocket
            If .LastError <> 0 Then
                pvSetError .LastError, , .GetErrorDescription(.LastError)
            End If
        End With
    #End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_lBufferSize = DEF_BUFFER_SIZE
End Sub

Private Sub Class_Terminate()
    Set m_oSocket = Nothing
End Sub
