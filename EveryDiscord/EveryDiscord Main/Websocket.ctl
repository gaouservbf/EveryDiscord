VERSION 5.00
Begin VB.UserControl Websocket 
   CanGetFocus     =   0   'False
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2805
   HasDC           =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   2805
   Windowless      =   -1  'True
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Closed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1590
   End
End
Attribute VB_Name = "Websocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================================

'VB6 Client Websocket UserControl
'Date created: July 25, 2021
'Author: Lewis Miller (aka vbLewis) dethbomb@hotmail.com
'License: portions of this source code fall under MIT license (see project directory)
'         everything else is submitted to public domain
'
'Special Thanks to:
'Vladimir Vissoultchev (wqweto@gmail.com) for VbAsyncSocket and zZipArchive Project
'Youran (chinese source) For source code and inspiration
'Matooh for creating the first vb6 secure websocket example
'=============================================================================

'VBWebsocket 1.10 - BETA
'release date: 4/10/2022

'please note that error handling is minimal to allow for finding errors and bugs more easily.
'after the code has been more thoroughly tested and matured, the BETA status will be removed and
'more aggressive error handling will be added.

'Alot of times i will code a procedure with single letter variables to save typing and then
'use the replace dialog to give the variable names better meaning. My apologies if i forgot
'to do so on some functions.

'NOTES
'the Websocket Usercontrol, mWebsocket module, and Project1 example are my primary
'contributions. The mWebsocket module is a heavily modified version of Youran's online
'source code found at https://www.cnblogs.com/xiii/p/7135233.html

'cAsyncSocket.cls,cTlsSocket.cls, cZipArchive, and mTlsThunks.bas are all
'related to wqweto's VBAsyncSocket and ZipArchive projects and remain largely unchanged. Please direct
'any questions related to those modules to him. (because I'm dumb and dont know anthing about SSL/TLS)
'Please note his license in the source code and give proper cedit.

'If you use this websocket in an application, credit would be appreciated but not needed.
'please report any bugs or feedback to the related post on the vbForums at
'https://www.vbforums.com/showthread.php?892835-VB6-Visual-Basic-6-Client-Websocket-Control
'==============================================================================

'see the IANA websocket registry for a list of protocols and extensions
'https://www.iana.org/assignments/websocket/websocket.xml

'see the rfc for the websocket protocol
'https://www.rfc-editor.org/rfc/rfc6455.txt


'the version of the protocol this websocket supports
Private Const PROTOCOL_VERSION As String = "13"
Private Enum UcsTlsLocalFeaturesEnum '--- bitmask
    ucsTlsSupportTls10 = 2 ^ 0
    ucsTlsSupportTls11 = 2 ^ 1
    ucsTlsSupportTls12 = 2 ^ 2
    ucsTlsSupportTls13 = 2 ^ 3
    ucsTlsIgnoreServerCertificateErrors = 2 ^ 4
    ucsTlsSupportAll = ucsTlsSupportTls10 Or ucsTlsSupportTls11 Or ucsTlsSupportTls12 Or ucsTlsSupportTls13
End Enum


'===================================
'USERCONTROL EVENTS
'===================================

'events from the server
Event OnMessage(ByVal Msg As Variant, ByVal OpCode As WebsocketOpCode)
Event OnConnect(ByVal RemoteHost As String, ByVal RemoteIP As String, ByVal RemotePort As String)
Event OnReConnect(ByVal newURI As String)
Event OnClose(ByVal eCode As WebsocketStatus, ByVal reason As String)
Event OnError(ByVal eCode As WebsocketStatus, ByVal reason As String)
Event OnPong(ByVal IncludedMsg As String)

'this event happens on large sends greater than the ChunkSize property, could be used to implement a status or progress bar
Event onProgress(ByVal bytesSent As Long, ByVal bytesMax As Long, Cancel As Boolean)


'=============================
'OBJECT VARIABLES
'=============================

'underlying TLS socket object
Private WithEvents TlsSock As cTlsSocket
Attribute TlsSock.VB_VarHelpID = -1

'deflate/inflate handler
Private Compressor As cZipArchive


'=====================================
'SERVER VARIABLES
'=====================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'websocket server information
Private ServerScheme As ProtocolScheme    'aka ws:// or wss://
Private ServerAddress As String
Private ServerPort As Long
Private ServerPath As String    'everything after the slash / and before the question mark ?
Private ServerUrl As String  'complete passed in uri
Private ServerVersion As String
Private ServerProtocols As String    'the desired protocols wanted by the client
Private ServerExtensions As String
Private ServerIP As String

'holds extra user supplied header lines
Private ServerHeaders As String

'if port is included in uri this holds that value
Private ServerPortParsed As String

'uri query paramters, raw and encoded in url encoding
Private ServerParams_Raw As String
Private ServerParams_Enc As String

'client secret key, raw and encoded in base64
Private ClientSecret_Raw As String
Private ClientSecret_Enc As String

'connection state of the websocket
Private ConnectionState As WebSocketState



'======================================
'BUFFER HANDLING
'======================================
'Two kinds of fragmenting can occur, websocket frames and winsock data
'1) winsock returns 16kb at most per data event, so a wait buffer has to
'   implemented for winsock if the data expected in a message exceeds that amount.
'2) totally seperate from the winsock buffer fragmenting is websocket frame fragmenting which
'   is outlined in the spec rfc6455. Both issues are handled by this control.

'congestion
'3) Another issue is multiple frames or messages can be contained in 1 winsock buffer (congested buffer)
'   this is also handled transparently to the user.

'buffer handling has been greatly simplified and changed.
'All incoming data is put onto the end of one continuous FIFO buffer, and frames
'are removed from the beginning of the buffer as they are completed.


'first in, first out
Private Type FIFOBuffer
    Count As Long
    B() As Byte
End Type

Private IncomingData As FIFOBuffer


'==============================
'PROPERTIES
'===============================
'true if the socket is in the process of sending/recieving data
Private I_AM_BUSY As Boolean

'holds the value for usecompression property
Private Use_Compression As Boolean

'holds the value for chunksize property
Private mvarChunkSize As Double


'internal use flags
Private DisconnectAttempts As Long
Private TrimmingData As Boolean




Private Sub UserControl_Initialize()
    'size the control
    UserControl.Width = Label1.Width
    UserControl.Height = Label1.Height

    'this is needed for random number generator
    Randomize Timer

    mvarChunkSize = DEFAULT_CHUNK_SIZE

    'this is needed to init the crypto subsystem for sha1 and sha256
    Set TlsSock = New cTlsSocket
    Set Compressor = New cZipArchive

End Sub


Private Sub SetStatus(ByVal StatusText As String, ByVal vbCol As ColorConstants)
    Label1.Caption = StatusText
    Label1.ForeColor = vbCol
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = Label1.Width
    UserControl.Height = Label1.Height
End Sub

'Connect()
'**********************************************************************************************************

'components surrounded by square brackets are optional, all shown are just examples an not actual values

'INPUT:  ws[s]://server.example.net[:80|443][/chat/path][?examplekey=YMTSSYASGTSRSEWSJKUASGS&notify_self=true]

'Uri - this should be at minimum the address of the server prefixed with "ws://" or "wss://", with an optional
'      port, path, and any query parameters


'[Port]- if not included in the uri, Port should be used as port 80 (non-secure) or 443 (secure) or a custom port.
'       any port found in the URI will override any passed in or default port. If no port is specified
'       a default port of 80 for "ws://" and 443 for "wss://" will be used.


'[SubProtocols] - Specify what protocols you wish to use (JSON, AMPQ, MSTT, STOMP, COAP etc) these are server specific.
'                 After the server handshake the 'Protocols' Property of the control will reflect
'                 what the server has accepted. Seperate multiple protocols with comma's like so: "chat, echo, json"

'[ProtocolExtensions] - specify what extensions you want to support, same format as subprotocols ex: "bbf-usp-protocol, permessage-deflate"
'                       Note: if you set useCompression to True before connecting the websocket will automatically handle the permessage-deflate
'                       header

'[Headers] - This is a collection of addition header lines to include in the initial GET header, see sample piesocket code for an example

'**********************************************************************************************************

'connect to a websocket server
Public Sub Connect(ByVal URI As String, Optional ByVal Port As String, Optional ByVal SubProtocols As String, Optional ByVal ProtocolExtensions As String, Optional Headers As Collection)

    Dim HeaderOpt As Variant

    On Error GoTo Error_Handle

    ClearLocalVariables

    If Not (Headers Is Nothing) Then
        If Headers.Count Then
            For Each HeaderOpt In Headers
                ServerHeaders = ServerHeaders & HeaderOpt & vbCrLf
            Next
        End If
    End If

    'sub-protocols
    If Len(SubProtocols) Then
        ServerProtocols = FormatProtocols(SubProtocols)    'make sure protocols are properly formatted
    End If

    'extensions use the same format as sub-protocols
    If Len(ProtocolExtensions) Then
        ServerExtensions = FormatProtocols(ProtocolExtensions)
    End If

    'check for port
    If Len(Port) Then
        Port = Trim$(Port)
        If IsNumeric(Port) Then
            ServerPort = CLng(Port)
        End If
    End If

    'parsing of URI has been separated to allow for multiuse
    If Not ParseServerUri(URI) Then
        GoTo Error_Handle
    End If

    'we now have all the info needed to connect
    ConnectionState = STATE_CONNECTING
    SetStatus "Connecting...", vbYellow

    'make sure socket is closed
    If Not (TlsSock Is Nothing) Then
        If Not TlsSock.IsClosed Then
            TlsSock.Close_
        End If
        Set TlsSock = Nothing
    End If

    'create a new instance (cTlsSocket doesnt seem to support re-use (or maybe im just dumb))
    Set TlsSock = New cTlsSocket

    If Not TlsSock.Connect(ServerAddress, ServerPort, ServerScheme, ucsTlsSupportAll Or ucsTlsIgnoreServerCertificateErrors) Then
        RaiseEvent OnError(InternalError, "Unable to connect.")
        GoTo Connect_Error
    End If

    Exit Sub

Error_Handle:
    RaiseEvent OnError(InternalError, "Invalid Connection Data or URL Parsing Error.")
Connect_Error:
    ForceShutdown
    SetStatus "Error", vbRed
    ClearLocalVariables

End Sub


'the intent of this method is to reconnect using the current state of the websocket
'without resetting everything. An optional new Uri can be passed in.
'this was created primarily for use by the location header redirect
Sub ReConnect(Optional ByVal newURI As String)

    If Len(newURI) Then
        If Not ParseServerUri(newURI) Then
            GoTo Error_Handle
        End If
    End If

    'we want to only clear the buffers, not server vars
    ClearBuffers

    'we have all the info needed to connect
    ConnectionState = STATE_CONNECTING
    SetStatus "Connecting...", vbYellow

    'make sure socket is closed
    If Not (TlsSock Is Nothing) Then
        If Not TlsSock.IsClosed Then
            TlsSock.Close_
        End If
        Set TlsSock = Nothing
    End If

    Set TlsSock = New cTlsSocket

    If Not TlsSock.Connect(ServerAddress, ServerPort, ServerScheme, ucsTlsSupportAll) Then
        RaiseEvent OnError(InternalError, "Unable to connect.")
        GoTo Connect_Error
    End If

    Exit Sub

Error_Handle:
    RaiseEvent OnError(InternalError, "Invalid Connection Data or URL Parsing Error. (" & newURI & ")")
Connect_Error:
    ForceShutdown
    SetStatus "Error", vbRed
    ClearLocalVariables
End Sub




'reset everything to default
Private Sub ClearLocalVariables()

    ConnectionState = STATE_CLOSED

    ServerScheme = ProtocolSchemeWS   'aka 0
    ServerPort = 0

    ServerUrl = ""
    ServerPath = ""
    ServerAddress = ""
    ServerParams_Raw = ""
    ServerParams_Enc = ""
    ServerPortParsed = ""
    ServerExtensions = ""
    ServerProtocols = ""
    ServerVersion = ""
    ServerIP = ""
    ClientSecret_Raw = ""
    ClientSecret_Enc = ""

    ServerHeaders = ""

    DisconnectAttempts = 0

    ClearBuffers
End Sub


Private Sub ClearBuffers()
    With IncomingData
        .Count = 0
        Erase .B
    End With
End Sub


'parsing of the connection URI is seperated to allow for use from different functions
' caller should call clearlocalvariables() if needed
Private Function ParseServerUri(ByVal URI As String) As Boolean

    Dim Pos As Long, sChar As String
    Dim Length As Long, sTemp As String

    ServerUrl = URI

    Length = Len(URI)

    'parse the uri (aka websocket url)
    'ws[s]://address[:port] [/path?[param=value]]
    If Length > 8 Then
        'parse scheme
        Pos = InStr(URI, ":")
        If Pos = 0 Or Pos > 4 Then
            Exit Function
        End If
        sChar = Left$(URI, Pos - 1)
        If StrComp(sChar, "ws", vbTextCompare) = 0 Then
            ServerScheme = ProtocolSchemeWS
        Else
            If StrComp(sChar, "wss", vbTextCompare) = 0 Then
                ServerScheme = ProtocolSchemeWSS
            Else
                Exit Function
            End If
        End If

        'parse server address
        Pos = InStr(URI, "//")
        If Pos = 0 Then
            Exit Function
        End If
        Pos = Pos + 2
        sChar = Mid$(URI, Pos, 1)
        ServerAddress = ""
        Do While (sChar <> ":") And (sChar <> "/")
            ServerAddress = ServerAddress & sChar
            Pos = Pos + 1
            If Pos > Length Then
                Exit Do
            End If
            sChar = Mid$(URI, Pos, 1)
        Loop


        'parse server port
        If sChar <> ":" Then    'no port
            If ServerPort = 0 Then
                If ServerScheme = ProtocolSchemeWS Then
                    ServerPort = 80
                Else
                    ServerPort = 443
                End If
            End If
        Else
            'parse server port included in server uri
            Pos = Pos + 1
            sChar = Mid$(URI, Pos, 1)
            Do While (sChar <> "/") And (sChar <> "?")
                sTemp = sTemp & sChar
                Pos = Pos + 1
                If Pos > Length Then
                    Exit Do
                End If
                sChar = Mid$(URI, Pos, 1)
            Loop

            'if any port was included in the uri it will override any passed in or default port
            ServerPortParsed = sTemp

            If Len(sTemp) Then
                If IsNumeric(sTemp) Then
                    ServerPort = CLng(sTemp)
                Else    'malformed uri
                    Exit Function
                End If
            Else
                Exit Function
            End If
        End If


        'parse any server path
        ServerPath = ""
        If sChar = "/" Then    'path
            Do While (sChar <> "?")
                ServerPath = ServerPath & sChar
                Pos = Pos + 1
                If Pos > Length Then
                    Exit Do
                End If
                sChar = Mid$(URI, Pos, 1)
            Loop
        Else
            'default path
            ServerPath = "/"
        End If


        'parse query params
        ServerParams_Raw = ""
        If sChar = "?" Then
            Pos = Pos + 1
            ServerParams_Raw = Mid$(URI, Pos)
            'we have to pass a flag to not encode ("=" and "&") for some non-standard compatability.
            ServerParams_Enc = URLEncode_UTF8(ServerParams_Raw, True)
        End If

        ParseServerUri = True
    End If

End Function



'ping the server
Sub Ping(Optional ByVal ExtraMsg As String)

    Dim Bytes() As Byte
    Dim Length As Long

    If Len(ExtraMsg) Then
        Bytes = ByteUTF8(ExtraMsg)
        Length = UBound(Bytes) + 1
    End If

    TlsSock.SendArray PingFrame(Bytes, Length)
End Sub


'VB6 reserves the Close keyword so we use Disconnect
Public Sub Disconnect()


    If ConnectionState <> STATE_CLOSED Then
        ConnectionState = STATE_CLOSING
    End If

    If Not (TlsSock.IsClosed = True And TlsSock.Socket.SocketHandle <> -1) Then
        'if we are connected
        'the specs want us to send a close frame and
        'the server responds with a closeframe to close the connection
        TlsSock.SendArray CompileMessage((vbNullChar), 0, False, opClose, True)
        DisconnectAttempts = DisconnectAttempts + 1
        If DisconnectAttempts > 1 Then
            ForceShutdown
            DisconnectAttempts = 0
        End If
    Else
        RaiseEvent OnClose(NormalClosure, GetStatusCodeText(NormalClosure))
        ForceShutdown
    End If

End Sub


Public Sub ShutDown()
    ConnectionState = STATE_CLOSED
    ForceShutdown
    ClearLocalVariables
End Sub



'(read-only) reflects the current connection state of the websocket
Public Property Get readyState() As WebSocketState
    readyState = ConnectionState
End Property

'(read-only) returns the protocols that the websocket server accepted
Public Property Get Protocols() As String
    Protocols = ServerProtocols
End Property

'(read-only) true if the server is sending or recieveing data
Public Property Get isBusy() As Boolean
    isBusy = I_AM_BUSY
End Property


'(read-write) set this to true if you want to use the 'permessage-deflate' compression extension
Public Property Let UseCompression(ByVal newValue As Boolean)

    ''this is disabled for now
    '    If ConnectionState <> STATE_CLOSED Then
    '         RaiseEvent OnError(InternalError, "You must set the compression state before connecting!")
    '         Exit Property
    '    End If

    Use_Compression = newValue
    If newValue = True Then
        If Compressor Is Nothing Then
            Set Compressor = New cZipArchive
        End If
    Else
        If Not (Compressor Is Nothing) Then
            Set Compressor = Nothing
        End If
    End If

End Property

Public Property Get UseCompression() As Boolean
    UseCompression = Use_Compression
End Property

'read this property to determine if the server accepted your requested extension
'note that websocket handles compression extension internally through the UseCompression property
'     which should be set before connecting
Public Property Get Extensions() As String
    Extensions = ServerExtensions
End Property


'this property lets you set the maximum chunk size used to send data. 4kb is default, max 16kb recommended.
'If for some reason you need to send data in a packet larger than 64kb you will need to use this to
'up the chunk size (max &H7FFFFFFF)  to use the larger packet send feature
Public Property Get ChunkSize() As Double
    ChunkSize = mvarChunkSize
End Property

'the reason we want to break up large sends into chunks is memory constraints... stuffing 16kb chunks into winsock
'is much faster then trying to stuff a large packet like 120 Mb's into winsock an letting the system try to handle our memory
'and winsocks memory buffer. it also creates 2 copies of our data which uses more memory.
Public Property Let ChunkSize(ByVal newValue As Double)

    If newValue > CDbl(&H7FFFFFFF) Then
        RaiseEvent OnError(InternalError, "Chunk size to large. Maximum supported chunk size is &H7FFFFFFF (2,147,483,647) bytes.")
        Exit Property
    End If

    mvarChunkSize = newValue
End Property






'Send()
'************************************************************************************************
'send data to the websocket server...
'data can be formatted in anyway that the server requires (plain text, JSON, XML etc).
' VB's unicode (normal string) is UTF16 and is converted to UTF8 encoding and sent as text.
' Byte Arrays are sent as Binary and unencoded.
'You can send a string as binary by setting NoUTF8Conversion to True, but the string must actually be binary not unicode
'************************************************************************************************
Public Sub Send(ByVal Data As Variant, Optional ByVal NoUTF8Conversion As Boolean)

    Dim Compressed() As Byte    'compressed data
    Dim Encoded() As Byte    'encoded data
    Dim Chunk() As Byte
    Dim ChunkLength As Long  'length of chunks
    Dim lByteCount As Long   'number of bytes
    Dim wOpCode As WebsocketOpCode

    Dim bSent As Boolean     'flag to indicate data has been sent
    Dim bCompressed As Boolean    'flag to indicate that data is compressed

    'used for onProgress event
    Dim lTotal As Long
    Dim lSent As Long
    Dim bCancel As Boolean

    'used to control frozen loops
    Dim LoopCount As Long
    Dim LoopMax As Long

    On Error GoTo ErrorHandle

    If Not TlsSock.IsClosed Then
        I_AM_BUSY = True

        'send data in chunks if need be, minus 8-14 bytes for overhead
        If mvarChunkSize < 126# Or mvarChunkSize > 2147483647# Then
            mvarChunkSize = DEFAULT_CHUNK_SIZE
        End If

        'allow for protocol overhead
        If mvarChunkSize > 65535# Then
            ChunkLength = CLng(mvarChunkSize - 14#)
        Else
            ChunkLength = CLng(mvarChunkSize - 8#)
        End If

        '---- version 2.0 ----
        'so much code between sending byte arrrays and strings is similar that I have
        'merged the 2 loops into one
        '---- version 3.0 ----
        ' all data is encoded and compresssed before entering send loop
        '============================================================================

        'data integrity checks
        If (VarType(Data) = vbString) Then

            If NoUTF8Conversion Then
                Encoded = Data
                Send Encoded, False
                Exit Sub
            End If

            'encode in UTF8
            Encoded = ByteUTF8(Data)
            wOpCode = opText
            Data = ""

        Else

            If Not IsArray(Data) Then
                RaiseEvent OnError(UnsupportedData, "Data in the Send() function contains an unsupported data type. Only Byte Arrays and Strings are supported.")
                GoTo ExitProc
            End If

            If VarType(Data) <> vbByte + vbArray Then
                'unsupported data type, you can of course write more code to handle other types such as long, currency, double etc
                RaiseEvent OnError(UnsupportedData, "Data in the Send() function contains an unsupported data type. Only Byte Arrays and Strings are supported.")
                GoTo ExitProc
            End If

            If NoUTF8Conversion = True Then
                RaiseEvent OnError(UnsupportedData, "You have specified the NoUTF8Conversion flag but data is not a String type.")
                GoTo ExitProc
            End If

            wOpCode = opBinary
            Encoded = Data
            Erase Data

        End If

        'compress data?
        If (Use_Compression = True And ServerSupportsCompression = True) Then
            If CompressData(Encoded, Compressed) Then
                bCompressed = True
                Encoded = Compressed
                Erase Compressed
            Else
                GoTo ExitProc
            End If
        End If

        'get length of data
        lTotal = ArrayCount(Encoded)
        If lTotal = 0 Then GoTo ExitProc

        lByteCount = lTotal

        'calculate loop max
        LoopMax = (lTotal / ChunkLength) + 1


        'send data in chunks if need be
        Do While (lByteCount > ChunkLength)
            'convert a chunk to byte array
            ReDim Chunk(ChunkLength - 1) As Byte
            CopyMemory Chunk(0), Encoded(0), ChunkLength

            If bSent Then
                TlsSock.SendArray CompileMessage(Chunk, UBound(Chunk) + 1, False, opContinue, False)
            Else
                TlsSock.SendArray CompileMessage(Chunk, UBound(Chunk) + 1, bCompressed, wOpCode, False)    'dont set FIN bit
            End If
            Erase Chunk

            'remove chunk from data
            lByteCount = (lByteCount - ChunkLength)
            CopyMemory Encoded(0), Encoded(ChunkLength), lByteCount
            ReDim Preserve Encoded(lByteCount - 1) As Byte

            bSent = True

            'check connection state before sending another chunk (server could disconnect)
            If ConnectionState <> STATE_OPEN Then    'immediate abort
                RaiseEvent OnError(InternalError, "Data send aborted. Websocket is not connected.")
                GoTo ExitProc
            End If

            lSent = lSent + ChunkLength
            RaiseEvent onProgress(lSent, lTotal, bCancel)

            'if the user has set bCancel to true then abort
            If bCancel Then GoTo ExitProc

            'loop safety (to prevent freezing app)
            LoopCount = LoopCount + 1
            If LoopCount > LoopMax Then
                Exit Do
            End If
        Loop

        'send final data
        If lByteCount Then
            ' note that the last optional parameter is not set because setfinbit is defualt true
            If bSent Then
                TlsSock.SendArray CompileMessage(Encoded, UBound(Encoded) + 1, False, opContinue)
            Else
                TlsSock.SendArray CompileMessage(Encoded, UBound(Encoded) + 1, bCompressed, wOpCode)
            End If
        Else
            'there is an infinitesimally small chance that the former buffer loop send came out even on 0,
            'so we should send an empty packet with the fin bit set
            If bSent Then    'data has been sent
                TlsSock.SendArray CompileMessage((vbNullChar), 0, False, opContinue)
            End If
        End If

        If bSent Then
            RaiseEvent onProgress(lTotal, lTotal, bCancel)
        End If


    Else
        'socket closed
        RaiseEvent OnError(InternalError, "Cant send data. Websocket is not connected.")
    End If

    I_AM_BUSY = False

    On Error GoTo 0
    Exit Sub
ErrorHandle:
    'MsgBox Err.Number & ":" & Err.Description, vbCritical
    RaiseEvent OnError(InternalError, Err.Description)
ExitProc:
    I_AM_BUSY = False
    On Error GoTo 0
End Sub



'SendAdvanced()
'************************************************************************************************
'Allows fine-grained control of the message construction, and sending extension data...
'for now we dont send in a loop or split into continue frames
'************************************************************************************************
Public Sub SendAdvanced(ByVal Data As Variant, ByVal OpCode As Long, Optional ByVal blnUTF8Encode As Boolean, Optional ByVal blnCompressData As Boolean, Optional ByVal blnSetFinBit As Boolean = True, Optional ByVal RSV1 As Boolean, Optional ByVal RSV2 As Boolean, Optional ByVal RSV3 As Boolean)

    Dim Bytes() As Byte    'temp data
    Dim Encoded() As Byte    'encoded data

    On Error GoTo ErrorHandle

    If OpCode > 15 Then
        RaiseEvent OnError(InvalidData, "Invalid opcode specified in the SendAdvanced() procedure!")
        Exit Sub
    End If

    If Not TlsSock.IsClosed Then
        I_AM_BUSY = True

        'data integrity checks
        If (VarType(Data) = vbString) Then

            If Len(Data) Then
                If blnUTF8Encode Then
                    Encoded = ByteUTF8(Data)
                Else
                    'we will assume a binary string
                    Encoded = Data
                End If
                'always free up memory as soon as possible
                Data = ""
            End If

        Else

            If Not IsArray(Data) Then
                RaiseEvent OnError(UnsupportedData, "Data in the SendAdvanced() function contains an unsupported data type. Only Byte Arrays and Strings are supported.")
                GoTo ExitProc
            End If

            If VarType(Data) <> vbByte + vbArray Then
                RaiseEvent OnError(UnsupportedData, "Data in the SendAdvanced() function contains an unsupported data type. Only Byte Arrays and Strings are supported.")
                GoTo ExitProc
            End If


            Encoded = Data
            Erase Data

            'should this be an error or should we encode...????
            If blnUTF8Encode = True Then
                'If ArrayCount(Encoded) Then
                'Bytes = ByteUTF16ToUTF8(Encoded)
                'Encoded = Bytes
                'Erase Bytes
                'End If

                RaiseEvent OnError(UnsupportedData, "You have specified the UTF8 Encode flag in the SendAdvanced() procedure, but data is not a String type.")
                GoTo ExitProc
            End If

        End If

        'compress data?
        If blnCompressData = True And ArrayCount(Encoded) > 0 Then
            If (ServerSupportsCompression = True) Then
                If CompressData(Encoded, Bytes) Then
                    Encoded = Bytes
                    Erase Bytes
                Else
                    RaiseEvent OnError(InternalError, "SendAdvanced(): Websocket was unable to compress data!")
                    GoTo ExitProc
                End If
            Else
                RaiseEvent OnError(InternalError, "Unable to compress data! You have specified the compression flag in SendAdvanced() but the compression option was not enabled by the server.")
                GoTo ExitProc
            End If
        End If


        'send final data
        Bytes = CompileMessage(Encoded, ArrayCount(Encoded), blnCompressData, OpCode, blnSetFinBit)
        Erase Encoded

        If RSV1 = True And blnCompressData = False Then
            Bytes(0) = CByte(Bytes(0) Or &H40)
        End If
        If RSV2 Then
            Bytes(0) = CByte(Bytes(0) Or &H20)
        End If
        If RSV3 Then
            Bytes(0) = CByte(Bytes(0) Or &H10)
        End If

        TlsSock.SendArray Bytes
    Else
        'socket closed
        RaiseEvent OnError(InternalError, "SendAdvanced(): Cant send data. Websocket is not connected.")
    End If

    I_AM_BUSY = False

    On Error GoTo 0
    Exit Sub
ErrorHandle:
    'MsgBox Err.Number & ":" & Err.Description, vbCritical
    RaiseEvent OnError(InternalError, "SendAdvanced(): " & Err.Description)
ExitProc:
    I_AM_BUSY = False
    On Error GoTo 0
End Sub





'======================================================
'PRIVATE FUNCTIONS
'=====================================================




'processing of http data and websocket frames is now seperate
'this function assumes we are in a connecting state
Private Sub ProcessHTTPData(ByVal buff As String)

    Dim FirstLine As String
    Dim Pos As Long
    Dim HttpCode As Long
    Dim URI As String

    Pos = InStr(buff, vbCrLf)
    If Pos Then
        FirstLine = Left$(buff, Pos - 1)
        If Len(FirstLine) Then
            If (UCase$(FirstLine) Like "HTTP/1.# ### *") Then

                Pos = InStr(buff, " ")
                HttpCode = CLng(Trim$(Mid$(buff, Pos + 1, 3)))

                Select Case HttpCode

                    Case 101    'switching protocols

                        'the protocol specs require these checks
                        If StrComp(getHeaderValue(buff, "upgrade"), "websocket", vbTextCompare) <> 0 Then
                            ConnectionState = STATE_CLOSING
                            RaiseEvent OnError(ProtocolError, "Invalid handshake recieved. Upgrade header not present.")
                            ForceShutdown
                            Exit Sub
                        End If
                        If StrComp(getHeaderValue(buff, "connection"), "upgrade", vbTextCompare) <> 0 Then
                            ConnectionState = STATE_CLOSING
                            RaiseEvent OnError(ProtocolError, "Invalid handshake recieved. Connection header not present.")
                            ForceShutdown
                            Exit Sub
                        End If

                        'here we validate the server, to be authentic
                        If MakeAcceptKey(ClientSecret_Enc) <> getHeaderValue(buff, "Sec-Websocket-Accept:") Then
                            ConnectionState = STATE_CLOSING
                            RaiseEvent OnError(ProtocolError, "The client-server secret key challenge failed.")
                            ForceShutdown
                            Exit Sub
                        End If

                        'set the server accepted protocols and extensions
                        ServerProtocols = FormatProtocols(getHeaderValue(buff, "Sec-WebSocket-Protocol:"))

                        'extensions are the same format as protocols
                        ServerExtensions = FormatProtocols(getHeaderValue(buff, "Sec-WebSocket-Extensions:"))

                        If Use_Compression Then
                            'todo: negotiate compression options
                            If InStr(1, ServerExtensions, "permessage-deflate", vbTextCompare) = 0 Then
                                RaiseEvent OnError(MandatoryExtension, "The server does not support deflate compression. Data will be sent uncompressed.")
                                Use_Compression = False
                            End If
                        End If

                        'we are officially connected and ready for business
                        ConnectionState = STATE_OPEN
                        RaiseEvent OnConnect(TlsSock.RemoteHostName, ServerIP, CStr(ServerPort))
                        SetStatus "Connected!", vbGreen

                    Case 301, 302, 303, 307, 308    'moved/redirect

                        'we make a weak attempt to support redirection
                        URI = getHeaderValue(buff, "Location:")
                        If Len(URI) > 8 Then
                            If StrComp(Left$(URI, 5), "https", vbTextCompare) Then
                                'secure socket
                                URI = "wss" & Mid$(URI, InStr(URI, ":"))
                            Else
                                If StrComp(Left$(URI, 4), "http", vbTextCompare) = 0 Then
                                    'normal socket
                                    URI = "ws" & Mid$(URI, InStr(URI, ":"))
                                End If
                            End If
                        Else
                            RaiseEvent OnError(BadGateWay, "A redirect by the server could not be processed.")
                            ForceShutdown
                            Exit Sub
                        End If

                        RaiseEvent OnReConnect(URI)
                        ReConnect URI

                    Case 305, 306    'use proxy
                        'todo: support proxies
                        'RaiseEvent OnError(BadGateWay, "The server requires a proxy but proxies are not supported in VBWebsocket.")
                        RaiseEvent OnClose(BadGateWay, "The server requires a proxy but the configuration is incorrect.")
                        Disconnect


                    Case 426  'Protocols not supported, upgrade required
                        'MsgBox "The server doesnt support the protocols ( " & ServerProtocols & " )", vbCritical
                        RaiseEvent OnError(ProtocolError, "The requested sub-protocol(s) are not supported.")
                        Disconnect

                    Case Else
                        'moved to a separate function to maintain clarity
                        '*****************************************************************************
                        RaiseGeneralHttpError HttpCode
                        Disconnect

                End Select

            End If
        End If
    Else
        'data is not http header
        RaiseEvent OnError(InvalidData, GetStatusCodeText(InvalidData))
        Disconnect
    End If

    If IncomingData.Count Then
        InspectBuffer
    End If


End Sub


'process complete websocket messages. Data may be compressed.
Private Sub ProcessWebSocketMessage(Data() As Byte, ByVal Length As Long, DF As DataFrame)

    Dim Bytes() As Byte      'temp buffer
    Dim E() As Byte      'temp err code buffer
    Dim eCode As Integer    'err code
    Dim eMsg As String   'err message


    If ConnectionState = STATE_CLOSING Or DF.OpCode = opClose Then
        'only process close frames if closing connection
        If DF.OpCode = opClose Then
            If DF.PayloadLen >= 2 Then
                'contains a 2 byte error code
                ReDim Bytes(1) As Byte
                'swap bytes
                Bytes(0) = Data(DF.DataOffset + 1)
                Bytes(1) = Data(DF.DataOffset)
                CopyMemory eCode, Bytes(0), 2

                'is there also an error message?
                If DF.PayloadLen > 2 Then
                    ReDim E(DF.PayloadLen - 3) As Byte
                    CopyMemory E(0), Data(DF.DataOffset + 2), DF.PayloadLen - 2
                    eMsg = StrConv(E, vbUnicode)
                Else
                    eMsg = GetStatusCodeText(eCode)
                End If

                RaiseEvent OnClose(eCode, eMsg)
            Else
                RaiseEvent OnClose(NormalClosure, GetStatusCodeText(NormalClosure))
            End If

            ForceShutdown
        End If

        Exit Sub

    ElseIf ConnectionState = STATE_OPEN Then

        'note: opClose is handled above
        'first handle control frames.
        'controls frames are not compressed

        'check for ping frame
        If DF.OpCode = OpPing Then
            'auto answer pings
            If DF.PayloadLen > 0 Then   'we need to send back payload data
                TlsSock.SendArray CompileMessage(ExtractPayload(Data, DF), DF.PayloadLen, False, opPong, True)
            Else
                TlsSock.SendArray PongFrame((vbNullChar), 0)
            End If

            'check for pong frame, these can be ignored
        ElseIf DF.OpCode = opPong Then
            'the server has just answered our ping.
            If DF.PayloadLen Then
                eMsg = StringUTF8(ExtractPayload(Data, DF))
            End If
            RaiseEvent OnPong(eMsg)

            'normal message
        Else

            'fragmented (opcontinue) frames are handled before they get here
            'we should have only complete websocket messages...
            'check for valid opcodes
            Select Case DF.OpCode
                Case opText, opBinary

                    'raisedata handles decompression and utf8 un-encoding
                    If Not RaiseData(Data, DF) Then
                        ConnectionState = STATE_CLOSING
                        Disconnect
                        Exit Sub
                    End If

                Case 3 To 7, 11 To 15    'extension opcodes
                    E = ExtractPayload(Data, DF)
                    RaiseEvent OnMessage(E, DF.OpCode)

                Case Else
                    ConnectionState = STATE_CLOSING
                    RaiseEvent OnError(ProtocolError, "Invalid opCode recieved. (" & DF.OpCode & ")")
                    Disconnect
                    Exit Sub
            End Select
        End If

    Else
        'invalid state for recieving data
        RaiseEvent OnError(InvalidData, "Data Received, but websocket state is invalid for receiving data.")
    End If

    'check for more data, this creates a recursive loop till all complete data is processed
    If ConnectionState <> STATE_CLOSED Then
        If IncomingData.Count Then
            InspectBuffer
        End If
    End If

End Sub




'force close the socket
Private Sub ForceShutdown()

    On Error Resume Next
    TlsSock.Close_
    TlsSock.ShutDown
    TlsSock.Socket.Close_
    SetStatus "Closed", vbWhite
    On Error GoTo 0
    ClearBuffers
    ConnectionState = STATE_CLOSED

End Sub

'generate the websocket handshake packet, the first thing that is sent
Private Function Handshake() As String

    Dim sPath As String
    Dim sAddr As String
    Dim sHttpOrigin As String
    Dim K() As Byte
    Dim SubProtoHeader As String
    Dim ExtensionHeader As String

    'GET /chat HTTP/1.1
    'Host: server.example.com
    'Upgrade: WebSocket
    'Connection: Upgrade
    'Origin: http://example.com
    'Sec-WebSocket-Key: dGhlIHNhbXBsZSBub25jZQ==
    'Sec-WebSocket-protocol: chat , superchat
    'Sec-WebSocket-Version: 13

    sPath = ServerPath

    If Len(ServerParams_Enc) Then
        sPath = sPath & "?" & ServerParams_Enc
    End If

    'make a random 16 byte key for client secret
    K = GenerateSecretKey(16)

    'store raw key in a string
    ClientSecret_Raw = StrConv(K, vbUnicode)

    'encode key with base64
    ClientSecret_Enc = EncBase64(K)

    sAddr = ServerAddress

    'if port is not standard add it to server address
    If Len(ServerPortParsed) Then
        If ServerPortParsed <> "80" And ServerPortParsed <> "443" Then
            sAddr = sAddr & ":" & ServerPortParsed
        End If
    End If


    If ServerScheme = ProtocolSchemeWS Then
        sHttpOrigin = "http://" & ServerAddress & sPath
    Else
        sHttpOrigin = "https://" & ServerAddress & sPath
    End If

    If Len(ServerProtocols) Then
        SubProtoHeader = "Sec-WebSocket-Protocol: " & ServerProtocols & vbCrLf
    End If

    'specify compression extension, if needed...
    If Use_Compression Then
        If Len(ServerExtensions) Then
            If InStr(1, ServerExtensions, "permessage-deflate", vbTextCompare) = 0 Then
                ServerExtensions = ServerExtensions & ", permessage-deflate; server_no_context_takeover"
            End If
        Else
            ServerExtensions = "permessage-deflate; server_no_context_takeover"
        End If
    End If

    If Len(ServerExtensions) Then
        ExtensionHeader = "Sec-Websocket-Extensions: " & ServerExtensions & vbCrLf
    End If

    'include an origin header if none already specified
    If InStr(1, ServerHeaders, "origin: ", vbTextCompare) = 0 Then
        ServerHeaders = ServerHeaders & "Origin: " & sHttpOrigin & vbCrLf
    End If

    Handshake = "GET " & sPath & " HTTP/1.1" & vbCrLf & _
                "Host: " & sAddr & vbCrLf & _
                "Upgrade: WebSocket" & vbCrLf & _
                "Connection: Upgrade" & vbCrLf & _
                ServerHeaders & _
                "Sec-WebSocket-Origin: " & sHttpOrigin & vbCrLf & _
                "Sec-WebSocket-Key: " & ClientSecret_Enc & vbCrLf & _
                "Sec-WebSocket-Version: " & PROTOCOL_VERSION & vbCrLf & _
                SubProtoHeader & _
                ExtensionHeader & _
                vbCrLf

End Function

'the winsock socket connection has closed
Private Sub TlsSock_OnClose()
    On Error Resume Next
    ConnectionState = STATE_CLOSED
    RaiseEvent OnClose(NormalClosure, GetStatusCodeText(NormalClosure))
    ForceShutdown
    On Error GoTo 0
End Sub

Private Sub TlsSock_OnConnect()

    'send the handshake packet
    If Not TlsSock.SendText(Handshake) Then
        'error
        RaiseEvent OnError(InternalError, "Couldnt send handshake to server.")
        Disconnect
    End If

    'we dont raise onConnect event till we recieve the "101 Switching Protocols" header from the server

End Sub


' socket error
Private Sub TlsSock_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)

    Dim s As String, E As String

    If EventMask = ucsSfdAccept Then
        s = "Event Type: Accept"
    ElseIf EventMask = ucsSfdAll Then
        s = "Event Type: Unknown/All"
    ElseIf EventMask = ucsSfdClose Then
        s = "Event Type: Close"
    ElseIf EventMask = ucsSfdConnect Then
        s = "Event Type: Connect"
    ElseIf EventMask = ucsSfdOob Then
        s = "Event Type: OOB Out-Of-Band"
    ElseIf EventMask = ucsSfdRead Then
        s = "Event Type: Recieve/Read"
    ElseIf EventMask = ucsSfdWrite Then
        s = "Event Type: Send/Write"
    Else
        s = "Event Type: Unknown (" & CStr(EventMask) & ")"
    End If

    E = TlsSock.GetErrorDescription(ErrorCode)
    s = s & " " & CStr(ErrorCode) & ":  Description: " & E

    ConnectionState = STATE_CLOSING
    RaiseEvent OnError(WinSockError, s)
    If Not TlsSock.IsClosed Then
        Disconnect
    Else
        ForceShutdown
    End If


End Sub


'raw incoming data starts here, all available winsock data is put into a FIFO buffer
'then processed by the InspectBuffer() function.
Private Sub TlsSock_OnReceive()
    Dim buff() As Byte
    Dim BuffLen As Long
    Dim T As Long

    I_AM_BUSY = True
    T = timeGetTime

    Do While (TlsSock.AvailableBytes > 0)
        If TlsSock.ReceiveArray(buff) Then

            'add to byte buffer
            BuffLen = UBound(buff) + 1

            'this prevents conflicts with the buffer resizing function
            Do While TrimmingData
                ' DoEvents
            Loop

            With IncomingData
                If .Count Then
                    ReDim Preserve .B((.Count + BuffLen) - 1) As Byte
                Else
                    ReDim .B(BuffLen - 1) As Byte
                End If
                'copy temp buffer to end of primary buffer
                CopyMemory .B(.Count), buff(0), BuffLen
                .Count = .Count + BuffLen
            End With
            BuffLen = 0
        Else
            Exit Do
        End If

        'keep app from freezing, do a doevents every 2 secs (adjust as needed or remove if problems occur on large sends)
        If (timeGetTime - T) > 2000 Then
            T = timeGetTime
            DoEvents
        End If

        'we approach maximum density (added loop safety)  (MAX_LONG - 16kb)
        If IncomingData.Count > 2147467263 Then
            Exit Do
        End If
    Loop

    InspectBuffer

    I_AM_BUSY = False

End Sub

'store server ip address
Private Sub TlsSock_OnResolve(IpAddress As String)
    ServerIP = IpAddress
End Sub


'the goal of this proc is to parse complete webwocket Messages from the
'incoming websocket Frames in the FIFO buffer before they are processed.
'It seperates the HTTP headers from the websocket frames. Complete messages are
'sent to ProcessWebsocketMessages() and continued (fragmented) frames are handled by CollapseFramesEx()
Private Sub InspectBuffer()
    Dim Bytes() As Byte    'temp buffer
    Dim PacketLen As Long
    Dim buff As String
    Dim BuffLen As Long
    Dim ContentLen As Long
    Dim HeaderEnd As Long
    Dim DF As DataFrame
    Dim HttpBuff As String

    With IncomingData
        If .Count < 2 Then
            Exit Sub
        End If

        If ConnectionState = STATE_CONNECTING Then
            'in most cases we can expect a complete http header in a single buffer.
            'but, Ive seen winsock buffers with the complete http header followed
            'by a websocket frame all in the same winsock buffer, which is why
            'this complicated code is needed...

            'convert to string to make parsing easier
            buff = StrConv(.B, vbUnicode)

            If Len(buff) < 12 Then
                RaiseEvent OnError(InvalidData, "Cannot connect! Incoming Header Data Corrupted. Please try again.")
                ForceShutdown
                Exit Sub
            End If

            'first do a simple check
            If (Right$(buff, 4) = vbCrLf & vbCrLf) Or (Right$(buff, 2) = vbLf & vbLf) Then    'all good
                Erase .B
                .Count = 0
                ProcessHTTPData buff
                Exit Sub
            End If

            'check for a content length
            ContentLen = StrToNum(getHeaderValue(buff, "Content-Length"))

            'get the header length
            HeaderEnd = InStr(buff, vbCrLf & vbCrLf)
            If HeaderEnd Then
                HeaderEnd = HeaderEnd + 3
            Else
                'look for unix line endings
                HeaderEnd = InStr(buff, vbLf & vbLf)
                If HeaderEnd Then
                    HeaderEnd = HeaderEnd + 1
                Else
                    'invalid data
                    RaiseEvent OnError(InvalidData, "Cannot connect! Incoming Header Data Corrupted. Please try again.")
                    ForceShutdown
                    Exit Sub
                End If
            End If

            'get the header and any content
            BuffLen = HeaderEnd + ContentLen
            HttpBuff = Left$(buff, BuffLen)

            If (Len(buff) - BuffLen) > 0 Then
                'we have to convert back to byte array to know the proper length to strip away
                'because conversion can cause different lengths
                Bytes = StrConv(HttpBuff, vbFromUnicode)
                TrimIncomingData ArrayCount(Bytes)
                Erase Bytes
            End If

            ProcessHTTPData HttpBuff

        ElseIf ConnectionState = STATE_CLOSED Then
            'invalid state for recieving data
            RaiseEvent OnError(InvalidData, "Data Received, but websocket state is invalid for receiving data.")

        ElseIf (ConnectionState = STATE_CLOSING) Or (ConnectionState = STATE_OPEN) Then

            'we need to do 2 things...
            '1. see if there is enough data for a complete packet
            '2. On opcontinue frames, see if all frames are in and collapse them into one.
            'isHeaderComplete is an error trapped function that checks to see if enough data is
            'in to analyze the header
            If IsHeaderComplete(.B, 0, .Count) Then
                DF = AnalyzeData(.B)
            Else
                Exit Sub
            End If

            'eat up all complete websocket frames that are NOT fragmented
            Do While DF.FIN = True And DF.OpCode <> opContinue
                BuffLen = DF.DataOffset + DF.PayloadLen
                If BuffLen > .Count Then
                    'wait for more data
                    Exit Sub
                End If

                'grab complete message and process
                ReDim Bytes(BuffLen - 1) As Byte
                CopyMemory Bytes(0), .B(0), BuffLen
                TrimIncomingData BuffLen
                ProcessWebSocketMessage Bytes, BuffLen, DF
                Erase Bytes

                'if 4 or more bytes left then analyze
                If IsHeaderComplete(.B, 0, .Count) Then
                    DF = AnalyzeData(.B)
                Else
                    Exit Sub
                End If

                If DF.OpCode > 15 Then
                    Erase .B
                    .Count = 0
                    Exit Sub
                End If

            Loop

            'eat up any fragmented frames, all parameters are byref
            'note to self: this should probably be a loop, but ProcessWebsocketMessage
            'is recursive so its probably ok as a 1 time check
            If CollapseFramesEx(Bytes, PacketLen, DF) Then
                ProcessWebSocketMessage Bytes, PacketLen, DF
            End If

        End If
    End With

End Sub

'removes a specified number of bytes from the fifo buffer, global variable 'trimmingdata' is set to avoid conflicts with other functions
Private Sub TrimIncomingData(ByVal Length As Long)
    Dim LeftOver As Long

    Do While TrimmingData
    Loop

    TrimmingData = True
    With IncomingData
        LeftOver = .Count - Length
        If LeftOver Then
            CopyMemory .B(0), .B(Length), LeftOver
        End If
        .Count = .Count - Length
        If .Count Then
            ReDim Preserve .B(.Count - 1) As Byte
        Else
            Erase .B
        End If
    End With
    TrimmingData = False

End Sub

'fragmented frames (from rfc 6455)
'---------------------------------
'      A fragmented message consists of a single frame with the FIN bit
'      clear and an opcode other than 0, followed by zero or more frames
'      with the FIN bit clear and the opcode set to 0, and terminated by
'      a single frame with the FIN bit set and an opcode of 0.  A
'      fragmented message is conceptually equivalent to a single larger
'      message whose payload is equal to the concatenation of the
'      payloads of the fragments in order; however, in the presence of
'      extensions, this may not hold true as the extension defines the
'      interpretation of the "Extension data" present.  For instance,
'      "Extension data" may only be present at the beginning of the first
'      fragment and apply to subsequent fragments, or there may be
'      "Extension data" present in each of the fragments that applies
'      only to that particular fragment.  In the absence of "Extension
'      Data ", the following example demonstrates how fragmentation works."
'
'      EXAMPLE: For a text message sent as three fragments, the first
'      fragment would have an opcode of 0x1 and a FIN bit clear, the
'      second fragment would have an opcode of 0x0 and a FIN bit clear,
'      and the third fragment would have an opcode of 0x0 and a FIN bit
'      that is set.
'
'      Control frames (see Section 5.5) MAY be injected in the middle of
'      a fragmented message.  Control frames themselves MUST NOT be
'      fragmented.

'version 2.0
'byref args: Data() is the return buffer with all the data, Length is the length of that buffer, FirstFrame is data frame of that buffer.
'this code hasnt been extensivly tested, but it seems to work...
Private Function CollapseFramesEx(Data() As Byte, Length As Long, FirstFrame As DataFrame) As Boolean

    Dim NextFrame As DataFrame
    Dim BuffLen As Long
    Dim eMsg As String

    With IncomingData

        BuffLen = FirstFrame.DataOffset + FirstFrame.PayloadLen
        If BuffLen > .Count Then
            'need more data
            Exit Function
        End If

        'copy the first frame in its entirety, it contains the opCode for the message
        ReDim Data(BuffLen - 1) As Byte
        CopyMemory Data(0), .B(0), BuffLen

        Length = BuffLen

        'is there enough data to process?
        If .Count - BuffLen < 2 Then
            Exit Function
        End If

        'loop thru all the remaining frames till we have the complete message
        Do Until (NextFrame.FIN = True) And (NextFrame.OpCode = opContinue)

            If IsHeaderComplete(.B, BuffLen, .Count) Then
                NextFrame = AnalyzeData(.B, BuffLen)
            Else
                Exit Function
            End If

            'total corruption, start over..this should be a fatal error
            If NextFrame.OpCode > 15 Then
                Erase .B
                .Count = 0
                Exit Function
            End If

            'the specs say we must handle control frames (opClose,opPing,opPong) stuffed in anywhere in the buffer
            If NextFrame.OpCode = opClose Then
                RaiseEvent OnClose(AbNormalClosure, "A Close frame was unexpectedly encountered while recieving data.")
                ForceShutdown
                Exit Function
            End If

            If NextFrame.OpCode = OpPing Then
                'auto answer pings
                If NextFrame.PayloadLen > 0 Then   'we need to send back payload data
                    TlsSock.SendArray CompileMessage(ExtractPayloadEx(.B, BuffLen, NextFrame), NextFrame.PayloadLen, False, opPong, True)
                Else
                    TlsSock.SendArray PongFrame((vbNullChar), 0)
                End If
                'now we need to remove this frame from the buffer, removeframe() resets the dataframe to the next frame.
                'if removeframe cannot set the next frame it returns false. We have to remove the frame from the buffer
                'instead of skipping past it (which would be way more efficient) because if there is not enough data to
                'complete the message we will run into these frames again on the next try...we do not need to adjust
                'bufflen because it will still point to the beginning of the next frame starting offset in the buffer
                If Not RemoveFrame(BuffLen, NextFrame) Then
                    Exit Function
                End If
            End If

            If NextFrame.OpCode = opPong Then
                'the server has just answered our ping.
                If NextFrame.PayloadLen Then
                    eMsg = StringUTF8(ExtractPayloadEx(.B, BuffLen, NextFrame))
                End If
                RaiseEvent OnPong(eMsg)
                'remove this frame from the buffer
                If Not RemoveFrame(BuffLen, NextFrame) Then
                    Exit Function
                End If
            End If

            'check to see if there is enough data to grab the next frame
            If (BuffLen + (NextFrame.DataOffset - BuffLen) + NextFrame.PayloadLen) > .Count Then
                Exit Function
            End If

            'we only get the payload of each new OpContinue frame not to include the header
            ReDim Preserve Data((Length + NextFrame.PayloadLen) - 1) As Byte
            CopyMemory Data(Length), .B(BuffLen + (NextFrame.DataOffset - BuffLen)), NextFrame.PayloadLen

            'the current length of copied data from the fifo buffer
            Length = Length + NextFrame.PayloadLen

            'the current length of all the inspected frames in the fifo buffer aka next frame offset
            BuffLen = BuffLen + (NextFrame.DataOffset - BuffLen) + NextFrame.PayloadLen

            If NextFrame.FIN = False And (.Count - BuffLen) < 2 Then
                Exit Function
            End If

        Loop

        'we have a complete message!
        TrimIncomingData BuffLen

        'adjust the outgoing byref dataframe
        FirstFrame.FIN = True
        FirstFrame.PayloadLen = Length - FirstFrame.DataOffset

    End With

    CollapseFramesEx = True

End Function


'removes a frame from the fifo buffer, returns true if another frame could be analyzed
Private Function RemoveFrame(ByVal StartIndex As Long, DF As DataFrame) As Boolean

    Dim LeftOver As Long
    Dim FrameLen As Long

    FrameLen = (DF.DataOffset - StartIndex) + DF.PayloadLen

    With IncomingData
        TrimmingData = True

        'scootch data left in the buffer if there is any past our current frame
        LeftOver = .Count - (StartIndex + FrameLen)
        If LeftOver Then
            CopyMemory .B(StartIndex), .B(StartIndex + FrameLen), LeftOver
        End If

        .Count = .Count - FrameLen
        ReDim Preserve .B(.Count - 1) As Byte

        TrimmingData = False

        If .Count - 1 <= StartIndex Then
            Exit Function
        End If

        If IsHeaderComplete(.B, StartIndex, .Count) Then
            DF = AnalyzeData(.B, StartIndex)
        Else
            Exit Function
        End If
    End With

    RemoveFrame = True

End Function




'===============================================================
'HELPER FUNCTIONS
'===============================================================


'determine of the server supports compression,
Private Function ServerSupportsCompression() As Boolean
    If Len(ServerExtensions) Then
        ServerSupportsCompression = (InStr(1, ServerExtensions, "permessage-deflate", vbTextCompare) <> 0)
    End If
End Function


'get a text representation of a websocket status code
'this function is accessible from the usercontrol
Public Function GetStatusCodeText(ByVal statcode As WebsocketStatus) As String
    Dim s As String

    If statcode = AbNormalClosure Then
        GetStatusCodeText = "Abnormal Closure or connection closed unexpectedly."
    ElseIf statcode = GoingAway Then
        GetStatusCodeText = "Server is Going Away."
    ElseIf statcode = InternalError Then
        GetStatusCodeText = "An Internal Error has Occurred."
    ElseIf statcode = InvalidData Then
        GetStatusCodeText = "Invalid Data was Received."
    ElseIf statcode = MandatoryExtension Then
        GetStatusCodeText = "The server has required a Mandatory Extension."
    ElseIf statcode = MessageToLarge Then
        GetStatusCodeText = "A Message was To Large or Corrupted."
    ElseIf statcode = NormalClosure Then
        GetStatusCodeText = "Normal Closure."
    ElseIf statcode = NoStatusReceived Then
        GetStatusCodeText = "No Status Received."
    ElseIf statcode = PolicyViolation Then
        GetStatusCodeText = "A Policy Violation has Occurred."
    ElseIf statcode = ProtocolError Then
        GetStatusCodeText = "There was a Protocol Error."
    ElseIf statcode = SslTlsHandshake Then
        GetStatusCodeText = "TLS Handshake Error."
    ElseIf statcode = StatusReserved Then
        GetStatusCodeText = "Reserved Error Code (should not be Used)"
    ElseIf statcode = UnsupportedData Then
        GetStatusCodeText = "Unsupported data format or encoding."
    ElseIf statcode = BadGateWay Then
        GetStatusCodeText = "Bad GateWay, DNS Error, or Server Not Responding."
    ElseIf statcode = TryAgainLater Then
        GetStatusCodeText = "Try Again Later."
    ElseIf statcode = ServiceRestart Then
        GetStatusCodeText = "Service is Restarting."
    ElseIf statcode = WinSockError Then
        GetStatusCodeText = "Winsock Socket Error."
    Else
        s = TlsSock.GetErrorDescription(statcode)
        If Len(s) Then
            GetStatusCodeText = s
        Else
            Select Case statcode
                Case 0 To 999, 1016 To 1999
                    GetStatusCodeText = "Unknown Error. Error number is in an Unused range."
                Case 2000 To 2999
                    GetStatusCodeText = "Websocket Extension Error."
                Case 3000 To 3999
                    GetStatusCodeText = "Library or Framework Error."
                Case 4000 To 4999
                    GetStatusCodeText = "Application Layer Error."
                Case Else
                    GetStatusCodeText = "Error code not recognized. Means nothing to WebSocket Protocol."
            End Select
        End If
    End If


End Function


'processes general http errors, some of these errors are/might be recoverable but isnt yet handled
Private Sub RaiseGeneralHttpError(ByVal HttpCode As Long)

    Select Case HttpCode
            '4xx - client errors
        Case 400, 406    ' bad request
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Bad Request.")

        Case 403, 405    'forbidden
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Request Forbidden.")

        Case 401, 561   'unauthorized
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Request Not Authorized.")

        Case 404, 410    'not found
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Resource Not Found.")

        Case 402    'payment required
            RaiseEvent OnError(PolicyViolation, "Payment Required.")

        Case 407    'proxy auth required
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Proxy Authorization Required.")

        Case 408, 522, 524  'timeout
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Request Timed Out.")

        Case 411    'length required
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Length required.")

        Case 412, 423, 424    'failed
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Request Failed.")

        Case 413    'data to large
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Data To Large.")

        Case 414    'uri to large
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": URI To Large.")

        Case 415    'unsupported
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Request is Not Supported.")

        Case 416, 417, 421, 422, 425, 428, 451    'cant complete request
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": The server cannot complete the request.")

        Case 418    'easter egg, "Im a tea pot" (in reply to "brew coffee")
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": The server got jokes.")

        Case 429    'to many requests, rate limit exceeded
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": To Many Requests, Rate Limit Exceeded.")

        Case 431, 431   'header fields to large
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Header Field To Large.")


            '5xx - server errors
        Case 500    'internal server error
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Internal Server Error.")

        Case 501    'not implemented
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": The Server has not implemented the requested function.")

        Case 502    'bad gateway
            RaiseEvent OnError(BadGateWay, CStr(HttpCode) & ": Bad Gateway.")

        Case 503    'service not available
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": The service is not available.")

        Case 504    'gateway timeout
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Gateway Timeout.")

        Case 505    'http version not supported
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": The HTTP Version used is not supported.")

        Case 506    'variant also negotiates
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Variant Also Negotiates.")

        Case 507    'insufficient storage
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": The server has insufficient storage to complete the request.")

        Case 508    'loop detected
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": The server has detected an Infinite Loop.")

        Case 510    'not extended, need more data
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Need more data.")

        Case 511    'network authentication required
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Network Authentication required. Read Wifi TOS?.")
            'todo: show wifi  tos page


            'expanded ssl/tls errors
        Case 495, 526    'ssl certificate error
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": SSL Certificate Error.")

        Case 496    'ssl certificate required
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": SSL Certificate Required.")

        Case 497    'http sent to https
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": HTTP sent to HTTPS")

        Case 499    'client closed request
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Client Closed Request.")

        Case 494    'request to large
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Request To Large.")

            'microsoft
        Case 440    'login timeout
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Login has timed out.")

        Case 449    'retry with
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Retry Login.")

        Case 451    'redirect
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Login redirect.")

            'cloud flare
        Case 520    'unknown error
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Cloud Flare - Unknown Error.")

        Case 521    'web server is down
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Cloud Flare - Web Server is Down.")

        Case 523    'origin unreachable
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Cloud Flare - Origin is Unreachable.")

        Case 525    'ssl handshake failed
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Cloud Flare - SSL Handshake Failed.")

        Case 527    'railgun server error
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Cloud Flare - Railgun Server Error.")

        Case 530    '
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Cloud Flare - Error.")

            'unofficial codes
        Case 103    'checkpoint
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Checkpoint Error.")

        Case 218    'this is fine
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": This is Fine.")

        Case 420    'smoke weed and remain calm (rate limit exceeded)
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Enhance Your Calm (Rate Limit Exceeded).")

        Case 450    'blocked by windows parental controls
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Blocked by Windows Parental Control.")

        Case 509    'bandwidth exceeded
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Bandwidth Exceeded.")

        Case 529    'site is overloaded
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Site is Overloaded.")

        Case 530    'site is frozen
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Site is Frozen.")

        Case 598    'network read timeout error
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Network Read Timeout Error.")

            'cache warning codes (proxies)
        Case 110    'response is stale
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Cache Warning - Response is Stale.")

        Case 111    'revalidation failed
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Cache Warning - Revalidation Failed.")

        Case 112    'disconnected operation
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Cache Warning - Disconnected Operation.")

        Case 113    'cache expired
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Cache Warning - Cache has Expired.")

        Case 199    'misc warning
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Cache Warning - Miscellaneous Warning.")

        Case 214    'transformation applied
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Cache Warning - Transformation Applied to Data.")

        Case 299   'misc persistent warning
            RaiseEvent OnError(ProtocolError, CStr(HttpCode) & ": Cache Warning - Miscellaneous Persistence Warning.")


        Case Else
            RaiseEvent OnError(UnsupportedData, "An HTTP code (" & CStr(HttpCode) & ") was recieived but could not be processed.")

    End Select

End Sub

'this procedure prepares the incoming data and sends it to the onMessage event
Private Function RaiseData(Data() As Byte, DF As DataFrame) As Boolean

    Dim Decompressed() As Byte, Bytes() As Byte

    'extract payload data
    Bytes = ExtractPayload(Data, DF)
    Erase Data

    If DF.RSV1 = True And Use_Compression = True Then     'the data is compressed
        If Not DecompressData(Bytes, Decompressed) Then
            Exit Function
        End If
    Else
        Decompressed = Bytes
    End If
    Erase Bytes

    If DF.OpCode = opText Then
        'text (utf8)
        RaiseEvent OnMessage(StringUTF8(Decompressed), opText)
    Else
        'binary
        RaiseEvent OnMessage(Decompressed, opBinary)
    End If

    RaiseData = True

End Function

'compress and decompress data helper functions
'Bytes() is the input data and Decompressed() is the output, while function returns success or failure
'**************************************************************************

Private Function DecompressData(Bytes() As Byte, Decompressed() As Byte) As Boolean
    Dim Length As Long

    Length = UBound(Bytes)

    'add trailing 4 bytes (00,00,FF,FF) if needed (per overly complicated specs)
    If Not (Bytes(Length) = 255 And Bytes(Length - 1) = 255 And Bytes(Length - 2) = 0 And Bytes(Length - 3) = 0) Then
        Length = Length + 4
        ReDim Preserve Bytes(Length) As Byte
        Bytes(Length) = CByte(255)
        Bytes(Length - 1) = CByte(255)
    End If

    If Not Compressor.Inflate(Bytes, Decompressed) Then
        RaiseEvent OnError(InternalError, "VBWebsocket was unable to decompress incoming data.")
        Exit Function
    End If

    DecompressData = True

End Function


Private Function CompressData(Bytes() As Byte, Compressed() As Byte) As Boolean
    Dim Length As Long

    If Not Compressor.Deflate(Bytes, Compressed) Then
        RaiseEvent OnError(InternalError, "VBWebsocket was unable to compress outgoing data.")
        Exit Function
    End If

    Length = UBound(Compressed)

    'need to make sure an empty deflate block BFINAL at the end
    If (Compressed(Length) <> 0) And (Compressed(Length) <> 255) Then
        Length = Length + 1
        ReDim Preserve Compressed(Length) As Byte
    End If

    'trim trailing bytes if needed (per overly complicated specs)
    If (Compressed(Length) = 255 And Compressed(Length - 1) = 255 And Compressed(Length - 2) = 0 And Compressed(Length - 3) = 0) Then
        ReDim Preserve Compressed(Length - 4) As Byte
    End If

    CompressData = True

End Function


'internal use, dumps a packet into a string for viewing
'useage: debug.print dumppacket(data,dataframe)
''Private Function DumpPacket(Data() As Byte, DF As DataFrame, Optional ByVal bDisplayAllData As Boolean, Optional ByVal bDisplayBinaryData As Boolean) As String
''
''    Dim S As String, A As String, Bytes() As Byte
''    Dim X As Long
''
''    With DF
''        If .FIN Then
''            S = "FIN "
''        Else
''            S = "NOFIN "
''        End If
''
''        If .RSV1 Then
''            S = S & "RSV1 "
''        Else
''            S = S & "NORSV1 "
''        End If
''
''        S = S & "OPCODE(" & CStr(.OpCode) & ") "
''
''        If .hasMASK Then
''            S = S & "MASK(" & CStr(Data(2)) & " " & CStr(Data(3)) & " " & CStr(Data(4)) & " " & CStr(Data(5)) & ") "
''        Else
''            S = S & "NOMASK "
''        End If
''
''        S = S & " PayLoadLen(" & CStr(.PayloadLen) & ") "
''        S = S & " DataOffSet(" & CStr(.DataOffset) & ") "
''
''        If .PayloadLen Then
''            Bytes = ExtractPayload(Data, DF)
''            If bDisplayBinaryData Then
''                If .OpCode = opText Then
''                    A = "Data String(" & StringUTF8(Bytes) & ") Binary: "
''                Else
''                    A = "Data: String(" & StrConv(Bytes, vbUnicode) & ") Binary: "
''                End If
''            End If
''
''            For X = .DataOffset To ((.PayloadLen + .DataOffset) - 1)
''                A = A & CStr(Data(X)) & "(" & Hex$(Data(X)) & ") "
''                If X > 256 Then
''                    If (Not bDisplayAllData) Then
''                        A = A & " ... (" & ((.PayloadLen + .DataOffset) - X) & " bytes not displayed.)"
''                        Exit For
''                    End If
''                End If
''            Next X
''            S = S & A
''        End If
''    End With
''
''    DumpPacket = S
''
''End Function
'''


'compression test
'============================================
''Public Sub Compression_Test()
''
''    Dim B() As Byte
''    Dim C() As Byte
''    Dim D() As Byte
''    Dim S As String
''    Dim F As String
''
''    Dim X As Long
''
''    Open App.Path & "\readme.txt" For Binary As #1
''    S = CStr(LOF(1)) & " Original - "
''    ReDim B(LOF(1) - 1) As Byte
''    Get #1, , B
''    F = Space$(LOF(1))
''    Get #1, 1, F
''    Close 1
''
''    If CompressData(B, C) Then
''        S = S & CStr(UBound(C) + 1) & " compressed: "
''        For X = 0 To UBound(C)
''            S = S & "&H" & Hex$(C(X)) & "(" & CStr(C(X)) & ") "
''        Next X
''Debug.Print S
''    Else
''        MsgBox "Compression failed"
''    End If
''
''    S = ""
''    If DecompressData(C, D) Then
''        S = StrConv(D, vbUnicode)
''        If S <> F Then
''            MsgBox "decompressed data not same as original!"
''        End If
''
''
''        If UBound(B) <> UBound(D) Then
''            MsgBox "decompressed data not same length as original!"
''        End If
''
''        For X = 0 To UBound(B)
''            If B(X) <> D(X) Then
''                MsgBox "decompressed data arrays not same as original array!"
''                Exit For
''            End If
''        Next X
''
''Debug.Print vbCrLf & Len(S) & " uncompressed: " & S
''    Else
''        MsgBox "decompression failed"
''    End If
''
''End Sub
''
















'this is it, go home, your children are waiting...




