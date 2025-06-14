Attribute VB_Name = "mWebSocket"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'VBWebsocket 1.10 BETA
' =============================================================
' -Original Header-
' By: Youran
' QQ: 2,860,898,817
' E-mail: ur1986@foxmail.com
' complete sample run discharge Q Group file sharing: 369088586
' ============================================================
' -Updated Header-
' Original Source: https://www.cnblogs.com/xiii/p/7135233.html
' 7/25/2021 - Chinese source translated with Google Translator
' Heavily Modified by Lewis Miller (vbLewis on VbForums.com)
' Note: some of the original comments remain
' email: dethbomb@hotmail.com
' Websocket Protocol Version 13
' rfc6455 (https://www.rfc-editor.org/rfc/rfc6455.txt)
'==============================================================

'notes:
' 1) future specs in the works include time-out handling, and multi-plexing support


'dont need synched socket
#Const ASYNC_NO_SYNC = 1

'dont need server side code
#Const ASYNCSOCKET_NO_TLSSERVER = 1

'enable in compiled app?
'#Const MST_NO_IDE_PROTECTION = 1

'the highest positive value a Long can hold
Public Const MAX_LONG As Long = &H7FFFFFFF  '2147483647

'the default size of data chunks that the websocket will send
'should be a even modulo of 1024 (1 kb)
Public Const DEFAULT_CHUNK_SIZE As Long = 4096&  '4kb


'============================
'ENUMS
'============================

'error codes
Public Enum WebsocketStatus
    NormalClosure = 1000
    GoingAway = 1001
    ProtocolError = 1002
    UnsupportedData = 1003
    StatusReserved = 1004
    NoStatusReceived = 1005
    AbNormalClosure = 1006
    InvalidData = 1007
    PolicyViolation = 1008
    MessageToLarge = 1009
    MandatoryExtension = 1010
    InternalError = 1011
    ServiceRestart = 1012
    TryAgainLater = 1013
    BadGateWay = 1014
    SslTlsHandshake = 1015    'added ssl because "tlshandshake" is a internal function call in the ssl/tls code
    WinSockError = 1016
End Enum
'these declarations preserve case of the enum members
#If False Then
Dim NormalClosure, ServiceRestart, GoingAway, BadGateWay, TryAgainLater, WinSockError, ProtocolError, _
    UnsupportedData, StatusReserved, NoStatusReceived, AbNormalClosure, InvalidData, PolicyViolation, _
    MessageToLarge, MandatoryExtension, InternalError, SslTlsHandshake
#End If

'enum to state whether url scheme is ws:// or wss://
Public Enum ProtocolScheme
    ProtocolSchemeWS
    ProtocolSchemeWSS
End Enum
#If False Then
Dim ProtocolSchemeWS, ProtocolSchemeWSS
#End If

'websocket connection state, these are from the java api (readyState)
Public Enum WebSocketState
    STATE_CLOSED
    STATE_OPEN
    STATE_CONNECTING
    STATE_CLOSING
End Enum
#If False Then
Dim STATE_CONNECTING, STATE_OPEN, STATE_CLOSING, STATE_CLOSED, STATE_SENDING
#End If

'websocket frame operation codes aka opcodes
Public Enum WebsocketOpCode
    opContinue = 0     ' successive message frame
    opText = 1        ' text message frame
    opBinary = 2     ' binary message frame
    opClose = 8       ' connection closing
    OpPing = 9       ' ping heartbeat check
    opPong = 10      ' ping heartbeat answer
End Enum
#If False Then
Dim opContinue, opText, opBinary, opClose, OpPing, opPong
#End If



'===============================
'TYPES
'===============================

'(internal use) holds info on an incoming websocket frame (old translated comments)
Public Type DataFrame
    FIN As Boolean              'if 0 indicates the current frame is part of a series, and there are more messages, 1 indicates that this is the last frame of the current message;
    RSV1 As Boolean             'if set indicates that the data is compressed, uses deflate to uncompress
    RSV2 As Boolean             ' 1 bit, if there is no custom extension, it must be 0, otherwise it must be disconnected.
    RSV3 As Boolean             ' 1 bit, if There is no custom extension, it must be 0, otherwise it must be disconnected.
    OpCode As WebsocketOpCode   ' 4-bit opcode, which defines the payload data. If an unknown opcode is received, the connection must be disconnected.
    hasMASK As Boolean          ' 1 bit defining whether a transmission mask, if so the mask is stored in MaskingKey. Masking keys are only used in outgoing data
    MaskingKey(3) As Byte       ' 32-bit mask
    PayloadLen As Long          ' length of the transmission data
    DataOffset As Long          ' The start bit of the data source
End Type


'misc win32 api calls
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function SafeArrayGetDim Lib "oleaut32" (ByVal pSA As Long) As Long

'this key is specified in rfc6455
Private Const MagicKey = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"

'used by the encbase64 function
Private Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="


'get a header value from a http header
Public Function getHeaderValue(ByVal strSearch As String, ByVal pName As String) As String

    Dim i As Long, j As Long


    If Len(pName) = 0 Or Len(strSearch) = 0 Then
        Exit Function
    End If

    pName = Trim$(pName)
    If Right$(pName, 1) <> ":" Then
        pName = pName & ":"
    End If

    i = InStr(1, strSearch, pName, vbTextCompare)
    If i > 0 Then
        i = i + Len(pName)
        j = InStr(i, strSearch, vbCrLf)
        If j > 0 Then
            getHeaderValue = Trim$(Mid$(strSearch, i, j - i))
        End If
    End If

End Function

'make the server challenge key
Public Function MakeAcceptKey(ByVal strKey As String) As String

    Dim Bytes() As Byte, retBytes() As Byte

    Bytes = StrConv(strKey & MagicKey, vbFromUnicode)

    If pvCryptoHashSha1(retBytes, Bytes) Then
        MakeAcceptKey = EncBase64(retBytes)
    End If

End Function

'base64 encoder
Function EncBase64(ByteArr() As Byte) As String

    Dim ByteBuf() As Byte, byteLen As Long, ModVal As Long
    Dim i As Long, X As Long

    On Error GoTo over

    ModVal = (UBound(ByteArr) + 1) Mod 3
    byteLen = UBound(ByteArr) + 1 - ModVal

    If ModVal <> 0 Then
        ReDim ByteBuf((byteLen / 3 * 4 + 4) - 1)
    Else
        ReDim ByteBuf((byteLen / 3 * 4) - 1)
    End If

    For i = 0 To byteLen - 1 Step 3
        ByteBuf(i / 3 * 4) = (ByteArr(i) And &HFC) / &H4
        ByteBuf(i / 3 * 4 + 1) = (ByteArr(i) And &H3) * &H10 + (ByteArr(i + 1) And &HF0) / &H10
        ByteBuf(i / 3 * 4 + 2) = (ByteArr(i + 1) And &HF) * &H4 + (ByteArr(i + 2) And &HC0) / &H40
        ByteBuf(i / 3 * 4 + 3) = ByteArr(i + 2) And &H3F
    Next

    If ModVal = 1 Then
        ByteBuf(byteLen / 3 * 4) = (ByteArr(byteLen) And &HFC) / &H4
        ByteBuf(byteLen / 3 * 4 + 1) = (ByteArr(byteLen) And &H3) * &H10
        ByteBuf(byteLen / 3 * 4 + 2) = 64
        ByteBuf(byteLen / 3 * 4 + 3) = 64
    ElseIf ModVal = 2 Then
        ByteBuf(byteLen / 3 * 4) = (ByteArr(byteLen) And &HFC) / &H4
        ByteBuf(byteLen / 3 * 4 + 1) = (ByteArr(byteLen) And &H3) * &H10 + (ByteArr(byteLen + 1) And &HF0) / &H10
        ByteBuf(byteLen / 3 * 4 + 2) = (ByteArr(byteLen + 1) And &HF) * &H4
        ByteBuf(byteLen / 3 * 4 + 3) = 64
    End If

    byteLen = UBound(ByteBuf) + 1
    EncBase64 = Space$(byteLen)
    X = 1
    For i = 0 To byteLen - 1
        Mid(EncBase64, X, 1) = Mid$(B64_CHAR_DICT, ByteBuf(i) + 1, 1)
        X = X + 1
    Next

over:
    On Error GoTo 0
End Function

'url query parameters encoder
Public Function URLEncode_UTF8(ByVal sText As String, Optional ByVal WeakEncode As Boolean = False) As String

    Dim x1 As Long
    Dim x2 As Long
    Dim chars() As Byte
    Dim Byte1 As Byte
    Dim Byte2 As Byte
    Dim UTF16 As Long
    Dim FinalPos As Long
    Dim TxtLen As Long
    Dim sHex As String
    Dim HexLen As Long

    TxtLen = Len(sText)
    URLEncode_UTF8 = Space$(TxtLen * 3)
    FinalPos = 1

    For x1 = 1 To TxtLen
        CopyMemory Byte1, ByVal StrPtr(sText) + ((x1 - 1) * 2), 1
        CopyMemory Byte2, ByVal StrPtr(sText) + ((x1 - 1) * 2) + 1, 1

        UTF16 = Byte2 * 256 + Byte1

        If UTF16 < &H80 Then
            ReDim chars(0) As Byte
            chars(0) = UTF16
        ElseIf UTF16 < &H800 Then
            ReDim chars(1) As Byte
            chars(1) = &H80 + (UTF16 And &H3F)
            UTF16 = UTF16 \ &H40
            chars(0) = &HC0 + (UTF16 And &H1F)
        Else
            ReDim chars(2) As Byte
            chars(2) = &H80 + (UTF16 And &H3F)
            UTF16 = UTF16 \ &H40
            chars(1) = &H80 + (UTF16 And &H3F)
            UTF16 = UTF16 \ &H40
            chars(0) = &HE0 + (UTF16 And &HF)
        End If

        For x2 = 0 To UBound(chars)
            Select Case chars(x2)
                Case 48 To 57, 65 To 90, 97 To 122
                    Mid(URLEncode_UTF8, FinalPos, 1) = Chr$(chars(x2))
                    FinalPos = FinalPos + 1
                Case 61, 38  ' "=" and "&"
                    If WeakEncode Then
                        Mid(URLEncode_UTF8, FinalPos, 1) = Chr$(chars(x2))
                        FinalPos = FinalPos + 1
                    Else
                        sHex = ("%" & Hex$(chars(x2)))
                        HexLen = Len(sHex)
                        Mid(URLEncode_UTF8, FinalPos, HexLen) = sHex
                        FinalPos = FinalPos + HexLen
                    End If
                Case Else
                    sHex = ("%" & Hex$(chars(x2)))
                    HexLen = Len(sHex)
                    Mid(URLEncode_UTF8, FinalPos, HexLen) = sHex
                    FinalPos = FinalPos + HexLen
            End Select
        Next
    Next

    URLEncode_UTF8 = Left$(URLEncode_UTF8, FinalPos - 1)

End Function



'incoming data can be incomplete, this special function makes sure there are enough bytes
'to be able to completely analyze the frame header without error. The payload is not examined.
'there must be a minimum of 2 bytes. StartIndex is 0 based
Function IsHeaderComplete(Data() As Byte, Optional ByVal StartIndex As Long, Optional ByVal Length As Long = -1) As Boolean

    Dim PacketType As Byte
    Dim PacketLen As Long

    If Length = -1 Then
        Length = UBound(Data) + 1
    End If

    If (StartIndex + 1) >= Length Then
        Exit Function
    End If

    PacketLen = Length - StartIndex
    If PacketLen < 2 Then
        Exit Function
    End If

    PacketType = Data(StartIndex + 1) And &H7F

    'this should never happen
    If PacketType > 127 Then
        Exit Function
    End If

    If PacketType > 125 Then
        If PacketType = 126 Then
            If PacketLen < 4 Then
                Exit Function
            End If

        ElseIf PacketType = 127 Then
            If PacketLen < 10 Then
                Exit Function
            End If
        End If
    End If

    IsHeaderComplete = True

End Function

'this function analyzes an incoming data packet and puts the resulting values in a dataframe structure.
'assumes header is complete. Optional StartIndex is 0 based.
Public Function AnalyzeData(ByteBuff() As Byte, Optional ByVal StartIndex As Long) As DataFrame

    Dim DF As DataFrame
    Dim PacketType As Byte
    Dim l(3) As Byte

    DF.FIN = ((ByteBuff(StartIndex) And &H80) = &H80)

    DF.RSV1 = ((ByteBuff(StartIndex) And &H40) = &H40)
    DF.RSV2 = ((ByteBuff(StartIndex) And &H20) = &H20)
    DF.RSV3 = ((ByteBuff(StartIndex) And &H10) = &H10)

    DF.OpCode = ByteBuff(StartIndex) And &H7F

    PacketType = ByteBuff(StartIndex + 1) And &H7F
    
    'this is only set on client sends never from server, but sometimes we need
    'to analyze client sends for debug purposes
    DF.hasMASK = ((ByteBuff(StartIndex + 1) And &H80) = &H80)
    If DF.hasMASK Then
        CopyMemory DF.MaskingKey(0), ByteBuff(StartIndex + 2), 4&
        'adjust startindex so that offset is correct
        StartIndex = StartIndex + 4
    End If


'    If PacketType > 127 Then    'protocol error
'        'todo: raise error
'    End If

    If PacketType < 126 Then    'aka 125 or less

        DF.PayloadLen = PacketType
        DF.DataOffset = StartIndex + 2

    ElseIf PacketType = 126 Then

        'payload length is in byte 2 and 3
        l(0) = ByteBuff(StartIndex + 3)
        l(1) = ByteBuff(StartIndex + 2)

        CopyMemory DF.PayloadLen, l(0), 4
        DF.DataOffset = StartIndex + 4

    ElseIf PacketType = 127 Then
        'if the server sends a packet > MAX_LONG 2147483647 (2 Gb)  we are screwed (but its highly unlikely)
        'bytes 2,3,4,5,6,7,8,9 hold the big-endian (network byte order) data length
        'we only get 4 bytes because that is all we will handle. If this becomes a major
        'problem in the future we can start using double to store the value
        l(0) = ByteBuff(StartIndex + 9)
        l(1) = ByteBuff(StartIndex + 8)
        l(2) = ByteBuff(StartIndex + 7)
        l(3) = ByteBuff(StartIndex + 6)

        CopyMemory DF.PayloadLen, l(0), 4&
        DF.DataOffset = StartIndex + 10

    End If

    AnalyzeData = DF

End Function


'removes the payload data from a frame, does NOT decompress because the frame could be part of a group of frames
Public Function ExtractPayload(ByteData() As Byte, DF As DataFrame) As Byte()

    Dim B() As Byte

    ReDim B(DF.PayloadLen - 1) As Byte

    CopyMemory B(0), ByteData(DF.DataOffset), DF.PayloadLen

    ExtractPayload = B

End Function

'version 2.0
Public Function ExtractPayloadEx(ByteData() As Byte, ByVal StartIndex As Long, DF As DataFrame) As Byte()

    Dim B() As Byte

    ReDim B(DF.PayloadLen - 1) As Byte

    CopyMemory B(0), ByteData(StartIndex + DF.DataOffset), DF.PayloadLen

    ExtractPayloadEx = B

End Function

'prepares a message for sending, prepends a header, masks the data and set flags
Public Function CompileMessage(ByteData() As Byte, ByVal ByteLength As Long, Optional ByVal isCompressed As Boolean, Optional ByVal OpCode As WebsocketOpCode = opText, Optional ByVal SetFINBit As Boolean = True) As Byte()

    Dim mKey(3) As Byte
    Dim i As Long, j As Long
    Dim ByteBuff() As Byte
    Dim l(3) As Byte


    'mask the data per rfc6455
    If ByteLength Then

        'the websocket protocol standard states that these should be randomized in every new frame
        mKey(0) = RandomNumber(1, 255)
        mKey(1) = RandomNumber(1, 255)
        mKey(2) = RandomNumber(1, 255)
        mKey(3) = RandomNumber(1, 255)

        For i = 0 To UBound(ByteData)
            ByteData(i) = ByteData(i) Xor mKey(j)
            j = j + 1
            If j = 4 Then j = 0
        Next i

    End If

    'create 1 of 3 types of packets depending on the length of the data
    If ByteLength < 126 Then    'small chatty packet

        ReDim ByteBuff(ByteLength + 5) As Byte    '6 byte overhead

        ByteBuff(0) = CByte(&H80& Or OpCode)

        If isCompressed Then
            ByteBuff(0) = CByte(ByteBuff(0) Or &H40)
        End If

        'the length is stored in 2nd byte
        ByteBuff(1) = CByte(ByteLength Or &H80&)

        If ByteLength Then
            CopyMemory ByteBuff(2), mKey(0), 4&
            CopyMemory ByteBuff(6), ByteData(0), ByteLength
        End If


    ElseIf ByteLength <= 65535 Then     'bigger packets

        ReDim ByteBuff(ByteLength + 7) As Byte   '8 byte overhead

        If SetFINBit Then
            ByteBuff(0) = CByte(&H80 Or OpCode)
        Else
            'the only time FIN (FINished) bit is *NOT* set is when a large
            'buffer is being sent in chunks (see notes below)
            ByteBuff(0) = CByte(OpCode)
        End If

        'set iscompressed bit
        If isCompressed Then
            ByteBuff(0) = CByte(ByteBuff(0) Or &H40)
        End If

        ByteBuff(1) = CByte(&HFE)       ' fixed mask bit + 126

        'copy low-order bytes from data length
        CopyMemory l(0), ByteLength, 2

        'swap in network byte order
        ByteBuff(2) = l(1)
        ByteBuff(3) = l(0)

        If ByteLength Then
            'copy the masking key we used
            CopyMemory ByteBuff(4), mKey(0), 4

            'copy remaining data
            CopyMemory ByteBuff(8), ByteData(0), ByteLength
        End If

    ElseIf ByteLength <= 2147483633 Then     'MAX_LONG = 2147483647 aka &H7FFFFFFF (less 14 bytes for overhead)

        ReDim ByteBuff(ByteLength + 13) As Byte    '14 byte overhead

        If SetFINBit Then
            ByteBuff(0) = CByte(&H80 Or OpCode)
        Else
            ByteBuff(0) = CByte(OpCode)
        End If

        If isCompressed Then
            ByteBuff(0) = CByte(ByteBuff(0) Or &H40)
        End If

        ByteBuff(1) = CByte(&HFF)       ' fixed mask bit + 127

        'copy length to temp 4 byte buffer
        CopyMemory l(0), ByteLength, 4&

        'swap the bytes (big-endian,network order)
        ByteBuff(9) = l(0)
        ByteBuff(8) = l(1)
        ByteBuff(7) = l(2)
        ByteBuff(6) = l(3)

        'bytes 2-5 are unused, because we dont handle data that large

        If ByteLength Then
            CopyMemory ByteBuff(10), mKey(0), 4
            CopyMemory ByteBuff(14), ByteData(0), ByteLength
        End If

    End If

    CompileMessage = ByteBuff

End Function



' ==================== =========================================
' control frames
' ================================= ============================
Public Function PingFrame(P() As Byte, ByVal pLen As Long) As Byte()
    If pLen = 0 Then
        PingFrame = CompileMessage((vbNullChar), 0, False, OpPing, True)
    Else
        PingFrame = CompileMessage(P, pLen, False, OpPing, True)
    End If
End Function


Public Function PongFrame(P() As Byte, ByVal pLen As Long) As Byte()

    If pLen = 0 Then
        PongFrame = CompileMessage((vbNullChar), 0, False, opPong, True)
    Else
        PongFrame = CompileMessage(P, pLen, False, opPong, True)
    End If
End Function



Public Function CloseFrame(P() As Byte, ByVal pLen As Long) As Byte()

    If pLen = 0 Then
        CloseFrame = CompileMessage((vbNullChar), 0, False, opClose, True)
    Else
        CloseFrame = CompileMessage(P, pLen, False, opClose, True)
    End If

End Function


'generate a random number
Function RandomNumber(Optional ByVal Minval As Long, Optional ByVal Maxval As Long) As Long

    'make sure to call randomize(timer) at least once on program start up
    'before using this function

    RandomNumber = ((Maxval - Minval) * Rnd) + Minval
    If RandomNumber > Maxval Or RandomNumber < Minval Then
        RandomNumber = RandomNumber(Minval, Maxval)
    End If

End Function


'generate the client secret key
Function GenerateSecretKey(ByVal nSize As Long) As Byte()

    Dim k() As Byte, X As Long

    ReDim k(nSize - 1) As Byte

    For X = 0 To nSize - 1
        k(X) = CByte(RandomNumber(33, 255))    'using (mostly) printable characters
    Next X

    GenerateSecretKey = k

    'could also use pvCryptoRandomBytes()

End Function


'makes sure user passed in protocol string is properly formatted
Function FormatProtocols(ByVal strProts As String) As String

    Dim X As Long
    Dim arr() As String

    strProts = Trim$(strProts)

    If Len(strProts) Then
        If InStr(strProts, ",") Then
            arr = Split(strProts, ",")
        Else
            ReDim arr(0) As String
            arr(0) = strProts
        End If

        For X = 0 To UBound(arr)
            FormatProtocols = FormatProtocols & Trim$(arr(X)) & ", "
        Next

        FormatProtocols = Left$(FormatProtocols, Len(FormatProtocols) - 2)
    End If

End Function

'utf8 byte array to vb string
Public Function StringUTF8(baText() As Byte, Optional ByVal CodePage As Long = 65001) As String
    Dim lSize As Long
    Dim lCount As Long

    On Error GoTo EH
    lCount = UBound(baText) + 1
    If lCount > 0 Then
        StringUTF8 = String$(2 * lCount, 0)
        lSize = MultiByteToWideChar(CodePage, 0, baText(0), lCount, StrPtr(StringUTF8), ((lCount * 2) + 1))
        If lSize <> Len(StringUTF8) Then
            StringUTF8 = Left$(StringUTF8, lSize)
        End If
    End If
EH:
    On Error GoTo 0

End Function

'vb string to utf8 byte array
Public Function ByteUTF8(ByVal sText As String, Optional ByVal CodePage As Long = 65001) As Byte()
    Dim baRetVal() As Byte
    Dim lSize As Long

    On Error GoTo EH
    If LenB(sText) <> 0 Then
        ReDim baRetVal(2 * Len(sText)) As Byte
        lSize = WideCharToMultiByte(CodePage, 0, StrPtr(sText), Len(sText), baRetVal(0), UBound(baRetVal) + 1, 0, 0)
        If lSize > 0 Then
            ReDim Preserve baRetVal(lSize - 1) As Byte
        Else
            lSize = WideCharToMultiByte(CodePage, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
            ReDim baRetVal(lSize - 1) As Byte
            lSize = WideCharToMultiByte(CodePage, 0, StrPtr(sText), Len(sText), baRetVal(0), UBound(baRetVal) + 1, 0, 0)
        End If
    Else
        baRetVal = vbNullString
    End If

    ByteUTF8 = baRetVal
EH:
    On Error GoTo 0
End Function

'utf16 byte array to utf8 byte array
Public Function ByteUTF16ToUTF8(Bytes() As Byte, Optional ByVal CodePage As Long = 65001) As Byte()
    Dim baRetVal() As Byte
    Dim lSize As Long

    On Error GoTo EH
    lSize = WideCharToMultiByte(CodePage, 0, Bytes(0), (UBound(Bytes) + 1), ByVal 0, 0, 0, 0)
    ReDim baRetVal(lSize - 1) As Byte
    lSize = WideCharToMultiByte(CodePage, 0, Bytes(0), (UBound(Bytes) + 1), baRetVal(0), UBound(baRetVal) + 1, 0, 0)

    ByteUTF16ToUTF8 = baRetVal
EH:
    On Error GoTo 0
End Function


'test a string to see if it is a number
Function IsNumber(ByVal strTest As String) As Boolean

    If Len(strTest) Then
        IsNumber = IsNumeric(strTest)
    End If

End Function

Function StrToNum(ByVal strNum As String, Optional ByVal DefaultNum As Long) As Long

    strNum = Trim$(strNum)

    If IsNumber(strNum) Then
        StrToNum = CLng(strNum)
    Else
        StrToNum = DefaultNum
    End If

End Function


'safely check byte array count
Function ArrayCount(arrBytes() As Byte) As Long
    Dim lPtr As Long

    '    On Error Resume Next
    '    ArrayCount = UBound(arrBytes) + 1
    '    On Error GoTo 0
    
    CopyMemory lPtr, ByVal ArrPtr(arrBytes), 4
    If lPtr Then
        If SafeArrayGetDim(lPtr) = 1 Then
            ArrayCount = UBound(arrBytes) + 1
        End If
    End If

End Function


