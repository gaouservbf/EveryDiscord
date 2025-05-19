VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "EveryDiscord"
   ClientHeight    =   8370
   ClientLeft      =   105
   ClientTop       =   555
   ClientWidth     =   10215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2520
      TabIndex        =   13
      Top             =   7560
      Width           =   435
   End
   Begin VB.Timer Timer4 
      Left            =   2220
      Top             =   870
   End
   Begin MSWinsockLib.Winsock wsGuild 
      Left            =   480
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsMessageFetch 
      Index           =   0
      Left            =   600
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   443
   End
   Begin MSWinsockLib.Winsock wsMessageFetch 
      Index           =   1
      Left            =   0
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   443
   End
   Begin MSWinsockLib.Winsock wsChannelFetch 
      Index           =   2
      Left            =   600
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   443
   End
   Begin MSWinsockLib.Winsock wsChannelFetch 
      Index           =   1
      Left            =   0
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   443
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   7995
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Connecting..."
            TextSave        =   "Connecting..."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsGateway 
      Left            =   4920
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   443
   End
   Begin VB.Timer Timer2 
      Interval        =   2500
      Left            =   9120
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   9240
      Top             =   240
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   2775
      TabIndex        =   5
      Top             =   7080
      Width           =   2775
      Begin VB.CommandButton Command2 
         Caption         =   "Config>>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         TabIndex        =   14
         Top             =   525
         Width           =   855
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   630
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   12
         Top             =   105
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1365
         TabIndex        =   7
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.ListBox lstChannel 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtCID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   3
      Text            =   "1355226543160557778"
      Top             =   6600
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.TextBox txtToken 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   2
      Text            =   "Token"
      Top             =   7080
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.TextBox txtMsg 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   7560
      Width           =   6375
   End
   Begin VB.CommandButton cmdSendMsg 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   0
      Top             =   7560
      Width           =   735
   End
   Begin MSWinsockLib.Winsock wscSocket 
      Index           =   0
      Left            =   3570
      Top             =   1230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   443
   End
   Begin BareboneThunks.ChatView ChatView1 
      Height          =   6975
      Left            =   3480
      TabIndex        =   9
      Top             =   0
      Width           =   6735
      _extentx        =   9763
      _extenty        =   8493
   End
   Begin MSWinsockLib.Winsock wsGuildIcon 
      Index           =   1
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   443
   End
   Begin MSWinsockLib.Winsock wsGuildIcon 
      Index           =   2
      Left            =   600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   443
   End
   Begin BareboneThunks.GuildView GuildView1 
      Height          =   6975
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   1335
      _extentx        =   2355
      _extenty        =   12303
   End
   Begin MSWinsockLib.Winsock wsMessage 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Guild Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin VB.Menu mnuMessages 
      Caption         =   "Messages"
      Visible         =   0   'False
      Begin VB.Menu delms 
         Caption         =   "Delete Message"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' Add this as a module or class-level variable

' WebSocket Control needed for Gateway connection

' WebSocket constants for Gateway connection
Private Const DISCORD_GATEWAY_URL As String = "gateway.discord.gg"
Private Const DISCORD_GATEWAY_VERSION As String = "10"
Private Const DISCORD_GATEWAY_PORT As Long = 443

Private m_lFreeSocketIndex As Long     ' Next available socket index

' Control array constants
Private Const MAX_CONNECTIONS As Long = 255
Private Const SOCKET_IDLE As Long = 0
Private Const SOCKET_CONNECTING As Long = 1
Private Const SOCKET_CONNECTED As Long = 2
Private Const SOCKET_RECEIVING As Long = 3

' TLS Context and Socket variables
Private m_GuildIds() As String
Private m_ChannelIds() As String
Private m_uCtx() As UcsTlsContext      ' Array of TLS contexts
Private m_sRequest() As String         ' Array of requests
Private m_sToken As String
Private m_sBaseUrl As String

' Socket state tracking
Private m_lSocketState() As Long       ' State of each socket
Private m_sResponseBuffer() As String  ' Response buffer for each socket
Private m_lContentLength() As Long     ' Content length for each socket
Private m_bReceivingData() As Boolean  ' Receiving flag for each socket
Private m_RequestType() As String      ' Request type for each socket
Private m_ExtraData() As String        ' Extra data for each socket
Private m_Index() As Long              ' Index for each socket
Private m_Target() As String           ' Target for each socket

' Gateway state tracking
Private m_sSessionId As String
Private m_lSequence As Long
Private m_lHeartbeatInterval As Long
Private m_dLastHeartbeat As Double
Private m_bGatewayConnected As Boolean
Private m_bIdentified As Boolean
Private m_sGatewayToken As String
Private m_sCurrentChannelId As String
Private WithEvents tmrHeartbeat As Timer
Attribute tmrHeartbeat.VB_VarHelpID = -1

' Add these to your declarations section
Private Type IconRequest
    GuildId As String
    iconHash As String
    guildIndex As Long
End Type

Private IconRequests() As IconRequest
Private IconRequestCount As Long
Private CurrentIconRequest As Long
Private m_bFetchingIcon As Boolean

Private Type RequestItem
    RequestType As String    ' "GuildList", "Channels", "Messages", "Icon", etc.
    Target As String         ' Guild ID, Channel ID, etc. depending on type
    Request As String        ' The full HTTP request
    ExtraData As String      ' Additional data (e.g., Icon hash, index)
    Index As Long            ' For indexing into arrays if needed
End Type

Private m_RequestQueue() As RequestItem
Private m_QueueCount As Long
Private m_ProcessingRequest As Boolean
Private m_CurrentRequestType As String
Private m_GatewayBuffer As String


Private Sub Connect(ByVal socketIndex As Long, ByVal sServer As String, ByVal lPort As Long)
    If socketIndex < 0 Or socketIndex >= MAX_CONNECTIONS Then
        Exit Sub  ' Invalid index
    End If
    
    ' Initialize TLS context for this connection
    Call TlsInitClient(m_uCtx(socketIndex), sServer)
    
    ' Close socket if already connected
    If m_lSocketState(socketIndex) <> SOCKET_IDLE Then
        wscSocket(socketIndex).Close
    End If
    
    ' Set state and connect
    m_lSocketState(socketIndex) = SOCKET_CONNECTING
    wscSocket(socketIndex).Connect sServer, lPort
End Sub
Private Sub SendData(ByVal socketIndex As Long, baData() As Byte)
    Dim baOutput() As Byte
    Dim lOutputPos As Long
    
    If socketIndex < 0 Or socketIndex >= MAX_CONNECTIONS Then
        Exit Sub  ' Invalid index
    End If
    
    If Not TlsSend(m_uCtx(socketIndex), baData, UBound(baData) + 1, baOutput, lOutputPos) Then
        OnError TlsGetLastError(m_uCtx(socketIndex)), "TlsSend", socketIndex
    End If
    
    If lOutputPos > 0 Then
        Debug.Assert UBound(baOutput) + 1 = lOutputPos
        wscSocket(socketIndex).SendData baOutput
    End If
End Sub

Private Sub Form_Resize()
    ChatView1.Width = Me.ScaleWidth - GuildView1.Width - lstChannel.Width
End Sub

Private Sub GuildView1_GuildSelected(ByVal Index As Long)
    Dim SelectedIndex As Long
    Dim GuildId As String
    
    SelectedIndex = Index
    
    If SelectedIndex >= 0 And SelectedIndex < UBound(m_GuildIds) + 1 Then
        ' Get the ID from our parallel array
        GuildId = m_GuildIds(SelectedIndex)
        
        ' Fetch channels for this guild
        FetchGuildChannels GuildId
        Label3.Caption = "<" & GuildView1.GetGuildName(Index) & ">"
        Me.Caption = "EveryDiscord - " & GuildView1.GetGuildName(Index)
    End If
End Sub

Private Sub Timer1_Timer()
   ' FetchUserGuilds
End Sub

Private Sub Timer2_Timer()
    'FetchChannelMessages txtCID.Text
End Sub

Private Sub wscSocket_Connect(Index As Integer)
    Dim baEmpty() As Byte
    Dim baOutput() As Byte
    Dim lOutputPos As Long
    
    On Error GoTo EH
    
    ' Update state
    m_lSocketState(Index) = SOCKET_CONNECTED
    
    ' Begin TLS handshake
    If Not TlsHandshake(m_uCtx(Index), baEmpty, -1, baOutput, lOutputPos) Then
        OnError TlsGetLastError(m_uCtx(Index)), "TlsHandshake", Index
    End If
    
    If lOutputPos > 0 Then
        Debug.Assert UBound(baOutput) + 1 = lOutputPos
        wscSocket(Index).SendData baOutput
    End If
    
    Exit Sub
EH:
    OnError Err.Description, "wscSocket_Connect", Index
End Sub


Private Sub wscSocket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next
    
    Dim bError As Boolean
    Dim baRecv() As Byte
    Dim baOutput() As Byte
    Dim lOutputPos As Long
    Dim baPlainText() As Byte
    Dim lSize As Long
    
    If TlsIsClosed(m_uCtx(Index)) Or bytesTotal = 0 Then
        Exit Sub
    End If
    
    ' Get received data
    wscSocket(Index).GetData baRecv
    
    ' Handle TLS handshake if not ready
    If Not TlsIsReady(m_uCtx(Index)) Then
        lOutputPos = 0
        bError = Not TlsHandshake(m_uCtx(Index), baRecv, -1, baOutput, lOutputPos)
        
        If lOutputPos > 0 Then
            Debug.Assert UBound(baOutput) + 1 = lOutputPos
            wscSocket(Index).SendData baOutput
        End If
        
        If bError Then
            OnError TlsGetLastError(m_uCtx(Index)), "TlsHandshake", Index
        End If
        
        If Not TlsIsReady(m_uCtx(Index)) Then
            Exit Sub
        End If
        
        ' Connection established, send request
        OnConnect Index
        
        ' Fall through to process any remaining data
        Erase baRecv
    End If
    
    ' Process data
    lOutputPos = 0
    bError = Not TlsReceive(m_uCtx(Index), baRecv, -1, baPlainText, lSize, baOutput, lOutputPos)
    
    If lOutputPos > 0 Then
        wscSocket(Index).SendData baOutput
    End If
    
    If lSize > 0 Then
        OnDataArrival Index, lSize, baPlainText
    End If
    
    If bError Then
        OnError TlsGetLastError(m_uCtx(Index)), "TlsReceive", Index
    End If
    
    If TlsIsClosed(m_uCtx(Index)) Then
        OnClose Index
    End If
End Sub
Private Sub wscSocket_Close(Index As Integer)
    If Not TlsIsClosed(m_uCtx(Index)) Then
        OnClose Index
    End If
End Sub

' Heartbeat timer
Private Sub tmrHeartbeat_Timer()
    
End Sub

' Add this to Form_Unload
Private Sub Form_Unload(Cancel As Integer)
    ' Clean up WebSocket and timers
    If Not wsGateway Is Nothing Then
        wsGateway.Close
    End If
    
    If Not tmrHeartbeat Is Nothing Then
        tmrHeartbeat.Enabled = False
        Set tmrHeartbeat = Nothing
    End If
End Sub
' Called when socket connected and TLS handshake completed
Private Sub OnConnect(ByVal socketIndex As Long)
    ' Send HTTP request
    SendData socketIndex, StrConv(m_sRequest(socketIndex), vbFromUnicode)
End Sub
Function ParseHttpResponse(rawData As String, bFetchingIcon As Boolean) As String
    ' For binary data (icons), don't apply text processing
    If bFetchingIcon Then
        ParseHttpResponse = rawData
        Exit Function
    End If
    
    ' Check if chunked encoding is used
    ' Default: Extract body after headers
    Dim headersEnd As Long
    Dim responseBody As String
    
    headersEnd = InStr(rawData, vbCrLf & vbCrLf)
    If headersEnd > 0 Then
        responseBody = Mid$(rawData, headersEnd + 4)
    Else
        responseBody = rawData
    End If
    
    ' Only for text responses, remove last characters if string is long enough
    If Len(responseBody) >= 3 Then
        ParseHttpResponse = Left$(responseBody, Len(responseBody) - 3)
    Else
        ParseHttpResponse = responseBody
    End If
End Function


Private Sub OnDataArrival(ByVal socketIndex As Long, ByVal bytesTotal As Long, baData() As Byte)
    Debug.Print "OnDataArrival, Socket=" & socketIndex & ", bytesTotal=" & bytesTotal, Timer
    
    ' Process response
    Dim sResponse As String
    sResponse = StrConv(baData, vbUnicode)
    
    ' Check if this is the start of a response or continuation
    If Not m_bReceivingData(socketIndex) Then
        m_bReceivingData(socketIndex) = True
        m_sResponseBuffer(socketIndex) = sResponse
        
        ' Extract Content-Length if available
        Dim clPos As Long
        clPos = InStr(1, sResponse, "Content-Length:", vbTextCompare)
        If clPos > 0 Then
            Dim clEnd As Long
            clEnd = InStr(clPos, sResponse, vbCrLf)
            If clEnd > 0 Then
                m_lContentLength(socketIndex) = Val(Mid$(sResponse, clPos + 15, clEnd - (clPos + 15)))
            End If
        End If
    Else
        ' Append to existing buffer
        m_sResponseBuffer(socketIndex) = m_sResponseBuffer(socketIndex) & sResponse
    End If
    
    ' Check if we have the complete response
    Dim headersEnd As Long
    Dim contentReceived As Long
    
    headersEnd = InStr(m_sResponseBuffer(socketIndex), vbCrLf & vbCrLf)
    If headersEnd > 0 Then
        contentReceived = Len(m_sResponseBuffer(socketIndex)) - (headersEnd + 3)
        
        ' For chunked encoding, check for final chunk
        If InStr(1, m_sResponseBuffer(socketIndex), "Transfer-Encoding: chunked", vbTextCompare) > 0 Then
            If InStr(m_sResponseBuffer(socketIndex), vbCrLf & "0" & vbCrLf & vbCrLf) > 0 Then
                ProcessHttpResponse socketIndex
            End If
        ' For Content-Length, check if we've received enough data
        ElseIf m_lContentLength(socketIndex) > 0 Then
            If contentReceived >= m_lContentLength(socketIndex) Then
                ProcessHttpResponse socketIndex
            End If
        ' If we can't determine length, assume this is all we'll get
        Else
            ProcessHttpResponse socketIndex
        End If
    End If
End Sub
Private Sub ResetSocket(ByVal socketIndex As Long)
    ' Skip if invalid index
    If socketIndex < 0 Or socketIndex >= MAX_CONNECTIONS Then
        Exit Sub
    End If
    
    ' Close the socket
    wscSocket(socketIndex).Close
    
    ' Reset state and buffers
    m_lSocketState(socketIndex) = SOCKET_IDLE
    m_sResponseBuffer(socketIndex) = ""
    m_lContentLength(socketIndex) = 0
    m_bReceivingData(socketIndex) = False
    m_RequestType(socketIndex) = ""
    m_ExtraData(socketIndex) = ""
    m_Index(socketIndex) = -1
    m_Target(socketIndex) = ""
End Sub

' Called when socket closed
Private Sub ProcessHttpResponse(ByVal socketIndex As Long)
    Dim sResponse As String
    Dim bFetchingIcon As Boolean
    
    ' Skip if invalid index
    If socketIndex < 0 Or socketIndex >= MAX_CONNECTIONS Then
        Exit Sub
    End If
    
    ' Get response from buffer
    sResponse = m_sResponseBuffer(socketIndex)
    
    ' Check if this is an icon request
    bFetchingIcon = (m_RequestType(socketIndex) = "Icon")
    
    ' Process based on request type
    Select Case m_RequestType(socketIndex)
        Case "Icon"
            ProcessIconResponse socketIndex, sResponse
            
        Case "GuildList"
            Dim sContent As String
            sContent = ParseHttpResponse(sResponse, bFetchingIcon)
            If InStr(sContent, vbCrLf) > 0 Then
                sContent = Mid$(sContent, InStr(sContent, vbCrLf) + 2)
            End If
            ProcessGuildsResponse sContent
            
        Case "Channels"
            Dim sChannelContent As String
            sChannelContent = ParseHttpResponse(sResponse, bFetchingIcon)
            If InStr(sChannelContent, vbCrLf) > 0 Then
                sChannelContent = Mid$(sChannelContent, InStr(sChannelContent, vbCrLf) + 2)
            End If
            ProcessChannelsResponse sChannelContent
            
        Case "Messages"
            Dim sMessageContent As String
            sMessageContent = ParseHttpResponse(sResponse, bFetchingIcon)
            If InStr(sMessageContent, vbCrLf) > 0 Then
                sMessageContent = Mid$(sMessageContent, InStr(sMessageContent, vbCrLf) + 2)
            End If
            ProcessMessagesResponse sMessageContent
            
        Case "SendMessage"
            ' No special handling needed for sent messages
            ' You could add code to refresh messages in current channel
            
        Case Else
            ' Unknown request type
            Debug.Print "Unknown request type: " & m_RequestType(socketIndex)
    End Select
    
    ' Reset socket state and buffers
    ResetSocket socketIndex
End Sub

Private Sub OnClose(ByVal socketIndex As Long)
    Debug.Print "OnClose, Socket=" & socketIndex, Timer
    If Len(m_sResponseBuffer(socketIndex)) > 0 Then
        ProcessHttpResponse socketIndex
    Else
        ResetSocket socketIndex
    End If
    DequeueAndProcessNext ' Add this to trigger next request
End Sub
' Called on socket or TLS error
Private Sub OnError(sDescription As String, sSource As String, socketIndex As Variant)
    Debug.Print "Critical error: " & sDescription & " in " & sSource & ", Socket=" & socketIndex, Timer
    MsgBox "Discord API Error: " & sDescription & " in " & sSource, vbCritical
    ResetSocket socketIndex
End Sub
Private Sub ProcessGuildsResponse(aJson As String)

    Dim parsed As ParseResult
    Dim i As Long
    Dim GuildCount As Long
    Dim sjson As String
    
    sjson = Left$(aJson, Len(aJson) - 5)
    ' Parse the JSON array
    parsed = Parse(sjson)
 
    If Not parsed.IsValid Then
        Exit Sub
    End If
    ' Clear existing guilds
    GuildView1.ClearGuilds
    
GuildView1.AddGuild "DMs", LoadPicture(App.Path & "\everydiscord.gif")
    ' Count guilds first to properly size the array
    GuildCount = 0
    For i = 0 To parsed.Value.Count
        On Error Resume Next
        Dim guildCheck As Object
        Set guildCheck = parsed.Value(i)
        If Not guildCheck Is Nothing Then GuildCount = GuildCount + 1
        On Error GoTo 0
    Next i
    
    ' Resize the array to match guild count
    ReDim m_GuildIds(0 To GuildCount - 1) As String
    
    ' Process each guild
    Dim validGuilds As Long
    validGuilds = 0
    
    For i = 1 To parsed.Value.Count
        On Error Resume Next
        Dim Guild As Object
        Set Guild = parsed.Value(i)
        
        ' Skip if no more guilds or error
        If Guild Is Nothing Then
            Exit For
        End If
        
        ' Extract guild details
        Dim sGuildName As String
        Dim sGuildId As String
        Dim sIconHash As String
        Dim guildIcon As StdPicture
        
        sGuildName = Guild("name")
        sGuildId = Guild("id")
        
        ' Get icon if available
        On Error Resume Next
        sIconHash = Guild("icon")
        On Error GoTo 0
        
        ' Add to GuildView with placeholder icon
        ' You'll need to implement icon fetching in a separate function
        Set guildIcon = LoadPicture() ' Default empty icon
        
        ' If icon hash exists, fetch it
        If Len(sIconHash) > 0 Then
            ' Queue icon for fetching (implement this separately)
            QueueGuildIconFetch sGuildId, sIconHash, validGuilds
        End If
        
        GuildView1.AddGuild sGuildName, guildIcon
        
        ' Store ID in parallel array
        m_GuildIds(validGuilds) = sGuildId
        validGuilds = validGuilds + 1
        
        ' Debug output
        Debug.Print "Added guild: " & sGuildName & " with ID: " & sGuildId
    Next i
End Sub
Private Sub ProcessChannelsResponse(aJson As String)
    Dim parsed As ParseResult
    Dim i As Long
    Dim channelCount As Long
    Dim sjson As String
    Dim validChannels As Long
    ' Clean up the JSON input string - remove trailing characters if present
    If Len(aJson) > 5 Then
        sjson = Left$(aJson, Len(aJson) - 5)
    Else
        sjson = aJson
    End If
    
    ' Parse the JSON array
    parsed = Parse(sjson)
 
    If Not parsed.IsValid Then
        MsgBox "Error parsing channel data: " & parsed.Error, vbExclamation
        Exit Sub
    End If
    
    ' Clear existing channels
    lstChannel.Clear
 
    For i = 1 To parsed.Value.Count  ' Increased to handle larger guilds
        On Error Resume Next
        Dim channel As Object
        Set channel = parsed.Value(i)
        
        
        ' Extract channel details
        Dim sChannelName As String
        Dim sChannelId As String
        Dim lChannelType As Long
        
        sChannelName = channel("name")
        sChannelId = channel("id")
        lChannelType = channel("type")
        
        ' Only add text channels (type 0)
        If lChannelType = 0 Then
            ' Add to listbox
            lstChannel.AddItem sChannelName
            
            ' Store ID in parallel array (check bounds first)
            If validChannels <= UBound(m_ChannelIds) Then
                m_ChannelIds(validChannels) = sChannelId
                validChannels = validChannels + 1
                
                ' Debug output
                Debug.Print "Added channel: " & sChannelName & " with ID: " & sChannelId
            End If
        End If
        On Error GoTo 0
    Next i
    
    ' If we found channels, select the first one
    If lstChannel.ListCount > 0 Then
        lstChannel.ListIndex = 0
    End If
End Sub

Private Sub lstGuild_Click()
    ' Get the selected guild ID
End Sub

Private Sub lstChannel_Click()
    ' Get the selected channel ID
    Dim SelectedIndex As Long
    Dim channelId As String
    MsgBox lstChannel.ListIndex
    SelectedIndex = lstChannel.ListIndex
    
    If SelectedIndex >= 0 And SelectedIndex < UBound(m_ChannelIds) + 1 Then
        ' Get the ID from our parallel array
        channelId = m_ChannelIds(SelectedIndex)
        Debug.Print "Selected channel ID: " & channelId
        
        ' Set the channel ID in the textbox
        txtCID.Text = channelId
        
        ' Store current channel ID for Gateway message filtering
        m_sCurrentChannelId = channelId
        
        ' Fetch messages for this channel
        FetchChannelMessages channelId
    End If
End Sub

Private Sub ProcessMessagesResponse(aJson As String)
    Dim parsed As ParseResult
    Dim i As Long
    Dim sOutput As String
    Dim fileNum As Integer
    Dim desktopPath As String
    Dim filePath As String
    Dim sjson As String
    sjson = Left$(aJson, Len(aJson) - 5)
    ' Parse the JSON array
    parsed = Parse(sjson)
 
    If Not parsed.IsValid Then
        Exit Sub
    End If
    
    ' Clear existing messages
    ChatView1.Clear
    ' Process each message
    For i = 20 To 1 Step -1
    On Error Resume Next
        Dim msg As Object
        Set msg = parsed.Value(i)
        
        ' Extract message details
        Dim sAuthor As String
        Dim sContent As String
        Dim sTimestamp As String
        
        On Error Resume Next ' Handle potential missing fields
        sAuthor = msg("author")("username") & "#" & msg("author")("discriminator")
        sContent = msg("content")
        sTimestamp = FormatDiscordTimestamp(msg("timestamp"))
        
        ' Format the output
        sOutput = "[" & sTimestamp & "] " & sAuthor & ": " & sContent
        
        ' Add to listbox
        ChatView1.AddMessage sAuthor, sContent
    Next i
End Sub

Private Sub cmdFetchMessages_Click()
    ' Validate inputs
    If Len(Trim(txtToken.Text)) = 0 Then
        MsgBox "Please enter your Discord user token", vbExclamation
        Exit Sub
    End If
    
    If Len(Trim(txtCID.Text)) = 0 Then
        MsgBox "Please enter a channel ID", vbExclamation
        Exit Sub
    End If
    
    m_sToken = txtToken.Text
    
    ' Save settings
    SaveSetting "DiscordClient", "Settings", "Token", txtToken.Text
    SaveSetting "DiscordClient", "Settings", "ChannelId", txtCID.Text
    
    ' Fetch messages
    FetchChannelMessages txtCID.Text
End Sub

' Add a button to connect to the Gateway
Private Sub cmdConnectGateway_Click()
    ' Validate token first
    If Len(Trim(txtToken.Text)) = 0 Then
        MsgBox "Please enter your Discord user token first", vbExclamation
        Exit Sub
    End If
    
    m_sToken = txtToken.Text
    
    ' Connect to Gateway
    
End Sub

Private Function FormatDiscordTimestamp(sTimestamp As String) As String
    ' Convert Discord's ISO 8601 timestamp to local time
    On Error Resume Next
    Dim dDate As Date
    Dim sDatePart As String
    Dim sTimePart As String
    
    ' Extract date and time parts (format: "2023-04-01T12:34:56.789000+00:00")
    sDatePart = Left$(sTimestamp, 10)
    sTimePart = Mid$(sTimestamp, 12, 8)
    
    ' Parse as local date/time
    dDate = CDate(sDatePart & " " & sTimePart)
    
    If Err.Number = 0 Then
        FormatDiscordTimestamp = Format$(dDate, "yyyy-mm-dd hh:nn:ss")
    Else
        FormatDiscordTimestamp = sTimestamp ' Return original if parsing fails
        Err.Clear
    End If
End Function



' Form initialization
Private Sub Form_Load()
    ' Initialize the form
    m_sBaseUrl = "discord.com"
    
    ' Initialize socket control array and related arrays
    InitializeSocketArray
    
    ' Initialize guild and channel arrays
    ReDim m_GuildIds(0) As String
    ReDim m_ChannelIds(0) As String
    
    ' Auto load token from settings if available
    If GetSetting("DiscordClient", "Settings", "Token", "") <> "" Then
        txtToken.Text = GetSetting("DiscordClient", "Settings", "Token", "")
        m_sToken = txtToken.Text
    End If
    
    ' Auto load channel from settings if available
    If GetSetting("DiscordClient", "Settings", "ChannelId", "") <> "" Then
        txtCID.Text = GetSetting("DiscordClient", "Settings", "ChannelId", "")
        
        ' Auto-fetch messages if we have both token and channel
        If Len(m_sToken) > 0 Then
            FetchChannelMessages txtCID.Text
            FetchUserGuilds
        End If
    End If
End Sub

Private Sub InitializeSocketArray()
    Dim i As Integer
    
    ReDim m_uCtx(0 To MAX_CONNECTIONS - 1)
    ReDim m_sRequest(0 To MAX_CONNECTIONS - 1)
    ReDim m_lSocketState(0 To MAX_CONNECTIONS - 1)
    ReDim m_sResponseBuffer(0 To MAX_CONNECTIONS - 1)
    ReDim m_lContentLength(0 To MAX_CONNECTIONS - 1)
    ReDim m_bReceivingData(0 To MAX_CONNECTIONS - 1)
    ReDim m_RequestType(0 To MAX_CONNECTIONS - 1)
    ReDim m_ExtraData(0 To MAX_CONNECTIONS - 1)
    ReDim m_Index(0 To MAX_CONNECTIONS - 1)
    ReDim m_Target(0 To MAX_CONNECTIONS - 1)
    
    ' Initialize each Winsock control
    For i = 0 To MAX_CONNECTIONS - 1  ' Start from 0, not 1
        If i > 0 Then Load wscSocket(i)  ' Only load if not the base control
        
        ' Set initial state
        m_lSocketState(i) = SOCKET_IDLE
        m_sResponseBuffer(i) = ""
        m_lContentLength(i) = 0
        m_bReceivingData(i) = False
        m_RequestType(i) = ""
        m_ExtraData(i) = ""
        m_Index(i) = i  ' Store the actual index
        m_Target(i) = ""
    Next i
    
    ' Set initial free socket index
    m_lFreeSocketIndex = 0
End Sub
Private Sub ProcessNextQueuedRequest()
    Dim socketIndex As Long
    
If m_QueueCount = 0 Then
    Exit Sub
End If
    ' Find a free socket
    socketIndex = GetFreeSocketIndex()
    If socketIndex < 0 Then
        ' No free sockets available, try again later
        Exit Sub
    End If
    
    ' Get the next request from the queue and mark as processing
    m_ProcessingRequest = True
    
    ' Set up the socket with the request data
    m_RequestType(socketIndex) = m_RequestQueue(0).RequestType
    m_sRequest(socketIndex) = m_RequestQueue(0).Request
    m_Target(socketIndex) = m_RequestQueue(0).Target
    m_ExtraData(socketIndex) = m_RequestQueue(0).ExtraData
    m_Index(socketIndex) = m_RequestQueue(0).Index
    
    ' Connect to the appropriate server based on request type
    If m_RequestType(socketIndex) = "Icon" Then
        Connect socketIndex, "cdn.discordapp.com", 443
    Else
        Connect socketIndex, m_sBaseUrl, 443
    End If
End Sub
Private Function GetFreeSocketIndex() As Long
    Dim i As Long
    
    For i = 0 To MAX_CONNECTIONS - 1
        If m_lSocketState(i) = SOCKET_IDLE Then
            GetFreeSocketIndex = i
            Exit Function
        End If
    Next i
    
    ' No free socket found
    GetFreeSocketIndex = -1
End Function
' Add a request to the queue
Private Sub QueueRequest(RequestType As String, Target As String, Request As String, Optional ExtraData As String = "", Optional Index As Long = -1)
    ' Resize the queue array
    If m_QueueCount = 0 Then
        ReDim m_RequestQueue(0)
    Else
        ReDim Preserve m_RequestQueue(m_QueueCount)
    End If
    
    ' Add the request
    With m_RequestQueue(m_QueueCount)
        .RequestType = RequestType
        .Target = Target
        .Request = Request
        .ExtraData = ExtraData
        .Index = Index
    End With
    ' Increment counter
    m_QueueCount = m_QueueCount + 1
    
        ProcessNextQueuedRequest
    
   ' DequeueAndProcessNext
End Sub

' Remove the first request from the queue and process the next one
Private Sub DequeueAndProcessNext()
    ' Make sure we have requests
    If m_QueueCount = 0 Then
        m_ProcessingRequest = False
        Exit Sub
    End If
    
    ' Remove the first request by shifting all others up
    If m_QueueCount > 1 Then
        Dim i As Long
        For i = 0 To m_QueueCount - 2
            m_RequestQueue(i) = m_RequestQueue(i + 1)
        Next i
    End If
    
    ' Decrease count
    m_QueueCount = m_QueueCount - 1
    
    ' Reset processing flag
    m_ProcessingRequest = False
    
    ' Process next request if any
    If m_QueueCount > 0 Then
        ProcessNextQueuedRequest
    End If
End Sub
' Process completed HTTP response for a specific socket
Private Sub ProcessCompleteResponse(ByVal socketIndex As Long)
    ' Skip if invalid index
    If socketIndex < 0 Or socketIndex >= MAX_CONNECTIONS Then
        Exit Sub
    End If
    
    Dim sResponse As String
    Dim bFetchingIcon As Boolean
    
    ' Get response from buffer
    sResponse = m_sResponseBuffer(socketIndex)
    
    ' Check request type
    bFetchingIcon = (m_RequestType(socketIndex) = "Icon")
    
    ' Process based on request type
    Select Case m_RequestType(socketIndex)
        Case "Icon"
            ProcessIconResponse socketIndex, sResponse
            
        Case "GuildList"
            Dim sContent As String
            sContent = ParseHttpResponse(sResponse, bFetchingIcon)
            If InStr(sContent, vbCrLf) > 0 Then
                sContent = Mid$(sContent, InStr(sContent, vbCrLf) + 2)
            End If
            ProcessGuildsResponse sContent
            
        Case "Channels"
            Dim sChannelContent As String
            sChannelContent = ParseHttpResponse(sResponse, bFetchingIcon)
            If InStr(sChannelContent, vbCrLf) > 0 Then
                sChannelContent = Mid$(sChannelContent, InStr(sChannelContent, vbCrLf) + 2)
            End If
            ProcessChannelsResponse sChannelContent
            
        Case "Messages"
            Dim sMessageContent As String
            sMessageContent = ParseHttpResponse(sResponse, bFetchingIcon)
            If InStr(sMessageContent, vbCrLf) > 0 Then
                sMessageContent = Mid$(sMessageContent, InStr(sMessageContent, vbCrLf) + 2)
            End If
            ProcessMessagesResponse sMessageContent
            
        Case "SendMessage"
            ' After sending a message, you might want to refresh the current channel
            If Len(m_sCurrentChannelId) > 0 Then
            End If
            
        Case Else
            ' Unknown request type
            Debug.Print "Unknown request type: " & m_RequestType(socketIndex)
    End Select
    
    ' Reset this socket
    ResetSocket socketIndex
    
    ' Process next request in queue
    m_ProcessingRequest = False
    
    ' Check if there are more requests to process
    If m_QueueCount > 0 Then
        ProcessNextQueuedRequest
    End If
End Sub

' Update existing methods to use the queue system
Private Sub FetchUserGuilds()
    On Error GoTo EH
    
    ' Create the API request path
    Dim sPath As String
    sPath = "api/v10/users/@me/guilds"
    
    ' Prepare the HTTP request
    Dim sRequest As String
    sRequest = "GET /" & sPath & " HTTP/1.1" & vbCrLf & _
              "Host: " & m_sBaseUrl & vbCrLf & _
              "Authorization: " & m_sToken & vbCrLf & _
              "Connection: close" & vbCrLf & vbCrLf
    
    ' Queue this request
    QueueRequest "GuildList", "", sRequest
    Exit Sub
EH:
    MsgBox "Error fetching guilds: " & Err.Description, vbCritical
End Sub

Private Sub FetchGuildChannels(ByVal sGuildId As String)
    On Error GoTo EH
    
    ' Create the API request path
    Dim sPath As String
    sPath = "api/v10/guilds/" & sGuildId & "/channels"
    
    ' Prepare the HTTP request
    Dim sRequest As String
    sRequest = "GET /" & sPath & " HTTP/1.1" & vbCrLf & _
              "Host: " & m_sBaseUrl & vbCrLf & _
              "Authorization: " & m_sToken & vbCrLf & _
              "Connection: close" & vbCrLf & vbCrLf
    
    ' Queue this request
    QueueRequest "Channels", sGuildId, sRequest
    Exit Sub
EH:
    MsgBox "Error fetching channels: " & Err.Description, vbCritical
End Sub

Private Sub FetchChannelMessages(ByVal sChannelId As String, Optional ByVal lLimit As Long = 10)
    On Error GoTo EH
    
    ' Create the API request path
    Dim sPath As String
    sPath = "api/v10/channels/" & sChannelId & "/messages?limit=" & lLimit
    
    ' Prepare the HTTP request
    Dim sRequest As String
    sRequest = "GET /" & sPath & " HTTP/1.1" & vbCrLf & _
              "Host: " & m_sBaseUrl & vbCrLf & _
              "Authorization: " & m_sToken & vbCrLf & _
              "Connection: close" & vbCrLf & vbCrLf
    
    ' Queue this request
    QueueRequest "Messages", sChannelId, sRequest
    Exit Sub
EH:
    MsgBox "Error fetching messages: " & Err.Description, vbCritical
End Sub

Private Sub SendDiscordMessage(ByVal sChannelId As String, ByVal sMessage As String)
    On Error GoTo EH
    
    ' Create the JSON payload
    Dim sPayload As String
    sPayload = "{""content"":""" & EscapeJsonString(sMessage) & """}"
    
    ' Create the API request path
    Dim sPath As String
    sPath = "api/v10/channels/" & sChannelId & "/messages"
    
    ' Prepare the HTTP request
    Dim sRequest As String
    sRequest = "POST /" & sPath & " HTTP/1.1" & vbCrLf & _
              "Host: " & m_sBaseUrl & vbCrLf & _
              "Authorization: " & m_sToken & vbCrLf & _
              "Content-Type: application/json" & vbCrLf & _
              "Content-Length: " & Len(sPayload) & vbCrLf & _
              "Connection: close" & vbCrLf & vbCrLf & _
              sPayload
    
    QueueRequest "SendMessage", sChannelId, sRequest
    Exit Sub
EH:
    MsgBox "Error sending message: " & Err.Description, vbCritical
End Sub

' Update icon fetching system to use the queue
Private Sub QueueGuildIconFetch(ByVal sGuildId As String, ByVal sIconHash As String, ByVal guildIndex As Long)
    ' Create the API request path for JPG icon
    Dim sPath As String
    sPath = "/icons/" & sGuildId & "/" & sIconHash & ".jpg?size=80&quality=lossless"
    
    ' Prepare the HTTP request
    Dim sRequest As String
    sRequest = "GET " & sPath & " HTTP/1.1" & vbCrLf & _
             "Host: cdn.discordapp.com" & vbCrLf & _
             "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36" & vbCrLf & _
             "Accept: image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8" & vbCrLf & _
             "Accept-Language: en-US,en;q=0.9" & vbCrLf & _
             "Referer: https://discord.com/" & vbCrLf & _
             "Authorization: " & m_sToken & vbCrLf & _
             "Connection: close" & vbCrLf & vbCrLf
             
    QueueRequest "Icon", sGuildId, sRequest, sIconHash, guildIndex
End Sub

Private Sub ProcessIconResponse(ByVal socketIndex As Long, ByVal sResponse As String)
    On Error GoTo EH

    Dim headerEnd As Long
    Dim tempFile As String
    Dim isAnimated As Boolean
    Dim guildIndex As Long
    Dim iconHash As String
    
    ' Extract details from current socket
    iconHash = m_ExtraData(socketIndex)
    guildIndex = m_Index(socketIndex)
    
    ' Detect animated icon via hash
    isAnimated = (Left$(iconHash, 2) = "a_")

    ' Find end of HTTP headers
    headerEnd = InStr(sResponse, vbCrLf & vbCrLf)

    If headerEnd > 0 Then
        Dim binaryData As String
        binaryData = Mid$(sResponse, headerEnd + 4)
        
        ' Use .gif for animated, .jpg otherwise
        Dim fileExt As String
        If isAnimated Then
            fileExt = ".gif"
        Else
            fileExt = ".jpg"
        End If
        
        tempFile = Environ$("TEMP") & "\discord_icon_" & Format$(Now, "yyyymmddhhmmss") & fileExt

        ' Save binary image data
        Dim fileNum As Integer
        fileNum = FreeFile
        Open tempFile For Binary As #fileNum
        Put #fileNum, , binaryData
        Close #fileNum

        ' Load into StdPicture (non-animated even if GIF)
        Dim icon As StdPicture
        Set icon = LoadPicture(tempFile)
        
        ' Update UI
        GuildView1.UpdateGuildIcon guildIndex + 1, icon

        ' Clean up temp
        Kill tempFile
    End If

    Exit Sub
EH:
    Debug.Print "Error processing icon: " & Err.Description
End Sub
Private Sub cmdSendMsg_Click()
    ' Validate inputs
    If Len(Trim(txtToken.Text)) = 0 Then
        MsgBox "Please enter your Discord user token", vbExclamation
        Exit Sub
    End If
    
    If Len(Trim(txtCID.Text)) = 0 Then
        MsgBox "Please enter a channel ID", vbExclamation
        Exit Sub
    End If
    
    If Len(Trim(txtMsg.Text)) = 0 Then
        MsgBox "Please enter a message to send", vbExclamation
        Exit Sub
    End If
    
    m_sToken = txtToken.Text
    
    ' Save settings
    SaveSetting "DiscordClient", "Settings", "Token", txtToken.Text
    SaveSetting "DiscordClient", "Settings", "ChannelId", txtCID.Text
    

    
    ' Send the message
    SendDiscordMessage txtCID.Text, txtMsg.Text
    
    ' Clear message textbox
    txtMsg.Text = ""
End Sub


Private Function EscapeJsonString(ByVal sText As String) As String
    ' Basic JSON string escaping
    Dim sEscaped As String
    
    sEscaped = Replace(sText, "\", "\\")
    sEscaped = Replace(sEscaped, """", "\""")
    sEscaped = Replace(sEscaped, vbCrLf, "\n")
    sEscaped = Replace(sEscaped, vbTab, "\t")
    
    EscapeJsonString = sEscaped
End Function




Public Function preg_replace(find_re As String, sText As String, Optional sReplace As String) As String
    preg_replace = pvInitRegExp(find_re).Replace(sText, sReplace)
End Function

Private Function pvInitRegExp(sPattern As String) As Object
    Dim lIdx            As Long

    Set pvInitRegExp = CreateObject("VBScript.RegExp")
    With pvInitRegExp
        .Global = True
        If Left$(sPattern, 1) = "/" Then
            lIdx = InStrRev(sPattern, "/")
            .Pattern = Mid$(sPattern, 2, lIdx - 2)
            .IgnoreCase = (InStr(lIdx, sPattern, "i") > 0)
            .MultiLine = (InStr(lIdx, sPattern, "m") > 0)
        Else
            .Pattern = sPattern
        End If
    End With
End Function



