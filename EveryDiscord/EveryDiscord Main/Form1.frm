VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "EveryDiscord"
   ClientHeight    =   8055
   ClientLeft      =   105
   ClientTop       =   555
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
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
      Height          =   6105
      Left            =   1320
      TabIndex        =   6
      Top             =   360
      Width           =   2175
   End
   Begin VB.ListBox lstGuild 
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
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1335
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
      Left            =   240
      TabIndex        =   3
      Text            =   "1355226543160557778"
      Top             =   6600
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
      Left            =   240
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
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   7560
      Width           =   9015
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   9240
      Top             =   240
   End
   Begin VB.CommandButton cmdSendMsg 
      Caption         =   "Send"
      Height          =   375
      Left            =   9360
      TabIndex        =   0
      Top             =   7560
      Width           =   735
   End
   Begin MSWinsockLib.Winsock wscSocket 
      Left            =   7965
      Top             =   1125
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstMessages 
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
      Left            =   3480
      TabIndex        =   4
      Top             =   0
      Width           =   6735
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

' TLS Context and Socket variables

Private m_GuildIds() As String
Private m_ChannelIds() As String
Private m_uCtx              As UcsTlsContext
Private m_sRequest          As String
Private m_sToken            As String
Private m_sBaseUrl          As String

' Add this as a module or class-level variable
Private m_sResponseBuffer As String
Private m_bReceivingData As Boolean
Private m_lContentLength As Long
' WebSocket constants for Gateway connection
Private Const DISCORD_GATEWAY_URL As String = "gateway.discord.gg"
Private Const DISCORD_GATEWAY_VERSION As String = "10"
Private Const DISCORD_GATEWAY_PORT As Long = 443

' Gateway OpCodes
Private Const GATEWAY_OP_DISPATCH As Long = 0        ' Events from Discord
Private Const GATEWAY_OP_HEARTBEAT As Long = 1       ' Keep connection alive
Private Const GATEWAY_OP_IDENTIFY As Long = 2        ' Authenticate connection
Private Const GATEWAY_OP_RESUME As Long = 6          ' Resume disconnected session
Private Const GATEWAY_OP_RECONNECT As Long = 7       ' Server asks client to reconnect
Private Const GATEWAY_OP_HELLO As Long = 10          ' Initial handshake
Private Const GATEWAY_OP_HEARTBEAT_ACK As Long = 11  ' Server ack of heartbeat

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
Function ParseHttpResponse(rawData As String) As String
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
    
    ' Remove last 3 characters if string is long enough
    If Len(responseBody) >= 3 Then
        ParseHttpResponse = Left$(responseBody, Len(responseBody) - 3)
    Else
        ParseHttpResponse = responseBody
    End If
End Function

Private Sub Connect(ByVal sServer As String, ByVal lPort As Long)
    Call TlsInitClient(m_uCtx, sServer)
    wscSocket.Close
    wscSocket.Connect sServer, lPort
End Sub

Private Sub SendData(baData() As Byte)
    Dim baOutput()          As Byte
    Dim lOutputPos          As Long
    
    If Not TlsSend(m_uCtx, baData, UBound(baData) + 1, baOutput, lOutputPos) Then
        OnError TlsGetLastError(m_uCtx), "TlsSend"
    End If
    If lOutputPos > 0 Then
        Debug.Assert UBound(baOutput) + 1 = lOutputPos
        wscSocket.SendData baOutput
    End If
End Sub

Private Sub Timer1_Timer()

    FetchChannelMessages txtCID.Text
    FetchUserGuilds
End Sub

Private Sub Timer2_Timer()
 If Timer1.Enabled = False Then
 Shape1.BackColor = vbRed
 Else
 Shape1.BackColor = vbGreen
 End If
End Sub

Private Sub wscSocket_Connect()
    Dim baEmpty()           As Byte
    Dim baOutput()          As Byte
    Dim lOutputPos          As Long
    
    On Error GoTo EH
    If Not TlsHandshake(m_uCtx, baEmpty, -1, baOutput, lOutputPos) Then
        OnError TlsGetLastError(m_uCtx), "TlsHandshake"
    End If
    If lOutputPos > 0 Then
        Debug.Assert UBound(baOutput) + 1 = lOutputPos
        wscSocket.SendData baOutput
    End If
    Exit Sub
EH:
    OnError Err.Description, "wscSocket_Connect"
End Sub

Private Sub wscSocket_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
    Dim bError              As Boolean
    Dim baRecv()            As Byte
    Dim baOutput()          As Byte
    Dim lOutputPos          As Long
    Dim baPlainText()       As Byte
    Dim lSize               As Long
    
    If TlsIsClosed(m_uCtx) Or bytesTotal = 0 Then
        Exit Sub
    End If
    wscSocket.GetData baRecv
    If Not TlsIsReady(m_uCtx) Then
        lOutputPos = 0
        bError = Not TlsHandshake(m_uCtx, baRecv, -1, baOutput, lOutputPos)
        If lOutputPos > 0 Then
            Debug.Assert UBound(baOutput) + 1 = lOutputPos
            wscSocket.SendData baOutput
        End If
        If bError Then
            OnError TlsGetLastError(m_uCtx), "TlsHandshake"
        End If
        If Not TlsIsReady(m_uCtx) Then
            Exit Sub
        End If
        OnConnect
        '--- fall-through to flush application data after TLS handshake (if any)
        Erase baRecv
    End If
    lOutputPos = 0
    bError = Not TlsReceive(m_uCtx, baRecv, -1, baPlainText, lSize, baOutput, lOutputPos)
    If lOutputPos > 0 Then
        wscSocket.SendData baOutput
    End If
    If lSize > 0 Then
        OnDataArrival lSize, baPlainText
    End If
    If bError Then
        OnError TlsGetLastError(m_uCtx), "TlsReceive"
    End If
    If TlsIsClosed(m_uCtx) Then
        OnClose
    End If
    Exit Sub
End Sub

Private Sub wscSocket_Close()
    If Not TlsIsClosed(m_uCtx) Then
        OnClose
    End If
End Sub

'= Gateway Functions ====================================================


'= callbacks =============================================================

Private Sub OnConnect()
    SendData StrConv(m_sRequest, vbFromUnicode)
End Sub


Private Sub OnDataArrival(ByVal bytesTotal As Long, baData() As Byte)
    Debug.Print "OnDataArrival, bytesTotal=" & bytesTotal, Timer
    
    ' Process response
    Dim sResponse As String
    sResponse = StrConv(baData, vbUnicode)
    ' Check if this is the start of a response or continuation
    If Not m_bReceivingData Then
        m_bReceivingData = True
        m_sResponseBuffer = sResponse
        
        ' Extract Content-Length if available
        Dim clPos As Long
        clPos = InStr(1, sResponse, "Content-Length:", vbTextCompare)
        If clPos > 0 Then
            Dim clEnd As Long
            clEnd = InStr(clPos, sResponse, vbCrLf)
            If clEnd > 0 Then
                m_lContentLength = Val(Mid$(sResponse, clPos + 15, clEnd - (clPos + 15)))
            End If
        End If
    Else
        ' Append to existing buffer
        m_sResponseBuffer = m_sResponseBuffer & sResponse
    End If
    
    ' Check if we have the complete response
    Dim headersEnd As Long
    Dim contentReceived As Long
    
    headersEnd = InStr(m_sResponseBuffer, vbCrLf & vbCrLf)
    If headersEnd > 0 Then
        contentReceived = Len(m_sResponseBuffer) - (headersEnd + 3)
        
        ' For chunked encoding, check for final chunk
        If InStr(1, m_sResponseBuffer, "Transfer-Encoding: chunked", vbTextCompare) > 0 Then
            If InStr(m_sResponseBuffer, vbCrLf & "0" & vbCrLf & vbCrLf) > 0 Then
                ProcessCompleteResponse
            End If
        ' For Content-Length, check if we've received enough data
        ElseIf m_lContentLength > 0 Then
            If contentReceived >= m_lContentLength Then
                ProcessCompleteResponse
            End If
        ' If we can't determine length, assume this is all we'll get
        Else
            ProcessCompleteResponse
        End If
    End If
End Sub



Private Sub FetchUserGuilds()
    On Error GoTo EH
    
    ' Create the API request path
    Dim sPath As String
    sPath = "api/v10/users/@me/guilds"
    
    ' Prepare the HTTP request
    m_sRequest = "GET /" & sPath & " HTTP/1.1" & vbCrLf & _
              "Host: " & m_sBaseUrl & vbCrLf & _
              "Authorization: " & m_sToken & vbCrLf & _
              "Connection: close" & vbCrLf & vbCrLf
    
    ' Connect to Discord API
    Connect m_sBaseUrl, 443
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
    m_sRequest = "GET /" & sPath & " HTTP/1.1" & vbCrLf & _
              "Host: " & m_sBaseUrl & vbCrLf & _
              "Authorization: " & m_sToken & vbCrLf & _
              "Connection: close" & vbCrLf & vbCrLf
    
    ' Connect to Discord API
    Connect m_sBaseUrl, 443
    Exit Sub
EH:
    MsgBox "Error fetching channels: " & Err.Description, vbCritical
End Sub
' Add these declarations at module level

Private Sub ProcessGuildsResponse(aJson As String)
    Dim parsed As ParseResult
    Dim i As Long
    Dim guildCount As Long
    Dim sjson As String
    
    sjson = Left$(aJson, Len(aJson) - 5)
    ' Parse the JSON array
    parsed = Parse(sjson)
 
    If Not parsed.IsValid Then
        Exit Sub
    End If
    
    ' Clear existing guilds
    lstGuild.Clear
    
    ' Count guilds first to properly size the array
    guildCount = 0
    For i = 0 To 60
        On Error Resume Next
        Dim guildCheck As Object
        Set guildCheck = parsed.Value(i)
        If Not guildCheck Is Nothing Then guildCount = guildCount + 1
        On Error GoTo 0
    Next i
    
    ' Resize the array to match guild count
    ReDim m_GuildIds(0 To guildCount - 1) As String
    
    ' Process each guild
    Dim validGuilds As Long
    validGuilds = 0
    
    For i = 0 To 60
        On Error Resume Next
        Dim guild As Object
        Set guild = parsed.Value(i)
        
        ' Skip if no more guilds or error
        
        ' Extract guild details
        Dim sGuildName As String
        Dim sGuildId As String
        
        sGuildName = guild("name")
        sGuildId = guild("id")
        
        ' Add to listbox
        lstGuild.AddItem sGuildName
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
    'MsgBox "hi"
    sjson = Left$(aJson, Len(aJson) - 5)
    ' Parse the JSON array
    parsed = Parse(sjson)
 
    If Not parsed.IsValid Then
    MsgBox "Oops.. Discord or EveryDiscord gone wrong! " & parsed.Error
        Exit Sub
    End If
    
    ' Clear existing channels
    lstChannel.Clear
    
    ' Count text channels first
    channelCount = 0
    For i = 0 To 60
        On Error Resume Next
        Dim channelCheck As Object
        Set channelCheck = parsed.Value(i)
        If Not channelCheck Is Nothing Then
            If channelCheck("type") = 0 Then ' Only count text channels
                channelCount = channelCount + 1
            End If
        End If
        On Error GoTo 0
    Next i
    
    ' Resize the array to match channel count
    ReDim m_ChannelIds(0 To channelCount - 1) As String
    
    ' Process each channel
    Dim validChannels As Long
    validChannels = 0
    
    For i = 0 To 15
        On Error Resume Next
        Dim channel As Object
        Set channel = parsed.Value(i)
        
        ' Skip if no more channels
        
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
            ' Store ID in parallel array
            m_ChannelIds(validChannels) = sChannelId
            validChannels = validChannels + 1
            
            ' Debug output
            Debug.Print "Added channel: " & sChannelName & " with ID: " & sChannelId
        End If
    Next i
End Sub

Private Sub lstGuild_Click()
    ' Get the selected guild ID
    Dim selectedIndex As Long
    Dim guildId As String
    
    selectedIndex = lstGuild.ListIndex
    
    If selectedIndex >= 0 And selectedIndex < UBound(m_GuildIds) + 1 Then
        ' Get the ID from our parallel array
        guildId = m_GuildIds(selectedIndex)
        
        ' Fetch channels for this guild
        FetchGuildChannels guildId
    End If
End Sub

Private Sub lstChannel_Click()
    ' Get the selected channel ID
    Dim selectedIndex As Long
    Dim channelId As String
    
    selectedIndex = lstChannel.ListIndex
    
    If selectedIndex >= 0 And selectedIndex < UBound(m_ChannelIds) + 1 Then
        ' Get the ID from our parallel array
        channelId = m_ChannelIds(selectedIndex)
        Debug.Print "Selected channel ID: " & channelId
        
        ' Set the channel ID in the textbox
        txtCID.Text = channelId
        
        ' Fetch messages for this channel
        FetchChannelMessages channelId
    End If
End Sub

' Update the OnDataArrival processing to handle channel responses
Private Sub ProcessCompleteResponse()
    ' Parse HTTP response
    Dim sContent As String
    sContent = ParseHttpResponse(m_sResponseBuffer)

    ' For message fetch responses, remove the first line if it exists
    If InStr(sContent, vbCrLf) > 0 Then
        sContent = Mid$(sContent, InStr(sContent, vbCrLf) + 2)
    End If
    
    ' Determine what type of response this is and process appropriately
    If InStr(m_sRequest, "api/v10/channels/") > 0 And InStr(m_sRequest, "/messages") > 0 Then
        ' This is a channel messages response
        ProcessMessagesResponse sContent
    ElseIf InStr(m_sRequest, "api/v10/users/@me/guilds") > 0 Then
        ' This is a guilds response
        ProcessGuildsResponse sContent
    ElseIf InStr(m_sRequest, "api/v10/guilds/") > 0 And InStr(m_sRequest, "/channels") > 0 Then
        ' This is a channels response
        ProcessChannelsResponse sContent
    End If
    
    ' Reset buffer and flags
    m_sResponseBuffer = ""
    m_bReceivingData = False
    m_lContentLength = 0
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

    End If
    
    ' Check if we got an array
    
    ' Clear existing messages
    lstMessages.Clear
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
        lstMessages.AddItem sOutput
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
    
    ' Clear result if available
    If Not txtResult Is Nothing Then
        txtResult.Text = vbNullString
    End If
    
    ' Fetch messages
    FetchChannelMessages txtCID.Text
End Sub

Private Sub FetchChannelMessages(ByVal sChannelId As String, Optional ByVal lLimit As Long = 20)
    On Error GoTo EH
    
    ' Create the API request path
    Dim sPath As String
    sPath = "api/v10/channels/" & sChannelId & "/messages?limit=" & lLimit
    
    ' Prepare the HTTP request
    m_sRequest = "GET /" & sPath & " HTTP/1.1" & vbCrLf & _
              "Host: " & m_sBaseUrl & vbCrLf & _
              "Authorization: " & m_sToken & vbCrLf & _
              "Connection: close" & vbCrLf & vbCrLf
    
    ' Connect to Discord API
    Connect m_sBaseUrl, 443
    Exit Sub
EH:
    MsgBox "Error fetching messages: " & Err.Description, vbCritical
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

Private Sub OnClose()
    Debug.Print "OnClose", Timer
End Sub

Private Sub OnError(sDescription As String, sSource As String)
    Debug.Print "Critical error: " & sDescription & " in " & sSource, Timer
    If Not txtResult Is Nothing Then
        pvAppendLogText txtResult, "Critical error: " & sDescription & " in " & sSource & vbCrLf & vbCrLf
    Else
        MsgBox "Discord API Error: " & sDescription & " in " & sSource, vbCritical
    End If
End Sub

'= form events ===========================================================

Private Sub Form_Load()
    ' Initialize the form
    m_sBaseUrl = "discord.com"
    
    ' Initialize arrays
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
        FetchChannelMessages txtCID.Text
        FetchUserGuilds
    End If
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

Private Sub SendDiscordMessage(ByVal sChannelId As String, ByVal sMessage As String)
    On Error GoTo EH
    
    ' Create the JSON payload
    Dim sPayload As String
    sPayload = "{""content"":""" & EscapeJsonString(sMessage) & """}"
    
    ' Create the API request path
    Dim sPath As String
    sPath = "api/v10/channels/" & sChannelId & "/messages"
    
    ' Prepare the HTTP request
    m_sRequest = "POST /" & sPath & " HTTP/1.1" & vbCrLf & _
              "Host: " & m_sBaseUrl & vbCrLf & _
              "Authorization: " & m_sToken & vbCrLf & _
              "Content-Type: application/json" & vbCrLf & _
              "Content-Length: " & Len(sPayload) & vbCrLf & _
              "Connection: close" & vbCrLf & vbCrLf & _
              sPayload
    
    ' Connect to Discord API
    Connect m_sBaseUrl, 443
    Exit Sub
EH:
    MsgBox "Error sending message: " & Err.Description, vbCritical
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

'= helpers ===============================================================

Private Sub pvAppendLogText(txtLog As TextBox, sValue As String)
    Const WM_SETREDRAW              As Long = &HB
    Const EM_SETSEL                 As Long = &HB1
    Const EM_REPLACESEL             As Long = &HC2
    Const WM_VSCROLL                As Long = &H115
    Const SB_BOTTOM                 As Long = 7
    Call SendMessage(txtLog.hWnd, WM_SETREDRAW, 0, ByVal 0)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, 0, ByVal -1)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, -1, ByVal -1)
    Call SendMessage(txtLog.hWnd, EM_REPLACESEL, 1, ByVal StrPtr(preg_replace("\r*\n", sValue, vbCrLf)))
    Call SendMessage(txtLog.hWnd, EM_SETSEL, 0, ByVal -1)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, -1, ByVal -1)
    Call SendMessage(txtLog.hWnd, WM_SETREDRAW, 1, ByVal 0)
    Call SendMessage(txtLog.hWnd, WM_VSCROLL, SB_BOTTOM, ByVal 0)
End Sub

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

