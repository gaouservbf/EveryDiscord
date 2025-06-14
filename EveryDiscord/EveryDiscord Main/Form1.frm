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
   Begin VB.CommandButton Emojiz 
      Height          =   435
      Left            =   9765
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7560
      Width           =   435
   End
   Begin VB.Timer tmrHeartbeat 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4515
      Top             =   3990
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   1320
      ScaleHeight     =   495
      ScaleWidth      =   8790
      TabIndex        =   11
      Top             =   0
      Width           =   8790
      Begin VB.Label lblChannel 
         BackStyle       =   0  'Transparent
         Caption         =   "<Channel Name>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   2520
         TabIndex        =   13
         Top             =   210
         Width           =   6270
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "<Guild Name>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   0
         TabIndex        =   12
         Top             =   210
         Width           =   2175
      End
   End
   Begin VB.CommandButton btnUpload 
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
      TabIndex        =   9
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
      Height          =   1080
      Left            =   0
      ScaleHeight     =   1080
      ScaleWidth      =   2775
      TabIndex        =   3
      Top             =   6975
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
         Left            =   1560
         TabIndex        =   10
         Top             =   630
         Width           =   855
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   525
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   8
         Top             =   210
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
         TabIndex        =   5
         Top             =   420
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
         TabIndex        =   4
         Top             =   210
         Width           =   1575
      End
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
      TabIndex        =   2
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
      TabIndex        =   1
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
      Left            =   2985
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   7560
      Width           =   6795
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
   Begin EveryDiscord.ChatView ChatView1 
      Height          =   7020
      Left            =   3480
      TabIndex        =   6
      Top             =   480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   12383
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
   Begin EveryDiscord.GuildView GuildView1 
      Height          =   6975
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   12303
   End
   Begin MSWinsockLib.Winsock wsMessage 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   7995
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
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
   Begin ComctlLib.TreeView lstChannel 
      Height          =   6495
      Left            =   1320
      TabIndex        =   16
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   11456
      _Version        =   327682
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin EveryDiscord.Websocket wsGateway 
      Height          =   465
      Left            =   5355
      Top             =   4410
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   820
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

DefObj A-Z
Private Const MODULE_NAME = "Form1"

#Const ImplUseDebugLog = (USE_DEBUG_LOG <> 0)

Private Const DISCORD_GATEWAY_URL As String = "wss://gateway.discord.gg/?v=10&encoding=json"
Private Const DISCORD_GATEWAY_VERSION As String = "10"
Private Const DISCORD_GATEWAY_PORT As Long = 443

Private WithEvents m_oHttpDownload As cHttpDownload
Attribute m_oHttpDownload.VB_VarHelpID = -1
Private m_lFreeSocketIndex As Long     ' Next available socket index

Private m_bGatewayReady As Boolean
' Control array constants
Private Const MAX_CONNECTIONS As Long = 100
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
    guildId As String
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
Private Sub btnUpload_Click()
    Dim commctrl As New CommonDialog
    Dim tlsSocket As New cTlsSocket
    Dim sRequest As String
    Dim sResponse As String
    Dim sFilePath As String
    Dim sFileName As String
    Dim fileBytes() As Byte
    Dim sBoundary As String
    Dim sUserHash As String
    Dim iFile As Integer
    Dim sFileUrl As String
    Dim sBody As String
    Dim iPos As Integer
    
    ' Show file dialog
    commctrl.ShowOpen
    sFilePath = commctrl.FileName
    
    If sFilePath = "" Then Exit Sub
    
    ' Extract filename from path
    sFileName = Right(sFilePath, Len(sFilePath) - InStrRev(sFilePath, "\"))
    
    ' Read file as byte array (don't convert to string for binary files)
    iFile = FreeFile
    Open sFilePath For Binary As #iFile
    ReDim fileBytes(0 To LOF(iFile) - 1)
    Get #iFile, , fileBytes
    Close #iFile
    
    ' Set your userhash here (optional)
    sUserHash = ""
    
    ' Generate boundary
    sBoundary = "----FormBoundary" & Format(Now, "yyyymmddhhmmss")
    
    ' Build request using byte arrays for binary files
    Dim requestData() As Byte
    requestData = BuildBinaryMultipartRequest(fileBytes, sFileName, sBoundary, sUserHash)
    
    ' Debug info
    Debug.Print "File size: " & UBound(fileBytes) + 1 & " bytes"
    Debug.Print "Total request size: " & UBound(requestData) + 1 & " bytes"
    
    ' Connect to server
    If Not tlsSocket.SyncConnect("catbox.moe", 443, , , ucsTlsSupportAll Or ucsTlsIgnoreServerCertificateErrors) Then
        MsgBox "Failed to connect to catbox.moe"
        Exit Sub
    End If
    
    ' Send the complete request as binary data
    If Not tlsSocket.SyncSendArray(requestData) Then
        MsgBox "Failed to send request"
        tlsSocket.Close_
        Exit Sub
    End If
    
    ' Receive response
    sResponse = ""
    Dim iAttempts As Integer
    For iAttempts = 1 To 20
        Dim sChunk As String
        sChunk = tlsSocket.SyncReceiveText(10000)
        If sChunk = "" Then Exit For
        sResponse = sResponse & sChunk
        DoEvents
        If InStr(sResponse, "https://files.catbox.moe/") > 0 Then Exit For
    Next iAttempts
    
    ' Parse response
    iPos = InStr(sResponse, vbCrLf & vbCrLf)
    If iPos > 0 Then
        sBody = Mid(sResponse, iPos + 4)
        
        ' Handle chunked encoding if present
        If InStr(sResponse, "Transfer-Encoding: chunked") > 0 Then
            sBody = ParseChunkedResponse(sBody)
        End If
        
        ' Extract file URL
        iPos = InStr(sBody, "https://files.catbox.moe/")
        If iPos > 0 Then
            Dim iEndPos As Integer
            iEndPos = InStr(iPos, sBody, vbCrLf)
            If iEndPos = 0 Then iEndPos = InStr(iPos, sBody, vbLf)
            If iEndPos = 0 Then iEndPos = Len(sBody) + 1
            
            sFileUrl = Mid(sBody, iPos, iEndPos - iPos)
            sFileUrl = Trim(sFileUrl)
        End If
    End If
    
    If sFileUrl <> "" Then
        SendDiscordMessage txtCID.Text, sFileUrl
        Clipboard.Clear
        Clipboard.SetText sFileUrl
    Else
        MsgBox "Upload failed: " & Left(sResponse, 500)
    End If
    
    tlsSocket.Close_
    Set tlsSocket = Nothing
End Sub

' Function to build multipart request as byte array for binary files
Private Function BuildBinaryMultipartRequest(fileData() As Byte, FileName As String, boundary As String, userHash As String) As Byte()
    Dim Result() As Byte
    Dim tempStr As String
    Dim tempBytes() As Byte
    Dim Pos As Long
    Dim i As Long
    
    ' Calculate total size needed
    Dim headerSize As Long
    Dim footerSize As Long
    
    ' Build header as string first
    tempStr = "POST /user/api.php HTTP/1.1" & vbCrLf & _
              "Host: catbox.moe" & vbCrLf & _
              "User-Agent: VB6-Uploader/1.0" & vbCrLf & _
              "Content-Type: multipart/form-data; boundary=" & boundary & vbCrLf
    
    ' Build form data header
    Dim formHeader As String
    formHeader = "--" & boundary & vbCrLf & _
                 "Content-Disposition: form-data; name=""reqtype""" & vbCrLf & vbCrLf & _
                 "fileupload" & vbCrLf
    
    If userHash <> "" Then
        formHeader = formHeader & "--" & boundary & vbCrLf & _
                     "Content-Disposition: form-data; name=""userhash""" & vbCrLf & vbCrLf & _
                     userHash & vbCrLf
    End If
    
    formHeader = formHeader & "--" & boundary & vbCrLf & _
                 "Content-Disposition: form-data; name=""fileToUpload""; filename=""" & FileName & """" & vbCrLf & _
                 "Content-Type: " & GetContentType(FileName) & vbCrLf & vbCrLf
    
    Dim formFooter As String
    formFooter = vbCrLf & "--" & boundary & "--" & vbCrLf
    
    ' Calculate content length
    Dim ContentLength As Long
    ContentLength = Len(formHeader) + UBound(fileData) + 1 + Len(formFooter)
    
    ' Complete the HTTP header
    tempStr = tempStr & "Content-Length: " & ContentLength & vbCrLf & vbCrLf
    
    ' Now build the complete request
    Dim totalSize As Long
    totalSize = Len(tempStr) + ContentLength
    ReDim Result(0 To totalSize - 1)
    
    Pos = 0
    
    ' Add HTTP headers
    tempBytes = StrConv(tempStr, vbFromUnicode)
    For i = 0 To UBound(tempBytes)
        Result(Pos) = tempBytes(i)
        Pos = Pos + 1
    Next i
    
    ' Add form header
    tempBytes = StrConv(formHeader, vbFromUnicode)
    For i = 0 To UBound(tempBytes)
        Result(Pos) = tempBytes(i)
        Pos = Pos + 1
    Next i
    
    ' Add file data (binary)
    For i = 0 To UBound(fileData)
        Result(Pos) = fileData(i)
        Pos = Pos + 1
    Next i
    
    ' Add form footer
    tempBytes = StrConv(formFooter, vbFromUnicode)
    For i = 0 To UBound(tempBytes)
        Result(Pos) = tempBytes(i)
        Pos = Pos + 1
    Next i
    
    BuildBinaryMultipartRequest = Result
End Function

' Helper function to get content type
Private Function GetContentType(FileName As String) As String
    Dim ext As String
    ext = LCase(Right(FileName, Len(FileName) - InStrRev(FileName, ".")))
    
    Select Case ext
        Case "gif"
            GetContentType = "image/gif"
        Case "jpg", "jpeg"
            GetContentType = "image/jpeg"
        Case "png"
            GetContentType = "image/png"
        Case "bmp"
            GetContentType = "image/bmp"
        Case "webp"
            GetContentType = "image/webp"
        Case "mp4"
            GetContentType = "video/mp4"
        Case "webm"
            GetContentType = "video/webm"
        Case "pdf"
            GetContentType = "application/pdf"
        Case "txt"
            GetContentType = "text/plain"
        Case "zip"
            GetContentType = "application/zip"
        Case Else
            GetContentType = "application/octet-stream"
    End Select
End Function

' Helper function to parse chunked response
Private Function ParseChunkedResponse(chunkedData As String) As String
    Dim Result As String
    Dim Pos As Integer
    Dim chunkSizeHex As String
    Dim ChunkSize As Long
    
    Result = ""
    Pos = 1
    
    Do
        Dim crlfPos As Integer
        crlfPos = InStr(Pos, chunkedData, vbCrLf)
        If crlfPos = 0 Then Exit Do
        
        chunkSizeHex = Trim(Mid(chunkedData, Pos, crlfPos - Pos))
        If chunkSizeHex = "" Then Exit Do
        
        ' Handle chunk size (ignore chunk extensions)
        Dim semicolonPos As Integer
        semicolonPos = InStr(chunkSizeHex, ";")
        If semicolonPos > 0 Then
            chunkSizeHex = Left(chunkSizeHex, semicolonPos - 1)
        End If
        
        On Error GoTo ChunkError
        ChunkSize = CLng("&H" & chunkSizeHex)
        On Error GoTo 0
        
        If ChunkSize = 0 Then Exit Do
        
        Pos = crlfPos + 2
        If Pos + ChunkSize - 1 <= Len(chunkedData) Then
            Result = Result & Mid(chunkedData, Pos, ChunkSize)
        End If
        
        Pos = Pos + ChunkSize + 2
    Loop
    
    ParseChunkedResponse = Result
    Exit Function
    
ChunkError:
    ParseChunkedResponse = chunkedData
End Function

Private Sub Command2_Click()
Form3.Show
End Sub


Private Sub Emojiz_Click()
Form4.Show
End Sub

Private Sub Form_Resize()
On Error Resume Next
    ChatView1.Width = Me.ScaleWidth - GuildView1.Width - lstChannel.Width
    GuildView1.Height = Me.ScaleHeight - Picture1.Height - StatusBar1.Height
      lstChannel.Height = Me.ScaleHeight - Picture1.Height - StatusBar1.Height - lstChannel.Top
    Picture3.Width = Me.ScaleWidth - Picture3.Left
    Picture1.Top = Me.ScaleHeight - Picture1.Height - StatusBar1.Height
    Emojiz.Top = Me.ScaleHeight - 810
    txtMsg.Top = Me.ScaleHeight - 810
    btnUpload.Top = Me.ScaleHeight - 810
    ChatView1.Height = Me.ScaleHeight - 1350
    Emojiz.Left = Me.ScaleWidth - Emojiz.Width
    txtMsg.Width = Me.ScaleWidth - txtMsg.Left - Emojiz.Width
End Sub

Private Sub GuildView1_GuildSelected(ByVal Index As Long)
If Index = 0 Then
lstChannel.Nodes.Clear
FetchUserDMs
Else
DoEvents
    Dim SelectedIndex As Long
    Dim guildId As String
    
    SelectedIndex = Index - 1
    
    If SelectedIndex >= 0 And SelectedIndex < UBound(m_GuildIds) + 1 Then
        ' Get the ID from our parallel array
        guildId = m_GuildIds(SelectedIndex)
        
        ' Fetch channels for this guild
DoEvents
        FetchGuildChannels guildId
DoEvents
        Label3.Caption = "<" & GuildView1.GetGuildName(Index) & ">"
        Me.Caption = "EveryDiscord - " & GuildView1.GetGuildName(Index)
    End If
    End If
End Sub


Private Sub Timer1_Timer()
   ' FetchUserGuilds
End Sub

Private Sub Timer2_Timer()
    'FetchChannelMessages txtCID.Text
End Sub



Private Sub txtMsg_Change()
    If Len(txtMsg.Text) >= 2 Then
        If Right$(txtMsg.Text, 2) = vbCrLf Then
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
        End If
    End If
End Sub

Function ParseChunkSize(rawData As String, currentPosition As Long) As Long
    ' Finds the chunk size indicator at the current position
    ' Returns the size in bytes, or -1 if it's the final chunk
    
    Dim chunkSizeEnd As Long
    Dim chunkSizeHex As String
    
    ' Find the end of the chunk size line (CRLF)
    chunkSizeEnd = InStr(currentPosition, rawData, vbCrLf)
    
    If chunkSizeEnd = 0 Then
        ' Invalid chunk format
        ParseChunkSize = -2
        Exit Function
    End If
    
    ' Extract the chunk size hex string
    chunkSizeHex = Mid$(rawData, currentPosition, chunkSizeEnd - currentPosition)
    
    ' Check for final chunk marker
    If chunkSizeHex = "0" Then
        ParseChunkSize = -1
        Exit Function
    End If
    
    ' Convert hex to decimal
    On Error Resume Next
    ParseChunkSize = Val("&h" & chunkSizeHex)
    If Err.Number <> 0 Then
        ParseChunkSize = -2 ' Invalid hex format
    End If
    On Error GoTo 0
End Function

Function MergeChunkedBody(rawData As String) As String
    ' Processes a chunked transfer-encoded response and merges it into a single body
    Dim mergedBody As String
    Dim currentPos As Long
    Dim ChunkSize As Long
    Dim chunkData As String
    
    currentPos = 1
    mergedBody = ""
    
    Do
        ' Get the size of the next chunk
        ChunkSize = ParseChunkSize(rawData, currentPos)
        
        If ChunkSize = -2 Then ' Invalid format
            Exit Do
        ElseIf ChunkSize = -1 Then ' Final chunk
            Exit Do
        ElseIf ChunkSize = 0 Then
            Exit Do
        End If
        
        ' Move to the start of the chunk data (after size + CRLF)
        currentPos = InStr(currentPos, rawData, vbCrLf) + 2
        
        ' Extract the chunk data
        If currentPos + ChunkSize - 1 > Len(rawData) Then
            ' Incomplete chunk
            Exit Do
        End If
        
        chunkData = Mid$(rawData, currentPos, ChunkSize)
        mergedBody = mergedBody & chunkData
        
        ' Move to next chunk (after data + CRLF)
        currentPos = currentPos + ChunkSize + 2
        
        ' Verify we have CRLF after chunk
        If Mid$(rawData, currentPos - 2, 2) <> vbCrLf Then
            ' Invalid chunk format
            Exit Do
        End If
    Loop
    
    MergeChunkedBody = mergedBody
End Function
Function ParseHttpResponse(rawData As String, bFetchingIcon As Boolean) As String
    ' For binary data (icons), don't apply text processing
    If bFetchingIcon Then
        ParseHttpResponse = rawData
        Exit Function
    End If
    
    ' Check if chunked encoding is used
    ' Default: Extract body after headers
    Dim headersEnd As Long
    Dim ResponseBody As String
    
    headersEnd = InStr(rawData, vbCrLf & vbCrLf)
    If headersEnd > 0 Then
        ResponseBody = Mid$(rawData, headersEnd + 4)
    Else
        ResponseBody = rawData
    End If
    
    ' Only for text responses, remove last characters if string is long enough
    If Len(ResponseBody) >= 3 Then
        ParseHttpResponse = Left$(ResponseBody, Len(ResponseBody) - 3)
    Else
        ParseHttpResponse = ResponseBody
    End If
End Function


Private Sub OnDataArrival(ByVal socketIndex As Long, ByVal BytesTotal As Long, baData() As Byte)
    Debug.Print "OnDataArrival, Socket=" & socketIndex & ", bytesTotal=" & BytesTotal, Timer
    
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
DoEvents
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
            
        Case "DMs"
            sContent = ParseHttpResponse(sResponse, bFetchingIcon)
            If InStr(sContent, vbCrLf) > 0 Then
                sContent = Mid$(sContent, InStr(sContent, vbCrLf) + 2)
            End If
            ProcessDMsResponse sContent
            
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

Private Sub FetchGuildIcon(ByVal sGuildId As String, ByVal sIconHash As String, ByVal guildIndex As Long)
   
End Sub
Private Sub ProcessGuildsResponse(aJson As String)
    Dim parsed As ParseResult
    Dim i As Long
    Dim GuildCount As Long
    Dim sjson As String
    Dim emojiList As String
    
    DoEvents
    'sjson = Left$(aJson, Len(aJson) - 5)
    sjson = aJson
    ' Parse the JSON array
    parsed = Parse(sjson)
 
    If parsed.IsValid = False Then
    MsgBox "naw" + " " + parsed.Error
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
    DoEvents
    ' Resize the array to match guild count

    ReDim m_GuildIds(0 To GuildCount - 1) As String
    
    DoEvents
    ' Process each guild
    Dim validGuilds As Long
    validGuilds = 0
    
    DoEvents
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
        Set guildIcon = LoadPicture() ' Default empty icon
        
    DoEvents
        ' If icon hash exists, fetch it
        If Len(sIconHash) > 0 Then ' Replace QueueGuildIconFetch calls with:
        
    DoEvents
FetchGuildIcon sGuildId, sIconHash, i
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
    On Error Resume Next
    Dim parsed As ParseResult
    Dim i As Long
    Dim sjson As String
    Dim categoryNode As Node
    Dim channelNode As Node
    Dim currentCategory As String
    Dim channelCount As Long
    
    sjson = aJson
    
    ' Parse the JSON array
    parsed = Parse(sjson)
    If parsed.IsValid = False Then
        MsgBox "Channel error!: " + parsed.Error
        Exit Sub
    End If
    
    ' Clear existing channels
    lstChannel.Nodes.Clear
    
    ' Resize the channel IDs array
    channelCount = 0
    For i = 1 To parsed.Value.Count
        On Error Resume Next
        Dim countChannel As Object
        Set countChannel = parsed.Value(i)
        If Not countChannel Is Nothing Then
            If countChannel("type") = 0 Or countChannel("type") = 2 Then ' Count text and voice channels
                channelCount = channelCount + 1
            End If
        End If
        On Error GoTo 0
    Next i
    
    ReDim m_ChannelIds(0 To channelCount - 1) As String
    
    ' Track categories and their nodes
    Dim categories As New Collection
    Dim categoryNodes As New Collection
    Dim channelIndex As Long
    channelIndex = 0
    
    ' First pass: Create category nodes
    For i = 1 To parsed.Value.Count
        On Error Resume Next
        Dim channel As Object
        Set channel = parsed.Value(i)
        
        If Not channel Is Nothing Then
            ' Only process category channels (type 4)
            If channel("type") = 4 Then
                Dim categoryName As String
                categoryName = channel("name")
                
                ' Add to categories collection if not already present
                On Error Resume Next
                categories.Add categoryName, categoryName
                If Err.Number = 0 Then ' Only if added successfully (not duplicate)
                    ' Add to TreeView
                    Set categoryNode = lstChannel.Nodes.Add(, , "cat_" & channel("id"), categoryName)
                    categoryNodes.Add categoryNode, categoryName
                End If
                On Error GoTo 0
            End If
        End If
    Next i
    
    ' Second pass: Add channels under their categories
    For i = 1 To parsed.Value.Count
        On Error Resume Next
        Set channel = parsed.Value(i)
        
        If Not channel Is Nothing Then
            Dim channelType As Long
            channelType = channel("type")
            
            ' Only process text (0) and voice (2) channels
            If channelType = 0 Or channelType = 2 Then
                Dim channelName As String
                Dim channelId As String
                Dim parentCategory As String
                Dim parentNode As Node
                
                channelName = channel("name")
                channelId = channel("id")
                
                ' Get parent category ID if exists
                On Error Resume Next
                parentCategory = channel("parent_id")
                If Err.Number <> 0 Then parentCategory = ""
                On Error GoTo 0
                
                ' Find the category node if parent exists
                Set parentNode = Nothing
                If parentCategory <> "" Then
                    For Each categoryNode In lstChannel.Nodes
                        If InStr(categoryNode.Key, parentCategory) > 0 Then
                            Set parentNode = categoryNode
                            Exit For
                        End If
                    Next
                End If
                
                ' Add channel to TreeView
                If Not parentNode Is Nothing Then
                    ' Add under category
                    Set channelNode = lstChannel.Nodes.Add(parentNode, tvwChild, "ch_" & channelId, channelName)
                Else
                    ' Add at root level
                    Set channelNode = lstChannel.Nodes.Add(, , "ch_" & channelId, channelName)
                End If
                
      
                
                ' Store channel ID in array
                If channelIndex <= UBound(m_ChannelIds) Then
                    m_ChannelIds(channelIndex) = channelId
                    channelIndex = channelIndex + 1
                End If
            End If
        End If
    Next i
    
    ' Expand all categories by default
    For Each categoryNode In lstChannel.Nodes
        If InStr(categoryNode.Key, "cat_") = 1 Then
            categoryNode.Expanded = True
        End If
    Next
    
    ' Select first channel if any exist
    If lstChannel.Nodes.Count > 0 Then
        For Each channelNode In lstChannel.Nodes
            If InStr(channelNode.Key, "ch_") = 1 Then
                lstChannel.SelectedItem = channelNode
                Exit For
            End If
        Next
    End If
End Sub

Private Sub ProcessDMsResponse(aJson As String)
    On Error Resume Next
    Dim parsed As ParseResult
    Dim i As Long
    Dim sjson As String
    Dim dmNode As Node
    Dim dmParentNode As Node
    Dim dmCount As Long
    
    sjson = aJson
    parsed = Parse(sjson)
    
    If Not parsed.IsValid Then
        ' Save the response to debug file
        Dim iFile As Integer
        iFile = FreeFile
        Open App.Path & "\dm_error.json" For Output As #iFile
        Print #iFile, sjson
        Close #iFile
        
        MsgBox "Error parsing DM/GC data: " & parsed.Error & vbCrLf & _
               "Raw response saved to dm_error.json", vbExclamation
        Exit Sub
    End If
    
    ' Count DMs and group DMs first
    dmCount = 0
    For i = 1 To parsed.Value.Count
        On Error Resume Next
        Dim countChannel As Object
        Set countChannel = parsed.Value(i)
        If Not countChannel Is Nothing Then
            If countChannel("type") = 1 Or countChannel("type") = 3 Then
                dmCount = dmCount + 1
            End If
        End If
    Next i
    
    ' Resize the channel IDs array
    ReDim m_ChannelIds(0 To dmCount - 1) As String
    
    ' Add DMs parent node if we have any DMs
    If dmCount > 0 Then
        Set dmParentNode = lstChannel.Nodes.Add(, , "dms_parent", "Direct Messages")
        dmParentNode.Image = "dms"
        dmParentNode.Expanded = True
    End If
    
    ' Process each DM/GC
    Dim validDms As Long
    validDms = 0
    
    For i = 1 To parsed.Value.Count
        On Error Resume Next
        Dim channel As Object
        Set channel = parsed.Value(i)
        
        If Not channel Is Nothing Then
            Dim channelType As Long
            channelType = channel("type")
            
            ' Only process DM (1) and group DM (3) channels
            If channelType = 1 Or channelType = 3 Then
                Dim channelName As String
                Dim channelId As String
                
                channelId = channel("id")
                
                ' Get channel name based on type
                If channelType = 1 Then ' DM Channel
                    On Error Resume Next
                    Dim recipients As Object
                    Set recipients = channel("recipients")
                    If Not recipients Is Nothing And recipients.Count > 0 Then
                        Dim recipientUser As Object
                        Set recipientUser = recipients(1) ' First recipient
                        If Not recipientUser Is Nothing Then
                            channelName = recipientUser("global_name")
                        Else
                            channelName = "DM with Unknown User"
                        End If
                    Else
                        channelName = "DM (No Recipient Info)"
                    End If
                    On Error GoTo 0
                Else ' Group DM Channel
                    On Error Resume Next
                    channelName = channel("name")
                    If Err.Number <> 0 Or channelName = "" Then
                        ' If no name, create from recipients
                        Set recipients = channel("recipients")
                        If Not recipients Is Nothing And recipients.Count > 0 Then
                            Dim usernames() As String
                            ReDim usernames(1 To recipients.Count)
                            Dim k As Long
                            For k = 1 To recipients.Count
                                Set recipientUser = recipients(k)
                                If Not recipientUser Is Nothing Then
                                    usernames(k) = recipientUser("global_name")
                                Else
                                    usernames(k) = "Unknown"
                                End If
                            Next k
                            channelName = "Group with " & Join(usernames, ", ")
                        Else
                            channelName = "Unnamed Group"
                        End If
                    End If
                    On Error GoTo 0
                End If
                
                ' Add to TreeView under DMs parent
                
        On Error Resume Next
                Set dmNode = lstChannel.Nodes.Add(dmParentNode, tvwChild, "dm_" & channelId, channelName)
              '  dmNode.Image = IIf(channelType = 1, "dm", "group")
                
                ' Store channel ID in array
                If validDms <= UBound(m_ChannelIds) Then
                    m_ChannelIds(validDms) = channelId
                    validDms = validDms + 1
                End If
            End If
        End If
    Next i
    
    ' Select first DM if any exist
    If dmCount > 0 Then
        For Each dmNode In lstChannel.Nodes
            If InStr(dmNode.Key, "dm_") = 1 Then
                lstChannel.SelectedItem = dmNode
                Exit For
            End If
        Next
    End If
End Sub

Private Sub lstChannel_Click()
    On Error Resume Next
    
    ' Make sure we have a selected item
    If lstChannel.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    Dim selectedNode As Node
    Set selectedNode = lstChannel.SelectedItem
    
    ' Only process channel nodes (not category or DM parent nodes)
    If Left$(selectedNode.Key, 3) = "ch_" Or Left$(selectedNode.Key, 3) = "dm_" Then
        ' Extract the channel ID from the node key
        Dim channelId As String
        channelId = Mid$(selectedNode.Key, 4) ' Remove the "ch_" or "dm_" prefix
        
        ' Update the channel ID textbox
        txtCID.Text = channelId
        
        ' Store current channel ID for Gateway message filtering
        m_sCurrentChannelId = channelId
        
        ' Update the label to show the channel name
        If Left$(selectedNode.Key, 3) = "ch_" Then
            ' For regular channels, show the parent category if available
            If Not selectedNode.Parent Is Nothing Then
                lblChannel.Caption = "<" & selectedNode.Parent.Text & " / " & selectedNode.Text & ">"
            Else
                lblChannel.Caption = "<" & selectedNode.Text & ">"
            End If
        Else
            ' For DMs, just show the name
            lblChannel.Caption = "<" & selectedNode.Text & ">"
        End If
        
        ' Fetch messages for this channel
        FetchChannelMessages channelId
    End If
    
    Exit Sub
End Sub

Private Sub ProcessMessagesResponse(aJson As String)
'On Error Resume Next
    Dim parsed As ParseResult
    Dim i As Long
    Dim sOutput As String
    Dim fileNum As Integer
    Dim desktopPath As String
    Dim filePath As String
    Dim sjson As String
'    sjson = Left$(aJson, Len(aJson) - 5)
    ' Parse the JSON array
    sjson = aJson
    parsed = Parse(sjson)
 
    If Not parsed.IsValid Then
       ' Exit Sub
    End If
    
    ' Clear existing messages
    ChatView1.Clear
    ' Process each message
    For i = 20 To 1 Step -1
    On Error Resume Next
        Dim Msg As Object
        Set Msg = parsed.Value(i)
        
        ' Extract message details
        Dim sAuthor As String
        Dim sContent As String
        Dim sTimestamp As String
        
        On Error Resume Next ' Handle potential missing fields
        sAuthor = Msg("author")("global_name")
        sContent = Msg("content")
        sTimestamp = FormatDiscordTimestamp(Msg("timestamp"))
        
        ' Format the output
        sOutput = "[" & sTimestamp & "] " & sAuthor & ": " & sContent
        
        ' Add to listbox
        ChatView1.AddMessage sAuthor, sContent
    Next i
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

Private Sub Form_Load()
    m_sBaseUrl = "discord.com"
    
    'InitializeSocketArray
    Emojiz.Picture = LoadPicture(App.Path & "\emoji\msn\slight_smile.gif")
    ReDim m_GuildIds(0) As String
    ReDim m_ChannelIds(0) As String
    If GetSetting("DiscordClient", "Settings", "Token", "") <> "" Then
        txtToken.Text = GetSetting("DiscordClient", "Settings", "Token", "")
        m_sToken = GetSetting("DiscordClient", "Settings", "Token", "")
    End If
    
    ' Auto load channel from settings if available
        'txtCID.Text = GetSetting("DiscordClient", "Settings", "ChannelId", "")
       
    m_bGatewayReady = False ' Add this line
        If Len(m_sToken) > 0 Then
            'FetchChannelMessages txtCID.Text
            FetchMeDetails
            
            FetchUserGuilds
            
        ConnectToGateway
        End If
End Sub
Private Sub ConnectToGateway()
    If Not wsGateway Is Nothing Then
        If wsGateway.readyState <> STATE_CLOSED Then
            wsGateway.Disconnect
        End If
    End If
    
    ' Set up headers (Discord requires Authorization header)
    Dim Headers As New Collection
    Headers.Add "Authorization: " & m_sToken
    
    ' Configure WebSocket
    'wsGateway.UseCompression = True ' Discord supports compression
    'wsGateway.ChunkSize = 4096
    
    ' Clear any existing messages
    
    ' Connect to Discord Gateway
    StatusBar1.Panels(1).Text = "Connecting to Discord Gateway..."
    wsGateway.Connect DISCORD_GATEWAY_URL, "", "", "", Headers
End Sub

Private Sub DisconnectFromGateway()
    If Not wsGateway Is Nothing Then
        If wsGateway.readyState <> STATE_CLOSED Then
            wsGateway.Disconnect
        End If
    End If
End Sub
Private Sub wsGateway_onClose(ByVal eCode As WebsocketStatus, ByVal reason As String)
MsgBox "Disconnected from Gateway: " & reason
    m_bGatewayConnected = False
    m_bIdentified = False
    
    ' Attempt to reconnect after a delay
    If m_sToken <> "" Then
      MsgBox "Attempting to reconnect in 5 seconds..."
    End If
End Sub

Private Sub wsGateway_OnConnect(ByVal RemoteHost As String, ByVal RemoteIP As String, ByVal RemotePort As String)

        Form1.tmrHeartbeat.Enabled = True
        Form1.tmrHeartbeat.Interval = 1000
   StatusBar1.Panels(1).Text = "Connected to Discord Gateway"
    m_bGatewayConnected = True
End Sub

Private Sub wsGateway_onError(ByVal eCode As WebsocketStatus, ByVal reason As String)
   MsgBox "Gateway Error: " & CStr(eCode) & ": " & reason
End Sub

Private Sub wsGateway_OnMessage(ByVal Msg As Variant, ByVal OpCode As WebsocketOpCode)
    ' Discord Gateway messages are always JSON text
    If OpCode = opText Then
        ProcessGatewayMessage CStr(Msg)
    Else
           StatusBar1.Panels(1).Text = "Received unexpected binary data from Gateway"
    End If
End Sub

Private Sub wsGateway_OnPong(ByVal IncludedMsg As String)
     StatusBar1.Panels(1).Text = "Received Pong from Gateway"
End Sub
Private Sub ProcessGatewayMessage(ByVal jsonMessage As String)
On Error Resume Next
    
    Dim parsed As ParseResult
    parsed = Parse(jsonMessage)
    
    If Not parsed.IsValid Then
         MsgBox "Invalid JSON from Gateway: " & jsonMessage
        Exit Sub
    End If
    
    Dim OpCode As Long
    Dim eventData As Object
    Dim seqNum As Long
    
    OpCode = parsed.Value("op")
    Set eventData = parsed.Value("d")
    
    seqNum = parsed.Value("s")
    
    If seqNum > 0 Then
        m_lSequence = seqNum ' Store sequence number for reconnects
    End If
    
    Select Case OpCode
        Case 0: ' Dispatch (event)
            ProcessGatewayEvent parsed.Value("t"), eventData
            
        Case 1: ' Heartbeat
            SendHeartbeat
            
        Case 7: ' Reconnect
          StatusBar1.Panels(1).Text = "Gateway requesting reconnect..."
            DisconnectFromGateway
            ConnectToGateway
            
        Case 9: ' Invalid Session
        StatusBar1.Panels(1).Text = "Gateway session invalidated"
            m_bIdentified = False
            If eventData Then ' Can resume?
             StatusBar1.Panels(1).Text = "Attempting to resume session..."
                SendResume
                
            Else
             StatusBar1.Panels(1).Text = "Starting new session..."
                SendIdentify
            End If
            
           Case 10: ' Hello (contains heartbeat interval)
        m_lHeartbeatInterval = eventData("heartbeat_interval")
        StatusBar1.Panels(1).Text = "Gateway Hello received, heartbeat interval: " & m_lHeartbeatInterval
        
        tmrHeartbeat.Interval = m_lHeartbeatInterval * 0.75
        
        tmrHeartbeat.Enabled = True
        
        SendHeartbeat
            
            If Not m_bIdentified Then
                SendIdentify
            End If
            
        Case 11: ' Heartbeat ACK
            ' Nothing to do, we got our ACK
         StatusBar1.Panels(1).Text = "Heartbeat acknowledged"
            
        Case Else
           StatusBar1.Panels(1).Text = "Unhandled Gateway opcode: " & OpCode
    End Select
    
    Exit Sub
EH:
MsgBox "Error processing Gateway message: " & Err.Description
End Sub

Private Sub ProcessGatewayEvent(ByVal eventType As String, ByVal eventData As Object)
    On Error Resume Next
    
    Select Case eventType
        Case "MESSAGE_CREATE"
            If eventData("channel_id") = m_sCurrentChannelId Then
                Dim author As String
                author = eventData("author")("global_name")
                Dim Content As String
                Content = eventData("content")
                
                ChatView1.AddMessage author, Content
            End If
            
        Case "READY"
            StatusBar1.Panels(1).Text = "Successfully identified with Gateway"
            m_bIdentified = True
            m_sSessionId = eventData("session_id")
            
            ' Clear existing guilds
           ' GuildView1.AddGuild "DMs", LoadPicture(App.Path & "\everydiscord.gif")

            
 Case "GUILD_CREATE"
 MsgBox "ping"
            ' This matches your original ProcessGuildsResponse structure
            Dim guildId As String
            Dim guildIndex As Long
            Dim sGuildName As String
            Dim sIconHash As String
            
            guildId = eventData("id")
            sGuildName = eventData("name")
            MsgBox sGuildName & " h"
            On Error Resume Next
            sIconHash = eventData("icon")
            On Error GoTo 0
            Dim i As Integer
            ' Find this guild in our array (same as REST version)
            guildIndex = -1
            For i = 0 To UBound(m_GuildIds)
                If m_GuildIds(i) = guildId Then
                    guildIndex = i
                    Exit For
                ElseIf m_GuildIds(i) = "" Then
                    ' Empty slot, use this
                    m_GuildIds(i) = guildId
                    guildIndex = i
                    Exit For
                End If
            Next i
            
            If guildIndex = -1 Then
                ' New guild - expand array (same as before)
                guildIndex = UBound(m_GuildIds) + 1
                ReDim Preserve m_GuildIds(guildIndex)
                m_GuildIds(guildIndex) = guildId
            End If
            
            ' Add to GuildView (same as REST version)
            MsgBox sGuildName
            GuildView1.AddGuild sGuildName, LoadPicture()
            
            ' Fetch icon if available (same as before)
            If Len(sIconHash) > 0 Then
                'QueueGuildIconFetch guildId, sIconHash, guildIndex
            End If
            
            ' Update status when all initial guilds are loaded
            If m_bGatewayReady Then
                Dim allLoaded As Boolean
                allLoaded = True
                
                For i = 0 To UBound(m_GuildIds)
                    If GuildView1.GetGuildName(i + 1) = "" Then
                        allLoaded = False
                        Exit For
                    End If
                Next i
                
                If allLoaded Then
                    StatusBar1.Panels(1).Text = "Loaded " & UBound(m_GuildIds) + 1 & " guilds"
                End If
            End If
            
        Case "GUILD_DELETE"
            ' We'll implement this later as requested
            ' Just log for now
            Debug.Print "Guild removed: " & eventData("id")
            
        Case Else
            StatusBar1.Panels(1).Text = "Unhandled event: " & eventType
    End Select
End Sub
Private Sub SendIdentify()
    Dim identifyPayload As String
    identifyPayload = "{""op"":2,""d"":{""token"":""" & m_sToken & """,""properties"":{""$os"":""windows"",""$browser"":""my_vb6_client"",""$device"":""my_vb6_client""},""compress"":false,""large_threshold"":250}}"
    
    wsGateway.SendAdvanced identifyPayload, 1, True, False, True, False, False, False
 StatusBar1.Panels(1).Text = "Sent Identify payload"
End Sub

Private Sub SendResume()
    Dim resumePayload As String
    resumePayload = "{""op"":6,""d"":{""token"":""" & m_sToken & """,""session_id"":""" & m_sSessionId & """,""seq"":" & m_lSequence & "}}"
    
    wsGateway.SendAdvanced resumePayload, 1, True, False, True, False, False, False
 StatusBar1.Panels(1).Text = "Sent Resume payload"
End Sub

Private Sub SendHeartbeat()
    If m_bGatewayConnected Then
        Dim heartbeatPayload As String
        heartbeatPayload = "{""op"":1,""d"":" & IIf(m_lSequence > 0, m_lSequence, "null") & "}"
        
        wsGateway.SendAdvanced heartbeatPayload, 1, True, False, True, False, False, False
        m_dLastHeartbeat = Now
       StatusBar1.Panels(1).Text = "Sent Heartbeat"
    End If
End Sub



Private Sub FetchUserGuilds()
    Dim tlsSocket As New cTlsSocket
    Dim sRequest As String
    Dim sResponse As String
    Dim sFullResponse As String
    Dim bChunked As Boolean
    Dim lContentLength As Long
    Dim headersEnd As Long
    Dim sHeaders As String
    Dim sBody As String
    
    
    
    ' Create the API request path
    Dim sPath As String
    sPath = "api/v10/users/@me/guilds"
    
    ' Prepare the HTTP request
    sRequest = "GET /" & sPath & " HTTP/1.1" & vbCrLf & _
              "Host: " & m_sBaseUrl & vbCrLf & _
              "Authorization: " & m_sToken & vbCrLf & _
              "Connection: close" & vbCrLf & vbCrLf
    
    ' Configure and connect socket
    tlsSocket.SyncConnect m_sBaseUrl, 443, , , ucsTlsSupportAll Or ucsTlsIgnoreServerCertificateErrors
    
    ' Send request
    tlsSocket.SyncSendText sRequest
    
    ' Receive initial response
    sResponse = tlsSocket.SyncReceiveText
    sFullResponse = sResponse
    
    ' Check if response is chunked
    headersEnd = InStr(sFullResponse, vbCrLf & vbCrLf)
    
    If headersEnd > 0 Then
        sHeaders = Left$(sFullResponse, headersEnd + 1)
        sBody = Mid$(sFullResponse, headersEnd + 4)
        
        ' Check for chunked transfer encoding
        bChunked = InStr(1, sHeaders, "Transfer-Encoding: chunked", vbTextCompare) > 0
        
        ' Check for Content-Length if not chunked
        If Not bChunked Then
            Dim contentLengthPos As Long
            contentLengthPos = InStr(1, sHeaders, "Content-Length:", vbTextCompare)
            If contentLengthPos > 0 Then
                Dim contentLengthEnd As Long
                contentLengthEnd = InStr(contentLengthPos, sHeaders, vbCrLf)
                lContentLength = Val(Mid$(sHeaders, contentLengthPos + 15, contentLengthEnd - (contentLengthPos + 15)))
            End If
        End If
    End If
    
    ' If chunked, keep receiving until we get all chunks
    If bChunked Then
        Do
            ' Check if we've received the final chunk (0\r\n\r\n)
            If InStr(sFullResponse, "0" & vbCrLf & vbCrLf) > 0 Then
                Exit Do
            End If
            
            ' Receive more data
            sResponse = tlsSocket.SyncReceiveText
            If Len(sResponse) = 0 Then Exit Do
            sFullResponse = sFullResponse & sResponse
        Loop
        
        ' Properly decode the chunked response
        sBody = DecodeChunkedResponse(sFullResponse)
    ElseIf lContentLength > 0 Then
        ' For non-chunked with Content-Length, receive until we have all data
        Do While Len(sFullResponse) - headersEnd - 3 < lContentLength
            sResponse = tlsSocket.SyncReceiveText
            If Len(sResponse) = 0 Then Exit Do
            sFullResponse = sFullResponse & sResponse
        Loop
        sBody = Mid$(sFullResponse, headersEnd + 4)
    Else
        ' Fallback for responses without Content-Length or chunked encoding
        sBody = Mid$(sFullResponse, headersEnd + 4)
    End If
    
    ' Process the JSON response
    ProcessGuildsResponse sBody
    
    ' Clean up
    tlsSocket.Close_
    Set tlsSocket = Nothing
    
    Exit Sub

End Sub

Private Sub FetchMePicture(sUserId As String, sAvatarHash As String)
'MsgBox "hi"
    Set m_oHttpDownload = New cHttpDownload

'm_oHttpDownload.DownloadFile "https://cdn.discordapp.com/avatars/" & sUserId & "/" & sAvatarHash & ".jpg", Environ$("TMP") & "\" & sUserId & ".jpg"

   
m_oHttpDownload.DownloadFile "https://cdn.discordapp.com/avatars/872926577858609182/4b696cfd0ac0c2ff9be940ca26881cfe.jpg?size=1024", Environ$("TMP") & "\" & "hiee" & ".jpg" '    m_oHttpDownload.DownloadFile IIf(chkUseHttps.Value = vbChecked, "https", "http") & "://dl.unicontsoft.com/upload/aaa.zip", Environ$("TMP") & "\aaa.zip"
   Exit Sub
End Sub
Private Sub m_oHttpDownload_DownloadComplete(ByVal LocalFileName As String)
    Const FUNC_NAME     As String = "m_oHttpDownload_DownloadComplete"
    
    MsgBox "Download to " & LocalFileName & " complete", vbExclamation
End Sub
Private Sub FetchMeDetails()
  
  Dim chtlsSocket As New cTlsSocket
    Dim sRequest As String
    Dim sResponse As String
    Dim sFullResponse As String
    Dim bChunked As Boolean
    Dim lContentLength As Long
    Dim headersEnd As Long
    Dim sHeaders As String
    Dim sBody As String
    
    chtlsSocket.SyncConnect m_sBaseUrl, 443, , , ucsTlsSupportAll Or ucsTlsIgnoreServerCertificateErrors
    
    ' Create the API request path
    Dim sPath As String
    sPath = "api/users/@me"
    
    ' Prepare the HTTP request
    sRequest = "GET /" & sPath & " HTTP/1.1" & vbCrLf & _
              "Host: " & m_sBaseUrl & vbCrLf & _
              "Authorization: " & m_sToken & vbCrLf & _
              "Connection: close" & vbCrLf & vbCrLf
    
    ' Send request
    chtlsSocket.SyncSendText sRequest
    
    ' Receive initial response
    sResponse = chtlsSocket.SyncReceiveText
    sFullResponse = sResponse
    
    ' Check if response is chunked
    headersEnd = InStr(sFullResponse, vbCrLf & vbCrLf)
    
    If headersEnd > 0 Then
        sHeaders = Left$(sFullResponse, headersEnd + 1)
        sBody = Mid$(sFullResponse, headersEnd + 4)
        
        ' Check for chunked transfer encoding
        bChunked = InStr(1, sHeaders, "Transfer-Encoding: chunked", vbTextCompare) > 0
        
        ' Check for Content-Length if not chunked
        If Not bChunked Then
            Dim contentLengthPos As Long
            contentLengthPos = InStr(1, sHeaders, "Content-Length:", vbTextCompare)
            If contentLengthPos > 0 Then
                Dim contentLengthEnd As Long
                contentLengthEnd = InStr(contentLengthPos, sHeaders, vbCrLf)
                lContentLength = Val(Mid$(sHeaders, contentLengthPos + 15, contentLengthEnd - (contentLengthPos + 15)))
            End If
        End If
    End If
    
    ' If chunked, keep receiving until we get all chunks
    If bChunked Then
        Do
            ' Check if we've received the final chunk (0\r\n\r\n)
            If InStr(sFullResponse, "0" & vbCrLf & vbCrLf) > 0 Then
                Exit Do
            End If
            
            ' Receive more data
            sResponse = chtlsSocket.SyncReceiveText
            If Len(sResponse) = 0 Then Exit Do
            sFullResponse = sFullResponse & sResponse
        Loop
        
        ' Properly decode the chunked response
        sBody = DecodeChunkedResponse(sFullResponse)
    ElseIf lContentLength > 0 Then
        ' For non-chunked with Content-Length, receive until we have all data
        Do While Len(sFullResponse) - headersEnd - 3 < lContentLength
            sResponse = chtlsSocket.SyncReceiveText
            If Len(sResponse) = 0 Then Exit Do
            sFullResponse = sFullResponse & sResponse
        Loop
        sBody = Mid$(sFullResponse, headersEnd + 4)
    Else
        ' Fallback for responses without Content-Length or chunked encoding
        sBody = Mid$(sFullResponse, headersEnd + 4)
    End If
    
    ' Process the JSON response
    Dim jparse As Dictionary
    Set jparse = Parse(sBody).Value
    Label1.Caption = jparse("global_name")
    
    ' Extract avatar information and fetch profile picture
    Dim sUserId As String
    Dim sAvatarHash As String
    
    sUserId = jparse("id")
    If Not IsNull(jparse("avatar")) Then
        sAvatarHash = jparse("avatar")
        ' Fetch the profile picture
        FetchMePicture sUserId, sAvatarHash
    End If
    
    ' Clean up
    chtlsSocket.Close_
    Set chtlsSocket = Nothing
    
    Exit Sub

End Sub
Private Sub FetchGuildChannels(ByVal sGuildId As String)
  
  Dim chtlsSocket As New cTlsSocket
    Dim sRequest As String
    Dim sResponse As String
    Dim sFullResponse As String
    Dim bChunked As Boolean
    Dim lContentLength As Long
    Dim headersEnd As Long
    Dim sHeaders As String
    Dim sBody As String
    
    chtlsSocket.SyncConnect m_sBaseUrl, 443, , , ucsTlsSupportAll Or ucsTlsIgnoreServerCertificateErrors
    
    ' Create the API request path
    Dim sPath As String
    sPath = "api/v10/guilds/" & sGuildId & "/channels"
    
    ' Prepare the HTTP request
    sRequest = "GET /" & sPath & " HTTP/1.1" & vbCrLf & _
              "Host: " & m_sBaseUrl & vbCrLf & _
              "Authorization: " & m_sToken & vbCrLf & _
              "Connection: close" & vbCrLf & vbCrLf
    
    ' Configure and connect socket
    
    ' Send request
    
    chtlsSocket.SyncSendText sRequest
    
    ' Receive initial response
    sResponse = chtlsSocket.SyncReceiveText
    sFullResponse = sResponse
    
    ' Check if response is chunked
    headersEnd = InStr(sFullResponse, vbCrLf & vbCrLf)
    
    If headersEnd > 0 Then
        sHeaders = Left$(sFullResponse, headersEnd + 1)
        sBody = Mid$(sFullResponse, headersEnd + 4)
        
        ' Check for chunked transfer encoding
        bChunked = InStr(1, sHeaders, "Transfer-Encoding: chunked", vbTextCompare) > 0
        
        ' Check for Content-Length if not chunked
        If Not bChunked Then
            Dim contentLengthPos As Long
            contentLengthPos = InStr(1, sHeaders, "Content-Length:", vbTextCompare)
            If contentLengthPos > 0 Then
                Dim contentLengthEnd As Long
                contentLengthEnd = InStr(contentLengthPos, sHeaders, vbCrLf)
                lContentLength = Val(Mid$(sHeaders, contentLengthPos + 15, contentLengthEnd - (contentLengthPos + 15)))
            End If
        End If
    End If
    
    ' If chunked, keep receiving until we get all chunks
    If bChunked Then
        Do
            ' Check if we've received the final chunk (0\r\n\r\n)
            If InStr(sFullResponse, "0" & vbCrLf & vbCrLf) > 0 Then
                Exit Do
            End If
            
            ' Receive more data
            sResponse = chtlsSocket.SyncReceiveText
            If Len(sResponse) = 0 Then Exit Do
            sFullResponse = sFullResponse & sResponse
        Loop
        
        ' Properly decode the chunked response
        sBody = DecodeChunkedResponse(sFullResponse)
    ElseIf lContentLength > 0 Then
        ' For non-chunked with Content-Length, receive until we have all data
        Do While Len(sFullResponse) - headersEnd - 3 < lContentLength
            sResponse = chtlsSocket.SyncReceiveText
            If Len(sResponse) = 0 Then Exit Do
            sFullResponse = sFullResponse & sResponse
        Loop
        sBody = Mid$(sFullResponse, headersEnd + 4)
    Else
        ' Fallback for responses without Content-Length or chunked encoding
        sBody = Mid$(sFullResponse, headersEnd + 4)
    End If
    
    ' Process the JSON response
    ProcessChannelsResponse sBody
    
    ' Clean up
    chtlsSocket.Close_
    Set chtlsSocket = Nothing
    
    Exit Sub

End Sub
Private Sub FetchUserDMs()

               
               
                Dim tlsSocket As New cTlsSocket
    Dim sRequest As String
    Dim sResponse As String
    Dim sFullResponse As String
    Dim bChunked As Boolean
    Dim lContentLength As Long
    Dim headersEnd As Long
    Dim sHeaders As String
    Dim sBody As String
    
    ' Create the API request path
    Dim sPath As String
    sPath = "api/v10/users/@me/channels" ' Changed from /guilds to /channels
    
    
    sRequest = "GET /" & sPath & " HTTP/1.1" & vbCrLf & _
               "Host: " & m_sBaseUrl & vbCrLf & _
               "Authorization: " & m_sToken & vbCrLf & _
               "Connection: close" & vbCrLf & vbCrLf
               
    ' Configure and connect socket
    tlsSocket.SyncConnect m_sBaseUrl, 443, , , ucsTlsSupportAll Or ucsTlsIgnoreServerCertificateErrors
    
    ' Send request
    tlsSocket.SyncSendText sRequest
    
    ' Receive initial response
    sResponse = tlsSocket.SyncReceiveText
    sFullResponse = sResponse
    
    ' Check if response is chunked
    headersEnd = InStr(sFullResponse, vbCrLf & vbCrLf)
    
    If headersEnd > 0 Then
        sHeaders = Left$(sFullResponse, headersEnd + 1)
        sBody = Mid$(sFullResponse, headersEnd + 4)
        
        ' Check for chunked transfer encoding
        bChunked = InStr(1, sHeaders, "Transfer-Encoding: chunked", vbTextCompare) > 0
        
        ' Check for Content-Length if not chunked
        If Not bChunked Then
            Dim contentLengthPos As Long
            contentLengthPos = InStr(1, sHeaders, "Content-Length:", vbTextCompare)
            If contentLengthPos > 0 Then
                Dim contentLengthEnd As Long
                contentLengthEnd = InStr(contentLengthPos, sHeaders, vbCrLf)
                lContentLength = Val(Mid$(sHeaders, contentLengthPos + 15, contentLengthEnd - (contentLengthPos + 15)))
            End If
        End If
    End If
    
    ' If chunked, keep receiving until we get all chunks
    If bChunked Then
        Do
            ' Check if we've received the final chunk (0\r\n\r\n)
            If InStr(sFullResponse, "0" & vbCrLf & vbCrLf) > 0 Then
                Exit Do
            End If
            
            ' Receive more data
            sResponse = tlsSocket.SyncReceiveText
            If Len(sResponse) = 0 Then Exit Do
            sFullResponse = sFullResponse & sResponse
        Loop
        
        ' Properly decode the chunked response
        sBody = DecodeChunkedResponse(sFullResponse)
    ElseIf lContentLength > 0 Then
        ' For non-chunked with Content-Length, receive until we have all data
        Do While Len(sFullResponse) - headersEnd - 3 < lContentLength
            sResponse = tlsSocket.SyncReceiveText
            If Len(sResponse) = 0 Then Exit Do
            sFullResponse = sFullResponse & sResponse
        Loop
        sBody = Mid$(sFullResponse, headersEnd + 4)
    Else
        ' Fallback for responses without Content-Length or chunked encoding
        sBody = Mid$(sFullResponse, headersEnd + 4)
    End If
    
    ' Process the JSON response
    'MsgBox sBody
    ProcessDMsResponse sBody
    
    ' Clean up
    tlsSocket.Close_
    Set tlsSocket = Nothing
    
    Exit Sub

End Sub
Private Sub FetchChannelMessages(ByVal sChannelId As String, Optional ByVal lLimit As Long = 30)
    Dim tlsSocket As New cTlsSocket
    Dim sRequest As String
    Dim sResponse As String
    Dim sFullResponse As String
    Dim bChunked As Boolean
    Dim lContentLength As Long
    Dim headersEnd As Long
    Dim sHeaders As String
    Dim sBody As String
    Dim vBody() As Byte
    Dim sUTF16Body As String
    
    ' Create the API request path
    Dim sPath As String
    sPath = "api/v10/channels/" & sChannelId & "/messages?limit=" & lLimit
    
    sRequest = "GET /" & sPath & " HTTP/1.1" & vbCrLf & _
              "Host: " & m_sBaseUrl & vbCrLf & _
              "Authorization: " & m_sToken & vbCrLf & _
              "Connection: close" & vbCrLf & vbCrLf
    
    ' Configure and connect socket
    tlsSocket.SyncConnect "discord.com", 443, , , ucsTlsSupportAll Or ucsTlsIgnoreServerCertificateErrors
    
    ' Send request
    tlsSocket.SyncSendText sRequest
    
    ' Receive initial response
    sResponse = tlsSocket.SyncReceiveText
    sFullResponse = sResponse
    
    ' Check if response is chunked
    headersEnd = InStr(sFullResponse, vbCrLf & vbCrLf)
    
    If headersEnd > 0 Then
        sHeaders = Left$(sFullResponse, headersEnd + 1)
        sBody = Mid$(sFullResponse, headersEnd + 4)
        
        ' Check for chunked transfer encoding
        bChunked = InStr(1, sHeaders, "Transfer-Encoding: chunked", vbTextCompare) > 0
        
        ' Check for Content-Length if not chunked
        If Not bChunked Then
            Dim contentLengthPos As Long
            contentLengthPos = InStr(1, sHeaders, "Content-Length:", vbTextCompare)
            If contentLengthPos > 0 Then
                Dim contentLengthEnd As Long
                contentLengthEnd = InStr(contentLengthPos, sHeaders, vbCrLf)
                lContentLength = Val(Mid$(sHeaders, contentLengthPos + 15, contentLengthEnd - (contentLengthPos + 15)))
            End If
        End If
    End If
    
    ' If chunked, keep receiving until we get all chunks
    If bChunked Then
        Do
            ' Check if we've received the final chunk (0\r\n\r\n)
            If InStr(sFullResponse, "0" & vbCrLf & vbCrLf) > 0 Then
                Exit Do
            End If
            
            ' Receive more data
            sResponse = tlsSocket.SyncReceiveText
            If Len(sResponse) = 0 Then Exit Do
            sFullResponse = sFullResponse & sResponse
        Loop
        
        ' Properly decode the chunked response
        sBody = DecodeChunkedResponse(sFullResponse)
    ElseIf lContentLength > 0 Then
        ' For non-chunked with Content-Length, receive until we have all data
        Do While Len(sFullResponse) - headersEnd - 3 < lContentLength
            sResponse = tlsSocket.SyncReceiveText
            If Len(sResponse) = 0 Then Exit Do
            sFullResponse = sFullResponse & sResponse
        Loop
        sBody = Mid$(sFullResponse, headersEnd + 4)
    Else
        ' Fallback for responses without Content-Length or chunked encoding
        sBody = Mid$(sFullResponse, headersEnd + 4)
    End If
    
    ' Convert the response body from UTF-8 to UTF-16 using the API function

    
    ' Process the JSON response
    ProcessMessagesResponse sBody
    
    ' Save the response as UTF-16 using the same conversion
    Dim iFile As Integer
    iFile = FreeFile
    Open App.Path & "\fcm.txt" For Output As #iFile
    Print #iFile, sBody
    Close #iFile
    
    ' Clean up
    tlsSocket.Close_
    Set tlsSocket = Nothing
End Sub

Private Function DecodeChunkedResponse(sResponse As String) As String
    Dim headersEnd As Long
    Dim sChunkedData As String
    Dim sResult As String
    Dim lChunkSize As Long
    Dim lPos As Long
    
    ' Find end of headers
    headersEnd = InStr(sResponse, vbCrLf & vbCrLf)
    If headersEnd = 0 Then
        DecodeChunkedResponse = sResponse
        Exit Function
    End If
    
    ' Get just the chunked data part
    sChunkedData = Mid$(sResponse, headersEnd + 4)
    sResult = ""
    lPos = 1
    
    Do While lPos <= Len(sChunkedData)
        ' Find the next line break
        Dim lLineEnd As Long
        lLineEnd = InStr(lPos, sChunkedData, vbCrLf)
        If lLineEnd = 0 Then Exit Do
        
        ' Get the chunk size line
        Dim sChunkSize As String
        sChunkSize = Mid$(sChunkedData, lPos, lLineEnd - lPos)
        
        ' Remove any extensions (like ;chunk-ext)
        If InStr(sChunkSize, ";") > 0 Then
            sChunkSize = Left$(sChunkSize, InStr(sChunkSize, ";") - 1)
        End If
        
        ' Convert hex to decimal
        On Error Resume Next
        lChunkSize = Val("&h" & sChunkSize)
        If Err.Number <> 0 Or lChunkSize < 0 Then
            Exit Do
        End If
        On Error GoTo 0
        
        ' Check for final chunk
        If lChunkSize = 0 Then Exit Do
        
        ' Move to start of chunk data
        lPos = lLineEnd + 2
        
        ' Get the chunk data
        If lPos + lChunkSize - 1 > Len(sChunkedData) Then
            Exit Do
        End If
        
        sResult = sResult & Mid$(sChunkedData, lPos, lChunkSize)
        
        ' Move to next chunk (after data + CRLF)
        lPos = lPos + lChunkSize + 2
    Loop
    
    DecodeChunkedResponse = Left$(sResult, Len(sResult) - 1)
End Function
Private Sub SendDiscordMessage(ByVal sChannelId As String, ByVal sMessage As String)

    
    Dim tlsSocket As New cTlsSocket
    Dim sPayload As String
    Dim sRequest As String
    Dim sResponse As String
    
    ' Create JSON payload
    sPayload = "{""content"":""" & EscapeJsonString(sMessage) & """}"
    
    ' Build HTTP request
    sRequest = "POST /api/v10/channels/" & sChannelId & "/messages HTTP/1.1" & vbCrLf & _
               "Host: discord.com" & vbCrLf & _
               "Authorization: " & m_sToken & vbCrLf & _
               "Content-Type: application/json" & vbCrLf & _
               "Content-Length: " & Len(sPayload) & vbCrLf & _
               "Connection: close" & vbCrLf & vbCrLf & _
               sPayload
    
    ' Configure and connect socket
    ' Allow TLS 1.2 or 1.3
    tlsSocket.SyncConnect "discord.com", 443, , , ucsTlsSupportAll Or ucsTlsIgnoreServerCertificateErrors
    
    ' Send request
    tlsSocket.SyncSendText sRequest
    
    ' Receive response (blocking)
    sResponse = tlsSocket.SyncReceiveText

    
    ' Clean up
    tlsSocket.Close_
    Set tlsSocket = Nothing


End Sub


Private Sub ProcessIconResponse(ByVal socketIndex As Long, ByVal sResponse As String)
    
DoEvents
    Dim HeaderEnd As Long
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
    HeaderEnd = InStr(sResponse, vbCrLf & vbCrLf)

    If HeaderEnd > 0 Then
        Dim binaryData As String
        binaryData = Mid$(sResponse, HeaderEnd + 4)
        
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
End Sub
