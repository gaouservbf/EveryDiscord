Attribute VB_Name = "Module2"

Public Const DISCORD_GATEWAY_URL As String = "wss://gateway.discord.gg/?v=10&encoding=json"
Public Const DISCORD_GATEWAY_VERSION As String = "10"
Public Const DISCORD_GATEWAY_PORT As Long = 443

Public m_lFreeSocketIndex As Long     ' Next available socket index

' Control array constants
Public Const MAX_CONNECTIONS As Long = 100
Public Const SOCKET_IDLE As Long = 0
Public Const SOCKET_CONNECTING As Long = 1
Public Const SOCKET_CONNECTED As Long = 2
Public Const SOCKET_RECEIVING As Long = 3

' TLS Context and Socket variables
Public m_GuildIds() As String
Public m_ChannelIds() As String
Public m_uCtx() As UcsTlsContext      ' Array of TLS contexts
Public m_sRequest() As String         ' Array of requests
Public m_sToken As String
Public m_sBaseUrl As String

' Socket state tracking
Public m_lSocketState() As Long       ' State of each socket
Public m_sResponseBuffer() As String  ' Response buffer for each socket
Public m_lContentLength() As Long     ' Content length for each socket
Public m_bReceivingData() As Boolean  ' Receiving flag for each socket
Public m_RequestType() As String      ' Request type for each socket
Public m_ExtraData() As String        ' Extra data for each socket
Public m_Index() As Long              ' Index for each socket
Public m_Target() As String           ' Target for each socket

' Gateway state tracking
Public m_sSessionId As String
Public m_lSequence As Long
Public m_lHeartbeatInterval As Long
Public m_dLastHeartbeat As Double
Public m_bGatewayConnected As Boolean
Public m_bIdentified As Boolean
Public m_sGatewayToken As String
Public m_sCurrentChannelId As String

' Add these to your declarations section
Public Type IconRequest
    guildId As String
    iconHash As String
    guildIndex As Long
End Type

Public IconRequests() As IconRequest
Public IconRequestCount As Long
Public CurrentIconRequest As Long
Public m_bFetchingIcon As Boolean

Public Type RequestItem
    RequestType As String    ' "GuildList", "Channels", "Messages", "Icon", etc.
    Target As String         ' Guild ID, Channel ID, etc. depending on type
    Request As String        ' The full HTTP request
    ExtraData As String      ' Additional data (e.g., Icon hash, index)
    Index As Long            ' For indexing into arrays if needed
End Type

Public Const MAX_QUEUE_SIZE As Long = 1000
Public m_RequestQueue(0 To MAX_QUEUE_SIZE - 1) As RequestItem
Public m_QueueHead As Long
Public m_QueueTail As Long
Public m_QueueCount As Long
Public m_bInProcessQueue As Boolean

Public m_ProcessingRequest As Boolean
Public m_CurrentRequestType As String
Public m_GatewayBuffer As String
