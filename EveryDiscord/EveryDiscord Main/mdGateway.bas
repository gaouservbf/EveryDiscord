Attribute VB_Name = "mdGateway"

' Gateway OpCodes
Private Const GATEWAY_OP_DISPATCH As Long = 0        ' Events from Discord
Private Const GATEWAY_OP_HEARTBEAT As Long = 1       ' Keep connection alive
Private Const GATEWAY_OP_IDENTIFY As Long = 2        ' Authenticate connection
Private Const GATEWAY_OP_RESUME As Long = 6          ' Resume disconnected session
Private Const GATEWAY_OP_RECONNECT As Long = 7       ' Server asks client to reconnect
Private Const GATEWAY_OP_HELLO As Long = 10          ' Initial handshake
Private Const GATEWAY_OP_HEARTBEAT_ACK As Long = 11  ' Server ack of heartbeat

