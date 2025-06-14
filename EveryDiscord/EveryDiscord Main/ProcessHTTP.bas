Attribute VB_Name = "ProcessHTTP"
Option Explicit

Public Sub ProcessGuildsResponse(aJson As String)

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
    Form1.GuildView1.ClearGuilds
    
Form1.GuildView1.AddGuild "DMs", LoadPicture(App.Path & "\everydiscord.gif")
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
Public Sub ProcessChannelsResponse(aJson As String)
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
        MsgBox "Error parsing channel data, JSON len: " & Len(aJson)
        Exit Sub
    End If
    
    ' Clear existing channels
    lstChannel.Clear
    
    ' Count text channels first (type 0)
    channelCount = 0
    For i = 1 To parsed.Value.Count
        On Error Resume Next
        Dim countChannel As Object
        Set countChannel = parsed.Value(i)
        
        If Not countChannel Is Nothing Then
            If countChannel("type") = 0 Then ' Only count text channels
                channelCount = channelCount + 1
            End If
        End If
        On Error GoTo 0
    Next i
    
    ' Resize the channel IDs array to hold all channels
    ReDim m_ChannelIds(0 To channelCount - 1) As String
    
    ' Reset valid channels counter
    validChannels = 0
 
    ' Process each channel
    For i = 1 To parsed.Value.Count
        On Error Resume Next
        Dim channel As Object
        Set channel = parsed.Value(i)
        
        If Not channel Is Nothing Then
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
                
                ' Store ID in array (we know it's properly sized now)
                If validChannels < channelCount Then
                    m_ChannelIds(validChannels) = sChannelId
                    validChannels = validChannels + 1
                    
                    ' Debug output
                    Debug.Print "Added channel: " & sChannelName & " with ID: " & sChannelId
                End If
            End If
        End If
        On Error GoTo 0
    Next i
    
    ' If we found channels, select the first one
    If lstChannel.ListCount > 0 Then
        lstChannel.ListIndex = 0
    End If
End Sub
