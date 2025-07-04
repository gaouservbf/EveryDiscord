VERSION 5.00
Begin VB.UserControl ChatView 
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FillColor       =   &H8000000E&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.VScrollBar vsbChat 
      Height          =   3495
      Left            =   4560
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "ChatView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'chatview.ctl
Private Type ChatMessage
    UserName As String
    UserId As String ' Added for status tracking
    Text As String
    Avatar As StdPicture ' optional
    RenderedHeight As Long ' cached height after word wrapping
    status As String ' "online", "idle", "dnd", "offline", "invisible"
End Type

Private Type TextElement
    ElementType As String ' "text", "emoji", "ping", "bold", "italic", "quote", "h1", "h2", "h3", "bullet", "small"
    Content As String
    X As Long
    Width As Long
    Picture As StdPicture ' for emojis
    IsBold As Boolean
    IsItalic As Boolean
End Type

Private Type userStatus
    UserId As String
    status As String ' "online", "idle", "dnd", "offline", "invisible"
End Type

Private TopIndex As Long
Private Messages() As ChatMessage
Private MessageCount As Long
Private UserStatuses() As userStatus ' Array to track user statuses
Private UserStatusCount As Long
Private Const BaseLineHeight As Long = 16
Private Const MessagePadding As Long = 4
Private EmojiCache As Collection ' Cache loaded emojis
Private IsUserScrolling As Boolean ' Track if user is manually scrolling
Private AutoScrollEnabled As Boolean ' Track if autoscroll should happen

' Status colors
Private Const StatusOnline As Long = vbGreen
Private Const StatusIdle As Long = vbYellow
Private Const StatusDnd As Long = vbRed
Private Const StatusOffline As Long = &H80848E ' Gray
Private Const StatusInvisible As Long = &H80848E ' Gray (same as offline)

Private Sub UserControl_Initialize()
    Set EmojiCache = New Collection
    IsUserScrolling = False
    AutoScrollEnabled = True
    UserStatusCount = 0
    ReDim UserStatuses(0)
End Sub

Public Sub AddMessage(ByVal UserName As String, ByVal UserId As String, ByVal Texts As String, Optional Avatar As StdPicture, Optional status As String = "offline")
    ' Check if we should autoscroll (only if user is at or near bottom)
    Dim ShouldAutoScroll As Boolean
    ShouldAutoScroll = AutoScrollEnabled And (vsbChat.value >= vsbChat.Max - (BaseLineHeight * 2))
    
    Dim Text() As String
    Text = Split(Texts, vbLf)
    Dim i As Integer
    For i = 0 To UBound(Text) Step 1
        ReDim Preserve Messages(MessageCount)
        Messages(MessageCount).UserName = UserName
        Messages(MessageCount).UserId = UserId
        Messages(MessageCount).Text = Text(i)
        Set Messages(MessageCount).Avatar = Avatar
        Messages(MessageCount).RenderedHeight = 0 ' Will be calculated during render
        Messages(MessageCount).status = status
        MessageCount = MessageCount + 1
    Next
    
    ' Update user status if not already tracked
    UpdateUserStatus UserId, status
    
    UpdateScrollbar
    
    ' Autoscroll to bottom if appropriate
    If ShouldAutoScroll Then
        ScrollToBottom
    End If
    
    UserControl.Refresh
End Sub

' Public method to get a user's current status
Public Function GetUserStatus(ByVal UserId As String) As String
    Dim i As Long
    For i = 0 To UserStatusCount - 1
        If UserStatuses(i).UserId = UserId Then
            GetUserStatus = UserStatuses(i).status
            Exit Function
        End If
    Next i
    GetUserStatus = "offline" ' Default if not found
End Function

Public Sub UpdateUserStatus(ByVal UserId As String, ByVal status As String)
    ' Check if user already exists in status array
    Dim i As Long
    For i = 0 To UserStatusCount - 1
        If UserStatuses(i).UserId = UserId Then
            UserStatuses(i).status = status
            Exit Sub
        End If
    Next i
    
    ' User not found, add new entry
    If UserStatusCount = 0 Then
        ReDim UserStatuses(0)
    Else
        ReDim Preserve UserStatuses(UserStatusCount)
    End If
    
    UserStatuses(UserStatusCount).UserId = UserId
    UserStatuses(UserStatusCount).status = status
    UserStatusCount = UserStatusCount + 1
    UserControl.Refresh
End Sub

Private Function GetStatusColor(ByVal status As String) As Long
    Select Case LCase(status)
        Case "online"
            GetStatusColor = StatusOnline
        Case "idle", "away"
            GetStatusColor = StatusIdle
        Case "dnd", "busy"
            GetStatusColor = StatusDnd
        Case "invisible"
            GetStatusColor = StatusInvisible
        Case Else ' "offline" or unknown
            GetStatusColor = StatusOffline
    End Select
End Function

Private Sub DrawStatusFrame(ByVal X As Long, ByVal Y As Long, ByVal Size As Long, ByVal status As String)
    Dim StatusColor As Long
    StatusColor = GetStatusColor(status)
    
    ' Don't draw frame for offline users (makes it cleaner)
    If LCase(status) = "offline" Then Exit Sub
    
    ' Draw outer frame (slightly larger than avatar)
    Dim FrameSize As Long
    FrameSize = Size + 3 ' 3 pixels border on each side
    Dim FrameX As Long, FrameY As Long
    FrameX = X - 2
    FrameY = Y - 2
    
    ' Draw the colored border as a thick rectangle outline
    UserControl.Line (FrameX, FrameY)-(FrameX + FrameSize, FrameY + 2), StatusColor, BF ' Top
    UserControl.Line (FrameX, FrameY)-(FrameX + 2, FrameY + FrameSize), StatusColor, BF ' Left
    UserControl.Line (FrameX + FrameSize - 2, FrameY)-(FrameX + FrameSize, FrameY + FrameSize), StatusColor, BF ' Right
    UserControl.Line (FrameX, FrameY + FrameSize - 2)-(FrameX + FrameSize, FrameY + FrameSize), StatusColor, BF ' Bottom
End Sub

Public Sub Clear()
    ReDim Messages(0)
    MessageCount = 0
    IsUserScrolling = False
    AutoScrollEnabled = True
    UserControl.Refresh
    vsbChat.Max = 0
End Sub

Private Sub ScrollToBottom()
    If vsbChat.Max > 0 Then
        vsbChat.value = vsbChat.Max
        TopIndex = vsbChat.value
    End If
End Sub

Private Sub UpdateScrollbar()
    Dim totalHeight As Long
    Dim i As Long
    
    For i = 0 To MessageCount - 1
        If Messages(i).RenderedHeight = 0 Then
            Messages(i).RenderedHeight = CalculateMessageHeight(i)
        End If
        totalHeight = totalHeight + Messages(i).RenderedHeight + MessagePadding
    Next i
    
    vsbChat.Max = totalHeight - UserControl.ScaleHeight
    If vsbChat.Max < 0 Then vsbChat.Max = 0
    vsbChat.LargeChange = UserControl.ScaleHeight \ 2
    vsbChat.SmallChange = BaseLineHeight
End Sub

Private Function CalculateMessageHeight(ByVal MessageIndex As Long) As Long
    Dim AvatarSize As Long: AvatarSize = 32
    Dim Margin As Long: Margin = 8
    Dim textWidth As Long
    Dim Elements As Collection
    Dim WrappedLines As Collection
    
    textWidth = UserControl.ScaleWidth - (Margin + AvatarSize + 16 + vsbChat.Width)
    Set Elements = ParseTextElements(Messages(MessageIndex).Text)
    Set WrappedLines = WrapTextElements(Elements, textWidth)
    
    Dim MinHeight As Long
    MinHeight = AvatarSize + 4  ' Reduced padding
    
    Dim textHeight As Long
    textHeight = 14 + (WrappedLines.Count * BaseLineHeight) ' Username (14px) + text lines
    
    ' Only use extra height if text actually wraps (more than 1 line) or has special formatting
    If WrappedLines.Count <= 1 And Not HasSpecialFormatting(Elements) Then
        CalculateMessageHeight = MinHeight
    Else
        CalculateMessageHeight = IIf(MinHeight > textHeight, MinHeight, textHeight)
    End If
End Function

Private Function HasSpecialFormatting(ByVal Elements As Collection) As Boolean
    Dim i As Long
    For i = 1 To Elements.Count
        Dim element As Collection
        Set element = Elements(i)
        Dim ElementType As String
        ElementType = element.Item(1) ' Get type from first item
        
        If ElementType <> "text" And ElementType <> "emoji" And ElementType <> "ping" Then
            HasSpecialFormatting = True
            Exit Function
        End If
    Next i
    HasSpecialFormatting = False
End Function

Private Function LoadEmoji(ByVal EmojiName As String) As StdPicture
'MsgBox EmojiName
    On Error GoTo ErrHandler
    
    ' Check cache first
    On Error Resume Next
    Set LoadEmoji = EmojiCache(EmojiName)
    If Not LoadEmoji Is Nothing Then Exit Function
    On Error GoTo ErrHandler
    
    ' Try to load emoji from file
    Dim EmojiPath As String
    EmojiPath = App.Path & "\emoji\msn\" & EmojiName & ".gif"
    
    If Dir(EmojiPath) <> "" Then
        Set LoadEmoji = LoadPicture(EmojiPath)
        ' Cache the loaded emoji
        EmojiCache.Add LoadEmoji, EmojiName
    Else
        ' Try .gif extension
        EmojiPath = App.Path & "\emoji\msn\" & EmojiName & ".gif"
        If Dir(EmojiPath) <> "" Then
            Set LoadEmoji = LoadPicture(EmojiPath)
            EmojiCache.Add LoadEmoji, EmojiName
        End If
    End If
    
    Exit Function
ErrHandler:
    Set LoadEmoji = Nothing
End Function

Private Function ParseTextElements(ByVal Text As String) As Collection
    Dim Elements As New Collection
    Dim i As Long
    Dim CurrentText As String
    Dim InEmoji As Boolean
    Dim InPing As Boolean
    Dim InBold As Boolean
    Dim InItalic As Boolean
    Dim EmojiName As String
    Dim PingContent As String
    
    ' Check for markdown prefixes first
    Dim TrimmedText As String
    TrimmedText = LTrim(Text)
    
    ' Handle quotes (>)
    If Left(TrimmedText, 1) = ">" Then
        Dim QuoteElement As New Collection
        QuoteElement.Add "quote" ' Type
        QuoteElement.Add Mid(TrimmedText, 2) ' Content
        Elements.Add QuoteElement
        Set ParseTextElements = Elements
        Exit Function
    End If
    
    ' Handle headers (###, ##, #) - check longest first
    If Left(TrimmedText, 3) = "###" Then
        Dim H3Element As New Collection
        H3Element.Add "h3" ' Type
        H3Element.Add Trim(Mid(TrimmedText, 4)) ' Content
        Elements.Add H3Element
        Set ParseTextElements = Elements
        Exit Function
    ElseIf Left(TrimmedText, 2) = "##" Then
        Dim H2Element As New Collection
        H2Element.Add "h2" ' Type
        H2Element.Add Trim(Mid(TrimmedText, 3)) ' Content
        Elements.Add H2Element
        Set ParseTextElements = Elements
        Exit Function
    ElseIf Left(TrimmedText, 1) = "#" Then
        Dim H1Element As New Collection
        H1Element.Add "h1" ' Type
        H1Element.Add Trim(Mid(TrimmedText, 2)) ' Content
        Elements.Add H1Element
        Set ParseTextElements = Elements
        Exit Function
    End If
    
    ' Handle bullet points (* )
    If Left(TrimmedText, 2) = "* " Then
        Dim BulletElement As New Collection
        BulletElement.Add "bullet" ' Type
        BulletElement.Add Trim(Mid(TrimmedText, 3)) ' Content
        Elements.Add BulletElement
        Set ParseTextElements = Elements
        Exit Function
    End If
    
    ' Handle small text (-#)
    If Left(TrimmedText, 2) = "-#" Then
        Dim SmallElement As New Collection
        SmallElement.Add "small" ' Type
        SmallElement.Add Trim(Mid(TrimmedText, 3)) ' Content
        Elements.Add SmallElement
        Set ParseTextElements = Elements
        Exit Function
    End If
    
    ' Parse inline formatting
    i = 1
    While i <= Len(Text)
        Dim Char As String
        Char = Mid(Text, i, 1)
        
        If Char = "*" And Not InEmoji And Not InPing Then
            ' Check for bold (**)
            If i < Len(Text) And Mid(Text, i + 1, 1) = "*" Then
                ' Bold formatting
                If InBold Then
                    ' End bold
                    If CurrentText <> "" Then
                        Dim BoldElement As New Collection
                        BoldElement.Add "bold" ' Type
                        BoldElement.Add CurrentText ' Content
                        Elements.Add BoldElement
                        CurrentText = ""
                    End If
                    InBold = False
                Else
                    ' Start bold - add any accumulated text first
                    If CurrentText <> "" Then
                        Dim TextElement As New Collection
                        TextElement.Add "text" ' Type
                        TextElement.Add CurrentText ' Content
                        Elements.Add TextElement
                        CurrentText = ""
                    End If
                    InBold = True
                End If
                i = i + 1 ' Skip the second *
            Else
                ' Italic formatting
                If InItalic Then
                    ' End italic
                    If CurrentText <> "" Then
                        Dim ItalicElement As New Collection
                        ItalicElement.Add "italic" ' Type
                        ItalicElement.Add CurrentText ' Content
                        Elements.Add ItalicElement
                        CurrentText = ""
                    End If
                    InItalic = False
                Else
                    ' Start italic - add any accumulated text first
                    If CurrentText <> "" Then
                        Dim TextElement2 As New Collection
                        TextElement2.Add "text" ' Type
                        TextElement2.Add CurrentText ' Content
                        Elements.Add TextElement2
                        CurrentText = ""
                    End If
                    InItalic = True
                End If
            End If
        ElseIf Char = ":" And Not InPing And Not InBold And Not InItalic Then
            ' Potential emoji start/end
            If InEmoji Then
                ' End of emoji
                Dim EmojiElement As New Collection
                EmojiElement.Add "emoji" ' Type
                EmojiElement.Add EmojiName ' Content
                Elements.Add EmojiElement
                EmojiName = ""
                InEmoji = False
            Else
                ' Start of emoji - add any accumulated text first
                If CurrentText <> "" Then
                    Dim TextElement3 As New Collection
                    TextElement3.Add "text" ' Type
                    TextElement3.Add CurrentText ' Content
                    Elements.Add TextElement3
                    CurrentText = ""
                End If
                InEmoji = True
            End If
        ElseIf Char = "<" And Mid(Text, i, 2) = "<@" And Not InEmoji And Not InBold And Not InItalic Then
            ' Start of ping
            If CurrentText <> "" Then
                Dim TextElement4 As New Collection
                TextElement4.Add "text" ' Type
                TextElement4.Add CurrentText ' Content
                Elements.Add TextElement4
                CurrentText = ""
            End If
            InPing = True
            i = i + 1 ' Skip the @
        ElseIf Char = ">" And InPing Then
            ' End of ping
            Dim PingElement As New Collection
            PingElement.Add "ping" ' Type
            PingElement.Add PingContent ' Content
            Elements.Add PingElement
            PingContent = ""
            InPing = False
        ElseIf InEmoji Then
            EmojiName = EmojiName & Char
        ElseIf InPing Then
            PingContent = PingContent & Char
        Else
            CurrentText = CurrentText & Char
        End If
        
        i = i + 1
    Wend
    
    ' Add any remaining text
    If CurrentText <> "" Then
        Dim FinalElement As New Collection
        If InBold Then
            FinalElement.Add "bold" ' Type
        ElseIf InItalic Then
            FinalElement.Add "italic" ' Type
        Else
            FinalElement.Add "text" ' Type
        End If
        FinalElement.Add CurrentText ' Content
        Elements.Add FinalElement
    End If
    
    ' Handle unclosed emoji
    If InEmoji And EmojiName <> "" Then
        Dim UnfinishedEmoji As New Collection
        UnfinishedEmoji.Add "text" ' Type
        UnfinishedEmoji.Add ":" & EmojiName ' Content
        Elements.Add UnfinishedEmoji
    End If
    
    ' Handle unclosed ping
    If InPing And PingContent <> "" Then
        Dim UnfinishedPing As New Collection
        UnfinishedPing.Add "text" ' Type
        UnfinishedPing.Add "<@" & PingContent ' Content
        Elements.Add UnfinishedPing
    End If
    
    Set ParseTextElements = Elements
End Function

Private Function WrapTextElements(ByVal Elements As Collection, ByVal MaxWidth As Long) As Collection
    Dim Lines As New Collection
    Dim CurrentLine As New Collection
    Dim CurrentWidth As Long
    Dim i As Long
    
    For i = 1 To Elements.Count
        Dim element As Collection
        Set element = Elements(i)
        
        Dim ElementType As String
        Dim Content As String
        ElementType = element.Item(1) ' Type is first item
        Content = element.Item(2) ' Content is second item
        
        Select Case ElementType
            Case "text", "bold", "italic"
                ' Handle word wrapping for text elements
                Dim words() As String
                words = Split(Content, " ")
                Dim j As Long
                
                For j = 0 To UBound(words)
                    Dim Word As String
                    Word = words(j)
                    If j < UBound(words) Then Word = Word & " "
                    
                    Dim WordWidth As Long
                    WordWidth = GetTextWidth(Word, ElementType)
                    
                    If CurrentWidth + WordWidth > MaxWidth And CurrentLine.Count > 0 Then
                        ' Start new line
                        Lines.Add CurrentLine
                        Set CurrentLine = New Collection
                        CurrentWidth = 0
                    End If
                    
                    ' Add word to current line - create new Collection for each word
                    Dim WordElement As Collection
                    Set WordElement = New Collection
                    WordElement.Add ElementType ' Type
                    WordElement.Add Word ' Content
                    CurrentLine.Add WordElement
                    CurrentWidth = CurrentWidth + WordWidth
                Next j
                
            Case Else ' emoji, ping, quote, h1, h2, h3, bullet, small
                Dim ElementWidth As Long
                ElementWidth = GetElementWidth(element)
                
                If CurrentWidth + ElementWidth > MaxWidth And CurrentLine.Count > 0 Then
                    ' Start new line
                    Lines.Add CurrentLine
                    Set CurrentLine = New Collection
                    CurrentWidth = 0
                End If
                
                CurrentLine.Add element
                CurrentWidth = CurrentWidth + ElementWidth
        End Select
    Next i
    
    ' Add the last line
    If CurrentLine.Count > 0 Then
        Lines.Add CurrentLine
    End If
    
    Set WrapTextElements = Lines
End Function

Private Function GetTextWidth(ByVal Text As String, ByVal ElementType As String) As Long
    Dim OldBold As Boolean
    Dim OldSize As Long
    
    OldBold = UserControl.FontBold
    OldSize = UserControl.FontSize
    
    Select Case ElementType
        Case "bold"
            UserControl.FontBold = True
        Case "h1"
            UserControl.FontSize = UserControl.FontSize + 4
            UserControl.FontBold = True
        Case "h2"
            UserControl.FontSize = UserControl.FontSize + 2
            UserControl.FontBold = True
        Case "h3"
            UserControl.FontBold = True
        Case "small"
            UserControl.FontSize = UserControl.FontSize - 2
    End Select
    
    GetTextWidth = UserControl.textWidth(Text)
    
    UserControl.FontBold = OldBold
    UserControl.FontSize = OldSize
End Function

Private Function GetElementWidth(ByVal element As Collection) As Long
    Dim ElementType As String
    Dim Content As String
    ElementType = element.Item(1) ' Type
    Content = element.Item(2) ' Content
    
    Select Case ElementType
        Case "emoji"
            GetElementWidth = 18
        Case "ping"
            GetElementWidth = UserControl.textWidth("@" & Content) + 4
        Case "quote"
            GetElementWidth = UserControl.textWidth("> " & Content)
        Case "bullet"
            GetElementWidth = UserControl.textWidth("• " & Content)
        Case "small"
            GetElementWidth = GetTextWidth(Content, ElementType)
        Case Else
            GetElementWidth = GetTextWidth(Content, ElementType)
    End Select
End Function

Private Sub DrawWrappedMessage(ByVal MessageIndex As Long, ByVal StartX As Long, ByVal StartY As Long)
    Dim Elements As Collection
    Dim WrappedLines As Collection
    Dim textWidth As Long
    
    textWidth = UserControl.ScaleWidth - (StartX + 8 + vsbChat.Width)
    Set Elements = ParseTextElements(Messages(MessageIndex).Text)
    Set WrappedLines = WrapTextElements(Elements, textWidth)
    
    Dim LineY As Long
    LineY = StartY
    Dim i As Long
    
    For i = 1 To WrappedLines.Count
        Dim Line As Collection
        Set Line = WrappedLines(i)
        
        Call DrawParsedLine(Line, StartX, LineY)
        LineY = LineY + BaseLineHeight
    Next i
End Sub

Private Sub DrawParsedLine(ByVal Line As Collection, ByVal StartX As Long, ByVal Y As Long)
    Dim CurrentX As Long
    CurrentX = StartX
    Dim i As Long
    
    For i = 1 To Line.Count
        Dim element As Collection
        Set element = Line(i)
        
        Dim ElementType As String
        Dim Content As String
        ElementType = element.Item(1) ' Type
        Content = element.Item(2) ' Content
        
        Select Case ElementType
            Case "text"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y - 2
                UserControl.ForeColor = vbBlack
                UserControl.FontBold = False
                UserControl.Print Content;
                CurrentX = UserControl.CurrentX
                
            Case "bold"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y - 2
                UserControl.ForeColor = vbBlack
                UserControl.FontBold = True
                UserControl.Print Content;
                UserControl.FontBold = False
                CurrentX = UserControl.CurrentX
                
            Case "italic"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y - 4
                UserControl.ForeColor = vbBlack
                UserControl.FontItalic = True
                UserControl.Print Content;
                UserControl.FontItalic = False
                CurrentX = UserControl.CurrentX
                
            Case "quote"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y - 2
                UserControl.ForeColor = &H808080
                UserControl.Print "> " & Content;
                CurrentX = UserControl.CurrentX
                
            Case "h1"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y - 2
                UserControl.ForeColor = vbBlack
                UserControl.FontBold = True
                UserControl.FontSize = UserControl.FontSize + 4
                UserControl.Print Content;
                UserControl.FontSize = UserControl.FontSize - 4
                UserControl.FontBold = False
                CurrentX = UserControl.CurrentX
                
            Case "h2"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y - 2
                UserControl.ForeColor = vbBlack
                UserControl.FontBold = True
                UserControl.FontSize = UserControl.FontSize + 2
                UserControl.Print Content;
                UserControl.FontSize = UserControl.FontSize - 2
                UserControl.FontBold = False
                CurrentX = UserControl.CurrentX
                
            Case "h3"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y - 2
                UserControl.ForeColor = vbBlack
                UserControl.FontBold = True
                UserControl.Print Content;
                UserControl.FontBold = False
                CurrentX = UserControl.CurrentX
                
            Case "small"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y - 2
                UserControl.ForeColor = vbBlack
                UserControl.FontSize = UserControl.FontSize - 2
                UserControl.Print Content;
                UserControl.FontSize = UserControl.FontSize + 2
                CurrentX = UserControl.CurrentX
                
            Case "bullet"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y - 2
                UserControl.ForeColor = vbBlack
                UserControl.Print "• " & Content;
                CurrentX = UserControl.CurrentX
                
            Case "emoji"
                Dim EmojiPic As StdPicture
                Set EmojiPic = LoadEmoji(Content)
                If Not EmojiPic Is Nothing Then
                    UserControl.PaintPicture EmojiPic, CurrentX, Y, 19, 19
                    CurrentX = CurrentX + 18
                Else
                    ' Fallback to text if emoji not found
                    UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y - 2
                    UserControl.ForeColor = vbBlack
                    UserControl.Print ":" & Content & ":";
                    CurrentX = UserControl.CurrentX
                End If
                
            Case "ping"
                ' Draw ping with system highlight color
                Dim PingText As String
                PingText = "@" & Content
                
                Dim PingWidth As Long
                PingWidth = UserControl.textWidth(PingText)
                
                ' Draw background with system highlight color
                UserControl.Line (CurrentX - 2, Y - 1)-(CurrentX + PingWidth + 2, Y + 13), &H8000000D, BF
                
                ' Draw text with system highlight text color
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y - 2
                UserControl.ForeColor = &H8000000E
                UserControl.FontBold = True
                UserControl.Print PingText;
                UserControl.FontBold = False
                CurrentX = UserControl.CurrentX + 2
        End Select
    Next i
End Sub

Private Sub UserControl_Resize()
    vsbChat.Move UserControl.ScaleWidth - vsbChat.Width, 0, vsbChat.Width, UserControl.ScaleHeight
    
    ' Remember current scroll position relative to bottom
    Dim WasAtBottom As Boolean
    WasAtBottom = (vsbChat.value >= vsbChat.Max - (BaseLineHeight * 2))
    
    ' Recalculate message heights since width changed
    Dim i As Long
    For i = 0 To MessageCount - 1
        Messages(i).RenderedHeight = 0
    Next i
    
    UpdateScrollbar
    
    ' If user was at bottom before resize, keep them at bottom
    If WasAtBottom Then
        ScrollToBottom
    End If
    
    UserControl.Refresh
End Sub

Private Sub vsbChat_Change()
    TopIndex = vsbChat.value
    UserControl.Refresh
    
    ' Check if user manually scrolled away from bottom
    If vsbChat.value < vsbChat.Max - (BaseLineHeight * 2) Then
        AutoScrollEnabled = False
    Else
        AutoScrollEnabled = True
    End If
End Sub

Private Sub vsbChat_Scroll()
    TopIndex = vsbChat.value
    UserControl.Refresh
    
    ' Check if user manually scrolled away from bottom
    If vsbChat.value < vsbChat.Max - (BaseLineHeight * 2) Then
        AutoScrollEnabled = False
    Else
        AutoScrollEnabled = True
    End If
End Sub
Private Sub UserControl_Paint()
    
    Dim i As Long
    Dim CurrentY As Long
    Dim AvatarSize As Long: AvatarSize = 32
    Dim Margin As Long: Margin = 8
    
    UserControl.Cls
    CurrentY = -TopIndex + Margin - 8
    
    For i = 0 To MessageCount - 1
        If Messages(i).RenderedHeight = 0 Then
            Messages(i).RenderedHeight = CalculateMessageHeight(i)
        End If
        
        If CurrentY + Messages(i).RenderedHeight > 0 And CurrentY < UserControl.ScaleHeight Then
            ' Draw status frame first (behind avatar)
            Call DrawStatusFrame(Margin, CurrentY, AvatarSize, GetUserStatus(Messages(i).UserId))
            
            ' Draw avatar
            If Not Messages(i).Avatar Is Nothing Then
            On Error Resume Next
                UserControl.PaintPicture Messages(i).Avatar, Margin, CurrentY, AvatarSize, AvatarSize
            Else
                UserControl.Line (Margin, CurrentY)-(Margin + AvatarSize, CurrentY + AvatarSize), vbGrayText, BF
            End If
            
            ' Draw username
            UserControl.CurrentX = Margin + AvatarSize + 8
            UserControl.CurrentY = CurrentY
            UserControl.ForeColor = vbBlack
            UserControl.FontBold = True
            UserControl.Print Messages(i).UserName
            
            ' Draw wrapped message text with markdown
            UserControl.FontBold = False
            Call DrawWrappedMessage(i, Margin + AvatarSize + 10, CurrentY + 16)
        End If
        
        CurrentY = CurrentY + Messages(i).RenderedHeight + MessagePadding
    Next i
    
End Sub

' Public method to manually scroll to bottom (useful for "scroll to bottom" button)
Public Sub ForceScrollToBottom()
    AutoScrollEnabled = True
    ScrollToBottom
    UserControl.Refresh
End Sub

' Public method to check if autoscroll is enabled
Public Property Get IsAutoScrollEnabled() As Boolean
    IsAutoScrollEnabled = AutoScrollEnabled
End Property

Public Function GetTrackedUserIds() As String()
    Dim UserIds() As String
    If UserStatusCount = 0 Then
        ReDim UserIds(0)
        GetTrackedUserIds = UserIds
        Exit Function
    End If
    
    ReDim UserIds(UserStatusCount - 1)
    Dim i As Long
    For i = 0 To UserStatusCount - 1
        UserIds(i) = UserStatuses(i).UserId
    Next i
    GetTrackedUserIds = UserIds
End Function

Public Function GetStatusCounts() As String
    Dim OnlineCount As Long, IdleCount As Long, DndCount As Long, OfflineCount As Long
    Dim i As Long
    
    For i = 0 To UserStatusCount - 1
        Select Case LCase(UserStatuses(i).status)
            Case "online"
                OnlineCount = OnlineCount + 1
            Case "idle", "away"
                IdleCount = IdleCount + 1
            Case "dnd", "busy"
                DndCount = DndCount + 1
            Case Else
                OfflineCount = OfflineCount + 1
        End Select
    Next i
    
    GetStatusCounts = "Online: " & OnlineCount & ", Idle: " & IdleCount & ", DND: " & DndCount & ", Offline: " & OfflineCount
End Function
