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
Private Type ChatMessage
    Username As String
    Text As String
    Avatar As StdPicture ' optional
    RenderedHeight As Long ' cached height after word wrapping
End Type

Private Type TextElement
    ElementType As String ' "text", "emoji", "ping", "bold", "italic", "quote", "h1", "h2", "bullet"
    Content As String
    X As Long
    Width As Long
    Picture As StdPicture ' for emojis
    IsBold As Boolean
    IsItalic As Boolean
End Type

Private TopIndex As Long
Private Messages() As ChatMessage
Private MessageCount As Long
Private Const BaseLineHeight As Long = 16
Private Const MessagePadding As Long = 4
Private EmojiCache As Collection ' Cache loaded emojis

Private Sub UserControl_Initialize()
    Set EmojiCache = New Collection
End Sub

Public Sub AddMessage(ByVal Username As String, ByVal Texts As String, Optional Avatar As StdPicture)
    Dim Text() As String
    Text = Split(Texts, vbLf)
    Dim i As Integer
    For i = 0 To UBound(Text) Step 1
        ReDim Preserve Messages(MessageCount)
        Messages(MessageCount).Username = Username
        Messages(MessageCount).Text = Text(i)
        Set Messages(MessageCount).Avatar = Avatar
        Messages(MessageCount).RenderedHeight = 0 ' Will be calculated during render
        MessageCount = MessageCount + 1
    Next
    UserControl.Refresh
    UpdateScrollbar
End Sub

Public Sub Clear()
    ReDim Messages(0)
    MessageCount = 0
    UserControl.Refresh
    vsbChat.Max = 0
End Sub

Private Sub UpdateScrollbar()
    Dim TotalHeight As Long
    Dim i As Long
    
    For i = 0 To MessageCount - 1
        If Messages(i).RenderedHeight = 0 Then
            Messages(i).RenderedHeight = CalculateMessageHeight(i)
        End If
        TotalHeight = TotalHeight + Messages(i).RenderedHeight + MessagePadding
    Next i
    
    vsbChat.Max = TotalHeight - UserControl.ScaleHeight
    If vsbChat.Max < 0 Then vsbChat.Max = 0
    vsbChat.LargeChange = UserControl.ScaleHeight \ 2
    vsbChat.SmallChange = BaseLineHeight
End Sub

Private Function CalculateMessageHeight(ByVal MessageIndex As Long) As Long
    Dim AvatarSize As Long: AvatarSize = 32
    Dim Margin As Long: Margin = 8
    Dim TextWidth As Long
    Dim Elements As Collection
    Dim WrappedLines As Collection
    
    TextWidth = UserControl.ScaleWidth - (Margin + AvatarSize + 16 + vsbChat.Width)
    Set Elements = ParseTextElements(Messages(MessageIndex).Text)
    Set WrappedLines = WrapTextElements(Elements, TextWidth)
    
    Dim MinHeight As Long
    MinHeight = AvatarSize + 4  ' Reduced padding
    
    Dim TextHeight As Long
    TextHeight = 14 + (WrappedLines.Count * BaseLineHeight) ' Username (14px) + text lines
    
    ' Only use extra height if text actually wraps (more than 1 line) or has special formatting
    If WrappedLines.Count <= 1 And Not HasSpecialFormatting(Elements) Then
        CalculateMessageHeight = MinHeight
    Else
        CalculateMessageHeight = IIf(MinHeight > TextHeight, MinHeight, TextHeight)
    End If
End Function

Private Function HasSpecialFormatting(ByVal Elements As Collection) As Boolean
    Dim i As Long
    For i = 1 To Elements.Count
        Dim Element As Collection
        Set Element = Elements(i)
        Dim ElementType As String
        ElementType = Element("type")
        
        If ElementType <> "text" And ElementType <> "emoji" And ElementType <> "ping" Then
            HasSpecialFormatting = True
            Exit Function
        End If
    Next i
    HasSpecialFormatting = False
End Function

Private Function LoadEmoji(ByVal EmojiName As String) As StdPicture
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
        QuoteElement.Add "quote", "type"
        QuoteElement.Add Mid(TrimmedText, 2), "content"
        Elements.Add QuoteElement
        Set ParseTextElements = Elements
        Exit Function
    End If
    
    ' Handle headers (# and ##)
    If Left(TrimmedText, 2) = "##" Then
        Dim H2Element As New Collection
        H2Element.Add "h2", "type"
        H2Element.Add Trim(Mid(TrimmedText, 3)), "content"
        Elements.Add H2Element
        Set ParseTextElements = Elements
        Exit Function
    ElseIf Left(TrimmedText, 1) = "#" Then
        Dim H1Element As New Collection
        H1Element.Add "h1", "type"
        H1Element.Add Trim(Mid(TrimmedText, 2)), "content"
        Elements.Add H1Element
        Set ParseTextElements = Elements
        Exit Function
    End If
    
    ' Handle bullet points (* )
    If Left(TrimmedText, 2) = "* " Then
        Dim BulletElement As New Collection
        BulletElement.Add "bullet", "type"
        BulletElement.Add Trim(Mid(TrimmedText, 3)), "content"
        Elements.Add BulletElement
        Set ParseTextElements = Elements
        Exit Function
    End If
    
    ' Handle small headers (-#)
    If Left(TrimmedText, 2) = "-#" Then
        Dim SmallHeaderElement As New Collection
        SmallHeaderElement.Add "h3", "type"
        SmallHeaderElement.Add Trim(Mid(TrimmedText, 3)), "content"
        Elements.Add SmallHeaderElement
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
                        BoldElement.Add "bold", "type"
                        BoldElement.Add CurrentText, "content"
                        Elements.Add BoldElement
                        CurrentText = ""
                    End If
                    InBold = False
                Else
                    ' Start bold - add any accumulated text first
                    If CurrentText <> "" Then
                        Dim TextElement As New Collection
                        TextElement.Add "text", "type"
                        TextElement.Add CurrentText, "content"
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
                        ItalicElement.Add "italic", "type"
                        ItalicElement.Add CurrentText, "content"
                        Elements.Add ItalicElement
                        CurrentText = ""
                    End If
                    InItalic = False
                Else
                    ' Start italic - add any accumulated text first
                    If CurrentText <> "" Then
                        Dim TextElement2 As New Collection
                        TextElement2.Add "text", "type"
                        TextElement2.Add CurrentText, "content"
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
                Dim Element As New Collection
                Element.Add "emoji", "type"
                Element.Add EmojiName, "content"
                Elements.Add Element
                EmojiName = ""
                InEmoji = False
            Else
                ' Start of emoji - add any accumulated text first
                If CurrentText <> "" Then
                    Dim TextElement3 As New Collection
                    TextElement3.Add "text", "type"
                    TextElement3.Add CurrentText, "content"
                    Elements.Add TextElement3
                    CurrentText = ""
                End If
                InEmoji = True
            End If
        ElseIf Char = "<" And Mid(Text, i, 2) = "<@" And Not InEmoji And Not InBold And Not InItalic Then
            ' Start of ping
            If CurrentText <> "" Then
                Dim TextElement4 As New Collection
                TextElement4.Add "text", "type"
                TextElement4.Add CurrentText, "content"
                Elements.Add TextElement4
                CurrentText = ""
            End If
            InPing = True
            i = i + 1 ' Skip the @
        ElseIf Char = ">" And InPing Then
            ' End of ping
            Dim PingElement As New Collection
            PingElement.Add "ping", "type"
            PingElement.Add PingContent, "content"
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
            FinalElement.Add "bold", "type"
        ElseIf InItalic Then
            FinalElement.Add "italic", "type"
        Else
            FinalElement.Add "text", "type"
        End If
        FinalElement.Add CurrentText, "content"
        Elements.Add FinalElement
    End If
    
    ' Handle unclosed emoji
    If InEmoji And EmojiName <> "" Then
        Dim UnfinishedEmoji As New Collection
        UnfinishedEmoji.Add "text", "type"
        UnfinishedEmoji.Add ":" & EmojiName, "content"
        Elements.Add UnfinishedEmoji
    End If
    
    ' Handle unclosed ping
    If InPing And PingContent <> "" Then
        Dim UnfinishedPing As New Collection
        UnfinishedPing.Add "text", "type"
        UnfinishedPing.Add "<@" & PingContent, "content"
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
        Dim Element As Collection
        Set Element = Elements(i)
        
        Dim ElementType As String
        Dim Content As String
        ElementType = Element("type")
        Content = Element("content")
        
        Select Case ElementType
            Case "text", "bold", "italic"
                ' Handle word wrapping for text elements
                Dim Words() As String
                Words = Split(Content, " ")
                Dim j As Long
                
                For j = 0 To UBound(Words)
                    Dim Word As String
                    Word = Words(j)
                    If j < UBound(Words) Then Word = Word & " "
                    
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
                    WordElement.Add ElementType, "type"
                    WordElement.Add Word, "content"
                    CurrentLine.Add WordElement
                    CurrentWidth = CurrentWidth + WordWidth
                Next j
                
            Case Else ' emoji, ping, quote, h1, h2, bullet
                Dim ElementWidth As Long
                ElementWidth = GetElementWidth(Element)
                
                If CurrentWidth + ElementWidth > MaxWidth And CurrentLine.Count > 0 Then
                    ' Start new line
                    Lines.Add CurrentLine
                    Set CurrentLine = New Collection
                    CurrentWidth = 0
                End If
                
                CurrentLine.Add Element
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
            UserControl.FontSize = UserControl.FontSize + 1
            UserControl.FontBold = True
    End Select
    
    GetTextWidth = UserControl.TextWidth(Text)
    
    UserControl.FontBold = OldBold
    UserControl.FontSize = OldSize
End Function

Private Function GetElementWidth(ByVal Element As Collection) As Long
    Dim ElementType As String
    Dim Content As String
    ElementType = Element("type")
    Content = Element("content")
    
    Select Case ElementType
        Case "emoji"
            GetElementWidth = 18
        Case "ping"
            GetElementWidth = UserControl.TextWidth("@" & Content) + 4
        Case "quote"
            GetElementWidth = UserControl.TextWidth("> " & Content)
        Case "bullet"
            GetElementWidth = UserControl.TextWidth("• " & Content)
        Case "h3"
            GetElementWidth = GetTextWidth(Content, ElementType)
        Case Else
            GetElementWidth = GetTextWidth(Content, ElementType)
    End Select
End Function

Private Sub DrawWrappedMessage(ByVal MessageIndex As Long, ByVal StartX As Long, ByVal StartY As Long)
    Dim Elements As Collection
    Dim WrappedLines As Collection
    Dim TextWidth As Long
    
    TextWidth = UserControl.ScaleWidth - (StartX + 8 + vsbChat.Width)
    Set Elements = ParseTextElements(Messages(MessageIndex).Text)
    Set WrappedLines = WrapTextElements(Elements, TextWidth)
    
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
        Dim Element As Collection
        Set Element = Line(i)
        
        Dim ElementType As String
        Dim Content As String
        ElementType = Element("type")
        Content = Element("content")
        
        Select Case ElementType
            Case "text"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y
                UserControl.ForeColor = vbBlack
                UserControl.FontBold = False
                UserControl.Print Content;
                CurrentX = UserControl.CurrentX
                
            Case "bold"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y
                UserControl.ForeColor = vbBlack
                UserControl.FontBold = True
                UserControl.Print Content;
                UserControl.FontBold = False
                CurrentX = UserControl.CurrentX
                
            Case "italic"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y
                UserControl.ForeColor = vbBlack
                UserControl.FontItalic = True
                UserControl.Print Content;
                UserControl.FontItalic = False
                CurrentX = UserControl.CurrentX
                
            Case "quote"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y
                UserControl.ForeColor = &H808080
                UserControl.Print "> " & Content;
                CurrentX = UserControl.CurrentX
                
            Case "h1"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y
                UserControl.ForeColor = vbBlack
                UserControl.FontBold = True
                UserControl.FontSize = UserControl.FontSize + 4
                UserControl.Print Content;
                UserControl.FontSize = UserControl.FontSize - 4
                UserControl.FontBold = False
                CurrentX = UserControl.CurrentX
                
            Case "h2"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y
                UserControl.ForeColor = vbBlack
                UserControl.FontBold = True
                UserControl.FontSize = UserControl.FontSize + 2
                UserControl.Print Content;
                UserControl.FontSize = UserControl.FontSize - 2
                UserControl.FontBold = False
                CurrentX = UserControl.CurrentX
                
            Case "h3"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y
                UserControl.ForeColor = vbBlack
                UserControl.FontBold = True
                UserControl.FontSize = UserControl.FontSize + 1
                UserControl.Print Content;
                UserControl.FontSize = UserControl.FontSize - 1
                UserControl.FontBold = False
                CurrentX = UserControl.CurrentX
                
            Case "bullet"
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y
                UserControl.ForeColor = vbBlack
                UserControl.Print "• " & Content;
                CurrentX = UserControl.CurrentX
                
            Case "emoji"
                Dim EmojiPic As StdPicture
                Set EmojiPic = LoadEmoji(Content)
                If Not EmojiPic Is Nothing Then
                    UserControl.PaintPicture EmojiPic, CurrentX, Y, 16, 16
                    CurrentX = CurrentX + 18
                Else
                    ' Fallback to text if emoji not found
                    UserControl.CurrentX = CurrentX
                    UserControl.CurrentY = Y
                    UserControl.ForeColor = vbBlack
                    UserControl.Print ":" & Content & ":";
                    CurrentX = UserControl.CurrentX
                End If
                
            Case "ping"
                ' Draw ping with system highlight color
                Dim PingText As String
                PingText = "@" & Content
                
                Dim PingWidth As Long
                PingWidth = UserControl.TextWidth(PingText)
                
                ' Draw background with system highlight color
                UserControl.Line (CurrentX - 2, Y - 1)-(CurrentX + PingWidth + 2, Y + 13), &H8000000D, BF
                
                ' Draw text with system highlight text color
                UserControl.CurrentX = CurrentX
                UserControl.CurrentY = Y
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
    
    ' Recalculate message heights since width changed
    Dim i As Long
    For i = 0 To MessageCount - 1
        Messages(i).RenderedHeight = 0
    Next i
    
    UpdateScrollbar
    UserControl.Refresh
End Sub

Private Sub vsbChat_Change()
    TopIndex = vsbChat.Value
    UserControl.Refresh
End Sub

Private Sub vsbChat_Scroll()
    TopIndex = vsbChat.Value
    UserControl.Refresh
End Sub

Private Sub UserControl_Paint()
    Dim i As Long
    Dim CurrentY As Long
    Dim AvatarSize As Long: AvatarSize = 32
    Dim Margin As Long: Margin = 8
    
    UserControl.Cls
    CurrentY = -TopIndex + Margin
    
    For i = 0 To MessageCount - 1
        If Messages(i).RenderedHeight = 0 Then
            Messages(i).RenderedHeight = CalculateMessageHeight(i)
        End If
        
        If CurrentY + Messages(i).RenderedHeight > 0 And CurrentY < UserControl.ScaleHeight Then
            ' Draw avatar
            If Not Messages(i).Avatar Is Nothing Then
                UserControl.PaintPicture Messages(i).Avatar, Margin, CurrentY, AvatarSize, AvatarSize
            Else
                UserControl.Line (Margin, CurrentY)-(Margin + AvatarSize, CurrentY + AvatarSize), vbGrayText, BF
            End If
            
            ' Draw username
            UserControl.CurrentX = Margin + AvatarSize + 8
            UserControl.CurrentY = CurrentY
            UserControl.ForeColor = vbBlack
            UserControl.FontBold = True
            UserControl.Print Messages(i).Username
            
            ' Draw wrapped message text with markdown
            UserControl.FontBold = False
            Call DrawWrappedMessage(i, Margin + AvatarSize + 8, CurrentY + 14)
        End If
        
        CurrentY = CurrentY + Messages(i).RenderedHeight + MessagePadding
    Next i
End Sub
