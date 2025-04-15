VERSION 5.00
Begin VB.UserControl GuildView 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1320
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   88
   Begin VB.VScrollBar VScroll1 
      Height          =   3615
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "GuildView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Events
Event GuildSelected(ByVal Index As Long)

' Guild structure
Private Type Guild
    Name As String
    icon As StdPicture
End Type

' Properties
Private Guilds() As Guild
Private GuildCount As Long
Private SelectedIndex As Long
Private HoverIndex As Long
Private Const IconSize As Long = 48
Private Const Padding As Long = 10

' Add a guild to the list
Public Sub AddGuild(ByVal Name As String, icon As StdPicture)
    ReDim Preserve Guilds(GuildCount)
    Guilds(GuildCount).Name = Name
    Set Guilds(GuildCount).icon = icon
    GuildCount = GuildCount + 1
    UpdateScrollSettings
    UserControl.Refresh
End Sub

' Clear all guilds
Public Sub ClearGuilds()
    Erase Guilds
    GuildCount = 0
    SelectedIndex = -1
    HoverIndex = -1
    VScroll1.Value = 0
    UserControl.Refresh
End Sub

' Update a guild's icon
Public Sub UpdateGuildIcon(ByVal Index As Long, icon As StdPicture)
    If Index >= 0 And Index < GuildCount Then
        Set Guilds(Index).icon = icon
        UserControl.Refresh
    End If
End Sub

' Get a guild name
Public Function GetGuildName(ByVal Index As Long) As String
    If Index >= 0 And Index < GuildCount Then
        GetGuildName = Guilds(Index).Name
    Else
        GetGuildName = ""
    End If
End Function

' Get the selected index
Public Property Get SelectedGuildIndex() As Long
    SelectedGuildIndex = SelectedIndex
End Property

' Set the selected index
Public Property Let SelectedGuildIndex(ByVal Value As Long)
    If Value >= -1 And Value < GuildCount Then
        SelectedIndex = Value
        EnsureVisible SelectedIndex
        UserControl.Refresh
    End If
End Property

' Ensure a guild is visible in the viewport
Private Sub EnsureVisible(ByVal Index As Long)
    If Index < 0 Or Index >= GuildCount Then Exit Sub
    
    Dim itemTop As Long
    itemTop = Index * (IconSize + Padding)
    
    Dim itemBottom As Long
    itemBottom = itemTop + IconSize
    
    Dim viewportHeight As Long
    viewportHeight = UserControl.ScaleHeight - IIf(VScroll1.Visible, VScroll1.Width, 0)
    
    ' If item is above the viewport
    If itemTop < VScroll1.Value Then
        VScroll1.Value = itemTop
    ' If item is below the viewport
    ElseIf itemBottom > VScroll1.Value + viewportHeight Then
        VScroll1.Value = itemBottom - viewportHeight
    End If
    
    UpdateScrollSettings
    UserControl.Refresh
End Sub

' Update scrollbar settings
Private Sub UpdateScrollSettings()
    Dim totalHeight As Long
    totalHeight = GuildCount * (IconSize + Padding) + Padding
    
    Dim viewportHeight As Long
    viewportHeight = UserControl.ScaleHeight
    
    ' Configure the scrollbar
    With VScroll1
        .Min = 0
        .Max = totalHeight - viewportHeight
        If .Max < 0 Then .Max = 0
        .SmallChange = IconSize + Padding
        .LargeChange = viewportHeight \ 2
        .Visible = (totalHeight > viewportHeight)
        
        ' Adjust width to account for scrollbar visibility
        If .Visible Then
            .Left = UserControl.ScaleWidth - .Width
            .Height = UserControl.ScaleHeight
        End If
    End With
End Sub

' Mouse down event - select a guild
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If HoverIndex >= 0 Then
        SelectedIndex = HoverIndex
        RaiseEvent GuildSelected(SelectedIndex)
        UserControl.Refresh
    End If
End Sub

' Mouse move event - highlight guild under mouse
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long, yPos As Long
    Dim oldHoverIndex As Long
    
    oldHoverIndex = HoverIndex
    yPos = Padding - VScroll1.Value
    HoverIndex = -1
    
    For i = 0 To GuildCount - 1
        If Y >= yPos And Y <= yPos + IconSize Then
            HoverIndex = i
            Exit For
        End If
        yPos = yPos + IconSize + Padding
    Next i
    
    ' Only refresh if hover index changed
    If oldHoverIndex <> HoverIndex Then
        UserControl.Refresh
    End If
End Sub

' Paint event - draw the guilds
Private Sub UserControl_Paint()
    Dim i As Long
    Dim Y As Long
    
    ' Clear the control
    UserControl.Cls
    
    ' Calculate visible area accounting for scrollbar
    Dim paintWidth As Long
    paintWidth = UserControl.ScaleWidth - IIf(VScroll1.Visible, VScroll1.Width, 0)
    
    ' Draw guilds
    Y = Padding - VScroll1.Value
    For i = 0 To GuildCount - 1
        ' Only draw if visible
        If Y + IconSize > 0 And Y < UserControl.ScaleHeight Then
            Dim bgColor As Long
            
            ' Determine background color based on state
            If i = SelectedIndex Then
                bgColor = RGB(88, 101, 242) ' Selected - Discord blue
            ElseIf i = HoverIndex Then
                bgColor = RGB(54, 57, 63) ' Hover - Discord dark gray
            Else
                bgColor = BackColor
            End If
            
            ' Draw background
            UserControl.Line (Padding, Y)-(Padding + IconSize, Y + IconSize), bgColor, BF
            
            ' Draw icon
            If Not Guilds(i).icon Is Nothing Then
            On Error Resume Next
                UserControl.PaintPicture Guilds(i).icon, Padding, Y, IconSize, IconSize
            Else
                ' Draw placeholder if no icon
                UserControl.Circle (Padding + IconSize / 2, Y + IconSize / 2), IconSize / 3, RGB(128, 128, 128)
            End If
        End If
        
        Y = Y + IconSize + Padding
    Next i
End Sub

' Mouse leave event - clear hover
Private Sub UserControl_MouseExit()
    HoverIndex = -1
    UserControl.Refresh
End Sub

' Initialize event
Private Sub UserControl_Initialize()
    SelectedIndex = -1
    HoverIndex = -1
    GuildCount = 0
    ReDim Guilds(0)
    'BackColor = RGB(32, 34, 37) ' Discord dark theme
    
    ' Initialize scrollbar
    With VScroll1
        .Min = 0
        .Max = 0
        .Value = 0
        .Visible = False
    End With
End Sub

' Resize event
Private Sub UserControl_Resize()
    ' Ensure minimum width
    If UserControl.Width < IconSize + (Padding * 2) + IIf(VScroll1.Visible, VScroll1.Width, 0) Then
        UserControl.Width = IconSize + (Padding * 2) + IIf(VScroll1.Visible, VScroll1.Width, 0)
    End If
    
    UpdateScrollSettings
    UserControl.Refresh
End Sub

' Scrollbar change event
Private Sub VScroll1_Change()
    UserControl.Refresh
End Sub

' Scrollbar scroll event (for smoother dragging)
Private Sub VScroll1_Scroll()
    UserControl.Refresh
End Sub

' Mouse wheel event - handle scrolling
Private Sub UserControl_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal X As Single, ByVal Y As Single)
    If Not VScroll1.Visible Then Exit Sub
    
    ' Adjust scroll position based on wheel rotation
    VScroll1.Value = VScroll1.Value - (Rotation / 120) * (IconSize + Padding) * 3 ' Scroll 3 items at a time
    
    ' Clamp scroll position (handled automatically by VScroll1)
    UserControl.Refresh
End Sub
