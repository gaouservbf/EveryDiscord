VERSION 5.00
Begin VB.UserControl ChatView 
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
End Type
Private TopIndex As Long

Private Messages() As ChatMessage
Private MessageCount As Long
Private Const LineHeight As Long = 48
Public Sub AddMessage(ByVal Username As String, ByVal Texts As String, Optional Avatar As StdPicture)
Dim Text() As String
Text = Split(Texts, vbLf)
Dim i As Integer
For i = 0 To UBound(Text) Step 1
    ReDim Preserve Messages(MessageCount)
    Messages(MessageCount).Username = Username
    Messages(MessageCount).Text = Text(i)
    Set Messages(MessageCount).Avatar = Avatar
    MessageCount = MessageCount + 1
    UserControl.Refresh
    vsbChat.Max = MessageCount - (UserControl.ScaleHeight \ LineHeight)
If vsbChat.Max < 0 Then vsbChat.Max = 0
vsbChat.LargeChange = (UserControl.ScaleHeight \ LineHeight)
vsbChat.SmallChange = 1
Next
End Sub
Public Sub Clear()
    ReDim Preserve Messages(MessageCount)
    Messages(MessageCount).Username = ""
    Messages(MessageCount).Text = ""
    Set Messages(MessageCount).Avatar = Nothing
    MessageCount = 0
    UserControl.Refresh
End Sub
Private Sub UserControl_Resize()
    vsbChat.Move UserControl.ScaleWidth - vsbChat.Width, 0, vsbChat.Width, UserControl.ScaleHeight
    vsbChat.Max = MessageCount - (UserControl.ScaleHeight \ LineHeight)
    If vsbChat.Max < 0 Then vsbChat.Max = 0
    vsbChat.LargeChange = (UserControl.ScaleHeight \ LineHeight)
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
    Dim Y As Long
    Dim AvatarSize As Long: AvatarSize = 32
    Dim Margin As Long: Margin = 8

UserControl.Cls

    For i = 0 To MessageCount - 1
       Y = (i - TopIndex) * LineHeight + Margin

If Y + LineHeight > 0 And Y < UserControl.ScaleHeight Then
    ' Draw avatar, username, text
    If Not Messages(i).Avatar Is Nothing Then
        UserControl.PaintPicture Messages(i).Avatar, Margin, Y, AvatarSize, AvatarSize
    Else
        UserControl.Line (Margin, Y)-(Margin + AvatarSize, Y + AvatarSize), vbGrayText, BF
    End If

        UserControl.FontBold = True
    UserControl.CurrentX = Margin + AvatarSize + 8
    UserControl.CurrentY = Y
    UserControl.ForeColor = vbBlack
    UserControl.Print Messages(i).Username

        UserControl.FontBold = False
    UserControl.CurrentX = Margin + AvatarSize + 8
    UserControl.CurrentY = Y + 14
    UserControl.ForeColor = vbBlack
    UserControl.Print Messages(i).Text
End If

        ' Draw avatar
        If Not Messages(i).Avatar Is Nothing Then
            UserControl.PaintPicture Messages(i).Avatar, Margin, Y, AvatarSize, AvatarSize
        Else
            UserControl.Line (Margin, Y)-(Margin + AvatarSize, Y + AvatarSize), vbGrayText, BF
        End If

        ' Draw username
        UserControl.CurrentX = Margin + AvatarSize + 8
        UserControl.CurrentY = Y
        UserControl.ForeColor = vbBlack
        UserControl.FontBold = True
        UserControl.Print Messages(i).Username

        ' Draw text below username
        
        UserControl.FontBold = False
        UserControl.CurrentX = Margin + AvatarSize + 8
        UserControl.CurrentY = Y + 14
        UserControl.ForeColor = vbBlack
        UserControl.Print Messages(i).Text
    Next i


End Sub

