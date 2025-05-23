VERSION 5.00
Begin VB.Form frmPieSocket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PieSocket Demo"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9165
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   59000
      Left            =   6825
      Top             =   75
   End
   Begin VB.ListBox List2 
      Height          =   3180
      Left            =   75
      TabIndex        =   5
      Top             =   600
      Width           =   1665
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   7140
   End
   Begin VB.TextBox Text3 
      Height          =   465
      Left            =   1800
      TabIndex        =   1
      Text            =   "Hello World"
      Top             =   3300
      Width           =   5865
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send"
      Height          =   540
      Left            =   7875
      TabIndex        =   0
      Top             =   3300
      Width           =   1065
   End
   Begin Project1.Websocket PieSocket 
      Height          =   465
      Left            =   7425
      Top             =   75
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   820
   End
   Begin VB.Label Label2 
      Caption         =   "Users:"
      Height          =   315
      Left            =   75
      TabIndex        =   6
      Top             =   300
      Width           =   1140
   End
   Begin VB.Label lblName 
      Caption         =   "Chatter"
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   300
      Width           =   2640
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "frmPieSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim Headers As Collection
    Dim WssUrl As String
    
    'this is needed for the random number generator
    Randomize Timer
    
    Set Headers = New Collection

    'required headers for piesocket
    Headers.Add "Content-Type: application/json"
    Headers.Add "Origin: localhost"

    'piesocket demo url
    '(note the notify_self and presence values - this sends us our own chat and lets us know when others leave or enter)
    WssUrl = "wss://demo.piesocket.com/v3/channel_1?api_key=oCdCMcMPQpbvNjUIzqtvF1d2X2okWpDQj4AwARJuAgtjhzKxVEjQU6IdCjwm&notify_self=true&presence=1"

    'connect to piesocket
    AddListItem "Connecting to piesocket..."
    
    PieSocket.Connect WssUrl, "443", , , Headers

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'proper disconnection
    If PieSocket.readyState <> STATE_CLOSED Then
        PieSocket.Send "{ ""event"":""member_leaving"",""sender"":""" & lblName & """ }"
        Do While PieSocket.isBusy
            DoEvents
        Loop

        PieSocket.Disconnect

        Do While PieSocket.readyState <> STATE_CLOSED
            DoEvents
        Loop
    End If

End Sub

Private Sub Command3_Click()

    'send chat message
    PieSocket.Send "{ ""event"": ""new_message"", ""sender"": """ & lblName & """, ""text"": """ & Text3 & """ }"
    Text3 = ""

End Sub

'if user hits enter key, send chat text
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Command3_Click
    End If
End Sub

Private Sub Timer1_Timer()
    'once a minute we send a ping to keep-alive
    If PieSocket.readyState = STATE_OPEN Then
        AddListItem "Ping"
        PieSocket.Ping
    End If
End Sub


'===========================================================
'PIESOCKET EVENTS
'===========================================================

Private Sub PieSocket_OnClose(ByVal eCode As WebsocketStatus, ByVal reason As String)
    AddListItem "Connection closed: " & reason
End Sub


Private Sub PieSocket_OnConnect(ByVal RemoteHost As String, ByVal RemoteIP As String, ByVal RemotePort As String)

    AddListItem "Connected! Setting name..."

    'create a new chatter name
    lblName = "Chatter_" & RandomNumber(1111, 9999)

    Caption = "PieSocket Demo - " & lblName

    PieSocket.Send "{ ""event"":""new_joining"",""sender"":""" & lblName & """ }"

End Sub

Private Sub PieSocket_OnMessage(ByVal Msg As Variant, ByVal OpCode As WebsocketOpCode)
    Dim JSon As Collection
    Dim X As Long

    'filter out piesocket demo spam
    If InStr(Msg, "Coordinated Universal Time") = 0 Then

        'if the message contains an event
        If InStr(Msg, """" & "event" & """") Then

            'parse the event
            Set JSon = FastParse(Msg)

            Select Case JSon("event")

                Case "new_joining"
                    'user joined the channel
                    List2.AddItem JSon("sender")

                Case "member_leaving"
                    'user left the channel
                    If List2.ListCount Then
                        For X = 0 To List2.ListCount - 1
                            If List2.List(X) = JSon("sender") Then
                                List2.RemoveItem X
                            End If
                        Next X
                    End If

                Case "new_message"
                    'chat message
                    AddListItem JSon("sender") & ": " & JSon("text")

                    'we can cheat and add a user to the list if they chat, if they are not already in the list
                    If List2.ListCount Then
                        For X = 0 To List2.ListCount - 1
                            If List2.List(X) = JSon("sender") Then    'already in the list
                                Exit Sub
                            End If
                        Next X
                    End If

                    'didnt find the user so add to the list
                    List2.AddItem JSon("sender")

                Case "system:member_list"
                    'when connected to a real piesocket and not a demo - here you would parse the member list
                    'and add it to the users list. in demo mode piesocket only lists all users as "anonymous"
                    'which is rather annoying
                    AddListItem "Recieved member list"


                Case "system:member_joined"
                    'when connected to a real piesocket and not a demo this is an official notification that
                    'a user has joined the channel
                    AddListItem "Somebody new joined the channel"

                Case "system:member_left"
                    'when connected to a real piesocket and not a demo this is an official notification that
                    'a user has left the channel
                    AddListItem "Somebody left the channel"

                Case Else
                    AddListItem Msg
            End Select

        Else
            'we will leave it as an exercise for the user to parse the rest of the json messages :)
            AddListItem Msg
        End If

    End If

End Sub

'the server answered a ping
Private Sub PieSocket_OnPong(ByVal IncludedMsg As String)
    AddListItem "Pong"
End Sub


'======================================================
'HELPER FUNCTIONS
'======================================================

'add an item to the list and scroll
Sub AddListItem(ByVal newVal As String)
    List1.AddItem newVal
    List1.ListIndex = List1.NewIndex    'scroll to new entry
End Sub


'generate a random number
Function RandomNumber(Optional ByVal Minval As Long, Optional ByVal Maxval As Long) As Long

    'make sure to call randomize(timer) at least once on program start up
    'before using this function

    RandomNumber = ((Maxval - Minval) * Rnd) + Minval
    If RandomNumber > Maxval Or RandomNumber < Minval Then
        RandomNumber = RandomNumber(Minval, Maxval)
    End If

End Function



'quick, nasty and dirty (very dirty) json parser.
'this is just for this example and you shouldnt use this in your own code
'because ... well, just because. Use a real json parser!
Function FastParse(ByVal strJSON As String) As Collection

    Dim Buff As String
    Dim X As Long
    Dim Char As String
    Dim nChar As String
    Dim pchar As String
    Dim nLen As Long
    Dim inQuote As Boolean
    Dim PrevVal As String
    Dim Keys As Collection

    Set Keys = New Collection
    Set FastParse = New Collection

    On Error Resume Next

    nLen = Len(strJSON)
    X = 1
    Do While X <= nLen
        pchar = Char
        Char = Mid$(strJSON, X, 1)
        If X < nLen Then
            nChar = Mid$(strJSON, X + 1, 1)
        Else
            nChar = ""
        End If

        If inQuote Then
            If Char <> """" Then
                Buff = Buff & Char
            Else
                inQuote = Not inQuote
            End If
        Else
            Select Case Char

                Case " "
                    If inQuote Then
                        Buff = Buff & Char
                    End If

                Case ",", "{", "[", "}", "]", vbCr, vbLf
                    If Len(PrevVal) Then
                        If Len(Buff) Then
                            FastParse.Add Buff, PrevVal
                            Keys.Add PrevVal
                            Buff = ""
                            PrevVal = ""
                        End If
                    Else
                        Select Case Char
                            Case "{", "[", ","
                                Buff = ""
                        End Select
                    End If


                Case """"
                    inQuote = Not inQuote


                Case ":"
                    PrevVal = Buff
                    Buff = ""

                Case Else    'usually numbers which are not quoted
                    Buff = Buff & Char
            End Select
        End If
        X = X + 1
    Loop

    'the first item in the returned collection is a collection of the keys
    FastParse.Add Keys, "Keys", 1

    On Error GoTo 0

End Function

