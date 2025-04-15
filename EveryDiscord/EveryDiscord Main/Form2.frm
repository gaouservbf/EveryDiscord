VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EveryDiscord"
   ClientHeight    =   7110
   ClientLeft      =   90
   ClientTop       =   435
   ClientWidth     =   4650
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtToken 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin VB.CommandButton btnLogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Your Discord Token"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "everydiscord"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form2 (Login Form)
Option Explicit

Private Sub btnLogin_Click()
    ' Validate token
    If Len(txtToken.Text) = 0 Then
        MsgBox "Please enter your Discord token", vbExclamation, "Login Error"
        Exit Sub
    End If
    
    ' Hide login form and show main form
    Me.Hide
    Form1.Show
    
    ' Trigger login in the main form
    Form1.txtToken.Text = txtToken.Text
    Unload Me
End Sub

Private Sub Form_Load()
    ' Center the form
        If GetSetting("DiscordClient", "Settings", "Token", "") <> "" Then
Unload Me
Form1.Show
    End If
    'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    ' Set focus to token field
End Sub

