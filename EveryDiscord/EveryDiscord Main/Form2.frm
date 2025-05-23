VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EveryDiscord"
   ClientHeight    =   7110
   ClientLeft      =   90
   ClientTop       =   435
   ClientWidth     =   4590
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnLogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   6600
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Via Tokens"
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   4095
      Begin VB.TextBox txtToken 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Your Discord Token"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3495
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5655
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9975
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Credentials"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
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
      TabIndex        =   1
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
    SaveSetting "DiscordClient", "Settings", "Token", txtToken.Text
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


