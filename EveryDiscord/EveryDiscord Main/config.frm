VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configure your client.."
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4605
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "EveryDiscord Specials"
      Height          =   5160
      Left            =   210
      TabIndex        =   3
      Top             =   525
      Width           =   4110
      Begin VB.CommandButton Command4 
         Caption         =   "Theming"
         Height          =   330
         Left            =   2835
         TabIndex        =   7
         Top             =   1365
         Width           =   1170
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Plugins.."
         Height          =   330
         Left            =   2835
         TabIndex        =   4
         Top             =   525
         Width           =   1170
      End
      Begin VB.Label Label2 
         Caption         =   "You can theme EveryDiscord, for example giving chat a nice background"
         Height          =   645
         Left            =   105
         TabIndex        =   6
         Top             =   1365
         Width           =   2745
      End
      Begin VB.Label Label1 
         Caption         =   "EveryDiscord lets you use plugins to enchance your experience"
         Height          =   645
         Left            =   105
         TabIndex        =   5
         Top             =   525
         Width           =   2745
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2415
      TabIndex        =   2
      Top             =   5985
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3465
      TabIndex        =   1
      Top             =   5985
      Width           =   960
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5790
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   10213
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   9
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Client Customization"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "My Account"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Profiles"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Content & Social"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Data & Privacy"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Family Center"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Authorized Apps"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Devices"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Connections"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click()
Form5.Show
End Sub

Private Sub Command4_Click()
Form6.Show
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem = 1 Then

End If
End Sub
