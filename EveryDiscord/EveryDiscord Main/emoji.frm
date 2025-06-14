VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form4 
   Caption         =   "Emojis"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4065
   LinkTopic       =   "Form4"
   ScaleHeight     =   3915
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   2
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1785
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error Resume Next
    Dim sTheFolder As String
    sTheFolder = App.Path + "\emoji\msn\"
    Dim sTheFile As String
    sTheFile = Dir$(sTheFolder & "*.*")
    Dim i As Integer
    i = 0
    Do While Len(sTheFile)
    
        i = i + 1
        sTheFile = Dir$
        
        ImageList1.ListImages.Add i, sTheFile, LoadPicture(App.Path + "\emoji\msn\" + sTheFile)
    Toolbar1.Buttons.Add i, ":" + Mid(sTheFile, 1, Len(sTheFile) - 4) + ":", StrConv(Replace(Mid(sTheFile, 1, Len(sTheFile) - 4), "_", " "), vbProperCase), tbrDefault, i
    Loop
End Sub

Private Sub Form_Resize()
Toolbar1.Height = Me.ScaleHeight
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Form1.txtMsg.Text = Form1.txtMsg.Text + Button.Key
End Sub
