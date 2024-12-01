VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "EveryDiscord"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OPENSSL_init_ssl Lib "libssl-3.dll" (ByVal options As Long, ByVal settings As Long) As Long

Private Declare Function SSLv23_client_method Lib "libssl-3.dll" () As Long
Private Declare Function SSL_CTX_new Lib "libssl-3.dll" (ByVal method As Long) As Long
Private Declare Function SSL_CTX_free Lib "libssl-3.dll" (ByVal ctx As Long) As Long
Private Declare Function SSL_new Lib "libssl-3.dll" (ByVal ctx As Long) As Long
Private Declare Function SSL_set_fd Lib "libssl-3.dll" (ByVal ssl As Long, ByVal fd As Long) As Long
Private Declare Function SSL_connect Lib "libssl-3.dll" (ByVal ssl As Long) As Long
Private Declare Function SSL_read Lib "libssl-3.dll" (ByVal ssl As Long, ByVal buf As String, ByVal num As Long) As Long
Private Declare Function SSL_write Lib "libssl-3.dll" (ByVal ssl As Long, ByVal buf As String, ByVal num As Long) As Long
Private Declare Function SSL_shutdown Lib "libssl-3.dll" (ByVal ssl As Long) As Long
Private Declare Function SSL_free Lib "libssl-3.dll" (ByVal ssl As Long) As Long
Private Declare Function ERR_get_error Lib "libssl-3.dll" () As Long
Private Declare Function ERR_error_string Lib "libssl-3.dll" (ByVal e As Long, ByVal buf As String) As String
Sub MakeHttpsRequest()
    Dim ctx As Long
    Dim ssl As Long
    Dim result As Long
    Dim buffer As String
    Dim errorMsg As String * 256

    ' Initialize the SSL library using OPENSSL_init_ssl
    Call OPENSSL_init_ssl(0, 0)

    ' Create a new SSL context using SSLv23_client_method
    ctx = SSL_CTX_new(SSLv23_client_method())
    If ctx = 0 Then
        MsgBox "Failed to create SSL context"
        Exit Sub
    End If

    ' Create a new SSL connection
    ssl = SSL_new(ctx)
    If ssl = 0 Then
        MsgBox "Failed to create SSL connection"
        SSL_CTX_free (ctx)
        Exit Sub
    End If

    ' Set the file descriptor (fd) for the SSL connection
    ' Example socket descriptor is 0 for simplicity
    result = SSL_set_fd(ssl, 0) ' Replace 0 with your socket descriptor
    If result = 0 Then
        MsgBox "Failed to set file descriptor for SSL connection"
        SSL_free (ssl)
        SSL_CTX_free (ctx)
        Exit Sub
    End If

    ' Perform SSL handshake and connect
    result = SSL_connect(ssl)
    If result <> 1 Then
        result = ERR_get_error()
        errorMsg = ERR_error_string(result, errorMsg)
        MsgBox "SSL connect failed: " & errorMsg
        SSL_free (ssl)
        SSL_CTX_free (ctx)
        Exit Sub
    End If

    ' Write data to the SSL connection
    buffer = "GET / HTTP/1.1" & vbCrLf & "Host: example.com" & vbCrLf & vbCrLf
    result = SSL_write(ssl, buffer, Len(buffer))
    If result <= 0 Then
        result = ERR_get_error()
        errorMsg = ERR_error_string(result, errorMsg)
        MsgBox "SSL write failed: " & errorMsg
    End If

    ' Read response from the SSL connection
    buffer = String(1024, Chr(0))
    result = SSL_read(ssl, buffer, Len(buffer))
    MsgBox Left(buffer, result)

    ' Shutdown and free the SSL connection
    SSL_shutdown (ssl)
    SSL_free (ssl)
    SSL_CTX_free (ctx)
End Sub



Private Sub Command1_Click()
MakeHttpsRequest
End Sub

