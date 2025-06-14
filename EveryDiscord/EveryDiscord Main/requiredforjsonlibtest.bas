Attribute VB_Name = "Module2"
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long



Public Function UTF8ToUTF16LE(ByRef textToConvert As String) As String
    Dim charCount As Long
    Const pageCodeUTF8 = 65001
    '
    charCount = MultiByteToWideChar(pageCodeUTF8, 0, StrPtr(textToConvert) _
                                  , LenB(textToConvert), 0, 0)
    If charCount = 0 Then Exit Function
    '
    UTF8ToUTF16LE = Space$(charCount)
    MultiByteToWideChar pageCodeUTF8, 0, StrPtr(textToConvert) _
                      , LenB(textToConvert), StrPtr(UTF8ToUTF16LE), charCount
End Function


