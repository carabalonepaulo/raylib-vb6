Attribute VB_Name = "modIni"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function INI_ReadString(ByVal app As String, ByVal key As String, ByVal path As String) As String
    Dim buffer As String * 255
    Dim ret As Long

    ret = GetPrivateProfileString(app, key, "", buffer, Len(buffer), path)
    INI_ReadString = Left$(buffer, ret)
End Function

Public Function INI_ReadByte(ByVal app As String, ByVal key As String, ByVal path As String) As Byte
    Dim val As String
    val = INI_ReadString(app, key, path)
    If IsNumeric(val) Then
        INI_ReadByte = CByte(val)
    Else
        INI_ReadByte = 0
    End If
End Function

Public Function INI_ReadInt(ByVal app As String, ByVal key As String, ByVal path As String) As Integer
    Dim val As String
    val = INI_ReadString(app, key, path)
    If IsNumeric(val) Then
        INI_ReadInt = CInt(val)
    Else
        INI_ReadInt = 0
    End If
End Function

Public Function INI_ReadLong(ByVal app As String, ByVal key As String, ByVal path As String) As Long
    Dim val As String
    val = INI_ReadString(app, key, path)
    If IsNumeric(val) Then
        INI_ReadLong = CLng(val)
    Else
        INI_ReadLong = 0
    End If
End Function

Public Function INI_ReadSingle(ByVal app As String, ByVal key As String, ByVal path As String) As Single
    Dim val As String
    val = INI_ReadString(app, key, path)
    If IsNumeric(val) Then
        INI_ReadSingle = CSng(val)
    Else
        INI_ReadSingle = 0#
    End If
End Function

Public Function INI_ReadBoolean(ByVal app As String, ByVal key As String, ByVal path As String) As Boolean
    Dim val As String
    val = INI_ReadString(app, key, path)
    If IsNumeric(val) Then
        INI_ReadBoolean = CBool(CInt(val))
    Else
        INI_ReadBoolean = (UCase$(val) = "TRUE")
    End If
End Function
