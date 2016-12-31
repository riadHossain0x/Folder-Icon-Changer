Imports System.Runtime.InteropServices

Namespace FolderIcons

    Public Class IniWriter

        <DllImport("kernel32")> _
        Private Shared Function WritePrivateProfileString(ByVal iniSection As String, ByVal iniKey As String, ByVal iniValue As String, ByVal iniFilePath As String) As Integer
        End Function

        Public Shared Sub WriteValue(ByVal iniSection As String, ByVal iniKey As String, ByVal iniValue As String, ByVal iniFilePath As String)
            WritePrivateProfileString(iniSection, iniKey, iniValue, iniFilePath)
        End Sub

    End Class
End Namespace