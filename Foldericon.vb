Imports System.IO

Namespace FolderIcons
    Public Class FolderIcon
        Private m_folderPath As String = ""
        Private m_iniPath As String = ""
        Public Sub New(ByVal folderPath As String)
            Me.FolderPath = folderPath
        End Sub

        Public Sub CreateFolderIcon(ByVal iconFilePath As String, ByVal infoTip As String)
            If CreateFolder() Then
                CreateDesktopIniFile(iconFilePath, infoTip)
                SetIniFileAttributes()
                SetFolderAttributes()
            End If
        End Sub

        Public Sub CreateFolderIcon(ByVal targetFolderPath As String, ByVal iconFilePath As String, ByVal infoTip As String)
            Me.FolderPath = targetFolderPath
            Me.CreateFolderIcon(iconFilePath, infoTip)
        End Sub


        Public Property FolderPath() As String
            Get
                Return Me.m_folderPath
            End Get
            Set(ByVal value As String)
                m_folderPath = value
                If Not Me.m_folderPath.EndsWith("\") Then
                    Me.m_folderPath += "\"
                End If
            End Set
        End Property

        Public Property IniPath() As String
            Get
                Return m_iniPath
            End Get
            Set(ByVal value As String)
                m_iniPath = value
            End Set
        End Property

        Private Function CreateFolder() As Boolean

            If Me.FolderPath.Length = 0 Then
                Return False
            End If


            If Directory.Exists(Me.FolderPath) Then
                Return True
            End If

            Try

                Dim di As DirectoryInfo = Directory.CreateDirectory(Me.FolderPath)
            Catch e As Exception
                Return False
            End Try

            Return True
        End Function

        Private Function CreateDesktopIniFile(ByVal iconFilePath As String, ByVal getIconFromDLL As Boolean, ByVal iconIndex As Integer, ByVal infoTip As String) As Boolean

            If Not Directory.Exists(Me.FolderPath) Then
                Return False
            End If


            If Not File.Exists(iconFilePath) Then
                Return False
            End If

            If Not getIconFromDLL Then
                iconIndex = 0
            End If


            Me.IniPath = Me.FolderPath & "desktop.ini"


            IniWriter.WriteValue(".ShellClassInfo", "IconFile", iconFilePath, Me.IniPath)
            IniWriter.WriteValue(".ShellClassInfo", "IconIndex", iconIndex.ToString(), Me.IniPath)
            IniWriter.WriteValue(".ShellClassInfo", "InfoTip", infoTip, Me.IniPath)

            Return True
        End Function

        Private Sub CreateDesktopIniFile(ByVal iconFilePath As String, ByVal infoTip As String)
            Me.CreateDesktopIniFile(iconFilePath, False, 0, infoTip)
        End Sub

        Private Function SetIniFileAttributes() As Boolean

            If Not File.Exists(Me.IniPath) Then
                Return False
            End If


            If (File.GetAttributes(Me.IniPath) And FileAttributes.Hidden) <> FileAttributes.Hidden Then
                File.SetAttributes(Me.IniPath, File.GetAttributes(Me.IniPath) Or FileAttributes.Hidden)
            End If


            If (File.GetAttributes(Me.IniPath) And FileAttributes.System) <> FileAttributes.System Then
                File.SetAttributes(Me.IniPath, File.GetAttributes(Me.IniPath) Or FileAttributes.System)
            End If

            Return True

        End Function

        Private Function SetFolderAttributes() As Boolean

            If Not Directory.Exists(Me.FolderPath) Then
                Return False
            End If

            If (File.GetAttributes(Me.FolderPath) And FileAttributes.System) <> FileAttributes.System Then
                File.SetAttributes(Me.FolderPath, File.GetAttributes(Me.FolderPath) Or FileAttributes.System)
            End If

            Return True

        End Function

    End Class
End Namespace
