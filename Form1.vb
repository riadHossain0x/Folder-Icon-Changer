
' |      ////////////////////////////////////////////////////////////// |
' |     //////////CODED BY RIYAD PATHAN (ARM@N_HACKTIVIST)////////////  |
' |    //////////////////////////////////////////////////////////////   |

Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data
Imports System.IO
Imports Riyad_Pathan.FolderIcons

Public Class Form1

    Private Sub ClsButtonGreen1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClsButtonGreen1.Click
        Dim FolderBrowserDialog1 As New FolderBrowserDialog
        FolderBrowserDialog1.SelectedPath = TextBox1.Text
        FolderBrowserDialog1.Description = "Select folder:"
        Dim result As DialogResult = FolderBrowserDialog1.ShowDialog()
        If result = DialogResult.OK Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub ClsButtonGreen2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClsButtonGreen2.Click
        Dim FileBrowser As New OpenFileDialog

        If TextBox1.Text.Length > 0 Then
            FileBrowser.InitialDirectory = TextBox1.Text
        End If

        If TextBox2.Text.Length > 0 Then
            FileBrowser.FileName = TextBox2.Text
        Else
            FileBrowser.FileName = Nothing
        End If

        FileBrowser.Filter = "Icon files (*.ico)|*.ico|All files (*.*)|*.*"
        FileBrowser.FilterIndex = 1
        FileBrowser.RestoreDirectory = False
        Dim result As DialogResult = FileBrowser.ShowDialog()

        If result = DialogResult.OK Then
            TextBox2.Text = FileBrowser.FileName
        End If
    End Sub

    Private Sub ClsButtonGreen3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClsButtonGreen3.Click
        If MessageBox.Show("Once Changed the original icon can not be recovered Do you wish to Continue?", "Warning", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Try
                If TextBox1.Text.Length > 0 Then
                    If File.Exists(TextBox2.Text) Then
                        Dim myFolderIcon As New FolderIcon(TextBox1.Text)
                        myFolderIcon.CreateFolderIcon(TextBox2.Text, "Forlder Icon Changer")
                        myFolderIcon = Nothing
                        MessageBox.Show("Successfully ICON Changed.", "Forlder Icon Changer", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("No ICON Selected.", "Folder Icon Changer", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If
                Else
                    MessageBox.Show("No Folder Selected.", "Folder Icon Changer", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            Catch ex As Exception
                MsgBox("Error!")
            End Try
        Else
        End If
    End Sub

    Private Sub ClsButtonBlue1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClsButtonBlue1.Click
        End
    End Sub
End Class
