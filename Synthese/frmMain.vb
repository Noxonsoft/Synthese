'* 6-8-2014
'*
'* ---------------------------
'* Synthese BETA | frmMain.vb
'* ---------------------------
'* This program is free software; you can redistribute it and/or modify
'* it under the terms of the Creative Common License as published by
'* the Creative Common Organization.
'*
'* Made by NoxonSoft

Imports IWshRuntimeLibrary
Imports System.Net

Public Class frmMain

    Dim icon_location As String

    Dim StartMenuFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu)
    Dim DesktopFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)

    Dim checkboxes As Boolean = False
    Dim icons As Boolean = False
    Dim text_url As Boolean = False
    Dim text_name As Boolean = False

    Dim winstyle As Integer = 1

    'Standard script to create a shortcut
    Public Sub CreateShortcut(ByVal FilePath As String, ByVal Shortcut_Directory As String, ByVal Shortcut_Description As String, ByVal Shortcut_Name As String, ByVal Shortcut_Icon As String, ByVal Windows_Style As Integer, ByVal URL As String)
        Dim new_shortcut As IWshShortcut
        Dim wsh As New WshShell

        new_shortcut = CType(wsh.CreateShortcut(Shortcut_Directory & "\" & Shortcut_Name & ".lnk"), IWshShortcut)

        With new_shortcut
            .TargetPath = FilePath
            .WindowStyle = Windows_Style
            .Description = Shortcut_Description
            .IconLocation = Shortcut_Icon & ", 0"
            .Arguments = Shortcut_Name & " " & URL
            .Save()
        End With
    End Sub

    'Small script to detect if a URL is valid (without webbrowser component!)
    Public Function CheckURL(ByVal URL As String) As Boolean
        Try
            Dim request As WebRequest = WebRequest.Create(URL)
            Dim response As WebResponse = request.GetResponse()
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    'Simple sub to check if we got all the info that we need to create a shortcut
    Public Sub IsReady()
        If checkboxes = True And icons = True And text_url = True And text_name = True Then
            Button2.Enabled = True
        Else
            Button2.Enabled = False
        End If
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Check if the launcher file still exists. If not, ask the user to download it.
        If My.Computer.FileSystem.FileExists(Application.StartupPath & "/launcher/launcher.exe") = False Then
            Dim result As Integer = MessageBox.Show("Launcher.exe is missing. Do you want to download it now?", "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
            If result = DialogResult.No Then
                Application.Exit()
            ElseIf result = DialogResult.Yes Then
                Try
                    'Download the file to the correct location
                    My.Computer.Network.DownloadFile("http://www.noxonsoft.com/Downloadable_Files/Launcher.exe", Application.StartupPath & "/launcher/launcher.exe")
                Catch ex As Exception
                    'The download failed. Restart application to try again.
                    MsgBox("Error while downloading file:" & vbCrLf & ex.ToString, MsgBoxStyle.Critical, "Error")
                    Application.Exit()
                End Try
            End If
        End If
        'Detect the OS version because Windows 8 and higher don't officialy have a Startmenu. So no support for 'fake' startmenu's.
        Dim OS_Version As Version = Environment.OSVersion.Version
        If OS_Version.Major = 6 And OS_Version.Minor = 1 Then
            'Windows 7
            CheckBox2.Enabled = True
            PictureBox5.Visible = False
        Else
            'Windows 8
            CheckBox2.Enabled = False
            PictureBox5.Visible = True
        End If

        'Make the images size to fit
        PictureBox4.SizeMode = PictureBoxSizeMode.StretchImage
        PictureBox2.SizeMode = PictureBoxSizeMode.StretchImage
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If RadioButton1.Checked = True Then
            winstyle = 1
        ElseIf RadioButton2.Checked = True Then
            winstyle = 3
        Else
            winstyle = 7
        End If

        'Check if the user has an image selected
        If ComboBox1.Text = "Custom..." And icon_location = "" Then
            MsgBox("You must select an image", MsgBoxStyle.Critical, "Error")
            TabControl1.SelectedIndex = 1
            Exit Sub
        Else
            GoTo a
        End If

a:
        Try
            'Create the shortcut(s) with CreateShortcut
            If CheckBox1.Checked = True Then
                CreateShortcut(Application.StartupPath & "\launcher\Launcher.exe", DesktopFolder, "Created with NoxonSoft Synthese", txt_name.Text, icon_location, winstyle, txt_url.Text)
            End If
            If CheckBox2.Checked = True Then
                CreateShortcut(Application.StartupPath & "\launcher\Launcher.exe", StartMenuFolder, "Created with NoxonSoft Synthese", txt_name.Text, icon_location, winstyle, txt_url.Text)
            End If
        Catch ex As Exception
            MsgBox("Cannot create shortcut:" & vbCrLf & ex.ToString, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt_url_Leave(sender As Object, e As EventArgs) Handles txt_url.Leave
        If Not txt_url.Text = "" Then
            If CheckURL("http://" & txt_url.Text) = True Then
                PictureBox4.Image = My.Resources.check
            Else
                PictureBox4.Image = My.Resources.cross
            End If
            text_url = True
            IsReady()
        Else
            text_url = False
            IsReady()
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "" Then
            icons = False
            IsReady()
        Else
            If ComboBox1.Text = "Custom..." Then
                Button1.Enabled = True
                icon_location = ""
            Else
                Button1.Enabled = False
                icon_location = Application.StartupPath & "/icons/default.ico"
            End If
            icons = True
            IsReady()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ofd As New OpenFileDialog
        ofd.Filter = "Icon file (*.ico)|*.ico"
        If ofd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Try
                PictureBox2.Load(ofd.FileName)
                icon_location = ofd.FileName
            Catch ex As Exception
                MsgBox("Cannot load icon. Please try again or try to use a different icon.", MsgBoxStyle.Exclamation, "Error")
            End Try
        End If
    End Sub

    Private Sub txt_url_Enter(sender As Object, e As EventArgs) Handles txt_url.Enter
        PictureBox4.Image = My.Resources.loading_gif
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox2.Checked = False Then
            If CheckBox1.Checked = False Then
                'Nothing checked
                checkboxes = False
                IsReady()
            Else
                'CheckBox1 checked
                checkboxes = True
                IsReady()
            End If
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub NewWindowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewWindowToolStripMenuItem.Click
        Dim new_window As New frmMain
        new_window.Show()
    End Sub

    Private Sub txt_name_Leave(sender As Object, e As EventArgs) Handles txt_name.Leave
        If txt_name.Text = "" Then
            text_name = False
            IsReady()
        Else
            text_name = True
            IsReady()
        End If
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        About.ShowDialog()
    End Sub

    Private Sub HelpToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem1.Click
        Help.ShowDialog()
    End Sub
End Class
