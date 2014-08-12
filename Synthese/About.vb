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

Public Class About

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Process.Start("http://creativecommons.org/licenses/by-nc-sa/4.0/")
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        MsgBox("Icons by Open Iconic (https://useiconic.com/open/)")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class