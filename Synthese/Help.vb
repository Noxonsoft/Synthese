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

Public Class Help

    Private Sub Help_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Load the help file
        WebBrowser1.Navigate(Application.StartupPath & "/help.html")
    End Sub
End Class