Public Class SplashScreen
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ProgressBar1.Value += 1
        If ProgressBar1.Value = 10 Then
            ToolStripStatusLabel1.Text = "Loading Forms"
        End If
        If ProgressBar1.Value = 20 Then
            ToolStripStatusLabel1.Text = "Loading Database"
        End If
        If ProgressBar1.Value = 30 Then
            ToolStripStatusLabel1.Text = "Loading Database Tables"
        End If
        If ProgressBar1.Value = 40 Then
            ToolStripStatusLabel1.Text = "Loading Database Values"
        End If
        If ProgressBar1.Value = 50 Then
            ToolStripStatusLabel1.Text = "Initializing Records"
        End If
        If ProgressBar1.Value = 75 Then
            ToolStripStatusLabel1.Text = "Starting Up Application"
        End If
        If ProgressBar1.Value = 100 Then
            Timer1.Enabled = False
            Login.Show()
            Me.Close()
        End If
    End Sub
End Class
