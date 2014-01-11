Public Class Login

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim username = TextBox1.Text
        Dim password = TextBox2.Text
        If User.Login(username, password) Then
            MsgBox("Login successfull. User is a " & TopLevelProperties.UserType, vbInformation, "Yay!")
            If TopLevelProperties.isTeacher Then
                Dim x As Teacher
                x = Teacher.find(TopLevelProperties.TeacherID)
                MsgBox("Teacher is " & x.ToString)
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim x = MsgBox("Are you sure you want to exit?", MsgBoxStyle.YesNo, "Message")
        If x = vbYes Then
            Me.Close()
        End If
    End Sub
End Class