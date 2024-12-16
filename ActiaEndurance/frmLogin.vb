Public Class frmLogin
    Private Sub btn_login_Click(sender As Object, e As EventArgs) Handles btn_login.Click
        Dim txtPassword As String

        txtPassword = txt_password.Text
        txtPassword = txtPassword.ToLower

        If cb_username.Text = "" And txtPassword = "" Then
            UserLevel = 1
            Me.Close()
            frmMain.Show()
        ElseIf cb_username.Text = "Engineer" And txtPassword = "engineer" Then
            UserLevel = 2
            Me.Close()
            frmMain.Show()
        ElseIf cb_username.Text = "Operator" And txtPassword = "" Then
            UserLevel = 3
            Me.Close()
            frmMain.Show()
        ElseIf cb_username.Text = "Quality" And txtPassword = "quality" Then
            UserLevel = 4
            Me.Close()
            frmMain.Show()
        Else
            lbl_incorrect.Visible = True
        End If
    End Sub

    Private Sub btn_exit_Click(sender As Object, e As EventArgs) Handles btn_exit.Click
        Close()
    End Sub
    Private Sub cb_username_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cb_username.KeyPress
        If e.KeyChar = Chr(13) Then
            btn_login.PerformClick()
        End If
    End Sub
    Private Sub txt_password_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_password.KeyPress
        If e.KeyChar = Chr(13) Then
            btn_login.PerformClick()
        End If
    End Sub

    Private Sub cb_username_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_username.SelectedIndexChanged
        lbl_incorrect.Visible = False
    End Sub

    Private Sub txt_password_TextChanged(sender As Object, e As EventArgs) Handles txt_password.TextChanged
        lbl_incorrect.Visible = False
    End Sub

    Private Sub LoginForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        txt_password.Text = ""
        cb_username.Text = ""
    End Sub
End Class