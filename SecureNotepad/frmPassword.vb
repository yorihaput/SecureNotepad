Public Class frmPassword
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        GlobalVariable.filePassword = TextBox1.Text
        GlobalVariable.frmPasswordCancel = False
        Me.Close()
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        GlobalVariable.frmPasswordCancel = True
        Me.Close()
    End Sub


    Private Sub frmPassword_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If GlobalVariable.filePassword = "" And GlobalVariable.frmPasswordCancel = True Then
            GlobalVariable.frmPasswordCancel = True
        End If
    End Sub

    Private Sub frmPassword_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""

    End Sub

    Private Sub frmPassword_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        TextBox1.Focus()
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1_Click(sender, e)
        End If
    End Sub
End Class