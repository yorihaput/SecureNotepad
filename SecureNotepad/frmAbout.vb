Public Class frmAbout
    Private Sub frmAbout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblAppname.Text = GlobalVariable.AppName
        lblVersion.Text = GlobalVariable.Version
    End Sub

    Private Sub frmAbout_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then Me.Close()
    End Sub
End Class