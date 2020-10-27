Imports System.IO

Public Class frmMain
    Dim is_save = True
    Dim filePath = ""
    Dim fileName = "Untitled"
    Private Sub SaveFile(ByVal content As String, ByVal filePath As String, Optional ByVal FilePassword As String = "")
        Dim fileWriter As StreamWriter
        Dim encrypted As String
        If FilePassword = "" Then
            encrypted = content
        Else
            encrypted = EncDec.Encrypt(content, FilePassword)
        End If
        Try
            fileWriter = New StreamWriter(filePath, False, System.Text.Encoding.Unicode)
            fileWriter.Write(encrypted)
            fileWriter.Close()
        Catch ex As Exception
            Dim eMsg As String
            eMsg = ex.Message
            Select Case True
                Case Else
                    MsgBox("Failed to save the file:" & Constants.vbNewLine & ex.Message, MsgBoxStyle.Critical, "Failed")
            End Select
        End Try
    End Sub

    Private Sub OpenFile(ByVal TextBox As TextBox, ByVal filePath As String, Optional ByVal FilePassword As String = "")
        Dim fileReader As StreamReader
        Dim decrypted As String = ""
        Dim content As String = ""
        Try
            fileReader = New StreamReader(filePath)
            content = fileReader.ReadToEnd()
            If FilePassword = "" Then
                decrypted = EncDec.Decrypt(content)
            Else
                decrypted = EncDec.Decrypt(content, FilePassword)
            End If

            TextBox.Text = decrypted
            fileReader.Close()
        Catch ex As Exception
            fileReader.Close()
            If Not Application.OpenForms().OfType(Of frmPassword).Any Then
                Dim eMsg As String
                eMsg = ex.Message
                Select Case True
                    Case eMsg.Contains("Could not find file")
                        MessageBox.Show("File " + filePath + " does not existed. Please enter the correct file", GlobalVariable.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Case eMsg.Contains("Padding is invalid and cannot be removed.")
                        frmPassword.lblSave.Visible = False
                        frmPassword.lblOpen.Visible = True
                        frmPassword.ShowDialog()
                        If GlobalVariable.frmPasswordCancel = False Then
                            OpenFile(TextBox, filePath, GlobalVariable.filePassword)
                        End If
                    Case Else
                        MsgBox("Failed to decrypt the file. Maybe the file not protected or corrupt.", MsgBoxStyle.Information, "Failed")
                        GlobalVariable.frmPasswordCancel = False
                        TextBox.Text = content
                End Select
            End If
        End Try
    End Sub
    Private Sub CheckIsSaveTitle()
        If is_save = False And Not Me.Text.Substring(0, 1) = "*" Then
            Me.Text = "*" + Me.Text
        ElseIf is_save = True And Me.Text.Substring(0, 1) = "*" Then
            Me.Text = Me.Text.Substring(1)
        End If

        If AdministratorPrivileg.IsUserAdministrator() = True And Not Me.Text.Substring(Me.Text.Length - 15, 15) = "[Administrator]" Then
            Me.Text = Me.Text + " [Administrator]"
        End If
    End Sub

    Private Sub CheckMenu()
        If txtMain.WordWrap = True Then
            WordWrapToolStripMenuItem.Checked = True
        Else
            WordWrapToolStripMenuItem.Checked = False
        End If

        If statusStripMain.Visible = True Then
            StatusToolStripMenuItem.Checked = True
        Else
            StatusToolStripMenuItem.Checked = False
        End If

        AssocFilesctxtToolStripMenuItem.Checked = Regedit.CheckFileAssoc
    End Sub
    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        frmAbout.ShowDialog()
    End Sub

    Private Sub WordWrapToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WordWrapToolStripMenuItem.Click
        If txtMain.WordWrap = True Then
            txtMain.WordWrap = False
        Else
            txtMain.WordWrap = True
        End If
        Call CheckMenu()
    End Sub

    Private Sub StatusToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StatusToolStripMenuItem.Click
        If statusStripMain.Visible = True Then
            statusStripMain.Visible = False
        Else
            statusStripMain.Visible = True
        End If
        Call CheckMenu()
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim s() As String = System.Environment.GetCommandLineArgs()
        If s.Length > 1 Then
            If Not s(1) Is Nothing Then
                filePath = s(1)
                Call OpenFile(txtMain, filePath, GlobalVariable.filePassword)
                If GlobalVariable.frmPasswordCancel = False Then
                    fileName = filePath.ToString().Split("\").Last()
                    Me.Text = fileName + " - " + GlobalVariable.AppName
                    is_save = True

                    Call CheckMenu()
                    Call CheckIsSaveTitle()
                    Exit Sub
                Else
                    Me.Close()
                End If
            End If
        End If

        

        Me.Text = fileName + " - " + GlobalVariable.AppName
        Call CheckMenu()
        Call CheckIsSaveTitle()
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        If Not txtMain.SelectedText = "" Then
            Clipboard.Clear()
            Clipboard.SetText(txtMain.SelectedText)
        End If
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripMenuItem.Click
        If Not txtMain.SelectedText = "" Then
            Clipboard.Clear()
            Clipboard.SetText(txtMain.SelectedText)
            txtMain.SelectedText = ""
        End If
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        Dim selectionIndex = txtMain.SelectionStart
        Dim insertText = Clipboard.GetText
        txtMain.Text = txtMain.Text.Insert(selectionIndex, insertText)
        txtMain.SelectionStart = selectionIndex + insertText.Length
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        txtMain.SelectedText = ""
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        If is_save = True Then
            txtMain.Text = ""
            is_save = True
        Else
            Dim saveFile = MessageBox.Show("Do you want to save the file?", GlobalVariable.AppName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If saveFile = vbYes Then
                SaveToolStripMenuItem_Click(sender, e)
                txtMain.Text = ""
                is_save = True
            ElseIf saveFile = vbNo Then
                txtMain.Text = ""
                is_save = True
            End If
        End If

        filePath = ""
        fileName = "Untitled"
        GlobalVariable.filePassword = ""
        GlobalVariable.frmPasswordCancel = True
        Me.Text = fileName + " - " + GlobalVariable.AppName
        Call CheckIsSaveTitle()
    End Sub

    Private Sub txtMain_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMain.TextChanged
        is_save = False
        Call CheckIsSaveTitle()
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        If filePath = "" Then
            If GlobalVariable.filePassword = "" Then
                frmPassword.lblSave.Visible = True
                frmPassword.lblOpen.Visible = False
                frmPassword.ShowDialog()
            End If

            If GlobalVariable.frmPasswordCancel = False Then
                If GlobalVariable.filePassword = "" Then
                    SaveFileDialog1.Filter = "Text File (*.txt)|*.txt|All files (*.*)|*.*"
                Else
                    SaveFileDialog1.Filter = "Secured Text File (*.sctxt)|*.sctxt|Text File (*.txt)|*.txt|All files (*.*)|*.*"
                End If
                If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                    filePath = SaveFileDialog1.FileName
                    Call SaveFile(txtMain.Text, filePath, GlobalVariable.filePassword)
                    fileName = filePath.ToString().Split("\").Last()
                End If
            End If
        Else
            Call SaveFile(txtMain.Text, filePath, GlobalVariable.filePassword)
        End If
        SaveFileDialog1.FileName = ""
        Me.Text = fileName + " - " + GlobalVariable.AppName
        is_save = True

        Call CheckIsSaveTitle()
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click
        If GlobalVariable.filePassword = "" Then
            frmPassword.lblSave.Visible = True
            frmPassword.lblOpen.Visible = False
            frmPassword.ShowDialog()
        End If

        If frmPasswordCancel = False Then
            If GlobalVariable.filePassword = "" Then
                SaveFileDialog1.Filter = "Text File (*.txt)|*.txt|All files (*.*)|*.*"
            Else
                SaveFileDialog1.Filter = "Secured Text File (*.sctxt)|*.sctxt|Text File (*.txt)|*.txt|All files (*.*)|*.*"
            End If
            If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                filePath = SaveFileDialog1.FileName
                Call SaveFile(txtMain.Text, filePath, GlobalVariable.filePassword)
                fileName = filePath.ToString().Split("\").Last()
            End If
        End If

        SaveFileDialog1.FileName = ""
        Me.Text = fileName + " - " + GlobalVariable.AppName
        is_save = True

        Call CheckIsSaveTitle()
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        Call NewToolStripMenuItem_Click(sender, e)
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Secured Text File (*.sctxt)|*.sctxt|Text File (*.txt)|*.txt|All files (*.*)|*.*"
        OpenFileDialog1.FilterIndex = 1
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            filePath = OpenFileDialog1.FileName
            Call OpenFile(txtMain, filePath, GlobalVariable.filePassword)
            fileName = filePath.ToString().Split("\").Last()

            If fileName = "" Then
                fileName = "Untitled"
            End If
        End If

        Me.Text = fileName + " - " + GlobalVariable.AppName
        is_save = True
        Call CheckIsSaveTitle()
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If is_save = False Then
            Dim saveFile = MessageBox.Show("Do you want to save the file?", GlobalVariable.AppName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If saveFile = vbYes Then
                SaveToolStripMenuItem_Click(sender, e)
            ElseIf saveFile = vbNo Then

            Else
                e.Cancel = True
            End If
        End If
    End Sub
    Private Sub AssocFilesctxtToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AssocFilesctxtToolStripMenuItem.Click
        If AssocFilesctxtToolStripMenuItem.Checked = True Then
            Call Regedit.UnregisterSecureNotepad()
        Else
            Call Regedit.RegisterSecureNotepad()
        End If

        Call Regedit.SHChangeNotify(
            GlobalVariable.HChangeNotifyEventID.SHCNE_ASSOCCHANGED,
            GlobalVariable.HChangeNotifyFlags.SHCNF_IDLIST,
            IntPtr.Zero,
            IntPtr.Zero
        )

        Call CheckMenu()
    End Sub
End Class
