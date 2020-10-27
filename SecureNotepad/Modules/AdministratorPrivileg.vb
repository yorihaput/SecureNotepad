Imports System.Security.Principal
Module AdministratorPrivileg
    Public Function IsUserAdministrator() As Boolean
        Dim isAdmin As Boolean
        Try
            Dim user As WindowsIdentity = WindowsIdentity.GetCurrent()
            Dim principal As WindowsPrincipal = New WindowsPrincipal(user)
            isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator)
        Catch ex As UnauthorizedAccessException
            isAdmin = False
            MessageBox.Show(ex.Message)
        Catch ex As Exception
            isAdmin = False
            MessageBox.Show(ex.Message)
        End Try

        Return isAdmin

    End Function
End Module
