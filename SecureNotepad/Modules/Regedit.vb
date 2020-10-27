Imports Microsoft.Win32
Imports System.Runtime.InteropServices
Module Regedit
    <DllImport("shell32.dll")> _
    Public Sub SHChangeNotify( _
        ByVal wEventID As GlobalVariable.HChangeNotifyEventID, _
        ByVal uFlags As GlobalVariable.HChangeNotifyFlags, _
        ByVal dwItem1 As IntPtr, _
        ByVal dwItem2 As IntPtr)
    End Sub
    Public Sub RegisterSecureNotepad()
        Dim MyKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Classes\SecureNotepad", True)
        If MyKey Is Nothing Then
            MyKey = Registry.CurrentUser.CreateSubKey("Software\Classes\SecureNotepad", RegistryKeyPermissionCheck.ReadWriteSubTree)
        End If
        MyKey.SetValue("", "Protected Text Document", RegistryValueKind.String)

        MyKey = Registry.CurrentUser.OpenSubKey("Software\Classes\SecureNotepad\DefaultIcon", True)
        If MyKey Is Nothing Then
            MyKey = Registry.CurrentUser.CreateSubKey("Software\Classes\SecureNotepad\DefaultIcon", RegistryKeyPermissionCheck.ReadWriteSubTree)
        End If
        MyKey.SetValue("", Application.ExecutablePath() + ",-105", RegistryValueKind.ExpandString)

        MyKey = Registry.CurrentUser.OpenSubKey("Software\Classes\SecureNotepad\shell\open\command", True)
        If MyKey Is Nothing Then
            MyKey = Registry.CurrentUser.CreateSubKey("Software\Classes\SecureNotepad\shell\open\command", RegistryKeyPermissionCheck.ReadWriteSubTree)
        End If
        MyKey.SetValue("", Application.ExecutablePath() + " %1", RegistryValueKind.ExpandString)

        MyKey = Registry.CurrentUser.OpenSubKey("Software\Classes\.sctxt", True)
        If MyKey Is Nothing Then
            MyKey = Registry.CurrentUser.CreateSubKey("Software\Classes\.sctxt", RegistryKeyPermissionCheck.ReadWriteSubTree)
        End If
        MyKey.SetValue("", "SecureNotepad", RegistryValueKind.String)
    End Sub

    Public Sub UnregisterSecureNotepad()
        Dim MyKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Classes\SecureNotepad", True)
        If Not MyKey Is Nothing Then
            Registry.CurrentUser.DeleteSubKeyTree("Software\Classes\SecureNotepad", True)
        End If

        MyKey = Registry.CurrentUser.OpenSubKey("Software\Classes\.sctxt", True)
        If Not MyKey Is Nothing Then
            Registry.CurrentUser.DeleteSubKeyTree("Software\Classes\.sctxt", True)
        End If
    End Sub

    Public Function CheckFileAssoc() As Boolean
        Dim MyKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Classes\SecureNotepad", True)
        Dim response As Boolean = True
        If MyKey Is Nothing Then
            response = False
        End If
        MyKey = Registry.CurrentUser.OpenSubKey("Software\Classes\SecureNotepad\DefaultIcon", True)
        If MyKey Is Nothing Then
            response = False
        End If
        MyKey = Registry.CurrentUser.OpenSubKey("Software\Classes\SecureNotepad\shell\open\command", True)
        If MyKey Is Nothing Then
            response = False
        End If
        MyKey = Registry.CurrentUser.OpenSubKey("Software\Classes\.sctxt", True)
        If MyKey Is Nothing Then
            response = False
        End If
        Return response
    End Function
End Module
