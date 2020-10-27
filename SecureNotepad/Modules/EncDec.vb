Imports System
Imports System.IO
Imports System.Text
Imports System.Security
Imports System.Security.Cryptography
Module EncDec
    Private Function CreateKey(ByVal strPassword As String) As Byte()
        Dim chrData() As Char = strPassword.ToCharArray
        Dim intLength As Integer = chrData.GetUpperBound(0)
        Dim bytDataToHash(intLength) As Byte
        For i As Integer = 0 To chrData.GetUpperBound(0)
            bytDataToHash(i) = CByte(Asc(chrData(i)))
        Next
        Dim SHA512 As New System.Security.Cryptography.SHA512Managed
        Dim bytResult As Byte() = SHA512.ComputeHash(bytDataToHash)
        Dim bytKey(31) As Byte
        For i As Integer = 0 To 31
            bytKey(i) = bytResult(i)
        Next
        Return bytKey
    End Function

    Private Function CreateIV(ByVal strPassword As String) As Byte()
        Dim chrData() As Char = strPassword.ToCharArray
        Dim intLength As Integer = chrData.GetUpperBound(0)
        Dim bytDataToHash(intLength) As Byte
        For i As Integer = 0 To chrData.GetUpperBound(0)
            bytDataToHash(i) = CByte(Asc(chrData(i)))
        Next
        Dim SHA512 As New System.Security.Cryptography.SHA512Managed
        Dim bytResult As Byte() = SHA512.ComputeHash(bytDataToHash)
        Dim bytIV(15) As Byte
        For i As Integer = 32 To 47
            bytIV(i - 32) = bytResult(i)
        Next
        Return bytIV
    End Function

    Public Function Encrypt(ByVal clearText As String, Optional ByVal FilePassword As String = "filePassword") As String
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        Dim aes As New RijndaelManaged
        bytKey = CreateKey(FilePassword)
        bytIV = CreateIV("secureNotepad")
        Using ms As New MemoryStream
            Using cs As New CryptoStream(ms, aes.CreateEncryptor(bytKey, bytIV), CryptoStreamMode.Write)
                Dim bytes() As Byte = Encoding.UTF8.GetBytes(clearText)
                cs.Write(bytes, 0, bytes.Length)
            End Using

            Return Convert.ToBase64String(ms.ToArray())
        End Using
    End Function

    Public Function Decrypt(ByVal cipherText As String, Optional ByVal FilePassword As String = "filePassword") As String
        Dim ms As New MemoryStream(Convert.FromBase64String(cipherText))
        Dim dataOffset As Integer = ms.Position
        Dim bytKey As Byte()
        Dim bytIV As Byte()

        Dim aes As New RijndaelManaged()
        bytKey = CreateKey(FilePassword)
        bytIV = CreateIV("secureNotepad")
        Using cs As New CryptoStream(ms, aes.CreateDecryptor(bytKey, bytIV), CryptoStreamMode.Read)
            Using sr As New StreamReader(cs, Encoding.UTF8)
                Return sr.ReadToEnd()
            End Using
        End Using
    End Function
End Module
