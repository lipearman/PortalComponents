Imports System
Imports System.IO
Imports System.Xml
Imports System.Text
Imports System.Security.Cryptography

Public Class Cryptography
    Private Shared key() As Byte = {}
    Private Shared IV() As Byte = {&H12, &H34, &H56, &H78, &H90, &HAB, &HCD, &HEF}
    Private Shared sEncryptionKey As String = "!#$+54?3"
    'We're going to use a hard-to-crack 8 byte key, !#$a54?3, to keep parts of our QueryString secret.
    Public Shared Function Decode(ByVal stringToDecrypt As String) As String
        stringToDecrypt = Replace(stringToDecrypt, " ", "+", 1, -1)
        'Dim inputByteArray(stringToDecrypt.Length) As Byte
        Try
            Dim inputByteArray(stringToDecrypt.Length) As Byte

            key = System.Text.Encoding.UTF8.GetBytes(Left(sEncryptionKey, 8))
            Dim des As New DESCryptoServiceProvider
            inputByteArray = Convert.FromBase64String(stringToDecrypt)
            Dim ms As New MemoryStream
            Dim cs As New CryptoStream(ms, des.CreateDecryptor(key, IV), _
                CryptoStreamMode.Write)
            cs.Write(inputByteArray, 0, inputByteArray.Length)
            cs.FlushFinalBlock()
            Dim encoding As System.Text.Encoding = System.Text.Encoding.UTF8
            Return encoding.GetString(ms.ToArray())
        Catch ex As Exception
            'Throw New ArgumentException(ex.ToString) 'New ArgumentException("Invalid String")
            Return ""
        End Try
    End Function
    Public Shared Function Encode(ByVal stringToEncrypt As String) As String
        Try
            key = System.Text.Encoding.UTF8.GetBytes(Left(sEncryptionKey, 8))
            Dim des As New DESCryptoServiceProvider
            Dim inputByteArray() As Byte = Encoding.UTF8.GetBytes( _
                stringToEncrypt)
            Dim ms As New MemoryStream
            Dim cs As New CryptoStream(ms, des.CreateEncryptor(key, IV), _
                CryptoStreamMode.Write)
            cs.Write(inputByteArray, 0, inputByteArray.Length)
            cs.FlushFinalBlock()
            Return Convert.ToBase64String(ms.ToArray())
        Catch ex As Exception
            'Throw New ArgumentException(ex.ToString) 'New ArgumentException("Invalid String")
            Return ""
        End Try
    End Function

End Class


