
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Globalization
Imports System.Security.Cryptography


Module SupEncrt

    Public AdminKeyAccept As String = "ALLANPHILIP"

    Public GetKeyNow As Boolean = False

#Region "KEYtoKEY"
    Public Function Encrypt(clearText As String) As String
        Dim EncryptionKey As String = "0LDH@B!TD!3H@RD2017"
        Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, _
             &H65, &H64, &H76, &H65, &H64, &H65, _
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                    cs.Write(clearBytes, 0, clearBytes.Length)
                    cs.Close()
                End Using
                clearText = Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
        Return clearText
    End Function

    Public Function Decrypt(cipherText As String) As String
        Dim EncryptionKey As String = "0LDH@B!TD!3H@RD2017"
        Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, _
             &H65, &H64, &H76, &H65, &H64, &H65, _
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                    cs.Write(cipherBytes, 0, cipherBytes.Length)
                    cs.Close()
                End Using
                cipherText = Encoding.Unicode.GetString(ms.ToArray())
            End Using
        End Using
        Return cipherText
    End Function
#End Region


#Region "Clear Cache"
    Declare Function SetProcessWorkingSetSize Lib "kernel32.dll" (ByVal process As IntPtr, ByVal minimumWorkingSetSize As Integer, ByVal maximumWorkingSetSize As Integer) As Integer
    Public Sub FlushMemory()
        Try
            GC.Collect()
            GC.WaitForPendingFinalizers()
            If (Environment.OSVersion.Platform = PlatformID.Win32NT) Then
                SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1)
                'Dim myProcesses As Process() = Process.GetProcessesByName("Applica­tionName")
                Dim myProcesses As Process() = Process.GetProcessesByName("PrjAutoBk.exe")
                Dim myProcess As Process 'Dim ProcessInfo As Process 
                For Each myProcess In myProcesses
                    SetProcessWorkingSetSize(myProcess.Handle, -1, -1)
                Next myProcess
            End If
        Catch ex As Exception
        End Try
    End Sub

#End Region

End Module
