Imports System
Imports System.IO
Imports System.Management
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports Microsoft.VisualBasic

<ComVisible(False)>
Friend Class ExecutionValidation

    Friend Class LicenseCheckResult
        Friend Property Status As LicenseStatus
        Friend Property ExpiryDate As String
    End Class

    Friend Enum LicenseStatus
        Valid
        Expired
        Unauthorized
        InvalidFormat
    End Enum

    Private ReadOnly aesKey As String
    Private ReadOnly aesIV As String
    Private ReadOnly fakeErrors As String()

    Public Sub New()
        aesKey = "AshuVedantRaghav"
        aesIV = "AshuVedantRaghav"
        fakeErrors = New String() {
            "Runtime Error 0x80070005: Access is denied.",
            "Unhandled Exception: System.IO.FileLoadException",
            "Fatal Error: Missing dependency 'mscorlib.dll'.",
            "Memory read violation at address 0x00000014.",
            "System.NullReferenceException at MainModule()",
            "Unexpected token in config file at line 42.",
            "Access violation while reading registry key.",
            "System.BadImageFormatException: Invalid executable format.",
            "Fatal: CLR runtime initialization failed.",
            "Exception code 0xc0000005: Access violation."
        }
    End Sub

    ' Validates the trial/license state. The C:\PlateNova\PlateNova.tst cache file stores the
    ' timestamp of the last successful check; if the system clock is ever behind that stored
    ' timestamp (a classic trial-reset trick), validation is rejected as tampering.
    Friend Function IsLicenceValid() As Boolean
        Try
            Dim now As DateTime = DateTime.Now
            Dim path As String = "C:\PlateNova\PlateNova.tst"

            If File.Exists(path) Then
                Dim encryptedBase As String = File.ReadAllText(path)
                Dim decrypted As String = AES_Decrypt(encryptedBase)
                Dim cachedDate As DateTime
                If Not DateTime.TryParse(decrypted, cachedDate) Then
                    ShowFakeError()
                    Return False
                ElseIf DateTime.Compare(now, cachedDate) < 0 Then
                    ' Clock moved backwards since the last successful check - treat as tampering
                    ShowFakeError()
                    Return False
                End If
                ' now >= cachedDate: fall through to the real validation below
            End If

            Dim biosSerial As String = GetBiosSerial().ToUpper()
            If String.IsNullOrWhiteSpace(biosSerial) Then
                ShowFakeError()
                Return False
            End If

            Dim isValid As Boolean
            Dim trialCutoff As New DateTime(2025, 12, 31)
            If DateTime.Compare(DateTime.Today, trialCutoff) <= 0 Then
                ' Within the free trial window - no license file required yet
                isValid = True
            Else
                Dim licPath As String = "C:\PlateNova\ExecutionFile.lic"
                If Not File.Exists(licPath) Then
                    ShowFakeError()
                    Return False
                End If

                Dim checkResult As LicenseCheckResult = ParseLicenseFile(licPath, biosSerial)
                If checkResult.Status <> LicenseStatus.Valid Then
                    ShowFakeError()
                    Return False
                End If
                isValid = True
            End If

            If isValid Then
                Try
                    Dim plainText As String = now.ToString("yyyy-MM-dd HH:mm:ss")
                    Dim encryptedNow As String = AES_Encrypt(plainText)
                    If Not String.IsNullOrEmpty(encryptedNow) Then
                        File.WriteAllText(path, encryptedNow)
                    End If
                Catch
                    ' Ignore cache-write failures; they don't affect the validity result
                End Try
            End If

            Return isValid

        Catch ex As Exception
            ShowFakeError()
            Return False
        End Try
    End Function

    ' Each line of the license file is "<BIOS-SERIAL>=<expiry-date | forever | permanent>",
    ' individually AES-encrypted. Returns the first line matching biosSerial, else Unauthorized.
    Private Function ParseLicenseFile(licPath As String, biosSerial As String) As LicenseCheckResult
        Try
            For Each rawLine As String In File.ReadAllLines(licPath)
                Dim decrypted As String = AES_Decrypt(rawLine).Trim()
                If String.IsNullOrWhiteSpace(decrypted) OrElse decrypted.StartsWith("#") OrElse
                   decrypted.StartsWith("//") OrElse Not decrypted.Contains("=") Then
                    Continue For
                End If

                Dim parts As String() = decrypted.Split("="c)
                If parts.Length <> 2 Then Continue For

                Dim key As String = parts(0).Trim().ToUpper()
                Dim value As String = parts(1).Trim().ToLower()

                If biosSerial = key Then
                    If value = "forever" OrElse value = "permanent" Then
                        Return New LicenseCheckResult With {.Status = LicenseStatus.Valid, .ExpiryDate = "Forever"}
                    End If

                    Dim expiry As DateTime
                    If Not DateTime.TryParse(value, expiry) Then
                        Return New LicenseCheckResult With {.Status = LicenseStatus.InvalidFormat, .ExpiryDate = value}
                    ElseIf DateTime.Compare(DateTime.Today, expiry) > 0 Then
                        Return New LicenseCheckResult With {.Status = LicenseStatus.Expired, .ExpiryDate = expiry.ToString("yyyy-MM-dd")}
                    Else
                        Return New LicenseCheckResult With {.Status = LicenseStatus.Valid, .ExpiryDate = expiry.ToString("yyyy-MM-dd")}
                    End If
                End If
            Next

            Return New LicenseCheckResult With {.Status = LicenseStatus.Unauthorized, .ExpiryDate = ""}

        Catch ex As Exception
            Return New LicenseCheckResult With {.Status = LicenseStatus.InvalidFormat, .ExpiryDate = ex.Message}
        End Try
    End Function

    Private Function AES_Decrypt(encryptedBase64 As String) As String
        Try
            Dim keyBytes As Byte() = Encoding.UTF8.GetBytes(aesKey)
            Dim ivBytes As Byte() = Encoding.UTF8.GetBytes(aesIV)
            Dim buffer As Byte() = Convert.FromBase64String(encryptedBase64)

            Using aes As Aes = Aes.Create()
                aes.Key = keyBytes
                aes.IV = ivBytes
                aes.Padding = PaddingMode.PKCS7
                Using transform As ICryptoTransform = aes.CreateDecryptor()
                    Using stream As New MemoryStream(buffer)
                        Using cryptoStream As New CryptoStream(stream, transform, CryptoStreamMode.Read)
                            Using reader As New StreamReader(cryptoStream)
                                Return reader.ReadToEnd()
                            End Using
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Function AES_Encrypt(plainText As String) As String
        Try
            Dim keyBytes As Byte() = Encoding.UTF8.GetBytes(aesKey)
            Dim ivBytes As Byte() = Encoding.UTF8.GetBytes(aesIV)

            Using aes As Aes = Aes.Create()
                aes.Key = keyBytes
                aes.IV = ivBytes
                aes.Padding = PaddingMode.PKCS7
                Using transform As ICryptoTransform = aes.CreateEncryptor()
                    Using memoryStream As New MemoryStream()
                        Using cryptoStream As New CryptoStream(memoryStream, transform, CryptoStreamMode.Write)
                            Using writer As New StreamWriter(cryptoStream)
                                writer.Write(plainText)
                            End Using
                        End Using
                        Return Convert.ToBase64String(memoryStream.ToArray())
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Function GetBiosSerial() As String
        Try
            Dim searcher As New ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BIOS")
            For Each item As ManagementObject In searcher.Get()
                Dim serial As String = item("SerialNumber").ToString().Trim()
                If Not String.IsNullOrWhiteSpace(serial) Then
                    Return serial
                End If
            Next
        Catch ex As Exception
            ' Ignore and fall through to empty result
        End Try
        Return ""
    End Function

    ' Shows a random fake system-error message (a deterrent, not a real crash) and, best-effort,
    ' closes the active workbook without saving so the user loses no work but can't continue.
    Private Sub ShowFakeError()
        Try
            Dim rnd As New Random()
            Dim prompt As String = fakeErrors(rnd.Next(fakeErrors.Length))
            Interaction.MsgBox(prompt, MsgBoxStyle.Critical, "Fatal Error")

            Try
                Dim excelApp As Object = Nothing
                Try
                    excelApp = Marshal.GetActiveObject("Excel.Application")
                Catch
                    ' Excel isn't running - nothing more to do
                End Try

                If excelApp IsNot Nothing Then
                    Dim activeWorkbook As Object = excelApp.ActiveWorkbook
                    If activeWorkbook IsNot Nothing Then
                        activeWorkbook.Close(False)
                    End If
                    Marshal.ReleaseComObject(excelApp)
                End If
            Catch
                ' Ignore cleanup errors
            End Try
        Catch
            ' This is a best-effort deterrent and must never throw
        End Try
    End Sub

    Friend Function GetCurrentBiosSerial() As String
        Return GetBiosSerial()
    End Function

    Friend Function ValidateLicenseFileFormat(licPath As String) As Boolean
        Try
            If Not File.Exists(licPath) Then Return False

            For Each rawLine As String In File.ReadAllLines(licPath)
                Dim decrypted As String = AES_Decrypt(rawLine).Trim()
                If String.IsNullOrWhiteSpace(decrypted) OrElse decrypted.StartsWith("#") OrElse Not decrypted.Contains("=") Then
                    Continue For
                End If

                Dim parts As String() = decrypted.Split("="c)
                If parts.Length = 2 AndAlso Not String.IsNullOrWhiteSpace(parts(0)) AndAlso Not String.IsNullOrWhiteSpace(parts(1)) Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class
