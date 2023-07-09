Public Class CertificateFunctions
    Public Shared Function GetCertificateList() As System.Data.DataTable
        Dim dtCertList As New System.Data.DataTable
        Dim strRegexPatternPID As String = "(?<=\()(.*?)(?=\)\,)"
        Dim strRegexPatternOU As String = "(?<=OU=)(.*?)(?=\,)"

        Dim store As System.Security.Cryptography.X509Certificates.X509Store = New System.Security.Cryptography.X509Certificates.X509Store(System.Security.Cryptography.X509Certificates.StoreName.My, System.Security.Cryptography.X509Certificates.StoreLocation.CurrentUser)
        store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.ReadOnly Or System.Security.Cryptography.X509Certificates.OpenFlags.OpenExistingOnly)

        If store.Certificates IsNot Nothing And store.Certificates.Count > 0 Then
            dtCertList.Columns.Add("PID")
            dtCertList.Columns.Add("TestCert", GetType(System.Boolean))
            dtCertList.Columns.Add("NVSCCert", GetType(System.Boolean))
            dtCertList.Columns.Add("ExpiryDate", GetType(System.DateTime))
            dtCertList.Columns.Add("Expired", GetType(System.Boolean))

            'Loop through list of User certificates
            Dim boolCertExpired, boolTestCert, boolNVSCCert As Boolean
            Dim strPID, strOU As String
            For Each cert As System.Security.Cryptography.X509Certificates.X509Certificate2 In store.Certificates
                boolCertExpired = False
                boolTestCert = False
                boolNVSCCert = False

                strPID = System.Text.RegularExpressions.Regex.Match(input:=cert.Subject.ToUpper, pattern:=strRegexPatternPID).Value.Trim
                strOU = System.Text.RegularExpressions.Regex.Match(input:=cert.Subject.ToUpper, pattern:=strRegexPatternOU).Value.Trim

                If Microsoft.VisualBasic.DateAndTime.DateDiff(Interval:=Microsoft.VisualBasic.DateInterval.Day, Date1:=cert.NotAfter.Date, Date2:=DateAndTime.Now.Date) > 0 Then
                    boolCertExpired = True
                End If

                If cert.Issuer.ToLower.Contains("text") Then
                    boolTestCert = True
                End If

                If strOU.ToUpper.Trim.Equals("VSC") Then
                    boolNVSCCert = True
                End If

                dtCertList.Rows.Add(strPID, boolTestCert, boolNVSCCert, cert.NotAfter, boolCertExpired)
            Next
            If dtCertList.Rows.Count > 0 Then
                dtCertList = dtCertList.DefaultView.ToTable(True)
            End If
        End If

        Return dtCertList
    End Function
End Class
