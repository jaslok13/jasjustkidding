Imports Microsoft.VisualBasic
Imports System.Security.Cryptography.X509Certificates
Imports System.Collections.Generic
Imports System.Linq

Module Certificates
    Public flag_Success As Boolean = False
    Public text_Remarks As String = String.Empty
    Sub Main()
        Try
            Get_PID_Certificate_List()


        Catch ex As Exception
            flag_Success = False
            text_Remarks = ex.ToString
        End Try
    End Sub


    Private Sub Get_PID_Certificate_List()
        Dim text_RegexPattern_PID As String = ""
        Dim text_RegexPattern_OU As String = ""
        Dim dt_PID_Cert_List As New System.Data.DataTable
        Dim store As System.Security.Cryptography.X509Certificates.X509Store = New System.Security.Cryptography.X509Certificates.X509Store(System.Security.Cryptography.X509Certificates.StoreName.My, System.Security.Cryptography.X509Certificates.StoreLocation.CurrentUser)
        Try
            store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.ReadOnly Or System.Security.Cryptography.X509Certificates.OpenFlags.OpenExistingOnly)
            If (store.Certificates IsNot Nothing And store.Certificates.Count > 0) Then
                dt_PID_Cert_List.Columns.Add("PID")
                dt_PID_Cert_List.Columns.Add("Test Certificate", GetType(System.Boolean))
                dt_PID_Cert_List.Columns.Add("NVSC Certificate", GetType(System.Boolean))
                dt_PID_Cert_List.Columns.Add("Expiry Date", GetType(System.DateTime))
                dt_PID_Cert_List.Columns.Add("Expired", GetType(System.Boolean))
                'Loop through the list of user certificates and filter CS certs
                Dim Cert_Expired, Text_Certificate, NVSC_Certificate As Boolean
                Dim PID, OU As String

                For Each pid_cert As System.Security.Cryptography.X509Certificates.X509Certificate2 In store.Certificates
                    If pid_cert.Issuer.Contains("Credit Suisse") Then
                        Cert_Expired = False
                        Text_Certificate = False
                        NVSC_Certificate = False
                        PID = System.Text.RegularExpressions.Regex.Match(input:=pid_cert.Subject.ToUpper, pattern:=text_RegexPattern_PID).Value.Trim
                        OU = System.Text.RegularExpressions.Regex.Match(input:=pid_cert.Subject.ToUpper, pattern:=text_RegexPattern_OU).Value.Trim

                        If Microsoft.VisualBasic.DateAndTime.DateDiff(Interval:=Microsoft.VisualBasic.DateInterval.Day, Date1:=pid_cert.NotAfter.Date, Date2:=DateAndTime.Now.Date) > 0 Then
                            Cert_Expired = True
                        End If

                        If pid_cert.Issuer.ToLower.Contains("test") Then
                            Text_Certificate = True
                        End If

                        If OU.Equals("VSC") Then
                            NVSC_Certificate = True
                        End If
                        dt_PID_Cert_List.Rows.Add(PID, Text_Certificate, NVSC_Certificate, pid_cert.NotAfter, Cert_Expired)
                    End If
                Next
                If dt_PID_Cert_List.Rows.Count > 0 Then
                    'Filter unique rows
                    dt_PID_Cert_List = dt_PID_Cert_List.DefaultView.ToTable(True)
                    flag_Success = True
                End If
            Else
                flag_Success = False
                text_Remarks = "No PID Certificates found"
            End If


        Catch ex As Exception
            flag_Success = False
            text_Remarks = ex.ToString
        End Try
    End Sub
End Module
