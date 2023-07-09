Imports JASRPAFunctions
Module TestDLL
    Dim boolSuccess As Boolean = False
    Dim strRemarks As String = String.Empty
    Sub Main()
        Try
            Dim dtCertList As System.Data.DataTable = JASRPAFunctions.CertificateFunctions.GetCertificateList
            Dim strFilePath As String = "D:\JAS\DBS Transactions.csv"
            Dim dtCSVData As System.Data.DataTable = JASRPAFunctions.TextFileFunctions.GetCSVData(strFilePath)
            Dim dtCSVText As String = JASRPAFunctions.TextFileFunctions.GetText(strFilePath)

            Console.ReadLine()
        Catch ex As Exception
            boolSuccess = False
            strRemarks = ex.ToString
        Finally
            Console.ReadLine()
        End Try
    End Sub
End Module
