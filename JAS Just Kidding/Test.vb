Imports JASRPAFunctions
Imports Microsoft.VisualBasic.Interaction
Module Test
    Dim boolSuccess As Boolean = False
    Dim strRemarks As String = String.Empty
    Sub Main()
        Try
            OLEDB.RunSelectSQLAccessDB("D:\Power Automate\Students.accdb", "Select* from Students")
            Dim resultDateTime As Date = JASRPAFunctions.DateTimeFunctions.TextToDate("12 May 2019", "dd MMM yyyy")
            Dim result = JASRPAFunctions.CertificateFunctions.CheckPIDCertificateExist(strPID:="M390538")
            Dim dtCertList As System.Data.DataTable = JASRPAFunctions.CertificateFunctions.GetCertificateList
            Dim strFilePath As String = "D:\JAS\DBS Transactions.csv"
            Dim dtCSVData As System.Data.DataTable = JASRPAFunctions.TextFileFunctions.GetCSVData(strFilePath)
            Dim dtCSVText As String = JASRPAFunctions.TextFileFunctions.GetText(strFilePath)
            Microsoft.VisualBasic.Interaction.MsgBox("Hello")
            Console.ReadLine()
        Catch ex As Exception
            boolSuccess = False
            strRemarks = ex.ToString
        Finally
            Console.ReadLine()
        End Try
    End Sub
End Module
