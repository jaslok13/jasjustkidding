Imports JasDLL
Module TestDLL
    Dim boolSuccess As Boolean = False
    Dim strRemarks As String = String.Empty
    Sub Main()
        Try
            'Dim dtCSVData As System.Data.DataTable = JasDLL.TextFile_Functions.Get_CSV_Data("D:\JAS\DBS Transactions.csv")
            Dim dtCSVText As String = JasDLL.TextFile_Functions.Get_Text("D:\JAS\DBS Transactions.csv")

            Console.ReadLine()
        Catch ex As Exception
            boolSuccess = False
            strRemarks = ex.ToString
        Finally
            Console.ReadLine()
        End Try
    End Sub
End Module
