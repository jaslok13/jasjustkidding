Module OLEDB
    Dim boolSuccess As Boolean = False
    Dim strRemarks As String = String.Empty
    Public Sub RunSelectSQLAccessDB(strFilePathAccessDB As String, strSelectSQL As String, Optional strDBPassword As String = "")
        Dim dtResults As New System.Data.DataTable
        Dim dbConnection As System.Data.OleDb.OleDbConnection = Nothing
        Try
            If Not String.IsNullOrEmpty(strFilePathAccessDB) Then
                If System.IO.File.Exists(strFilePathAccessDB) Then
                    If Not String.IsNullOrEmpty(strSelectSQL) Then
                        Dim strConnection As String = ""
                        If String.IsNullOrEmpty(strDBPassword) Then
                            strConnection = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & strFilePathAccessDB & ";"
                        Else
                            strConnection = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & strFilePathAccessDB & ";Jet OLEDB:Database Password=" & strDBPassword & ";"
                        End If

                        'Open connection
                        dbConnection = New System.Data.OleDb.OleDbConnection(strConnection)
                        dbConnection.Open()

                        'Run Query
                        'Dim dbDataAdapter As System.Data.OleDb.OleDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter(selectCommandText:=strSelectSQL, selectConnection:=dbConnection)
                        Dim dbDataAdapter As New System.Data.OleDb.OleDbDataAdapter(selectCommandText:=strSelectSQL, selectConnection:=dbConnection)
                        Dim dsResults As New System.Data.DataSet
                        dbDataAdapter.Fill(dsResults)
                        dtResults = dsResults.Tables(0)
                    Else
                        strRemarks = "SQL is blank"
                    End If
                Else
                    strRemarks = "File not found: " & strFilePathAccessDB
                End If
            Else
                strRemarks = "Path for Database is blank"
            End If
        Catch ex As Exception
            strRemarks = ex.ToString
        Finally
            Try
                'Try to close the DB connection
                dbConnection.Close()
                dbConnection.Dispose()
            Catch ex As Exception
                'Eat the Exception
            End Try
        End Try
    End Sub

    Public Function RunInsertUpdateDeleteSQLAccessDB() As Integer
        Dim intResult As Integer = 0

        Return intResult
    End Function
End Module
