Public Class DataTableFunctions
    Public Shared Function GetRowCount(dtInput As System.Data.DataTable) As Integer
        Return dtInput.Rows.Count
    End Function
    Public Shared Function GetColumnCount(dtInput As System.Data.DataTable) As Integer
        Return dtInput.Columns.Count
    End Function
    Public Shared Function GetDistinctRows(dtInput As System.Data.DataTable) As System.Data.DataTable
        If dtInput.Rows.Count > 0 Then
            Return dtInput.DefaultView.ToTable(True)
        Else
            Return dtInput
        End If
    End Function
    Public Function FindAndGetFirstRowIndex(dtInput As System.Data.DataTable, strFilterQuery As String) As Integer
        Dim intRowIndex As Integer = -1
        Return intRowIndex
    End Function
    Public Shared Function FindAndGetRowCount(dtInput As System.Data.DataTable, strFilterQuery As String) As Integer
        Dim intRowCount As Integer = 0
        If dtInput.Rows.Count > 0 Then
            intRowCount = dtInput.Select("" & strFilterQuery & "").Count
        Else
            Throw New System.Exception("No rows in input DataTable")
        End If
        Return intRowCount
    End Function
End Class
