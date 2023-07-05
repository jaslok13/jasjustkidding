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
    Public Shared Function FindAndGetFirstRowIndex(dtInput As System.Data.DataTable, strFilterQuery As String) As Integer
        Dim intRowIndex As Integer = -1
        Return intRowIndex
    End Function
End Class
