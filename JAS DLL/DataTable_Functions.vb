Public Class DataTable_Functions
    Public Shared Function Get_RowCount(dtInput As System.Data.DataTable) As Integer
        Return dtInput.Rows.Count
    End Function
    Public Shared Function Get_ColumnCount(dtInput As System.Data.DataTable) As Integer
        Return dtInput.Columns.Count
    End Function
End Class
