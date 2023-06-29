Public Class TextFile_Functions
    Public Shared Function Get_CSV_Data(strCSVFilePath As String) As System.Data.DataTable
        Dim dtCSV_Data As New System.Data.DataTable
        If System.IO.File.Exists(strCSVFilePath) Then
            Dim tfp As Microsoft.VisualBasic.FileIO.TextFieldParser = New Microsoft.VisualBasic.FileIO.TextFieldParser(strCSVFilePath)
            'tfp.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
            tfp.SetDelimiters(",")

            Dim intRowIndex As Integer = 0
            While Not tfp.EndOfData
                intRowIndex = intRowIndex + 1
                Dim intFieldIndex As Integer = 0
                For Each field As String In tfp.ReadFields
                    If intRowIndex = 1 Then
                        dtCSV_Data.Columns.Add(field)
                    Else
                        If intFieldIndex = 0 Then
                            dtCSV_Data.Rows.Add()
                        Else
                            dtCSV_Data.Rows(dtCSV_Data.Rows.Count - 1).Item(intFieldIndex) = field
                        End If
                    End If
                    intFieldIndex = intFieldIndex + 1
                Next
            End While
        Else
            Throw New Exception("File not found: " & strCSVFilePath)
        End If
        Return dtCSV_Data
    End Function

    Public Shared Function Get_Text(strTextFilePath As String) As String
        Dim strCSV_Text As String = String.Empty
        If System.IO.File.Exists(strTextFilePath) Then
            strCSV_Text = My.Computer.FileSystem.ReadAllText(strTextFilePath)
        Else
            Throw New Exception("File not found: " & strTextFilePath)
        End If
        Return strCSV_Text
    End Function
End Class

