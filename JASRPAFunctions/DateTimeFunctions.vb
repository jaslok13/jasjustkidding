Public Class DateTimeFunctions
    Public Shared Function TextToDate(strDate As String, strDateFormat As String) As Date
        Return System.DateTime.ParseExact(s:=strDate, format:=strDateFormat, provider:=System.Globalization.DateTimeFormatInfo.InvariantInfo, style:=System.Globalization.DateTimeStyles.AssumeUniversal)
    End Function
End Class
