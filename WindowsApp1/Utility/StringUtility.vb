Public Class StringUtility

    Public Shared Function ByteLeft(value As String, len As Integer) As String
        If len < 0 Then
            Return String.Empty
        End If
        Dim byteSJIS As Byte() = System.Text.Encoding.GetEncoding("Shift-JIS").GetBytes(value)
        Return System.Text.Encoding.Default.GetString(byteSJIS, 0, len)
    End Function

End Class
