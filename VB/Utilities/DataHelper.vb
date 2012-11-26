Public Class DataHelper
    Public Shared Function DbObjectToString(dbValue As Object) As String
        If TypeOf dbValue Is DBNull Then
            Return dbValue.ToString()
        Else
            Return String.Empty
        End If
    End Function
End Class
