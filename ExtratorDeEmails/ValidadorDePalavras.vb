Public Class ValidadorDePalavras
    Public Function ehEmail(ByVal Palavra As String) As Boolean
        Dim resultadoValidacao As Boolean = Palavra.EndsWith(".br") AndAlso Palavra.IndexOf("@") > 0
        Return resultadoValidacao
    End Function


End Class
