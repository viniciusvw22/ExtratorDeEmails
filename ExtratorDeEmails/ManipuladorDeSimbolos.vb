Public Class DelimitadorExtraido
    Private ExtraidoComSucesso As Boolean = False
    Private Valor As String = ""

    Public Sub DefinirValores(ByVal extraidoComSucesso As Boolean, ByVal valor As String)
        Me.ExtraidoComSucesso = extraidoComSucesso
        Me.Valor = valor
    End Sub
    Public Function foiExtraidoComSucesso() As Boolean
        Return ExtraidoComSucesso
    End Function
    Public Function getValor() As String
        Return Valor
    End Function
End Class

Public Class PesquisadorDeSimbolos
    Private TextoAnalisado As String = ""
    Private DelimitadoresProcurados() As String = {}

    Sub New(ByVal textoAnalisado As String, ByVal caracteresProcurados() As String)
        Me.TextoAnalisado = textoAnalisado
        Me.DelimitadoresProcurados = caracteresProcurados
    End Sub

    Public Function TentarExtrairDelimitador(ByVal posi As Integer) As DelimitadorExtraido
        Dim delimExtraido As New DelimitadorExtraido()
        Dim tam As Integer = 0
        Dim tam_texto As Integer = TextoAnalisado.Length
        For Each delim As String In DelimitadoresProcurados
            tam = delim.Length
            If tam + posi <= tam_texto AndAlso TextoAnalisado.Substring(posi, tam) = delim Then
                delimExtraido.DefinirValores(extraidoComSucesso:=True, valor:=delim)
                Exit For
            End If
        Next
        Return delimExtraido
    End Function

    Public Function PosicaoContemDelimitador(ByVal posi As Integer) As Boolean
        Dim tam As Integer = 0
        Dim tam_meuTexto As Integer = TextoAnalisado.Length
        For Each delim As String In DelimitadoresProcurados
            tam = delim.Length
            If tam + posi <= tam_meuTexto AndAlso TextoAnalisado.Substring(posi, tam) = delim Then
                Return True
            End If
        Next
        Return False
    End Function
End Class
