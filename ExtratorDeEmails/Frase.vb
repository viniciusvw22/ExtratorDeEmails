Public Class Frase
    Private meuTexto As String = ""
    Private delimitadoresDePalavras() As String = {}
    Private pesquisadorSimbolos As PesquisadorDeSimbolos = Nothing
    Private delimitadorPalavra As DelimitadorExtraido = Nothing

    Sub New(ByVal frase As String)
        Me.meuTexto = frase
        Me.delimitadoresDePalavras = {"Email:", "|", " "}
        Me.pesquisadorSimbolos = New PesquisadorDeSimbolos(meuTexto, delimitadoresDePalavras)

    End Sub

    Public Function encontrarInicioDaFrase(ByVal PosInicial As Integer) As Integer
        'Busca o início da frase, isto é, qualquer caractere que não seja um delimitador.
        Dim IndexInicioDaFrase As Integer = -1
        Dim i As Integer = PosInicial
        While i <= meuTexto.Length - 1
            delimitadorPalavra = pesquisadorSimbolos.TentarExtrairDelimitador(i)
            If delimitadorPalavra.foiExtraidoComSucesso() Then
                'Continuar procurando, logo depois do fim do delimitador.
                i += delimitadorPalavra.getValor().ToString().Length - 1
            Else
                'Significa que é uma parte da palavra, logo deverá sair do Loop. 
                IndexInicioDaFrase = i
                Exit While
            End If
            i += 1
        End While
        delimitadorPalavra = Nothing
        Return IndexInicioDaFrase
    End Function

    Public Function encontrarFimDaFrase(ByVal PosInicial As Integer) As Integer
        'Encontra a posição do delimitador do email buscando a partir de (PosInicial).
        'Se não houver delimitadores, retorna -1.
        Dim IndexFinalDaFrase As Integer = -1

        Dim i As Integer = PosInicial
        While i <= meuTexto.Length - 1
            delimitadorPalavra = pesquisadorSimbolos.TentarExtrairDelimitador(i)
            If delimitadorPalavra.foiExtraidoComSucesso() Then
                IndexFinalDaFrase = i
                Exit While
            Else
                i += 1
            End If
        End While
        Return IndexFinalDaFrase
    End Function

    Public Sub ExtrairEmails()
        Dim PosFimEmail As Integer = 0
        Dim PosInicioEmail As Integer = 0
        Dim Auxiliar As Integer = 0
        Dim Email As String = ""
        Dim verificarPalavra As New ValidadorDePalavras()
        Do
            PosInicioEmail = encontrarInicioDaFrase(Auxiliar)
            If PosInicioEmail = -1 Then
                Exit Do
            Else
                PosFimEmail = encontrarFimDaFrase(PosInicioEmail)
                If PosFimEmail = -1 Then
                    Email = meuTexto.Substring(PosInicioEmail)
                Else
                    Email = meuTexto.Substring(PosInicioEmail, PosFimEmail - PosInicioEmail)
                End If
                If verificarPalavra.ehEmail(Email) Then
                    Console.WriteLine(Email)
                End If
                Auxiliar = PosFimEmail
            End If
        Loop While PosInicioEmail <> -1 AndAlso PosFimEmail <> -1
    End Sub

End Class
