Module teclaspermitidas
    Function numerosspermitidos(ByVal KeyAscii As Integer) As Integer
        'Transformar letras minusculas em Maiúsculas
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        ' Intercepta um código ASCII recebido e admite somente letras
        If InStr("0123456789,.", Chr(KeyAscii)) = 0 Then
            numerosspermitidos = 0
        Else
            numerosspermitidos = KeyAscii
        End If
        ' teclas adicionais permitidas
        If KeyAscii = 8 Then numerosspermitidos = KeyAscii ' Backspace
        If KeyAscii = 127 Then numerosspermitidos = KeyAscii ' Delete
    End Function
End Module
