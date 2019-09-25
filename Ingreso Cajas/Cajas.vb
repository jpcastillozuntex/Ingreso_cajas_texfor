Module cajas_01
    Dim obj As New empresas
    Dim cnn As New SqlClient.SqlConnection

    Private Function recorta(ByVal texto As String, ByVal t As Integer) As String
        Dim tx As String
        tx = Mid(texto, 1, t)
        Return tx
    End Function


    Public Function modifica_caja(ByRef co As DataTable, ByVal caja As String, ByVal tipo As String, ByVal talla As String, ByVal uni As Integer, ByRef orden As Integer, ByRef unidad As Integer) As Boolean
        Dim dd As DataRow()
        Dim dr As DataRow = Nothing
        Dim ok As Boolean = False
        dd = co.Select("CAJA = '" & caja & "' AND TIPO = '" & tipo & "' AND TALLA = '" & talla & "'")
        If dd.Length > 0 Then
            dr = dd(0)
            dr("UNIDADES") = dr("UNIDADES") - uni
            unidad = dr("UNIDADES")
            orden = dr("ORDEN")
        End If
        Return ok
    End Function

    Public Function saldo_caja(ByRef co As DataTable, ByVal caja As String, ByVal tipo As String, ByVal talla As String, ByRef orden As String, ByRef fecha As String) As Integer
        Dim dd As DataRow()
        Dim dr As DataRow = Nothing
        Dim unidades As Integer = 0
        dd = co.Select("CAJA ='" & caja & "' AND TIPO = '" & tipo & "' AND TALLA = '" & talla & "'")
        If dd.Length > 0 Then
            dr = dd(0)
            unidades = dr("UNIDADES")
            orden = dr("ORDEN")
            fecha = Format(dr("FECHA"), "yyyy-MM-dd HH:mm:ss")
        End If
        Return unidades
    End Function

End Module
