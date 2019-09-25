Imports C1.Win.C1FlexGrid
Imports System.Drawing.Printing
Public Class Cajas99
    Dim dt As New DataTable
    Dim strsql As String
    Dim cnn As New SqlClient.SqlConnection
    Dim con(3) As String
    Dim obj As New empresas
    Dim seccion As String = obj.seccion
    Dim sec As String = "('" & obj.seccion & "','TEXFOR" & Mid(obj.seccion, 7) & "')"
    Private Sub Plan_costura_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AddHandler caja.KeyPress, AddressOf keypressed1
        conexiones(con)
        Label2.Text = seccion
        setea_fg()
        caja.Focus()
    End Sub
    Private Sub setea_fg()
        fg.Rows.Count = 1
        fg.Rows(0).Height = 30
    End Sub
    Private Sub limpia_caja()
        busca_produccion()
        caja.Text = ""
        caja.Focus()
    End Sub

    Private Sub busca_produccion()
        Dim i As Integer
        Dim j As Integer = 1
        Dim k As Integer = 1
        Dim fecha As String = Format(Today, "yyyy-MM-dd")
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim sec As String = obj.seccion
        Dim strsql As String = "SELECT SECCION,CAJAS04.CORTE,SUM(CAJAS04.UNIDADES) AS PRENDAS FROM CAJAS04 LEFT JOIN CAJAS01 ON CAJAS04.CAJA = CAJAS01.CAJA AND CAJAS04.CORTE = CAJAS01.CORTE AND CAJAS04.TIPO = CAJAS01.TIPO AND CAJAS04.TALLA = CAJAS01.TALLA WHERE CONVERT(date,CAJAS04.FECHA) = '" & fecha & "' AND SECCION = '" & sec & "' GROUP BY SECCION,CAJAS04.CORTE"
        '  llena_tablas(dt, strsql, cnn)
        jg.Rows.Count = 1
        For i = 1 To 3
            Try
                dt = New DataTable
                cnn.ConnectionString = con(i)
                llena_tablas_e(dt, strsql, cnn)
                For Each dr In dt.Rows
                    jg.Rows.Count = j + 1
                    jg(j, 1) = dr("CORTE")
                    jg(j, 2) = dr("PRENDAS")
                    j = j + 1
                Next
            Catch
            End Try
        Next
    End Sub
    Private Sub chequea_datos()
        Dim ok As Boolean
        busca_caja(dt, ok)
        If ok Then
            actualiza(dt)
        Else
            limpia_caja()
        End If
    End Sub
    Private Sub actualiza(ByVal dt As DataTable)
        Dim l As Integer = fg.Rows.Count
        Dim ok As Boolean
        Graba_datos(dt, ok)
        fg.Rows.Count = fg.Rows.Count + dt.Rows.Count
            For Each dr In dt.Rows
                fg(l, 1) = caja.Text
                fg(l, 2) = dr("CORTE")
                fg(l, 3) = dr("TALLA")
                fg(l, 4) = dr("UNIDADES")
            fg(l, 5) = dr("SECCION")
            If ok Then
                fg(l, 6) = True
            Else
                fg(l, 6) = False
            End If
            fg(l, 7) = dr("FECHA")
            l = l + 1
        Next
        fg.Sort(SortFlags.Descending, 7)
        limpia_caja()

    End Sub
    Private Sub busca_caja(ByRef dt As DataTable, ByRef ok As Boolean)
        Dim up As New DataTable
        ok = False
        dt = New DataTable
        Dim dr As DataRow = Nothing
        Dim strsql As String = "SELECT * FROM CAJAS01 LEFT JOIN CORTES ON CAJAS01.CORTE = CORTES.CORTE WHERE CAJA = '" & caja.Text & "'"
        llena_tablas(dt, strsql, cnn)
        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
        Else
            Exit Sub
        End If
        llena_tablas(up, "SELECT * FROM UPC WHERE CLIENTE = '" & dr("CLIENTE") & "' AND ESTILO = '" & dr("ESTILO") & "' AND COLOR = '" & dr("COLOR") & "'", cnn)
        If up.Rows.Count > 0 Then
            MsgBox("Este Corte tiene UPC registrado. Use el programa de scan & Pack !!!", MsgBoxStyle.Critical, "Por favor revise.")
        Else
            If InStr(sec, dr("SECCION")) = False Then
                MsgBox("Este Corte es de otra Sección. !!!", MsgBoxStyle.Critical, "Por favor revise.")
            Else

                ok = True
            End If
        End If
    End Sub

    Private Sub keypressed1(ByVal o As [Object], ByVal e As KeyPressEventArgs)
        Dim cajas As String = caja.Text
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            If cajas.Length = 9 Then
                My.Computer.Audio.Play("c:\scan\beep.wav")
                caja.Text = UCase(caja.Text)
                chequea_datos()
            Else
                MsgBox("Datos incorrectos !!!!!", MsgBoxStyle.Critical, "Por favor revise !!!")
                limpia_caja()
            End If
        End If
    End Sub 'keypressed
    Private Sub Graba_datos(ByVal dt As DataTable, ByRef ok As Boolean)
        Dim afectados As Integer = 0
        Dim strsql As String
        Dim dr As DataRow
        Dim obj As New empresas
        ok = False
        cnn.Open()
        ' Comienza la transacción
        Dim transaccion As SqlClient.SqlTransaction = cnn.BeginTransaction()
        ' Crea el comando para la transacción
        Dim comando As SqlClient.SqlCommand = cnn.CreateCommand()
        comando.Transaction = transaccion

        Try
            For Each dr In dt.Rows
                Try
                    strsql = "INSERT INTO CAJAS04 (CAJA,CORTE,TALLA,TIPO,UNIDADES,FECHA ,QUIEN) " & _
                                 "VALUES ( '" & dr("CAJA") & "','" & _
                                                dr("CORTE") & "','" & _
                                                dr("TALLA") & "','" & _
                                                dr("TIPO") & "','" & _
                                                dr("UNIDADES") & "',GETDATE(),'OPERADOR')"
                    comando.CommandText = strsql
                    afectados = afectados + comando.ExecuteNonQuery()
                    ok = True
                Catch
                End Try
            Next
            If afectados > 0 Then
                strsql = "UPDATE CAJAS01 SET ESTADO = 'P' WHERE CAJA = '" & caja.Text & "'"
                comando.CommandText = strsql
                afectados = comando.ExecuteNonQuery()
            End If
            transaccion.Commit()

        Catch e As Exception
            Try
                MsgBox("Inconsistencia en Datos", MsgBoxStyle.Critical, "Por favor revise !!!!")
                transaccion.Rollback()
            Catch ex As SqlClient.SqlException
                If Not transaccion.Connection Is Nothing Then
                    MsgBox("Ocurrió un error de tipo " & ex.GetType().ToString() & _
                                      " se generó cuando se trato de eliminar la transacción.", MsgBoxStyle.Critical, "Error")
                End If
            End Try
        Finally
            cnn.Close()
        End Try
    End Sub

    
End Class

