Imports System.IO
Imports C1.Win.C1FlexGrid
Imports System.Drawing.Printing
Imports C1.C1Excel
Imports System
Imports System.Text.RegularExpressions
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.Net.Mail
Imports System.Net
Module Module1
    Public _styles As Hashtable
    Public Sub setea_empresa(ByVal empresa As Integer)
        Dim obj As New empresas
        Dim cnstr As C1cajasLib_SIF.util = Nothing
        If empresa = 1 Then
            obj.conexion = "inventarios"
        ElseIf empresa = 2 Then
            obj.conexion = "TRECENTO"
        Else
            obj.conexion = "zuntex"
        End If
        obj.constr = cnstr.con_string(empresa - 1)
        obj.conole = cnstr.con_ole(empresa - 1)
    End Sub
    Public Sub def_sql(ByRef cnn As System.Data.SqlClient.SqlConnection)
        Dim obj As New empresas()
        cnn = New System.Data.SqlClient.SqlConnection()
        cnn.ConnectionString = obj.constr
    End Sub

    Public Sub Estado_Tela(ByVal con As Object, ByVal C4 As System.Windows.Forms.ComboBox)
        Dim dt As New DataTable()
        Dim dr As DataRow
        Dim cnn As New SqlClient.SqlConnection()
        Dim strSQL As String = "SELECT * FROM ESTADOS_TELA"
        llena_tablas(dt, strSQL, cnn)
        For Each dr In dt.Rows
            C4.Items.Add(dr("ESTADO_TELA"))
        Next
        C4.Items.Add("TODOS")
    End Sub

    Public Sub a_excel(ByVal fg As C1FlexGrid, ByVal path As String, ByRef ok As Boolean)
        Try
            Dim ch As Char = Microsoft.VisualBasic.Chr(9)
            If File.Exists(path) Then
                File.Delete(path)
            End If
            fg.SaveGrid(path, FileFormatEnum.Excel, FileFlags.IncludeFixedCells + FileFlags.VisibleOnly)
            System.Diagnostics.Process.Start(path)
            ok = True
        Catch
            MsgBox("Por favor cierre todas sus Hojas de Excel y vuelva a tratar. Gracias", MsgBoxStyle.OkOnly, "Atencion ")
            ok = False
        End Try
        If ok Then
            MsgBox("Sus datos fueron trasladados a Excel en el directorio: " + path, MsgBoxStyle.OkOnly, "TRASLADO DE DATOS ")
        End If
    End Sub

    Public Sub con_string(ByVal e As Integer, ByRef constr As String)
        Dim cnstr As C1cajasLib_SIF.util = Nothing
        constr = cnstr.con_string(e)
    End Sub
    Public Sub llena_clientes(ByRef C4 As System.Windows.Forms.ComboBox)
        Dim cnn As New System.Data.SqlClient.SqlConnection()
        Dim dt As New DataTable()
        Dim dr As DataRow
        Dim strsql As String = "SELECT CLIENTE FROM CLIENTES ORDER BY CLIENTE"
        C4.Items.Add("TODOS")
        llena_tablas(dt, strsql, cnn)
        For Each dr In dt.Rows
            C4.Items.Add(Trim(dr("CLIENTE")))
        Next
    End Sub
    Public Function get_codigo_color(ByVal cliente As String, ByVal color As String) As String
        Dim cnn As New SqlClient.SqlConnection
        Dim dt As New DataTable
        Dim dr As DataRow = Nothing
        Dim codigo As String = ""
        llena_tablas(dt, "SELECT * FROM COLORES WHERE CLIENTE = '" & cliente & "' AND COLOR = '" & color & "'", cnn)
        For Each dr In dt.Rows
            Try
                codigo = dr("CODIGO_C")
            Catch ex As Exception
            End Try
        Next
        Return codigo
    End Function
    Public Sub tallas_cortes(ByVal cpo As String, ByVal estilo As String, ByVal colo As String, ByRef fg As C1.Win.C1FlexGrid.C1FlexGrid, ByVal fil As Integer, ByVal col As Integer)
        Dim i As Integer
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim cnn As New SqlClient.SqlConnection
        Dim escala As String
        Try
            llena_tablas(dt, "SELECT * FROM CPO_D WHERE CPO = '" & cpo & "' AND ESTILO = '" & estilo & "' AND COLOR = '" & colo & "'", cnn)
            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                escala = dr("ESCALA")
                llena_tablas(dt, "SELECT * FROM E_TALLAS WHERE ESCALA = '" & escala & "'", cnn)
                If dt.Rows.Count > 0 Then
                    dr = dt.Rows(0)
                    For i = 1 To 10
                        fg(fil, col) = dr(i + 1)
                        col = col + 1
                    Next
                End If

            End If
        Catch
        End Try
    End Sub

    Public Sub get_escalas(ByRef ge As String)
        Dim cnn As New SqlClient.SqlConnection
        Dim dt As New DataTable
        Dim dr As DataRow
        llena_tablas(dt, "SELECT * FROM E_TALLAS ", cnn)
        For Each dr In dt.Rows
            ge = ge + dr("NOMBRE") + "|"
        Next
    End Sub
    Public Function lector_dc() As String
        Dim dc As String = "Data Source=JT;Initial Catalog=lector;Persist Security Info=True;User ID=user_l;Password=Lector_01"
        Return dc
    End Function
    Public Function get_tallas(ByVal escala As String) As DataRow
        Dim cnn As New SqlClient.SqlConnection
        Dim dt As New DataTable
        Dim dr As DataRow = Nothing
        llena_tablas(dt, "SELECT * FROM E_TALLAS WHERE ESCALA = '" & Format(CInt(escala), "00") & "'", cnn)
        For Each dr In dt.Rows
            get_tallas = dr
        Next
        Return dr
    End Function
    Public Sub talla_Grid(ByRef fg As C1.Win.C1FlexGrid.C1FlexGrid, ByVal col As Integer, ByVal adulto As Boolean)
        Dim t As String = "XS|S|M|L|XL|2XL|3XL|4XL|5XL|6XL"
        Dim co As String()
        Dim tit As String
        Dim i As Integer
        co = Split(t, "|")
        For i = col To col + 9
            If adulto Then
                tit = co(i - col)
            Else
                tit = "T" + Format(i - col + 1, "00")
            End If
            fg(0, i) = tit
            fg.Cols(i).TextAlign = TextAlignEnum.CenterCenter
            fg.Cols(i).TextAlignFixed = TextAlignEnum.CenterCenter
        Next
    End Sub
    Public Sub talla_setea(ByVal ta As C1.Win.C1FlexGrid.C1FlexGrid, ByVal escala As Integer)
        Dim t As String = Format(escala, "00")
        Dim dr As DataRow
        Dim i As Integer
        dr = get_tallas(t)
        For i = 1 To 10
            t = dr(i + 1)
            ta(0, i) = t
            ta(1, i) = 0
            If Trim(t) = "" Then
                ta.Cols(i).AllowEditing = False
            Else
                ta.Cols(i).AllowEditing = True
            End If
        Next
        ta(0, 11) = "TOTALES"
        ta(1, 11) = 0
        ta.SetCellStyle(1, 11, ta.Styles("amarillo"))
        ta.Enabled = False
    End Sub

    Public Sub llena_tipos_stock(ByRef C4 As System.Windows.Forms.ComboBox, ByRef tips As C1.Win.C1FlexGrid.C1FlexGrid)
        Dim tps As New C1.Win.C1FlexGrid.C1FlexGrid()
        Dim dt As New DataTable()
        Dim dr As DataRow
        Dim cnn As New System.Data.SqlClient.SqlConnection()
        Dim strsql As String = "SELECT * FROM TIPO_STOCK ORDER BY TIPO"
        Dim l As Integer
        llena_tablas(dt, strsql, cnn)
        tps.Clear()
        tps.Rows.Count = 0
        tps.Cols.Count = 2
        For Each dr In dt.Rows
            tps.Rows.Count = tps.Rows.Count + 1
            l = tps.Rows.Count - 1
            C4.Items.Add(dr("DESCRIPCION"))
            tps(l, 0) = dr("TIPO")
            tps(l, 1) = dr("DESCRIPCION")
        Next
        tips = tps
    End Sub

    Public Sub llena_salas(ByVal sala As System.Windows.Forms.ComboBox, ByVal cnn As SqlClient.SqlConnection, ByVal todas As String)
        Dim dt As New DataTable()
        Dim dr As DataRow
        sala.Items.Clear()
        If todas = "S" Then
            sala.Items.Add("TODAS")
        End If
        Dim strSQL As String = "SELECT * FROM SALAS"
        llena_tablas(dt, strSQL, cnn)
        For Each dr In dt.Rows
            sala.Items.Add(dr("SALA"))
        Next
        If sala.Items.Count > 0 Then
            sala.SelectedIndex = 0
        End If
    End Sub

    Public Sub flex_a_dt(ByRef fg As C1.Win.C1FlexGrid.C1FlexGrid, ByRef dt As DataTable)
        Dim dr As DataRow
        Dim dc As DataColumn
        Dim i As Integer
        Dim j As Integer
        Dim colnom As String
        dt = New DataTable()
        For j = 1 To fg.Cols.Count - 1
            colnom = fg(0, j)
            dt.Columns.Add(colnom)
            dc = New DataColumn(colnom)
            dc.DataType = fg.Cols(j).GetType      'System.Type.GetType("System.String")
        Next
        For i = 1 To fg.Rows.Count - 1
            dr = dt.NewRow
            For j = 1 To fg.Cols.Count - 1
                dr.Item(j - 1) = fg(i, j)
            Next
            dt.Rows.Add(dr)
        Next
    End Sub
    Public Sub llena_combos_e(ByVal combo As System.Windows.Forms.ComboBox, ByVal e As String, ByVal strsql As String, ByVal campo As String)
        Dim dt As New DataTable()
        Dim dr As DataRow
        Dim cnn As New SqlClient.SqlConnection()
        Dim constr As String = ""
        con_string(e, constr)
        cnn.ConnectionString = constr
        llena_tablas_e(dt, strsql, cnn)
        combo.Items.Clear()
        combo.Text = ""
        Try
            For Each dr In dt.Rows
                combo.Items.Add(dr(campo))
            Next
            If combo.Items.Count > 0 Then
                combo.SelectedIndex = 0
            End If
        Catch
        End Try
    End Sub
    Public Sub llena_combos_d(ByVal combo As System.Windows.Forms.ComboBox, ByVal dt As DataTable, ByVal campo As String)
        Dim dr As DataRow
        combo.Items.Clear()
        combo.Text = ""
        Try
            For Each dr In dt.Rows
                combo.Items.Add(dr(campo))
            Next
            If combo.Items.Count > 0 Then
                combo.SelectedIndex = 0
            End If
        Catch
        End Try
    End Sub
    Public Sub llena_combos(ByRef combo As System.Windows.Forms.ComboBox, ByVal strsql As String, ByVal campo As String)
        Dim dt As New DataTable()
        Dim dr As DataRow
        Dim cnn As New SqlClient.SqlConnection()
        llena_tablas(dt, strsql, cnn)
        combo.Items.Clear()
        Try
            For Each dr In dt.Rows
                combo.Items.Add(dr(campo))
            Next
            combo.SelectedIndex = 0
        Catch
        End Try
    End Sub

    Public Sub llena_tablas(ByRef dt As DataTable, ByVal strSql As String, ByRef cnn As SqlClient.SqlConnection)
        Dim obj As empresas
        Dim da As System.Data.SqlClient.SqlDataAdapter
        cnn = New System.Data.SqlClient.SqlConnection()
        obj = New empresas()
        cnn.ConnectionString = obj.constr
        Dim ds As New DataSet()
        da = New System.Data.SqlClient.SqlDataAdapter(strSql, cnn)
        Try
            da.Fill(ds)
            dt = ds.Tables(0)
        Catch
        End Try
    End Sub
    Public Sub llena_tablas_e(ByRef dt As DataTable, ByVal strSql As String, ByRef cnn As SqlClient.SqlConnection)
        Dim da As System.Data.SqlClient.SqlDataAdapter
        Dim ds As New DataSet()
        da = New System.Data.SqlClient.SqlDataAdapter(strSql, cnn)
        Try
            da.Fill(ds)
            dt = ds.Tables(0)
        Catch
        End Try
    End Sub
    Public Sub llena_tablas_con(ByRef dt As DataTable, ByVal con() As String, ByVal strSql As String)
        Dim cnn As New SqlClient.SqlConnection
        Dim da As System.Data.SqlClient.SqlDataAdapter
        Dim ta As New DataTable
        Dim i As Integer
        Dim Col As New DataColumn
        dt = New DataTable
        For i = 1 To con.Length - 2
            cnn.ConnectionString = con(i)
            Dim ds As New DataSet()
            da = New System.Data.SqlClient.SqlDataAdapter(strSql, cnn)
            Try
                da.Fill(ds)
                ta = New DataTable
                ta = ds.Tables(0)
                Col = New DataColumn
                With Col
                    .ColumnName = "EMPRESA"
                    .DataType = System.Type.GetType("System.String")
                    .DefaultValue = CStr(i)
                End With
                ta.Columns.Add(Col)


                'Col = ta.Columns.Add("EMPRESA", Type.GetType("System.String"))
                'Col.DefaultValue = CStr(i)

                dt.Merge(ta)
            Catch
            End Try
        Next
    End Sub
    Public Sub llena_tablas_con1(ByRef dt As DataTable, ByVal con() As String, ByVal strSql As String)
        Dim cnn As New SqlClient.SqlConnection
        Dim da As System.Data.SqlClient.SqlDataAdapter
        Dim ta As New DataTable
        Dim i As Integer
        Dim Col As New DataColumn
        For i = 1 To con.Length - 2
            cnn.ConnectionString = con(i)
            Dim ds As New DataSet()
            da = New System.Data.SqlClient.SqlDataAdapter(strSql, cnn)
            Try
                da.Fill(ds)
                ta = New DataTable
                ta = ds.Tables(0)
                Col = New DataColumn
                With Col
                    .ColumnName = "EMPRESA"
                    .DataType = System.Type.GetType("System.String")
                    .DefaultValue = CStr(i)
                End With
                ta.Columns.Add(Col)

                dt.Merge(ta)
            Catch
            End Try
        Next
    End Sub

    Public Sub lee_cortes_pad(ByVal strsql As String, ByRef dr As DataRow, ByRef ep As String, ByRef es As String, ByRef rp As String, ByRef rs As String, ByRef ok As Boolean)
        Dim te(40) As String
        Dim cnn As New System.Data.SqlClient.SqlConnection()
        Dim dt As New DataTable()
        Dim i As Integer
        Dim lin As String
        Dim val As Integer
        Dim tas As String = "EP0,EP1,EP2,EP3,EP4,EP5,EP6,EP7,EP8,EP9,ES0,ES1,ES2,ES3,ES4,ES5,ES6,ES7,ES8,ES9,RP0,RP1,RP2,RP3,RP4,RP5,RP6,RP7,RP8,RP9,RS0,RS1,RS2,RS3,RS4,RS5,RS6,RS7,RS8,RS9"
        ok = False
        ep = ""
        es = ""
        rp = ""
        rs = ""
        llena_tablas(dt, strsql, cnn)
        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            te = Split(tas, ","c)
            For i = 0 To 39
                lin = te(i)
                val = dr(lin)
                If i < 10 Then
                    ep = ep + CStr(val) + ","
                ElseIf i < 20 Then
                    es = es + CStr(val) + ","
                ElseIf i < 30 Then
                    rp = rp + CStr(val) + ","
                Else
                    rs = rs + CStr(val) + ","
                End If
            Next
            ep = Mid(ep, 1, Len(ep) - 1)
            es = Mid(es, 1, Len(es) - 1)
            rp = Mid(rp, 1, Len(rp) - 1)
            rs = Mid(rs, 1, Len(rs) - 1)
            ok = True
        End If
    End Sub

    Public Sub codif_tela(ByVal codt As DataTable, ByVal cod As String, ByRef descr As String, ByRef ok As Boolean)
        Dim i As Integer
        Dim p As String
        Dim r As Integer
        Dim d As String
        Dim t(4) As String
        Dim dr As DataRow
        Dim res As DataRow()
        descr = ""
        t = Split("ORDEN,FABRIC,WEIGHT,CONTENID,FIBRA", ","c)
        ok = True
        If cod.Length <> 8 Then
            descr = "MAL CODIGO"
            ok = False
            Exit Sub
        End If
        For i = 1 To 8 Step 2
            r = (i - 1) / 2 + 1
            p = Mid(cod, i, 2)
            res = codt.Select("ORDEN = '" & p & "'")
            Try
                dr = res(0)
                d = Trim(dr(t(r)))
                If d <> Nothing Then
                    descr = descr & d & " "
                Else
                    ok = False
                End If
            Catch
                ok = False
            End Try
        Next
        If Not ok Then
            descr = "MAL CODIGO"
        End If
    End Sub

    Public Sub llena_clientes_cpo(ByRef C4 As System.Windows.Forms.ComboBox)
        Dim cnn As New System.Data.SqlClient.SqlConnection()
        Dim dt As New DataTable()
        Dim td As New DataTable()
        Dim dr As DataRow
        Dim tr As DataRow
        Dim obj As empresas
        Dim strsql As String
        obj = New empresas()
        C4.Items.Clear()
        strsql = "SELECT * FROM USUARIO_CLIENTE WHERE USUARIO = '" & obj.clave & "' ORDER BY CLIENTE"
        llena_tablas(dt, strsql, cnn)
        For Each dr In dt.Rows
            If dr("CLIENTE") = "TODOS" Then
                strsql = "SELECT CLIENTE FROM CLIENTES ORDER BY CLIENTE"
                llena_tablas(td, strsql, cnn)
                For Each tr In td.Rows
                    C4.Items.Add(tr("CLIENTE"))
                Next
                Try
                    C4.SelectedIndex = 0
                Catch
                End Try
                Exit Sub
            End If
            C4.Items.Add(dr("CLIENTE"))
        Next
        Try
            C4.SelectedIndex = 0
        Catch
        End Try
    End Sub

    Public Sub grabar_sql(ByVal strsql As String, ByRef cnn As SqlClient.SqlConnection, ByRef afectados As Integer)
        Try
            afectados = 0
            Dim obj As New empresas()
            cnn.ConnectionString = obj.constr
            cnn.Open()

            Dim graba As New SqlClient.SqlCommand(strsql, cnn)
            afectados = graba.ExecuteNonQuery()
        Catch
        Finally
            cnn.Close()
        End Try
    End Sub

    Public Sub llena_fpos_rec(ByVal cliente As String, ByRef fpo As DataTable, ByRef f_rec As DataTable)
        Dim cnn As New SqlClient.SqlConnection()
        Dim fpol As String
        Dim fpoes As String
        Dim dr As DataRow
        Dim p As Integer
        Dim strSQl As String = "SELECT * FROM FPO WHERE CPO IN (SELECT DISTINCT CPO FROM CPO_D WHERE ESTADO = 'A') AND CLIENTE = '" & cliente & "' AND TIPO = 'TELA' ORDER BY CPO,COLOR,CODIGO"
        llena_tablas(fpo, strSQl, cnn)
        fpol = "("
        For Each dr In fpo.Rows
            fpoes = "'" & dr("FPO") & "'"
            p = fpol.IndexOf(fpoes)
            If p = -1 Then
                fpol = fpol & fpoes & ","
            End If
        Next
        fpol = Mid(fpol, 1, Len(fpol) - 1) + ")"
        fecha_recepcion(f_rec, fpol)
    End Sub

    Public Sub quita_caracter(ByRef texto As String, ByVal caracter As String)
        Dim p As Integer = texto.IndexOf(caracter)
        While p > 0
            texto = Mid(texto, 1, p) + Mid(texto, p + 2)
            p = texto.IndexOf(caracter)
        End While
    End Sub

    Public Sub adios(ByRef nombre As String)
        quita_caracter(nombre, "/")
        quita_caracter(nombre, "\")
    End Sub

    Public Sub graba_t(ByVal cnn As SqlClient.SqlConnection, ByVal strsql As String)
        Dim afectados As Integer
        cnn.Open()
        ' Comienza la transacción
        Dim transaccion As SqlClient.SqlTransaction = cnn.BeginTransaction()

        ' Crea el comando para la transacción
        Dim comando As SqlClient.SqlCommand = cnn.CreateCommand()
        comando.Transaction = transaccion
        Try
            comando.CommandText = strsql
            afectados = comando.ExecuteNonQuery()

            transaccion.Commit()

        Catch e As Exception
            Try
                MsgBox("Inconsistencia en Datos, no se pudo actualizar.", MsgBoxStyle.Critical, "Por favor revise !!!!")
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

    Public Sub fecha_recepcion(ByRef f_rec As DataTable, ByRef fpol As String)
        f_rec = New DataTable()
        Dim strsql As String
        Dim cnn As New SqlClient.SqlConnection
        strsql = "SELECT DISTINCT FECHA, DMOVTO.CPO,DMOVTO.FPO, TIPO, CODIGO,COLOR, LOTE FROM DMOVTO, ROLLOS WHERE DMOVTO.BATCH = ROLLOS.BATCH AND DMOVTO.ROLLO = ROLLOS.ROLLO AND TMOVTO = '1' AND DMOVTO.FPO IN " & fpol
        llena_tablas(f_rec, strsql, cnn)
    End Sub

    'Public Class myPrinter
    '    Friend TextToBePrinted As String
    '    Public Sub prt(ByVal text As String)
    '        TextToBePrinted = text
    '        Dim prn As New Printing.PrintDocument()
    '        Try
    '            prn.PrinterSettings.PrinterName = "PrinterName"
    '            AddHandler prn.PrintPage, AddressOf Me.PrintPageHandler
    '            prn.Print()
    '            RemoveHandler prn.PrintPage, AddressOf Me.PrintPageHandler
    '        Catch
    '        End Try
    '    End Sub
    '    Private Sub PrintPageHandler(ByVal sender As Object, _
    '       ByVal args As Printing.PrintPageEventArgs)
    '        Dim myFont As New Font("Microsoft San Serif", 10)
    '        args.Graphics.DrawString(TextToBePrinted, _
    '           New Font(myFont, FontStyle.Regular), _
    '           Brushes.Black, 50, 50)
    '    End Sub
    'End Class

    Public Sub AutoCompletar(ByRef cb As ComboBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim strFindStr As String
        Dim i As Integer = cb.SelectedIndex
        If e.KeyChar = Chr(8) Then  'Backspace
            If cb.SelectionStart <= 1 Then
                cb.SelectionStart = 1
                ' Exit Sub
            End If
            If cb.SelectionLength = 0 Then
                strFindStr = cb.Text.Substring(0, cb.Text.Length - 1)
            Else
                strFindStr = cb.Text.Substring(0, cb.SelectionStart - 1)
            End If
        Else
            If cb.SelectionLength = 0 Then
                strFindStr = cb.Text & e.KeyChar
            Else
                strFindStr = cb.Text.Substring(0, cb.SelectionStart) & e.KeyChar
            End If
        End If

        Dim intIdx As Integer = -1

        ' Busca el string en el combobox 
        intIdx = cb.FindString(strFindStr)

        If intIdx <> -1 Then ' String encontrado
            cb.SelectedText = ""
            cb.SelectedIndex = intIdx
            cb.SelectionStart = strFindStr.Length
            cb.SelectionLength = cb.Text.Length
            e.Handled = True
        Else
            e.Handled = True
        End If
    End Sub


    Public Sub llena_clientes_usuario(ByRef c4 As ComboBox, ByVal usuario As String, ByVal tipo As String)
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim todos As Boolean = False
        Dim strsql As String
        Dim clientes As String = "("
        revisa_acceso_usuario(usuario, dt)
        For Each dr In dt.Rows
            If dr(tipo) = "S" Then
                If dr("CLIENTE") = "TODOS" Then
                    todos = True
                Else
                    clientes = clientes & "'" & dr("CLIENTE") & "',"
                End If
            End If
        Next
        If todos Then
            strsql = "SELECT * FROM CLIENTES ORDER BY CLIENTE"
        Else
            If clientes.Length > 0 Then
                clientes = Mid(clientes, 1, clientes.Length - 1) + ")"
            End If
            strsql = "SELECT * FROM CLIENTES WHERE CLIENTE IN " & clientes & " ORDER BY CLIENTE"
        End If
        llena_combos(c4, strsql, "CLIENTE")
    End Sub
    Public Sub revisa_acceso_usuario(ByRef usuario As String, ByRef dt As DataTable)
        Dim cnn As New SqlClient.SqlConnection
        llena_tablas(dt, "SELECT * FROM USUARIO_CLIENTE WHERE USUARIO = '" & usuario & "'", cnn)
    End Sub
    Public Sub envia_corrreo_bom(ByVal cliente As String, ByVal asunto As String, ByVal mensaje As String, ByVal path As String)

        Dim dt As New DataTable
        Dim dr As DataRow
        Dim cnn As New SqlClient.SqlConnection
        Dim strsql As String = "SELECT * FROM USUARIO_CLIENTE WHERE SEGUIMIENTO_BOM = 'S' AND (CLIENTE = '" & cliente & "' OR CLIENTE = 'TODOS') ORDER BY USUARIO"
        llena_tablas(dt, strsql, cnn)

        Try
            Dim attachFile As New Attachment(path)
            Dim SmtpServer As New System.Net.Mail.SmtpClient
            Dim mail As New System.Net.Mail.MailMessage
            Dim correo As String = ""
            SmtpServer.Credentials = New Net.NetworkCredential("jcperez@pcs.com.gt", "Cnmrs98s")
            SmtpServer.Port = 25 ' == puerto smtp 587
            SmtpServer.Host = "pop.emailsrvr.com"
            'mtpServer.EnableSsl = True
            mail = New MailMessage()
            mail.From = New MailAddress("jcperez@pcs.com.gt")
            mail.Attachments.Add(attachFile)

            For Each dr In dt.Rows
                Try
                    correo = dr("CORREO")
                    mail.To.Add(correo)
                Catch ex As Exception
                    MsgBox("Error al enviar a la dir'eccion " & correo, MsgBoxStyle.Critical, "Error en el envío del correo.")
                End Try
            Next
            mail.Subject = asunto
            mail.Body = mensaje  ' === "Cuerpo del Mensaje"
            SmtpServer.Send(mail)
        Catch ex As Exception
            MsgBox("No puede enviar el correo de Autorización", MsgBoxStyle.Critical, "Error en el envío del correo.")
        End Try
    End Sub
    Public Sub envia_corrreo_split(ByVal asunto As String, ByVal mensaje As String, ByRef ok As Boolean)
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim cnn As New SqlClient.SqlConnection
        Dim strsql As String = "SELECT * FROM CORTE_SPLIT_C ORDER BY CORREO"
        llena_tablas(dt, strsql, cnn)
        ok = False
        Try
            Dim SmtpServer As New System.Net.Mail.SmtpClient
            Dim mail As New System.Net.Mail.MailMessage
            Dim correo As String = ""
            SmtpServer.Credentials = New Net.NetworkCredential("ricoh@pcs.com.gt", "Djk2e39df")
            SmtpServer.Port = 25 ' == puerto smtp 587
            SmtpServer.Host = "pop.emailsrvr.com"
            'mtpServer.EnableSsl = True
            mail = New MailMessage()
            mail.From = New MailAddress("ricoh@pcs.com.gt", "Sistemas Integrados Server")
            'mail.Attachments.Add(attachFile)

            For Each dr In dt.Rows
                Try
                    correo = dr("CORREO")
                    mail.To.Add(correo)
                Catch ex As Exception
                    MsgBox("Error al enviar a la direccion " & correo, MsgBoxStyle.Critical, "Error en el envío del correo.")
                End Try
            Next
            mail.Subject = asunto
            mail.Body = mensaje  ' === "Cuerpo del Mensaje"
            SmtpServer.Send(mail)
            ok = True
        Catch ex As Exception
            MsgBox("No puede enviar el correo de Autorización", MsgBoxStyle.Critical, "Error en el envío del correo.")
        End Try
    End Sub
    Public Sub envia_corrreo_estilo(ByVal asunto As String, ByVal mensaje As String, ByRef ok As Boolean)
        Dim dt As New DataTable
        Dim cnn As New SqlClient.SqlConnection
        Dim strsql As String = "SELECT * FROM CORTE_SPLIT_C ORDER BY CORREO"
        llena_tablas(dt, strsql, cnn)
        ok = False
        Try
            Dim SmtpServer As New System.Net.Mail.SmtpClient
            Dim mail As New System.Net.Mail.MailMessage
            Dim correo As String = ""
            SmtpServer.Credentials = New Net.NetworkCredential("ricoh@pcs.com.gt", "Djk2e39df")
            SmtpServer.Port = 25 ' == puerto smtp 587
            SmtpServer.Host = "pop.emailsrvr.com"
            'mtpServer.EnableSsl = True
            mail = New MailMessage()
            mail.From = New MailAddress("ricoh@pcs.com.gt", "Sistemas Integrados Server")
            'mail.Attachments.Add(attachFile)
            mail.To.Add("amata@pcs.com.gt")
            mail.To.Add("rbarillas@pcs.com.gt")
            mail.Subject = asunto
            mail.Body = mensaje  ' === "Cuerpo del Mensaje"
            SmtpServer.Send(mail)
            ok = True
        Catch ex As Exception
            MsgBox("No puede enviar el correo de Autorización", MsgBoxStyle.Critical, "Error en el envío del correo.")
        End Try
    End Sub

    Public Sub busca_descripciones(ByVal cu As DataTable, ByVal codigo As String, ByRef descripcion As String, ByRef ok As Boolean)
        Dim dw As DataRow()
        Dim dr As DataRow
        descripcion = ""
        ok = False
        Try
            dw = cu.Select("CODIGO = '" & codigo & "'")
            If dw.Length > 0 Then
                dr = dw(0)
                descripcion = dr("DESCRIPCION")
                ok = True
            End If
        Catch
        End Try
    End Sub
    Public Sub busca_descripciones1(ByVal cu As DataTable, ByVal llave As String, ByVal campo As String, ByVal codigo As String, ByRef descripcion As String)
        Dim dw As DataRow()
        Dim dr As DataRow
        Try
            dw = cu.Select(llave & " = '" & codigo & "'")
            If dw.Length > 0 Then
                dr = dw(0)
                descripcion = dr(campo)
            End If
        Catch
            descripcion = ""
        End Try
    End Sub

    Public Sub llena_tipos_Telas(ByRef cu As DataTable, ByRef codigo As ComboBox)
        Dim dr As DataRow
        codigo.Items.Clear()
        For Each dr In cu.Rows
            codigo.Items.Add(dr("CODIGO"))
        Next
        If cu.Rows.Count > 0 Then
            codigo.SelectedIndex = 0
        End If
    End Sub

    Public Sub conexiones(ByRef con() As String)
        Dim cnstr As New C1cajasLib_SIF.util
        Dim i As Integer
        ReDim con(2)
        For i = 0 To 2
            con(i) = cnstr.con_string(i)
        Next
    End Sub


    Public Sub calcula_fecha_corte(ByVal ec As DataTable, ByVal estilo As String, ByVal fcostura As Date, ByRef fcorte As Date)
        Dim dd As DataRow()
        Dim dr As DataRow
        Dim d As Integer
        Dim dias As Integer = 15
        dd = ec.Select("ESTILO = '" & estilo & "'")
        If dd.Length > 0 Then
            dr = dd(0)
            dias = dr("DIAS_CORTE")
        End If
        Try
            fcorte = fcostura.AddDays(-dias)
            d = fcorte.DayOfWeek
            If d = 0 Then
                fcorte = fcorte.AddDays(-2)
            ElseIf d = 6 Then
                fcorte = fcorte.AddDays(-2)
            End If
        Catch
            fcostura = Today
            fcorte = fcostura.AddDays(-dias)
        End Try

    End Sub


    Public Sub determina_columna_tela(ByVal fei As Date, ByVal fecha As Date, ByVal coli As Integer, ByVal colmax As Integer, ByRef col As Integer)
        Dim d As Integer
        Dim res As Decimal
        d = DateDiff(DateInterval.Day, fei, fecha)
        res = (d / 7)
        If d > 0 Then
            d = Fix(d / 7)
            res = res - d
            If res > 0 Then
                d = d + 1
            End If
        Else
            d = 0
        End If
        col = coli + d
        If col < coli Then
            col = coli
        End If
        If col > colmax Then
            col = colmax
        End If
    End Sub


    Public Sub ajusta_fecha_produccion(ByRef va As DataTable, ByVal sec As String, ByVal fechai As Date, ByVal dias As Decimal, ByRef fechaf As Date)
        Dim d As Integer = 0
        Dim dd As DataRow()
        Dim fecha As New Date

        Dim h As String = "yyyy-MM-dd"
        Dim fei As Date
        Dim fef As Date
        Dim ok As Boolean
        Dim sabdom As String = "06"
        Dim diasem As Integer

        fechaf = fechai.AddDays(dias)
        fei = Format(fechai, h)
        fef = Format(fechaf, h)
        ' ======================== VERIFICA ASUETOS-VACACIONES ==================
        ok = True
        'Try
        Do While ok
            dd = va.Select("TIPO = 'D' AND FECHA = '" & Format(fechaf, "yyyy-MM-dd") & "' AND SECCION = '" & sec & "'")
            If dd.Length > 0 Then
                fechaf = fechaf.AddDays(1)
            Else
                ok = False
                ' ====================== VERIFICA FIN DE SEMANA =========================
                diasem = fechaf.DayOfWeek
                If sabdom.IndexOf(diasem) > -1 Then
                    dd = va.Select("TIPO = 'F' AND SECCION = '" & sec & "' AND FECHA = '" & Format(fechaf, "yyyy-MM-dd") & "'")
                    If dd.Length = 0 Then
                        fechaf = fechaf.AddDays(1)
                        ok = True
                    End If
                End If
            End If
        Loop
        'Catch
        'End Try

    End Sub
    Public Sub crea_sub_inventario(ByRef ruta As String)
        Dim obj As New empresas
        ruta = "c:\telas\inventarios"
        Try
            If Not Directory.Exists(ruta) Then
                Directory.CreateDirectory(ruta)
            End If
        Catch ex As Exception
        End Try
        ruta = ruta + "\" + obj.nombre
        Try
            If Not Directory.Exists(ruta) Then
                Directory.CreateDirectory(ruta)
            End If
        Catch ex As Exception
        End Try
    End Sub
    Public Function semana_fecha(ByVal ano As Integer, ByRef semana As Integer, ByRef fecha As Date)
        Dim inicio As Date = CDate(Format(ano, "0000") + "-01-01")
        Dim dia As Integer = inicio.DayOfWeek
        If semana = 0 Then
            semana = 99
            fecha = CDate("9999-12-31")
        Else
            dia = 7 - dia
            inicio = inicio.AddDays(dia)
            dia = (semana - 1) * 7
            fecha = inicio.AddDays(dia)
        End If
        Return fecha
    End Function

    Public Sub llena_combos_d(ByVal combo As System.Windows.Forms.ComboBox, ByVal strsql As String, ByRef dt As DataTable, ByVal dm As String, ByVal vm As String)
        'Dim dr As DataRow
        Dim cnn As New SqlClient.SqlConnection()
        llena_tablas(dt, strsql, cnn)
        Try
            combo.DataSource = dt
            combo.DisplayMember = dm
            combo.ValueMember = vm
        Catch ex As Exception
        End Try
        If combo.Items.Count > 0 Then
            combo.SelectedIndex = 0
        End If
    End Sub

    Public Function d_estado(ByVal dt As DataTable, ByVal des As String) As String
        Dim dd As DataRow()
        Dim dr As DataRow
        Dim res As String
        dd = dt.Select("ESTADO_TELA = '" & des & "'")
        If dd.Length > 0 Then
            dr = dd(0)
            res = dr("CODIGO")
        Else
            res = ""
        End If
        Return res
    End Function

    Public Sub busca_registro(ByRef dt As DataTable, ByVal campob As String, ByVal busca As String, ByVal campo As String, ByRef resultado As String)
        Dim dd As DataRow()
        Dim dr As DataRow
        resultado = ""
        dd = dt.Select(campob & " = '" & busca & "'")
        If dd.Length > 0 Then
            dr = dd(0)
            resultado = dr(campo)
        End If
    End Sub

    Public Function embelishment(ByVal dd As DataRow()) As String
        Dim dr As DataRow
        Dim res As String = "N/A"

        If dd.Length > 0 Then
            dr = dd(0)
            If dr("O2") Then
                res = "BORDADO"
            Else
                If dr("O3") Then
                    res = "SERIGRAFIA"
                End If
            End If
            If dr("O2") And dr("O3") Then
                res = "BORDADO/SERIGRA"
            End If
        End If
        Return res
    End Function

    Public Function fpo_ofecha(ByVal cliente) As Date
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim fecha As Date = Nothing
        Dim cnna As New SqlClient.SqlConnection
        Dim strsql As String = "SELECT TOP 1 FECHA FROM FPO_OFECHA ORDER BY FECHA DESC"
        llena_tablas(dt, strsql, cnna)
        For Each dr In dt.Rows
            fecha = dr("fecha")
        Next
        Return fecha
    End Function

    Public Function descarga_lector(ByVal c As Integer) As DataTable
        Dim cnn As New SqlClient.SqlConnection
        Dim dt As New DataTable
        Dim tipo(3) As String
        tipo(0) = "CODIGO"
        tipo(1) = "ROLLOS_UBICA"
        tipo(2) = "DESP_T2"
        cnn.ConnectionString = lector_dc()
        llena_tablas_e(dt, "SELECT * FROM " & tipo(c), cnn)
        Return dt
    End Function

    Public Function imprime_etiquetas_tela(ByRef empresa As String, ByVal batch As String, ByVal rollo As String, ByVal knit As String, ByVal colort As String, ByVal barra As String, ByVal yardas As Decimal, ByVal libras As Decimal) As Boolean
        Dim ok As Boolean = True
        Dim ipAddress As String = "192.9.200.28"
        Dim port As Integer = 9100
        Dim logo As String = ""
        Dim ZPLString As String
        Dim ti As String = ""
        Dim enca As String
        Dim cliente As String = ""
        Dim h As String = "##0.00"
        Dim ya As String = Format(yardas, h)
        Dim li As String = Format(libras, h)

        ya = Space(6 - Len(ya)) + ya
        li = Space(6 - Len(li)) + li
        '       Try
        Dim client As New System.Net.Sockets.TcpClient()
        client.Connect(ipAddress, port)

        ' Write ZPL String to connection
        Dim writer As New System.IO.StreamWriter(client.GetStream())

        If empresa = "1" Then
            logo = "^FO10,10^GFA,2205,2205,15,,:::O03FE,O07FF8,N01IFC,N03IFE,N07JF,N07JF8,:N0KFC,::::::N0KF8,N07JF8,N07JF,N03IFE,N01IFC,O0IF8,O03FE,,:::M01WFE,:::::::::::::::::::::U03KFC,U01KFC,U03KFC,U03KFC10FE,U03KFC38FE,:U03KFC30C6,M01KFE03KFC38C6,M01KFE03KFC3806,M01KFE03KFC380E,M01KFE03KFC1C1E,M01KFE03KFC1FFC,M01KFE03KFC0FFC,M01KFE03KFC07F,M01KFE03KFC,::M01KFE03KFC3FFE,M01KF800KFC3FFE,M01JFCI01JFC183E,M01JFK0JFC007C,M01IFEK03IFC01F,M01IFCK01IFC07C,M01IF8004I0IFC1F,M01IF003F8007FFC3E,M01FFE00FFE003FFC3FFE,M01FFC00F1E003FFC3FFE,M01FFC01E0E001FFC1FFE,M01FF801E0F001FFC,M01FF801E0FI0FFC,M01FF801E1EI0FFC3FFE,M01FF001F1EI0FFC3FFE,M01FFI0F3CI07FC3FFE,M01FFI0FF8I07FC,M01FFI07E03807FC,M01FF001FE07807FC01C,M01FF003FF07807FC07F,M01FF007CF87807FC0FFC,M01FF00787CF007FC1FFC,M01FF00F87CF007FC1C1E,M01FF00F03FE007FC380E,M01FF00F81FE00FFC3806,M01FF80F80FC00FFC3006,M01FF80FC0FC00FFC3006,M01FF807F3FE01FFC3FFE,M01FFC03JF01FFC3FFE,M01FFE01FFCF83FFC3FFE,M01FFE007E0787FFC,M01IFM07FFC,M01IF8L0IFC,M01IFCK01IFCI0E,M01JFK07IFC007E,M01JF8J0JFC03FE,M01KFI07JFC1FF,M01KFC01KFC3F3,M03KFC01KFC383,M03KFC03KFC3F3,M03KFC03KFC1FF8,M03KFC03KFC03FE,M07KFC03KFC007E,M0LFC03KFCI0E,M0LFC03KFC,L03LFC03KFC,L0MF803KFC,K07MF803KFC0E0E,I0PF803KFC1F3E,I0PF803KFC1FFE,I0PF003KFC39F8,I0PF003KFC30E,I0PF003KFC30C,I0OFE003KFC38C,I0OFE003KFC3FFE,I0OFC003KFC3FFE,I0OF8003KFC1FFE,I0OF8003KFC,I0OFI03KFC,I0OFI03KFC38,I0NFEI03KFC38,I0NFCI03KFC38,I0NF8I03KFC38,I0NFJ03KFC3FFE,I0MFEJ03KFC3FFE,I0MF8J03KFC3A6E,I0MFK03KFC38,I0LF8K03KFC38,I0KFEL03KFC38,I0KFM03KFC,I0IFEN03KFC,I0FQ01KF8,,::::::^FS"
        ElseIf empresa = "3" Then
            logo = "^FO15,45^GFA,1480,1480,20,,:::Q0F8,P0IFC,O07IFC,I07FE001FFD2,001JF03FF8,007MFC,I093LF8,K03LF,K01LFC,K0JFE67E,J03KF003,I01LFC,I03JF7FE,I07JF3FF,001KF9FF8,003F4IF9DFC3FC,007E1IF8E7E3FF8003,00FE3FBF8E0F01FF0IF,00F23FBF070F003KF,01C07F3F030383JF8,01007C3F03838KF,0100FC3F03819F3IF8,I01E03F01CI03JF,I01F03E01EI07F7FFC,I01C03F00E001FEIFE,I03C03E00F001ECFFDF,I03C03E007003CCFFCF8,I03801D00700718FBE5C,I03001C00380E1879E0C,I06001E00380C30F8F04,I06001F001C0830F87,I04001E001C00607838,L01E001E00E07838,L01CI0E00E07838,M0EI0E01C0380C,:M0AI0F01807004,M02I0703803,Q0783003,Q0787003,Q0787002,Q078E,:::Q079C,:Q07FCJ01FF8L03FC,Q07FCJ07FF8L0FFC,Q07F8J07FFL01FF8,Q07F8J0F3EL01FE,Q07F8K0380CK03E00C,Q07F8K0700E006003C03E706,Q07F8K0E03E00E00780FE79E,Q07F8J03C07C71C70F81CE3FC,Q07F8J07807CF3CF1F03DC3F8,Q0FF8J0F00F1E7FE1F03983F,Q0FF8I03E00F7E7FE3E0F383E,Q0FF8I07C19EFCFFE3C0F603E,Q0FF8I0FCF1FFCFBC7C1F40FE,Q0FF8001FFE3F79F3CF81F73FF,O018FF8803FFE3E79E3CF01FE7FF,P09FF8801FF07879838E01F871F,P09FF88,O02BFFC,O01JFC,O03JFE,O0LF8,,:^FS"
        End If

        enca = "^XA" + logo
        ZPLString = enca
        ti = "^FO290,60^A0,50,50^FDBATCH:^FS"
        ti = ti + "^FO460,60^A0,50,50^FD" + batch + "^FS"
        ti = ti + "^FO290,120^A0,50,50^FDROLLO:^FS"
        ti = ti + "^FO460,120^A0,50,50^FD" + rollo + "^FS"
        ti = ti + "^FO30,190^A0,45,45^FDKNIT:^FS"
        ti = ti + "^FO30,230^A0,35,35^FD" + knit + "^FS"
        ti = ti + "^FO30,290^A0,45,45^FDCOLOR:^FS"
        ti = ti + "^FO30,330^A0,35,35^FD" + colort + "^FS"
        ti = ti + "^FO640,290^AD,30,14^FDYds.:^FS"
        ti = ti + "^FO730,290^AD,30,14^FD" & ya & "^FS"
        ti = ti + "^FO640,330^AD,30,14^FDLbs.:^FS"
        ti = ti + "^FO730,330^AD,30,14^FD" & li & "^FS"

        ZPLString = ZPLString + ti + "^BY5,2,120" + "^FO75,410^BC^FD" + barra & "^FS" + "^XZ"
        writer.Write(ZPLString)
        writer.Flush()

        writer.Close()
        client.Close()
        ' Catch
        '    ok = False
        ' End Try
        Return ok
    End Function

    Public Function get_seccion(ByVal corte As String) As String
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim seccion As String = ""
        Dim cnn As New SqlClient.SqlConnection
        Dim strsql As String = "SELECT SECCION FROM CORTES WHERE CORTE =  '" & corte & "'"
        llena_tablas(dt, strsql, cnn)
        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            seccion = dr("SECCION")
        End If
        Return seccion
    End Function


    Public Function DefaultPrinterName() As String
        Dim oPS As New System.Drawing.Printing.PrinterSettings

        Try
            MessageBox.Show(oPS.PrinterName)
            DefaultPrinterName = oPS.PrinterName
        Catch ex As System.Exception
            DefaultPrinterName = ""
        Finally
            oPS = Nothing
        End Try
    End Function


    Public Function revisa_existencias(ByVal caja As String, ByVal tipo As String, ByVal talla As String) As Integer
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim cnn As New SqlClient.SqlConnection
        Dim unidades As Integer = 0
        Dim strsql As String = "SELECT * FROM CAJAS01 WHERE CAJA = '" & caja & "' AND TIPO = '" & tipo & "' AND TALLA = '" & talla & "'"
        llena_tablas(dt, strsql, cnn)
        '''Selecciona el escalar de cuantas unidades existentes hay
        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            unidades = dr("UNIDADES")
        End If
        Return unidades
    End Function

    Public Function revisa_datos(ByVal fg As C1.Win.C1FlexGrid.C1FlexGrid, ByVal c As Integer) As Boolean
        Dim ok As Boolean = True
        Dim unidad As Integer
        Dim i As Integer
        For i = 1 To fg.Rows.Count - 1
            unidad = revisa_existencias(fg(i, 1), fg(i, 2), fg(i, 3))
            If unidad <> fg(i, c) Then
                MsgBox("Las unidades de la caja " & fg(i, 1) & "Han cambiado" + vbLf + "No se puede efectuar esta operacion.", MsgBoxStyle.Critical, "Trate de nuevo.")
                ok = False
            End If
        Next
        Return ok
    End Function


    Public Function GetIpV4() As String
        Dim myHost As String = Dns.GetHostName
        Dim ipEntry As IPHostEntry = Dns.GetHostEntry(myHost)
        Dim ip As String = ""

        For Each tmpIpAddress As IPAddress In ipEntry.AddressList
            If tmpIpAddress.AddressFamily = Sockets.AddressFamily.InterNetwork Then
                Dim ipAddress As String = tmpIpAddress.ToString
                ip = ipAddress
                Exit For
            End If
        Next

        If ip = "" Then
            Throw New Exception("No 10. IP found!")
        End If

        Return ip
    End Function
End Module
