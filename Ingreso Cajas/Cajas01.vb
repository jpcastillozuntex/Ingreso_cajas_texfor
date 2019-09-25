Imports C1.Win.C1FlexGrid

Public Class Cajas01
    Inherits System.Windows.Forms.Form
    Dim cnn As New SqlClient.SqlConnection()
    Dim dt As New DataTable()
    Friend WithEvents graba As System.Windows.Forms.Button
    Friend WithEvents fg As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents CORTE As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents S1 As System.Windows.Forms.Button
    Dim dr As DataRow
    Dim es As New DataTable
    Dim e(10) As String
    Dim t(10) As Integer
    Dim p(10)
    Dim escala As String = ""
    Dim inicial As String = ""
    Dim final As String = ""
    Dim seccion As String = ""
    Dim obj As New empresas
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Cajas01))
        Me.CORTE = New System.Windows.Forms.ComboBox()
        Me.graba = New System.Windows.Forms.Button()
        Me.fg = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.S1 = New System.Windows.Forms.Button()
        CType(Me.fg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CORTE
        '
        Me.CORTE.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.CORTE.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CORTE.Items.AddRange(New Object() {"JT", "TRECENTO", "ZUNTEX"})
        Me.CORTE.Location = New System.Drawing.Point(105, 23)
        Me.CORTE.Name = "CORTE"
        Me.CORTE.Size = New System.Drawing.Size(216, 33)
        Me.CORTE.TabIndex = 93
        '
        'graba
        '
        Me.graba.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.graba.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.graba.ForeColor = System.Drawing.Color.Black
        Me.graba.Image = CType(resources.GetObject("graba.Image"), System.Drawing.Image)
        Me.graba.Location = New System.Drawing.Point(1088, 4)
        Me.graba.Name = "graba"
        Me.graba.Size = New System.Drawing.Size(68, 51)
        Me.graba.TabIndex = 91
        Me.graba.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.graba.UseVisualStyleBackColor = False
        '
        'fg
        '
        Me.fg.AllowFiltering = True
        Me.fg.ColumnInfo = resources.GetString("fg.ColumnInfo")
        Me.fg.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.fg.Location = New System.Drawing.Point(13, 84)
        Me.fg.Name = "fg"
        Me.fg.Rows.DefaultSize = 21
        Me.fg.Size = New System.Drawing.Size(1333, 782)
        Me.fg.StyleInfo = resources.GetString("fg.StyleInfo")
        Me.fg.TabIndex = 92
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(23, 20)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(119, 41)
        Me.Label3.TabIndex = 94
        Me.Label3.Text = "Corte:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'S1
        '
        Me.S1.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.S1.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.S1.ForeColor = System.Drawing.Color.Black
        Me.S1.Image = CType(resources.GetObject("S1.Image"), System.Drawing.Image)
        Me.S1.Location = New System.Drawing.Point(982, 2)
        Me.S1.Name = "S1"
        Me.S1.Size = New System.Drawing.Size(68, 51)
        Me.S1.TabIndex = 95
        Me.S1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.S1.UseVisualStyleBackColor = False
        '
        'Cajas01
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 19)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1190, 696)
        Me.Controls.Add(Me.S1)
        Me.Controls.Add(Me.CORTE)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.fg)
        Me.Controls.Add(Me.graba)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Red
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Cajas01"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Impresión de Cajas por corte"
        CType(Me.fg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub Cortes_status(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sec As String = obj.seccion
        llena_tablas(es, "SELECT * FROM E_TALLAS", cnn)
        llena_combos(CORTE, "SELECT CORTE FROM CORTES WHERE CORTE NOT IN (SELECT DISTINCT CORTE FROM PROD_DIARIA) AND CORTE NOT IN (SELECT DISTINCT CORTE FROM CAJAS01) AND SECCION = '" & sec & "' AND TOTAL > 0 AND EXPORTADO <> 'S' ORDER BY CORTE", "CORTE")
        Try
            CORTE.SelectedIndex = 0
        Catch
        End Try
        setea_fg()
    End Sub

    Private Sub setea_fg()
        fg.Rows.Count = 1
        fg.Rows(0).Height = 30
    End Sub
    Private Sub llena_fg(ByRef dr As DataRow)
        Dim l As Integer = 0
        Dim i As Integer
        Dim r As Integer
        Dim f As Integer
        Dim fil As Integer
        Dim ok As Boolean
        Dim strsql As String = "SELECT CORTE,CORTES.CPO,CORTES.ESTILO,CORTES.COLOR,CLIENTE,CORTES.XS,CORTES.S,CORTES.M,CORTES.L,CORTES.XL,CORTES.XL2,CORTES.XL3,CORTES.XL4,CORTES.XL5,CORTES.XL6,CORTES.TOTAL,ESCALA,CORTES.* FROM CORTES LEFT JOIN CPO_D ON CORTES.CPO = CPO_D.CPO AND CORTES.ESTILO = CPO_D.ESTILO AND CORTES.COLOR = CPO_D.COLOR WHERE CORTE = '" & CORTE.Text & "'"
        llena_tablas(dt, strsql, cnn)
        If dt.Rows.Count = 0 Then
            Exit Sub
        End If
        dr = dt.Rows(0)
        llena_tallas_escala(dr)
        llena_empaque(dr("ESTILO"), ok)
        If ok Then

            For l = 0 To 9
                If t(l) > 0 Then
                    i = Fix(t(l) / p(l))

                    r = t(l) - (i * p(l))
                    f = i
                    If r > 0 Then
                        f = f + 1
                    End If
                    fil = fg.Rows.Count
                    fg.Rows.Count = fil + f
                    For f = 1 To i
                        fg(fil, 1) = fil
                        fg(fil, 2) = dr("CPO")
                        fg(fil, 3) = dr("ESTILO")
                        fg(fil, 4) = dr("COLOR")
                        fg(fil, 5) = dr("CLIENTE")
                        fg(fil, 6) = e(l)
                        fg(fil, 7) = p(l)
                        fg(fil, 8) = l
                        fil = fil + 1
                    Next
                    If r > 0 Then
                        fg(fil, 1) = fil
                        fg(fil, 2) = dr("CPO")
                        fg(fil, 3) = dr("ESTILO")
                        fg(fil, 4) = dr("COLOR")
                        fg(fil, 5) = dr("CLIENTE")
                        fg(fil, 6) = e(l)
                        fg(fil, 7) = r
                        fg(fil, 8) = l
                    End If
                End If
            Next
        Else
            Close()
        End If

    End Sub
    Private Sub llena_tallas_escala(ByVal dr As DataRow)
        ReDim e(10)
        ReDim t(10)
        escala = dr("ESCALA")
        Dim dd As DataRow()
        Dim i As Integer
        For i = 0 To 10
            t(i) = dr(i + 5)
        Next
        dd = es.Select("ESCALA = '" & dr("ESCALA") & "'")
        If dd.Length > 0 Then
            dr = dd(0)
            For i = 0 To 9
                e(i) = dr(i + 2)
            Next
        End If
    End Sub
    Private Sub llena_empaque(ByRef estilo As String, ByRef ok As Boolean)
        Dim dt As New DataTable
        Dim dr As DataRow
        ReDim p(10)
        Dim strsql As String = "SELECT * FROM ESTILO_EMP WHERE ESTILO = '" & estilo & "'"
        Dim i As Integer
        ok = False
        llena_tablas(dt, strsql, cnn)
        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            For i = 0 To 9
                p(i) = dr(i + 1)
            Next
            ok = True
        Else
            MsgBox("Aún no han ingresado el número de unidades por Caja.", MsgBoxStyle.Critical, "Por favor ingresar Empaque por Estilo!!!")
        End If
    End Sub
    Private Sub chequea_upc(ByVal dr As DataRow, ByRef ok As Boolean)
        Dim da As New DataTable
        ok = False
        llena_tablas(dA, "SELECT * FROM UPC WHERE CLIENTE = '" & dr("CLIENTE") & "' AND ESTILO = '" & dr("ESTILO") & "' AND COLOR = '" & dr("COLOR") & "'", cnn)
        If dA.Rows.Count > 0 Then
            MsgBox("Este corte tiene UPC.", MsgBoxStyle.Critical, "No se puede imprimir las Etiquetas.")
        Else
            ok = True
        End If
    End Sub
    Private Sub graba_datos()
        Dim strsql As String
        Dim afectados As Integer
        Dim corre As Integer
        Dim pre As String = ""
        Dim j As Object
        Dim obj As New empresas
        Dim h As String = "0000000"
        If obj.numero = "1" Then
            pre = "JT"
        ElseIf obj.numero = "3" Then
            pre = "ZU"
        End If

        cnn.Open()
        ' Comienza la transacción
        Dim transaccion As SqlClient.SqlTransaction = cnn.BeginTransaction()

        ' Crea el comando para la transacción
        Dim comando As SqlClient.SqlCommand = cnn.CreateCommand()
        comando.Transaction = transaccion

        Try
            strsql = "SELECT CORRELATIVO FROM CAJAS02"
            comando.CommandText = strsql
            j = comando.ExecuteScalar()
            corre = j.ToString
            inicial = pre + Format(corre + 1, h)

            For i = 1 To fg.Rows.Count - 1
                strsql = "INSERT INTO CAJAS01 (CAJA,CORTE,TALLA,TIPO,UNIDADES,CLIENTE,UBICACION,FECHA,ESTADO,ESCALA,ORDEN,IMPRESO,TIPO_SEG,SECCION) VALUES ('" & _
                                               pre + Format(corre + i, h) & "','" & CORTE.Text & "','" & _
                                               fg(i, 6) & "','P','" & _
                                               fg(i, 7) & "','" & fg(i, 5) & "','00',GETDATE(),'A','" & _
                                               escala & "','" & fg(i, 8) & "','" & obj.usuario & "','0','" & seccion & "')"
                comando.CommandText = strsql
                afectados = comando.ExecuteNonQuery()

                strsql = "UPDATE CAJAS02 SET CORRELATIVO = CORRELATIVO + 1"
                comando.CommandText = strsql
                afectados = comando.ExecuteNonQuery()

                final = pre + Format(corre + i, h)

            Next

            transaccion.Commit()
            MsgBox("Grabacion Exitosa", MsgBoxStyle.Exclamation, "Grabaciones")
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

    Private Sub graba_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles graba.Click
        Dim ok As Boolean
        Dim seguro As MsgBoxResult
        Dim pr As New C1cajas.prt
        seguro = MsgBox("Seguro de Grabar las Cajas?  ", MsgBoxStyle.YesNo, "Actualizacion de Datos !!!")
        If seguro = MsgBoxResult.Yes Then
            graba_datos()
            MsgBox(pr.imprime_cajas_s(inicial, final, obj.seccion, obj.numero, obj.constr))
            Close()
        End If
    End Sub
    Private Sub empresa_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles CORTE.KeyPress
        AutoCompletar(CORTE, e)
    End Sub
    Private Sub S1_Click(sender As System.Object, e As System.EventArgs) Handles S1.Click
        Dim ok As Boolean
        Dim dr As DataRow
        dr = Nothing
        llena_fg(dr)
        chequea_upc(dr, ok)
        If ok Then
            seccion = get_seccion(CORTE.Text)
            S1.Visible = False
        Else
            setea_fg()
            CORTE.Focus()
        End If

    End Sub
End Class

