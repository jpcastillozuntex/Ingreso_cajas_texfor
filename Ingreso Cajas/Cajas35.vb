Imports C1.Win.C1FlexGrid

Public Class Cajas35
    Inherits System.Windows.Forms.Form
    Dim cnn As New SqlClient.SqlConnection()
    Dim dt As New DataTable
    Dim co As New DataTable
    Dim es As New DataTable
    Friend WithEvents graba As System.Windows.Forms.Button
    Friend WithEvents fg As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents cliente As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Cancela As System.Windows.Forms.Button
    Dim dr As DataRow
    Dim tp As New DataTable
    Dim dj As DataRow
    Dim ta As String = "|XS|S|M|L|XL|XL2|XL3|XL4|XL5|XL6"
    Friend WithEvents upc As C1.Win.C1FlexGrid.C1FlexGrid
    Dim ts(10) As String
    Friend WithEvents si As System.Windows.Forms.Button
    Dim dato As String
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Cajas35))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cliente = New System.Windows.Forms.ComboBox()
        Me.Cancela = New System.Windows.Forms.Button()
        Me.si = New System.Windows.Forms.Button()
        Me.graba = New System.Windows.Forms.Button()
        Me.fg = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.upc = New C1.Win.C1FlexGrid.C1FlexGrid()
        CType(Me.fg, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.upc, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cliente
        '
        Me.cliente.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.cliente.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cliente.Items.AddRange(New Object() {"JT", "TRECENTO", "ZUNTEX"})
        Me.cliente.Location = New System.Drawing.Point(201, 21)
        Me.cliente.Name = "cliente"
        Me.cliente.Size = New System.Drawing.Size(222, 28)
        Me.cliente.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.cliente, "Cliente.")
        '
        'Cancela
        '
        Me.Cancela.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.Cancela.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cancela.ForeColor = System.Drawing.Color.Black
        Me.Cancela.Image = CType(resources.GetObject("Cancela.Image"), System.Drawing.Image)
        Me.Cancela.Location = New System.Drawing.Point(814, 12)
        Me.Cancela.Name = "Cancela"
        Me.Cancela.Size = New System.Drawing.Size(60, 40)
        Me.Cancela.TabIndex = 104
        Me.Cancela.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.Cancela, "Presione este Boton para Cancelar toda la operación y limpiar todos los datos sin" & _
        " Grabar.")
        Me.Cancela.UseVisualStyleBackColor = False
        '
        'si
        '
        Me.si.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.si.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.si.ForeColor = System.Drawing.Color.Black
        Me.si.Image = CType(resources.GetObject("si.Image"), System.Drawing.Image)
        Me.si.Location = New System.Drawing.Point(814, 12)
        Me.si.Name = "si"
        Me.si.Size = New System.Drawing.Size(60, 40)
        Me.si.TabIndex = 1
        Me.si.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.si, "Chequeo de Datos")
        Me.si.UseVisualStyleBackColor = False
        '
        'graba
        '
        Me.graba.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.graba.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.graba.ForeColor = System.Drawing.Color.Black
        Me.graba.Image = CType(resources.GetObject("graba.Image"), System.Drawing.Image)
        Me.graba.Location = New System.Drawing.Point(1043, 593)
        Me.graba.Name = "graba"
        Me.graba.Size = New System.Drawing.Size(60, 40)
        Me.graba.TabIndex = 9
        Me.graba.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.graba, "Graba datos.")
        Me.graba.UseVisualStyleBackColor = False
        '
        'fg
        '
        Me.fg.AllowEditing = False
        Me.fg.AllowFiltering = True
        Me.fg.ColumnInfo = resources.GetString("fg.ColumnInfo")
        Me.fg.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.fg.Location = New System.Drawing.Point(7, 65)
        Me.fg.Name = "fg"
        Me.fg.Rows.DefaultSize = 21
        Me.fg.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row
        Me.fg.Size = New System.Drawing.Size(1120, 371)
        Me.fg.StyleInfo = resources.GetString("fg.StyleInfo")
        Me.fg.TabIndex = 92
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.SteelBlue
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(12, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(173, 32)
        Me.Label3.TabIndex = 94
        Me.Label3.Text = "Cliente:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'upc
        '
        Me.upc.AllowFiltering = True
        Me.upc.AutoClipboard = True
        Me.upc.ColumnInfo = resources.GetString("upc.ColumnInfo")
        Me.upc.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.upc.HighLight = C1.Win.C1FlexGrid.HighLightEnum.Never
        Me.upc.Location = New System.Drawing.Point(7, 471)
        Me.upc.Name = "upc"
        Me.upc.Rows.DefaultSize = 21
        Me.upc.Size = New System.Drawing.Size(1124, 92)
        Me.upc.StyleInfo = resources.GetString("upc.StyleInfo")
        Me.upc.TabIndex = 130
        '
        'Cajas35
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1139, 656)
        Me.Controls.Add(Me.upc)
        Me.Controls.Add(Me.si)
        Me.Controls.Add(Me.cliente)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.fg)
        Me.Controls.Add(Me.graba)
        Me.Controls.Add(Me.Cancela)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Cajas35"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Revisa UPC"
        CType(Me.fg, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.upc, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub Cajas35_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        AddHandler cliente.KeyPress, AddressOf keypressed1
        upc.DragMode = DragModeEnum.AutomaticCopy
        upc.DropMode = DropModeEnum.Automatic
        llena_tablas(es, "SELECT * FROM E_TALLAS", cnn)
        llena_combos(cliente, "SELECT  DISTINCT CLIENTE FROM CLIENTES WHERE ACTIVO = 'S'", "CLIENTE")
        Try
            cliente.SelectedIndex = 0
        Catch
        End Try
        ts = ta.Split("|")
        limpia_variables()
    End Sub
    Private Sub limpia_variables()
        setea_grid()
        cliente.Enabled = True
        si.Visible = True
        cliente.Focus()
    End Sub
    Private Sub habilita()
        si.Visible = False
        graba.Visible = True
        cliente.Enabled = False
    End Sub
    Private Sub setea_grid()
        fg.Rows.Count = 1
        fg.Rows.Fixed = 1
        fg.Rows(0).Height = 30
        upc.Rows.Count = 1
        upc.Rows.Count = 2
        upc.Rows(0).Height = 30
        upc.Rows(1).Height = 30
    End Sub
    Private Sub si_Click(sender As System.Object, e As System.EventArgs) Handles si.Click
        llena_corte()
        si.Visible = False
        cliente.Enabled = False
        Cancela.Visible = True
    End Sub

    Private Sub llena_corte()
        Dim dr As DataRow
        Dim l As Integer = 1
        Dim i As Integer
        llena_tablas(co, "select DISTINCT CPO_D.ESTILO + CPO_D.COLOR ,CPO_D.ESTILO,CPO_D.COLOR,CPO_D.ESCALA ,UPC.* FROM CPO_D LEFT JOIN CPO ON CPO_D.CPO = CPO.CPO LEFT JOIN UPC ON CPO.CLIENTE = UPC.CLIENTE AND CPO_D.ESTILO = UPC.ESTILO AND CPO_D.COLOR = UPC.COLOR WHERE  ESTADO = 'A' AND CPO.CLIENTE = '" & cliente.Text & "'", cnn)
        fg.Rows.Count = co.Rows.Count + 1
        For Each dr In co.Rows
            fg(l, 1) = dr("ESTILO")
            fg(l, 2) = dr("COLOR")
            fg(l, 3) = dr("ESCALA")
            For i = 0 To 9
                fg(l, i + 4) = dr(i + 9)
            Next
            l = l + 1
        Next
    End Sub
    Private Sub muestra_upc()
        Dim i As Integer
        Dim f As Integer = fg.RowSel
        Dim dj As DataRow = Nothing
        llena_talla(fg(f, 3), dj)
        For i = 1 To 10
            upc(0, i) = dj(i + 1)
            upc(1, i) = fg(f, i + 3)
        Next
    End Sub
    Private Sub llena_talla(ByVal escala As String, ByRef dr As DataRow)
        Dim dd As DataRow()
        dd = es.Select("ESCALA = '" & escala & "'")
        If dd.Length > 0 Then
            dr = dd(0)
        Else
            dr = Nothing
        End If
    End Sub
    '================================== HANDLERS ================================
    Private Sub keypressed1(ByVal o As [Object], ByVal e As KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            If Trim(cliente.Text) <> "" Then
                si.Focus()
            End If
        End If
    End Sub 'keypressed
    Private Sub revisa_upc(ByRef ok As Boolean)
        Dim i As Integer
        ok = False
        For i = 1 To 10
            fg(1, i) = Trim(fg(1, i))
            If fg(1, i) <> Nothing Then
                ok = True
            End If
        Next
    End Sub

    Private Sub graba_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles graba.Click
        Dim seguro As MsgBoxResult
        Dim ok As Boolean
        seguro = MsgBox("Seguro de Actualizar todos los Cambios Efectuados?  ", MsgBoxStyle.YesNo, "Actualizacion de Datos !!!")
        If seguro = MsgBoxResult.Yes Then
            revisa_upc(ok)
            If ok Then
                graba_datos()
                setea_grid()
                llena_corte()
            Else
                MsgBox("Al menos debe ingresear un Código de UPC.", MsgBoxStyle.Critical, "Por favor revise !!!")
            End If
        Else
            MsgBox("Aún no ha ingresado unidades a la nueva Caja", MsgBoxStyle.Critical, "Por favor revise !!!")
        End If

    End Sub
    Private Sub cliente_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cliente.KeyPress
        AutoCompletar(cliente, e)
    End Sub

    Private Sub Cancela_Click(sender As System.Object, e As System.EventArgs) Handles Cancela.Click
        limpia_variables()
        cliente.Focus()
    End Sub

    Private Sub fg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles upc.KeyDown
        Dim tarr() As String
        tarr = Clipboard.GetText().Split(Environment.NewLine)
        If e.Control Then
            Select Case e.KeyCode
                Case Keys.Insert, Keys.C
                    'copy
                    'System.Windows.Forms.Clipboard.SetDataObject(fg.Clip, True)
                    dato = upc(upc.RowSel, upc.ColSel)
                    Exit Select
                Case Keys.V 'paste
                    Dim rg As CellRange = upc.Selection
                    If dato <> "" Then
                        rg.Data = tarr
                    End If
            End Select
        End If
    End Sub
    '============================= Actualiza la Base de Datos =============================
    Private Sub graba_datos()
        Dim afectados As Integer
        Dim strsql As String
        Dim f As Integer = fg.RowSel
        Dim material As String
        Dim obj As New empresas()
        cnn.Open()
        ' Comienza la transacción
        Dim transaccion As SqlClient.SqlTransaction = cnn.BeginTransaction()

        ' Crea el comando para la transacción
        Dim comando As SqlClient.SqlCommand = cnn.CreateCommand()
        comando.Transaction = transaccion

        material = fg(f, 1) + fg(f, 2)
        If material.Length > 0 Then
            material = Mid(material, 1, 30)
        End If
        Try

            '============== actualiza  =======================
            strsql = "UPDATE UPC SET T1 ='" & upc(1, 1) & "', T2 = '" & upc(1, 2) & "', T3 ='" & upc(1, 3) & "', T4 ='" & upc(1, 4) & "', T5 ='" & upc(1, 5) & "', T6 = '" & upc(1, 6) & "', T7 ='" & upc(1, 7) & "', T8 ='" & upc(1, 8) & "', T9 ='" & upc(1, 9) & "', T0 = '" & upc(1, 10) & "' " & _
                            " WHERE CLIENTE = '" & cliente.Text & "' AND ESTILO = '" & fg(f, 1) & "' AND COLOR = '" & fg(f, 2) & "'"
            comando.CommandText = strsql
            afectados = comando.ExecuteNonQuery()

            If afectados = 0 Then
                strsql = "INSERT INTO UPC (CLIENTE,ESTILO,COLOR,MATERIAL,ESCALA,T1,T2,T3,T4,T5,T6,T7,T8,T9,T0,USUARIO,FECHA) VALUES ('" & _
                          cliente.Text & "','" & _
                          fg(f, 1) & "','" & _
                          fg(f, 2) & "','" & _
                          material & "','" & _
                          fg(f, 3) & "','" & _
                          upc(1, 1) & "','" & _
                          upc(1, 2) & "','" & _
                          upc(1, 3) & "','" & _
                          upc(1, 4) & "','" & _
                          upc(1, 5) & "','" & _
                          upc(1, 6) & "','" & _
                          upc(1, 7) & "','" & _
                          upc(1, 8) & "','" & _
                          upc(1, 9) & "','" & _
                          upc(1, 10) & "','" & _
                          obj.usuario & "',GETDATE() )"
                comando.CommandText = strsql
                comando.ExecuteNonQuery()
            End If
            transaccion.Commit()
            MsgBox("Actualización Exitosa.", MsgBoxStyle.Exclamation, "Datos Actualizados.")
        Catch e As Exception
            Try
                MsgBox("Inconsistencia en Datos", MsgBoxStyle.Critical, "Por favor revise !!!!")
                transaccion.Rollback()
            Catch ex As SqlClient.SqlException
                If Not transaccion.Connection Is Nothing Then
                    Console.WriteLine("Ocurrió un error de tipo " & ex.GetType().ToString() & _
                                      " se generó cuando se trato de eliminar la transacción.")
                End If
            End Try
        Finally
            cnn.Close()
        End Try
    End Sub

    Private Sub fg_Click(sender As System.Object, e As System.EventArgs) Handles fg.SelChange
        Try
            muestra_upc()
        Catch
        End Try
    End Sub
End Class

