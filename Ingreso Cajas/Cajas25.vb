Imports C1.Win.C1FlexGrid

Public Class Cajas25
    Inherits System.Windows.Forms.Form
    Dim cnn As New SqlClient.SqlConnection()
    Dim dt As New DataTable
    Dim co As New DataTable
    Dim pp As New DataTable
    Friend WithEvents graba As System.Windows.Forms.Button
    Friend WithEvents fg As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents corte As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents si As System.Windows.Forms.Button
    Friend WithEvents Cancela As System.Windows.Forms.Button
    Dim dr As DataRow
    Dim cliente As String
    Dim tp As New DataTable
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents escala As System.Windows.Forms.Label
    Friend WithEvents estilo As System.Windows.Forms.Label
    Friend WithEvents colo As System.Windows.Forms.Label
    Dim dj As DataRow
    Dim ta As String = "|XS|S|M|L|XL|XL2|XL3|XL4|XL5|XL6"
    Friend WithEvents material As System.Windows.Forms.TextBox
    Friend WithEvents upc As C1.Win.C1FlexGrid.C1FlexGrid
    Dim ts(10) As String
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Cajas25))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.corte = New System.Windows.Forms.ComboBox()
        Me.si = New System.Windows.Forms.Button()
        Me.Cancela = New System.Windows.Forms.Button()
        Me.graba = New System.Windows.Forms.Button()
        Me.fg = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.escala = New System.Windows.Forms.Label()
        Me.estilo = New System.Windows.Forms.Label()
        Me.colo = New System.Windows.Forms.Label()
        Me.material = New System.Windows.Forms.TextBox()
        Me.upc = New C1.Win.C1FlexGrid.C1FlexGrid()
        CType(Me.fg, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.upc, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'corte
        '
        Me.corte.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.corte.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.corte.Items.AddRange(New Object() {"JT", "TRECENTO", "ZUNTEX"})
        Me.corte.Location = New System.Drawing.Point(201, 21)
        Me.corte.Name = "corte"
        Me.corte.Size = New System.Drawing.Size(222, 28)
        Me.corte.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.corte, "No. Corte.")
        '
        'si
        '
        Me.si.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.si.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.si.ForeColor = System.Drawing.Color.Black
        Me.si.Image = CType(resources.GetObject("si.Image"), System.Drawing.Image)
        Me.si.Location = New System.Drawing.Point(909, 19)
        Me.si.Name = "si"
        Me.si.Size = New System.Drawing.Size(60, 40)
        Me.si.TabIndex = 1
        Me.si.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.si, "Chequeo de Datos")
        Me.si.UseVisualStyleBackColor = False
        '
        'Cancela
        '
        Me.Cancela.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.Cancela.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cancela.ForeColor = System.Drawing.Color.Black
        Me.Cancela.Image = CType(resources.GetObject("Cancela.Image"), System.Drawing.Image)
        Me.Cancela.Location = New System.Drawing.Point(977, 19)
        Me.Cancela.Name = "Cancela"
        Me.Cancela.Size = New System.Drawing.Size(60, 40)
        Me.Cancela.TabIndex = 104
        Me.Cancela.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.Cancela, "Presione este Boton para Cancelar toda la operación y limpiar todos los datos sin" & _
        " Grabar.")
        Me.Cancela.UseVisualStyleBackColor = False
        '
        'graba
        '
        Me.graba.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.graba.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.graba.ForeColor = System.Drawing.Color.Black
        Me.graba.Image = CType(resources.GetObject("graba.Image"), System.Drawing.Image)
        Me.graba.Location = New System.Drawing.Point(909, 19)
        Me.graba.Name = "graba"
        Me.graba.Size = New System.Drawing.Size(60, 40)
        Me.graba.TabIndex = 9
        Me.graba.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.graba.UseVisualStyleBackColor = False
        '
        'fg
        '
        Me.fg.AllowEditing = False
        Me.fg.AllowFiltering = True
        Me.fg.ColumnInfo = resources.GetString("fg.ColumnInfo")
        Me.fg.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.fg.HighLight = C1.Win.C1FlexGrid.HighLightEnum.Never
        Me.fg.Location = New System.Drawing.Point(7, 222)
        Me.fg.Name = "fg"
        Me.fg.Rows.DefaultSize = 21
        Me.fg.Size = New System.Drawing.Size(1120, 97)
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
        Me.Label3.Text = "Corte:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.SteelBlue
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(12, 166)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(173, 32)
        Me.Label1.TabIndex = 107
        Me.Label1.Text = "Material:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.SteelBlue
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(12, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(173, 32)
        Me.Label2.TabIndex = 108
        Me.Label2.Text = "Estilo:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.SteelBlue
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(12, 93)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(173, 32)
        Me.Label4.TabIndex = 109
        Me.Label4.Text = "Color:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.SteelBlue
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(12, 129)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(173, 32)
        Me.Label5.TabIndex = 110
        Me.Label5.Text = "Escala:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'escala
        '
        Me.escala.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.escala.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.escala.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.escala.Location = New System.Drawing.Point(202, 131)
        Me.escala.Name = "escala"
        Me.escala.Size = New System.Drawing.Size(92, 32)
        Me.escala.TabIndex = 111
        Me.escala.Text = " "
        Me.escala.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'estilo
        '
        Me.estilo.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.estilo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.estilo.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.estilo.Location = New System.Drawing.Point(202, 56)
        Me.estilo.Name = "estilo"
        Me.estilo.Size = New System.Drawing.Size(428, 32)
        Me.estilo.TabIndex = 112
        Me.estilo.Text = " "
        Me.estilo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'colo
        '
        Me.colo.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.colo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.colo.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.colo.Location = New System.Drawing.Point(202, 93)
        Me.colo.Name = "colo"
        Me.colo.Size = New System.Drawing.Size(428, 32)
        Me.colo.TabIndex = 113
        Me.colo.Text = " "
        Me.colo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'material
        '
        Me.material.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.material.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.material.Location = New System.Drawing.Point(201, 171)
        Me.material.Name = "material"
        Me.material.Size = New System.Drawing.Size(429, 26)
        Me.material.TabIndex = 129
        '
        'upc
        '
        Me.upc.AllowFiltering = True
        Me.upc.AutoClipboard = True
        Me.upc.ColumnInfo = resources.GetString("upc.ColumnInfo")
        Me.upc.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.upc.HighLight = C1.Win.C1FlexGrid.HighLightEnum.Never
        Me.upc.Location = New System.Drawing.Point(7, 346)
        Me.upc.Name = "upc"
        Me.upc.Rows.DefaultSize = 21
        Me.upc.Size = New System.Drawing.Size(1124, 92)
        Me.upc.StyleInfo = resources.GetString("upc.StyleInfo")
        Me.upc.TabIndex = 130
        '
        'Cajas25
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1139, 603)
        Me.Controls.Add(Me.upc)
        Me.Controls.Add(Me.material)
        Me.Controls.Add(Me.colo)
        Me.Controls.Add(Me.estilo)
        Me.Controls.Add(Me.escala)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Cancela)
        Me.Controls.Add(Me.si)
        Me.Controls.Add(Me.corte)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.fg)
        Me.Controls.Add(Me.graba)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Cajas25"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Revisa UPC"
        CType(Me.fg, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.upc, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    Private Sub Cajas05_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        AddHandler corte.KeyPress, AddressOf keypressed1
        upc.DragMode = DragModeEnum.AutomaticCopy
        upc.DropMode = DropModeEnum.Automatic
        llena_combos(corte, "SELECT  DISTINCT CORTE FROM CORTES WHERE CORTE NOT IN (SELECT DISTINCT CORTE FROM PROD_DIARIA) AND EXPORTADO <> 'S' ORDER BY CORTE", "CORTE")
        Try
            corte.SelectedIndex = 0
        Catch
        End Try
        ts = ta.Split("|")
        limpia_variables()
    End Sub
    Private Sub limpia_variables()
        setea_GRID()
        corte.Enabled = True
        escala.Text = ""
        estilo.Text = ""
        colo.Text = ""
        material.Text = ""
        si.Visible = True
        graba.Visible = False
        corte.Focus()
    End Sub
    Private Sub habilita()
        si.Visible = False
        graba.Visible = True
        corte.Enabled = False
    End Sub
    Private Sub setea_grid()
        fg.Rows.Count = 1
        fg.Rows.Fixed = 1
        fg.Rows.Count = 2
        fg.Rows(0).Height = 30
        fg.Rows(1).Height = 30
        upc.Rows.Count = 1
        upc.Rows.Count = 2
        upc.Rows(0).Height = 30
        upc.Rows(1).Height = 30
    End Sub
    Private Sub si_Click(sender As System.Object, e As System.EventArgs) Handles si.Click
        llena_corte()
    End Sub

    Private Sub llena_corte()
        Dim i As Integer
        llena_tablas(co, "SELECT CORTES.*,CPO_D.ESCALA, E_TALLAS.*,UPC.* FROM CORTES LEFT JOIN CPO_D ON CPO_D.CPO = CORTES.CPO AND CPO_D.ESTILO = CORTES.ESTILO AND CPO_D.COLOR = CORTES.COLOR LEFT JOIN E_TALLAS ON CPO_D.ESCALA = E_TALLAS.ESCALA LEFT JOIN UPC ON CORTES.ESTILO = UPC.ESTILO AND CORTES.COLOR = UPC.COLOR AND CORTES.CLIENTE = UPC.CLIENTE WHERE CORTE = '" & corte.Text & "'", cnn)
        If co.Rows.Count > 0 Then
            dr = co.Rows(0)
            dj = dr
            escala.Text = dr("CPO")
            estilo.Text = dr("ESTILO")
            colo.Text = dr("COLOR")
            cliente = dr("CLIENTE")
            escala.Text = dr("ESCALA")
            Try
                material.Text = dr("MATERIAL")
            Catch
                material.Text = estilo.Text & "_" & colo.Text
                material.Text = Mid(material.Text, 1, 30)
            End Try

            fg(1, 1) = "Cortado"
            For i = 1 To 10
                If escala.Text = "00" Then
                    dj(i + 24) = ts(i)
                End If
                fg(0, i + 1) = dj(i + 24)
                upc(0, i) = dj(i + 24)
                fg(1, i + 1) = dj(i + 5)
                upc(1, i) = dj(i + 39)
            Next
        End If
        habilita()
    End Sub

    '================================== HANDLERS ================================
    Private Sub keypressed1(ByVal o As [Object], ByVal e As KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            If Trim(corte.Text) <> "" Then
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
                limpia_variables()
            Else
                MsgBox("Al menos debe ingresear un Código de UPC.", MsgBoxStyle.Critical, "Por favor revise !!!")
            End If
        Else
            MsgBox("Aún no ha ingresado unidades a la nueva Caja", MsgBoxStyle.Critical, "Por favor revise !!!")
        End If

    End Sub
    Private Sub corteS_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles corte.KeyPress
        AutoCompletar(corte, e)
    End Sub

    Private Sub Cancela_Click(sender As System.Object, e As System.EventArgs) Handles Cancela.Click
        limpia_variables()
        corte.Focus()
    End Sub

    Private Sub fg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles upc.KeyDown
        Dim tarr() As String
        tArr = Clipboard.GetText().Split(Environment.NewLine)
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
        Dim obj As New empresas()
        cnn.Open()
        ' Comienza la transacción
        Dim transaccion As SqlClient.SqlTransaction = cnn.BeginTransaction()

        ' Crea el comando para la transacción
        Dim comando As SqlClient.SqlCommand = cnn.CreateCommand()
        comando.Transaction = transaccion

        Try

            '============== actualiza  =======================
            strsql = "UPDATE UPC SET T1 ='" & upc(1, 1) & "', T2 = '" & upc(1, 2) & "', T3 ='" & upc(1, 3) & "', T4 ='" & upc(1, 4) & "', T5 ='" & upc(1, 5) & "', T6 = '" & upc(1, 6) & "', T7 ='" & upc(1, 7) & "', T8 ='" & upc(1, 8) & "', T9 ='" & upc(1, 9) & "', T0 = '" & upc(1, 10) & "' " & _
                            " WHERE CLIENTE = '" & cliente & "' AND ESTILO = '" & estilo.Text & "' AND COLOR = '" & colo.Text & "'"
            comando.CommandText = strsql
            afectados = comando.ExecuteNonQuery()

            If afectados = 0 Then
                strsql = "INSERT INTO UPC (CLIENTE,ESTILO,COLOR,MATERIAL,ESCALA,T1,T2,T3,T4,T5,T6,T7,T8,T9,T0,USUARIO,FECHA) VALUES ('" & _
                          cliente & "','" & _
                          estilo.Text & "','" & _
                          colo.Text & "','" & _
                          material.Text & "','" & _
                          escala.Text & "','" & _
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

End Class

