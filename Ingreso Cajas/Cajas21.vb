Imports C1.Win.C1FlexGrid

Public Class Cajas21
    Inherits System.Windows.Forms.Form
    Dim cliente As String
    Dim escala As String
    Dim fila As Integer
    Dim dt As New DataTable()
    Dim mt As New DataTable()
    Dim es As New DataTable()
    Dim dr As DataRow
    Dim cnn As New SqlClient.SqlConnection()
    Friend WithEvents ec As C1.Win.C1FlexGrid.C1FlexGrid
    Dim lineas As Integer
    Dim empre As New empresas
    Dim clave As String
    Friend WithEvents estil As System.Windows.Forms.Label
    Friend WithEvents colo As System.Windows.Forms.Label
    Dim ac As New DataTable
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
    Friend WithEvents graba As System.Windows.Forms.Button
    Friend WithEvents quita As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents fg As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents grupo As System.Windows.Forms.GroupBox
    Friend WithEvents material As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Cajas21))
        Me.fg = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.graba = New System.Windows.Forms.Button()
        Me.quita = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.grupo = New System.Windows.Forms.GroupBox()
        Me.estil = New System.Windows.Forms.Label()
        Me.colo = New System.Windows.Forms.Label()
        Me.material = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ec = New C1.Win.C1FlexGrid.C1FlexGrid()
        CType(Me.fg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grupo.SuspendLayout()
        CType(Me.ec, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fg
        '
        Me.fg.AllowDragging = C1.Win.C1FlexGrid.AllowDraggingEnum.None
        Me.fg.AllowEditing = False
        Me.fg.AllowResizing = C1.Win.C1FlexGrid.AllowResizingEnum.None
        Me.fg.ColumnInfo = "10,1,0,0,0,95,Columns:"
        Me.fg.FocusRect = C1.Win.C1FlexGrid.FocusRectEnum.None
        Me.fg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.fg.ForeColor = System.Drawing.Color.Black
        Me.fg.HighLight = C1.Win.C1FlexGrid.HighLightEnum.WithFocus
        Me.fg.Location = New System.Drawing.Point(8, 543)
        Me.fg.Name = "fg"
        Me.fg.Rows.DefaultSize = 19
        Me.fg.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row
        Me.fg.Size = New System.Drawing.Size(1027, 88)
        Me.fg.TabIndex = 22
        Me.ToolTip1.SetToolTip(Me.fg, "Area de Datos.")
        '
        'graba
        '
        Me.graba.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.graba.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.graba.ForeColor = System.Drawing.Color.Black
        Me.graba.Location = New System.Drawing.Point(837, 19)
        Me.graba.Name = "graba"
        Me.graba.Size = New System.Drawing.Size(60, 40)
        Me.graba.TabIndex = 8
        Me.graba.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.graba, "Presione este Boton para Grabar o Actualizar los Datos de la FPO.")
        Me.graba.UseVisualStyleBackColor = False
        '
        'quita
        '
        Me.quita.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.quita.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.quita.ForeColor = System.Drawing.Color.Black
        Me.quita.Location = New System.Drawing.Point(837, 19)
        Me.quita.Name = "quita"
        Me.quita.Size = New System.Drawing.Size(60, 40)
        Me.quita.TabIndex = 21
        Me.quita.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.quita, "Presione este Boton para Borrar el Registro Seleccionado.")
        Me.quita.UseVisualStyleBackColor = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.Button3.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.Location = New System.Drawing.Point(903, 19)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(60, 40)
        Me.Button3.TabIndex = 19
        Me.Button3.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.Button3, "Presione este Boton para Refrescar la pantalla sin efectuar modificaciones.")
        Me.Button3.UseVisualStyleBackColor = False
        '
        'grupo
        '
        Me.grupo.BackColor = System.Drawing.Color.White
        Me.grupo.Controls.Add(Me.estil)
        Me.grupo.Controls.Add(Me.colo)
        Me.grupo.Controls.Add(Me.material)
        Me.grupo.Controls.Add(Me.Label4)
        Me.grupo.Controls.Add(Me.Label2)
        Me.grupo.Controls.Add(Me.Label3)
        Me.grupo.Controls.Add(Me.Button3)
        Me.grupo.Controls.Add(Me.graba)
        Me.grupo.Controls.Add(Me.quita)
        Me.grupo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grupo.ForeColor = System.Drawing.Color.Black
        Me.grupo.Location = New System.Drawing.Point(8, 414)
        Me.grupo.Name = "grupo"
        Me.grupo.Size = New System.Drawing.Size(1027, 122)
        Me.grupo.TabIndex = 1
        Me.grupo.TabStop = False
        '
        'estil
        '
        Me.estil.BackColor = System.Drawing.Color.White
        Me.estil.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.estil.Location = New System.Drawing.Point(168, 16)
        Me.estil.Name = "estil"
        Me.estil.Size = New System.Drawing.Size(273, 24)
        Me.estil.TabIndex = 54
        Me.estil.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'colo
        '
        Me.colo.BackColor = System.Drawing.Color.White
        Me.colo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.colo.Location = New System.Drawing.Point(167, 45)
        Me.colo.Name = "colo"
        Me.colo.Size = New System.Drawing.Size(275, 24)
        Me.colo.TabIndex = 53
        Me.colo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'material
        '
        Me.material.Location = New System.Drawing.Point(168, 83)
        Me.material.MaxLength = 25
        Me.material.Name = "material"
        Me.material.Size = New System.Drawing.Size(273, 21)
        Me.material.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.material, "Color")
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.SteelBlue
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(16, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(144, 24)
        Me.Label4.TabIndex = 46
        Me.Label4.Text = "Material:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.SteelBlue
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(16, 45)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(144, 24)
        Me.Label2.TabIndex = 45
        Me.Label2.Text = "Color:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.SteelBlue
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(16, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(144, 24)
        Me.Label3.TabIndex = 44
        Me.Label3.Text = "Estilo:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ec
        '
        Me.ec.AllowDragging = C1.Win.C1FlexGrid.AllowDraggingEnum.None
        Me.ec.AllowEditing = False
        Me.ec.AllowResizing = C1.Win.C1FlexGrid.AllowResizingEnum.None
        Me.ec.ColumnInfo = "10,1,0,0,0,95,Columns:"
        Me.ec.FocusRect = C1.Win.C1FlexGrid.FocusRectEnum.None
        Me.ec.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.ec.Location = New System.Drawing.Point(8, 12)
        Me.ec.Name = "ec"
        Me.ec.Rows.DefaultSize = 19
        Me.ec.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row
        Me.ec.Size = New System.Drawing.Size(1027, 396)
        Me.ec.TabIndex = 44
        Me.ToolTip1.SetToolTip(Me.ec, "Area de Datos.")
        '
        'Cajas21
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1059, 696)
        Me.Controls.Add(Me.ec)
        Me.Controls.Add(Me.grupo)
        Me.Controls.Add(Me.fg)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Cajas21"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Maestro de UPC"
        CType(Me.fg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grupo.ResumeLayout(False)
        CType(Me.ec, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub BOM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AddHandler material.KeyPress, AddressOf keypressed1
        llena_tablas(es, "SELECT * FROM E_TALLAS", cnn)
        llena_combos(material, "select DISTINCT MATERIAL FROM UPC_C", "MATERIAL")
        setea_ec()
        Try
            selecciona_ec(1)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub setea_ec()
        ec.Rows.Count = 1
        ec.Rows(0).Height = 30
        llena_ec()
    End Sub
    Private Sub llena_ec()
        Dim dr As DataRow
        Dim dt As New DataTable
        Dim strsql As String = "SELECT DISTINCT CORTES.ESTILO+CORTES.COLOR, CORTES.CLIENTE,CORTES.ESTILO,CORTES.COLOR, CPO_D.ESCALA, MATERIAL FROM CORTES LEFT JOIN CPO_D ON CORTES.CPO = CPO_D.CPO AND CORTES.ESTILO = CPO_D. ESTILO AND CORTES.COLOR = CPO_D.COLOR LEFT JOIN UPC ON CORTES.ESTILO = UPC.ESTILO AND CORTES.COLOR =UPC.COLOR AND CPO_D.ESCALA = UPC.ESCALA WHERE CORTE NOT IN (SELECT DISTINCT CORTE FROM PROD_DIARIA) ORDER BY CORTES.CLIENTE,CORTES.ESTILO,CORTES.COLOR"
        Dim l As Integer = 1
        llena_tablas(dt, strsql, cnn)
        ec.Rows.Count = dt.Rows.Count + 1
        For Each dr In dt.Rows
            ec(l, 1) = dr("CLIENTE")
            ec(l, 2) = dr("ESTILO")
            ec(l, 3) = dr("COLOR")
            ec(l, 4) = dr("ESCALA")
            ec(l, 5) = dr("MATERIAL")
            l = l + 1
        Next
    End Sub

    Private Sub setea_grid()
        Dim i As Integer
        Dim j As Integer
        fg.Rows.Count = 2
        fg.Rows(0).Height = 30
        For i = 0 To 1
            For j = 1 To 10
                fg(i, j) = ""
            Next
        Next
        llena_grid()
    End Sub

    Private Sub llena_grid()
        Dim dd As DataRow()
        Dim dr As DataRow
        Dim dt As New DataTable
        Dim i As Integer
        Dim t As String = "|"
        Dim talla() As String = Nothing
        dd = es.Select("ESCALA = '" & escala & "'")
        If dd.Length > 0 Then
            dr = dd(0)
            For i = 1 To 10
                t = t + dr(i + 1) + "|"
                fg(0, i) = dr(i + 1)
            Next
            t = Mid(t, 1, t.Length - 1)
            talla = Split(t, "|")
        End If
        Dim strsql As String = "SELECT * FROM UPC_C WHERE MATERIAL = '" & material.Text & "'"
        llena_tablas(dt, strsql, cnn)
        For Each dr In dt.Rows
            Dim ta As String = dr("TALLA")
            i = Array.IndexOf(talla, ta)
            If i > 0 Then
                fg(1, i) = dr("UPC")
            Else
                If ta = "XXL" And escala = "00" Then
                    fg(1, 6) = dr("UPC")
                End If
            End If
        Next
    End Sub
    '================================== HANDLERS ================================

    Private Sub keypressed1(ByVal o As [Object], ByVal e As KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            setea_grid()
        End If
    End Sub 'keypressed
    Private Sub mensaje(ByVal var As String)
        MsgBox("Por favor revise " + var, MsgBoxStyle.OkOnly, var + " NO VALIDO !!!! ")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles graba.Click
        graba_registros()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles quita.Click
        Dim seguro As MsgBoxResult
        Dim elimino As Boolean = False
        seguro = MsgBox("Seguro de Querer Eliminar?  ", MsgBoxStyle.YesNo, "Eliminando Registro !!!")
        If seguro = MsgBoxResult.Yes Then
            elimina()
        End If
    End Sub
    Private Sub mate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles material.KeyPress
        AutoCompletar(material, e)
    End Sub

    Private Sub Colo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        llena_grid()
    End Sub

    REM =======================================================================================
    REM =                                                                                     =
    REM =                      GRABACIONES                                                    =
    REM =                                                                                     =
    REM =======================================================================================

    Private Sub graba_registros()
        Dim afectados As Integer
        Dim strsql As String
        Dim tipo As String = "N"
        Dim obj As New empresas()
        Dim fechas As String = Format(Now, "yyyy-MM-dd hh:mm:ss") + ".000"
        Dim cliente As String = ec(ec.RowSel, 1)
        cnn.Open()
        ' Comienza la transacción
        Dim transaccion As SqlClient.SqlTransaction = cnn.BeginTransaction()

        ' Crea el comando para la transacción
        Dim comando As SqlClient.SqlCommand = cnn.CreateCommand()
        comando.Transaction = transaccion

        Try
 
            '============== actualiza  =======================
            strsql = "UPDATE UPC SET T1 ='" & fg(1, 1) & "', T2 = '" & fg(1, 2) & "', T3 ='" & fg(1, 3) & "', T4 ='" & fg(1, 4) & "', T5 ='" & fg(1, 5) & "', T6 = '" & fg(1, 6) & "', T7 ='" & fg(1, 7) & "', T8 ='" & fg(1, 8) & "', T9 ='" & fg(1, 9) & "', T0 = '" & fg(1, 10) & "' " & _
                            " WHERE CLIENTE = '" & cliente & "' AND ESTILO = '" & estil.Text & "' AND COLOR = '" & colo.Text & "'"
            comando.CommandText = strsql
            afectados = comando.ExecuteNonQuery()

            If afectados = 0 Then
                strsql = "INSERT INTO UPC (CLIENTE,ESTILO,COLOR,MATERIAL,ESCALA,T1,T2,T3,T4,T5,T6,T7,T8,T9,T0,USUARIO,FECHA) VALUES ('" & _
                          cliente & "','" & _
                          estil.Text & "','" & _
                          colo.Text & "','" & _
                          material.Text & "','" & _
                          escala & "','" & _
                          fg(1, 1) & "','" & _
                          fg(1, 2) & "','" & _
                          fg(1, 3) & "','" & _
                          fg(1, 4) & "','" & _
                          fg(1, 5) & "','" & _
                          fg(1, 6) & "','" & _
                          fg(1, 7) & "','" & _
                          fg(1, 8) & "','" & _
                          fg(1, 9) & "','" & _
                          fg(1, 10) & "','" & _
                          obj.usuario & "',GETDATE() )"
                comando.CommandText = strsql
                comando.ExecuteNonQuery()
            End If
            transaccion.Commit()
            ec(ec.RowSel, 5) = material.Text
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

    REM =======================================================================================
    REM =                                                                                     =
    REM =                     ELIMINA REGISTSRO                                               =
    REM =                                                                                     =
    REM =======================================================================================
    Private Sub elimina()
        Dim afectados As Integer
        Dim strsql As String
        Dim obj As New empresas()
        Dim fechas As String = Format(Now, "yyyy-MM-dd hh:mm:ss")
        cnn.Open()
        ' Comienza la transacción
        Dim transaccion As SqlClient.SqlTransaction = cnn.BeginTransaction()

        ' Crea el comando para la transacción
        Dim comando As SqlClient.SqlCommand = cnn.CreateCommand()
        comando.Transaction = transaccion

        Try
            '============== ELIMINA  =======================
            strsql = "DELETE UPC WHERE CLIENTE = '" & cliente & "' AND ESTILO = '" & estil.Text & "' AND COLOR = '" & colo.Text & "'"
            comando.CommandText = strsql
            afectados = comando.ExecuteNonQuery()

            transaccion.Commit()
            ec(ec.RowSel, 5) = ""
        Catch e As Exception
            Try
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

    Private Sub ec_Click(sender As System.Object, e As System.EventArgs) Handles ec.Click
        Dim f As Integer = ec.RowSel
        selecciona_ec(f)
    End Sub

    Private Sub ec_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles ec.KeyDown
        If e.KeyCode = Keys.Up Then
            selecciona_ec(ec.Row - 1)
        End If

        If e.KeyCode = Keys.Down Then
            selecciona_ec(ec.Row + 1)
        End If
    End Sub

    Private Sub selecciona_ec(ByVal f As Integer)
        Dim i As Integer
        If f > 0 Then
            Try
                cliente = ec(f, 1)
                estil.Text = ec(f, 2)
                colo.Text = ec(f, 3)
                escala = ec(f, 4)
                i = material.FindString(ec(f, 5))
                If Trim(ec(f, 5)) <> "" Then
                    material.SelectedIndex = material.FindString(ec(f, 5))
                    graba.Visible = False
                    quita.Visible = True
                Else
                    graba.Visible = True
                    quita.Visible = False
                End If
                setea_grid()
                llena_grid()
            Catch
            End Try
        End If
    End Sub

    Private Sub material_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles material.SelectedIndexChanged
        setea_grid()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        graba.Visible = True
        quita.Visible = False
    End Sub
End Class
