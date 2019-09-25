Imports C1.Win.C1FlexGrid

Public Class Cajas22
    Inherits System.Windows.Forms.Form
    Dim cnn As New SqlClient.SqlConnection()
    Dim dt As New DataTable()
    Friend WithEvents fg As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents CORTE As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents S1 As System.Windows.Forms.Button
    Friend WithEvents impre As System.Windows.Forms.Button
    Dim seccion As String = ""
    Friend WithEvents CheckBox1 As CheckBox
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Cajas22))
        Me.CORTE = New System.Windows.Forms.ComboBox()
        Me.fg = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.S1 = New System.Windows.Forms.Button()
        Me.impre = New System.Windows.Forms.Button()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        CType(Me.fg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CORTE
        '
        Me.CORTE.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.CORTE.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CORTE.Items.AddRange(New Object() {"JT", "TRECENTO", "ZUNTEX"})
        Me.CORTE.Location = New System.Drawing.Point(92, 18)
        Me.CORTE.Name = "CORTE"
        Me.CORTE.Size = New System.Drawing.Size(189, 28)
        Me.CORTE.TabIndex = 93
        '
        'fg
        '
        Me.fg.AllowFiltering = True
        Me.fg.ColumnInfo = resources.GetString("fg.ColumnInfo")
        Me.fg.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.fg.Location = New System.Drawing.Point(11, 66)
        Me.fg.Name = "fg"
        Me.fg.Rows.DefaultSize = 21
        Me.fg.Size = New System.Drawing.Size(1141, 618)
        Me.fg.StyleInfo = resources.GetString("fg.StyleInfo")
        Me.fg.TabIndex = 92
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(20, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 32)
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
        Me.S1.Location = New System.Drawing.Point(911, 12)
        Me.S1.Name = "S1"
        Me.S1.Size = New System.Drawing.Size(60, 40)
        Me.S1.TabIndex = 95
        Me.S1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.S1.UseVisualStyleBackColor = False
        '
        'impre
        '
        Me.impre.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.impre.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.impre.ForeColor = System.Drawing.Color.Black
        Me.impre.Image = CType(resources.GetObject("impre.Image"), System.Drawing.Image)
        Me.impre.Location = New System.Drawing.Point(977, 12)
        Me.impre.Name = "impre"
        Me.impre.Size = New System.Drawing.Size(60, 40)
        Me.impre.TabIndex = 129
        Me.impre.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.impre.UseVisualStyleBackColor = False
        Me.impre.Visible = False
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.ForeColor = System.Drawing.Color.Black
        Me.CheckBox1.Location = New System.Drawing.Point(1080, 40)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(72, 20)
        Me.CheckBox1.TabIndex = 130
        Me.CheckBox1.Text = "Todas"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Cajas22
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1157, 696)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.S1)
        Me.Controls.Add(Me.CORTE)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.fg)
        Me.Controls.Add(Me.impre)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Red
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Cajas22"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Re-impresion de Etiquetas"
        CType(Me.fg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    Private Sub Cortes_status(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim strsql As String = "SELECT CORTE FROM CORTES WHERE EXPORTADO  <> 'S' AND CORTE IN (SELECT CORTE FROM CAJAS01) ORDER BY CORTE"
        llena_combos(CORTE, strsql, "CORTE")
        Try
            CORTE.SelectedIndex = 0
        Catch
        End Try
        setea_fg()
    End Sub

    Private Sub setea_fg()
        fg.Rows.Count = 1
        fg.Rows(0).Height = 30
        fg.Enabled = False
    End Sub
    Private Sub llena_fg()
        Dim fil As Integer = 1
        Dim dr As DataRow
        Dim strsql As String = "SELECT CAJA,CPO,ESTILO,COLOR,TALLA,CAJAS01.CLIENTE,SUM(UNIDADES) AS UNIDADES FROM CAJAS01 LEFT JOIN CORTES ON CAJAS01.CORTE = CORTES.CORTE WHERE CAJAS01.CORTE = '" & CORTE.Text & "' GROUP BY CAJA,CPO,ESTILO,COLOR,TALLA,CAJAS01.CLIENTE"
        llena_tablas(dt, strsql, cnn)
        fg.Rows.Count = dt.Rows.Count + 1
        For Each dr In dt.Rows
            fg(fil, 1) = dr("CAJA")
            fg(fil, 2) = dr("CPO")
            fg(fil, 3) = dr("ESTILO")
            fg(fil, 4) = dr("COLOR")
            fg(fil, 5) = dr("CLIENTE")
            fg(fil, 6) = dr("TALLA")
            fg(fil, 7) = dr("UNIDADES")
            fg(fil, 8) = False
            fil = fil + 1
        Next
    End Sub

    Private Sub empresa_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles CORTE.KeyPress
        AutoCompletar(CORTE, e)
    End Sub
    Private Sub corte_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CORTE.KeyDown
        If e.KeyCode = Keys.Delete Then
            e.Handled = True
        End If
    End Sub

    Private Sub impre_Click(sender As System.Object, e As System.EventArgs) Handles impre.Click
        Dim i As Integer
        Dim pr As New C1cajas.prt
        Dim obj As New empresas
        For i = 1 To fg.Rows.Count - 1
            If fg(i, 8) = True Then
                pr.imprime_cajas_s(fg(i, 1), fg(i, 1), obj.seccion, obj.numero, obj.constr)
            End If
        Next
        Close()
    End Sub

    Private Sub S1_Click_1(sender As System.Object, e As System.EventArgs) Handles S1.Click
        llena_fg()
        S1.Visible = False
        impre.Visible = True
        fg.Enabled = True
        CORTE.Enabled = False
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim ok As Boolean = False
        Dim i As Integer
        If CheckBox1.Checked = True Then
            ok = True
        End If
        For i = 1 To fg.Rows.Count - 1
            fg(i, 8) = ok
        Next
    End Sub
End Class

