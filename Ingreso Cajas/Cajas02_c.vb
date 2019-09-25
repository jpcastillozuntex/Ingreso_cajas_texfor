Imports C1.Win.C1FlexGrid

Public Class Cajas02_c
    Inherits System.Windows.Forms.Form
    Dim cnn As New SqlClient.SqlConnection()
    Dim dt As New DataTable()
    Dim co As New DataTable
    Friend WithEvents graba As System.Windows.Forms.Button
    Friend WithEvents fg As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents corte As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cajad As System.Windows.Forms.ComboBox
    Friend WithEvents tallad As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents tipon As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents sigue As System.Windows.Forms.Button
    Friend WithEvents unidn As System.Windows.Forms.TextBox
    Friend WithEvents si As System.Windows.Forms.Button
    Friend WithEvents Cancela As System.Windows.Forms.Button
    Friend WithEvents unid As System.Windows.Forms.Label
    Friend WithEvents tipod As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Dim dr As DataRow
    Dim caja_nueva As String
    Dim cliente As String
    Dim escala As String
    Friend WithEvents orden As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Dim tallap As String
    Dim seccion As String = ""
    Friend WithEvents fechad As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Cajas02_c))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.corte = New System.Windows.Forms.ComboBox()
        Me.cajad = New System.Windows.Forms.ComboBox()
        Me.sigue = New System.Windows.Forms.Button()
        Me.si = New System.Windows.Forms.Button()
        Me.Cancela = New System.Windows.Forms.Button()
        Me.tallad = New System.Windows.Forms.ComboBox()
        Me.tipon = New System.Windows.Forms.ComboBox()
        Me.tipod = New System.Windows.Forms.ComboBox()
        Me.unidn = New System.Windows.Forms.TextBox()
        Me.graba = New System.Windows.Forms.Button()
        Me.fg = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.fechad = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.orden = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.unid = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        CType(Me.fg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'corte
        '
        Me.corte.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.corte.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.corte.Location = New System.Drawing.Point(223, 19)
        Me.corte.Name = "corte"
        Me.corte.Size = New System.Drawing.Size(189, 28)
        Me.corte.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.corte, "No. corte.")
        '
        'cajad
        '
        Me.cajad.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.cajad.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cajad.Items.AddRange(New Object() {"JT", "TRECENTO", "ZUNTEX"})
        Me.cajad.Location = New System.Drawing.Point(199, 31)
        Me.cajad.Name = "cajad"
        Me.cajad.Size = New System.Drawing.Size(189, 28)
        Me.cajad.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.cajad, " No. de Caja.")
        '
        'sigue
        '
        Me.sigue.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.sigue.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sigue.ForeColor = System.Drawing.Color.Black
        Me.sigue.Image = CType(resources.GetObject("sigue.Image"), System.Drawing.Image)
        Me.sigue.Location = New System.Drawing.Point(402, 102)
        Me.sigue.Name = "sigue"
        Me.sigue.Size = New System.Drawing.Size(60, 40)
        Me.sigue.TabIndex = 7
        Me.sigue.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.sigue, "Chequeo de Datos")
        Me.sigue.UseVisualStyleBackColor = False
        '
        'si
        '
        Me.si.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.si.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.si.ForeColor = System.Drawing.Color.Black
        Me.si.Image = CType(resources.GetObject("si.Image"), System.Drawing.Image)
        Me.si.Location = New System.Drawing.Point(857, 19)
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
        Me.Cancela.Location = New System.Drawing.Point(925, 19)
        Me.Cancela.Name = "Cancela"
        Me.Cancela.Size = New System.Drawing.Size(60, 40)
        Me.Cancela.TabIndex = 104
        Me.Cancela.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.Cancela, "Presione este Boton para Cancelar toda la operación y limpiar todos los datos sin" &
        " Grabar.")
        Me.Cancela.UseVisualStyleBackColor = False
        '
        'tallad
        '
        Me.tallad.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.tallad.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tallad.Items.AddRange(New Object() {"JT", "TRECENTO", "ZUNTEX"})
        Me.tallad.Location = New System.Drawing.Point(199, 102)
        Me.tallad.Name = "tallad"
        Me.tallad.Size = New System.Drawing.Size(189, 28)
        Me.tallad.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.tallad, "Talla.")
        '
        'tipon
        '
        Me.tipon.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.tipon.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tipon.Location = New System.Drawing.Point(199, 31)
        Me.tipon.Name = "tipon"
        Me.tipon.Size = New System.Drawing.Size(189, 28)
        Me.tipon.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.tipon, "Tipo de Prenda.")
        '
        'tipod
        '
        Me.tipod.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.tipod.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tipod.Items.AddRange(New Object() {"JT", "TRECENTO", "ZUNTEX"})
        Me.tipod.Location = New System.Drawing.Point(199, 66)
        Me.tipod.Name = "tipod"
        Me.tipod.Size = New System.Drawing.Size(189, 28)
        Me.tipod.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.tipod, "Tipo de Prenda.")
        '
        'unidn
        '
        Me.unidn.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.unidn.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.unidn.Location = New System.Drawing.Point(199, 67)
        Me.unidn.MaxLength = 20
        Me.unidn.Name = "unidn"
        Me.unidn.Size = New System.Drawing.Size(189, 26)
        Me.unidn.TabIndex = 6
        Me.unidn.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'graba
        '
        Me.graba.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.graba.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.graba.ForeColor = System.Drawing.Color.Black
        Me.graba.Image = CType(resources.GetObject("graba.Image"), System.Drawing.Image)
        Me.graba.Location = New System.Drawing.Point(857, 19)
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
        Me.fg.Location = New System.Drawing.Point(23, 326)
        Me.fg.Name = "fg"
        Me.fg.Rows.DefaultSize = 21
        Me.fg.Size = New System.Drawing.Size(962, 335)
        Me.fg.StyleInfo = resources.GetString("fg.StyleInfo")
        Me.fg.TabIndex = 92
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.SteelBlue
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(29, 16)
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
        Me.Label1.Location = New System.Drawing.Point(6, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(173, 28)
        Me.Label1.TabIndex = 95
        Me.Label1.Text = "Caja:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.SteelBlue
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(6, 102)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(173, 28)
        Me.Label2.TabIndex = 96
        Me.Label2.Text = "Talla:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.SteelBlue
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(6, 173)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(173, 28)
        Me.Label5.TabIndex = 98
        Me.Label5.Text = "Unidades:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Gainsboro
        Me.GroupBox1.Controls.Add(Me.fechad)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.orden)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.tipod)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.unid)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.tallad)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.cajad)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.ForeColor = System.Drawing.Color.Black
        Me.GroupBox1.Location = New System.Drawing.Point(24, 73)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(468, 247)
        Me.GroupBox1.TabIndex = 101
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Caja Donante"
        '
        'fechad
        '
        Me.fechad.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fechad.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.fechad.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.fechad.Location = New System.Drawing.Point(196, 208)
        Me.fechad.Name = "fechad"
        Me.fechad.Size = New System.Drawing.Size(189, 28)
        Me.fechad.TabIndex = 105
        Me.fechad.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.SteelBlue
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.White
        Me.Label11.Location = New System.Drawing.Point(6, 208)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(173, 28)
        Me.Label11.TabIndex = 104
        Me.Label11.Text = "Fecha:"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'orden
        '
        Me.orden.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.orden.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.orden.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.orden.Location = New System.Drawing.Point(196, 138)
        Me.orden.Name = "orden"
        Me.orden.Size = New System.Drawing.Size(189, 28)
        Me.orden.TabIndex = 103
        Me.orden.Text = "0"
        Me.orden.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.SteelBlue
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.White
        Me.Label10.Location = New System.Drawing.Point(6, 138)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(173, 28)
        Me.Label10.TabIndex = 102
        Me.Label10.Text = "Orden"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.SteelBlue
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(6, 66)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(173, 28)
        Me.Label4.TabIndex = 101
        Me.Label4.Text = "Tipo:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'unid
        '
        Me.unid.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.unid.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.unid.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.unid.Location = New System.Drawing.Point(196, 173)
        Me.unid.Name = "unid"
        Me.unid.Size = New System.Drawing.Size(189, 28)
        Me.unid.TabIndex = 99
        Me.unid.Text = "0"
        Me.unid.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Gainsboro
        Me.GroupBox2.Controls.Add(Me.unidn)
        Me.GroupBox2.Controls.Add(Me.sigue)
        Me.GroupBox2.Controls.Add(Me.tipon)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.ForeColor = System.Drawing.Color.Black
        Me.GroupBox2.Location = New System.Drawing.Point(517, 73)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(468, 247)
        Me.GroupBox2.TabIndex = 102
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Nueva Caja"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.SteelBlue
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(6, 31)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(173, 28)
        Me.Label8.TabIndex = 97
        Me.Label8.Text = "Tipo:"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.SteelBlue
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.White
        Me.Label9.Location = New System.Drawing.Point(6, 67)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(173, 28)
        Me.Label9.TabIndex = 98
        Me.Label9.Text = "Unidades:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(24, 664)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(385, 18)
        Me.Label6.TabIndex = 105
        Me.Label6.Text = "Double click elimina línea."
        '
        'Cajas02_c
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1007, 696)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Cancela)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.corte)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.fg)
        Me.Controls.Add(Me.si)
        Me.Controls.Add(Me.graba)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Cajas02_c"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Segundas a Primeras"
        CType(Me.fg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub Cortes_status(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AddHandler corte.KeyPress, AddressOf keypressed1
        AddHandler cajad.KeyPress, AddressOf keypressed2
        AddHandler tipod.KeyPress, AddressOf keypressed2_1
        AddHandler tallad.KeyPress, AddressOf keypressed3
        AddHandler tipon.KeyPress, AddressOf keypressed4
        AddHandler unidn.KeyPress, AddressOf keypressed5
        tipod.Items.Clear()
        tipon.Items.Add("Primeras")
        tipon.SelectedIndex = 0
        limpia_variables()
    End Sub
    Private Sub limpia_variables()
        setea_fg()
        corte.Enabled = True
        cajad.Items.Clear()
        tallad.Items.Clear()
        unid.Text = "0"
        unidn.Text = "0"
        orden.Text = "0"
        fechad.Text = ""
        si.Visible = True
        graba.Visible = False
        cajad.Enabled = False
        tallad.Enabled = False
        unid.Enabled = False
        tipon.Enabled = False
        unidn.Enabled = False
        corte.Focus()
    End Sub
    Private Sub habilita()
        si.Visible = False
        graba.Visible = True
        corte.Enabled = False
        cajad.Enabled = True
        tallad.Enabled = True
        unid.Enabled = True
        tipon.Enabled = True
        unidn.Enabled = True
        Try
            tipon.SelectedIndex = 0
        Catch ex As Exception

        End Try
        cajad.Focus()
    End Sub
    Private Sub setea_fg()
        Dim sec As String = obj.seccion
        fg.Rows.Count = 1
        fg.Rows.Fixed = 1
        fg.Rows(0).Height = 30
        llena_combos(corte, "SELECT DISTINCT CORTE FROM CAJAS01 WHERE ESTADO IN ('A','P') AND TIPO_SEG > 0  ORDER BY CORTE", "CORTE")
    End Sub
    Private Sub si_Click(sender As System.Object, e As System.EventArgs) Handles si.Click
        llena_corte()
        seccion = get_seccion(corte.Text)
    End Sub

    Private Sub llena_corte()
        llena_tablas(co, "SELECT * FROM CAJAS01 WHERE CORTE = '" & corte.Text & "'", cnn)
        llena_combos(cajad, "SELECT DISTINCT CAJA FROM CAJAS01 WHERE CORTE = '" & corte.Text & "' AND UNIDADES > 0 AND TIPO_SEG > 0 ORDER BY CAJA", "CAJA")
        If co.Rows.Count > 0 Then
            dr = co.Rows(0)
            cliente = dr("CLIENTE")
            escala = dr("ESCALA")
            tallap = dr("ORDEN")
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

    Private Sub keypressed2(ByVal o As [Object], ByVal e As KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            If Trim(cajad.Text) <> "" Then
                tipod.Focus()
            End If
        End If
    End Sub

    Private Sub keypressed2_1(ByVal o As [Object], ByVal e As KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            If Trim(tipod.Text) <> "" Then
                tallad.Focus()
            End If
        End If
    End Sub
    Private Sub keypressed3(ByVal o As [Object], ByVal e As KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            unid.Text = saldo_caja(co, cajad.Text, tipod.Text, tallad.Text, orden.Text, fechad.Text)
            tipon.Focus()
        End If
    End Sub 'keypressed

    Private Sub keypressed4(ByVal o As [Object], ByVal e As KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            If Trim(tipon.Text) <> "" Then
                unidn.Focus()
            End If
        End If
    End Sub 'keypressed

    Private Sub keypressed5(ByVal o As [Object], ByVal e As KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            If Trim(unidn.Text) <> "" Then
                sigue.Focus()
            End If
        End If
    End Sub 'keypressed


    Private Sub graba_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles graba.Click
        Dim ok As Boolean
        Dim seguro As MsgBoxResult
        seguro = MsgBox("Seguro de Actualizar todos los Cambios Efectuados?  ", MsgBoxStyle.YesNo, "Actualizacion de Datos !!!")
        If seguro = MsgBoxResult.Yes Then
            If fg.Rows.Count > 1 Then
                ok = revisa_datos(fg, 10)
                If ok Then
                    graba_datos()
                    imprime_etiquetas()
                    Close()
                Else
                    Close()
                End If
            Else
                MsgBox("Aún no ha ingresado unidades a la nueva Caja", MsgBoxStyle.Critical, "Por favor revise !!!")
            End If
        End If
    End Sub
    Private Sub imprime_etiquetas()
        Dim ok As Boolean
        Dim i As Integer
        Dim etique As String
        Dim pr As New C1cajas.prt
        For i = 1 To caja_nueva.Length Step 9
            Try
                etique = Mid(caja_nueva, i, 9)
                ok = pr.imprime_cajas_s(etique, etique, obj.seccion, obj.numero, obj.constr)
            Catch
            End Try
        Next
    End Sub

    Private Sub corteS_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles corte.KeyPress
        AutoCompletar(corte, e)
    End Sub
    Private Sub cajad_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cajad.KeyPress
        AutoCompletar(cajad, e)
    End Sub
    Private Sub tipod_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tipod.KeyPress
        AutoCompletar(tipod, e)
    End Sub
    Private Sub tipon_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tipon.KeyPress
        AutoCompletar(tipon, e)
    End Sub
    Private Sub talla_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tallad.KeyPress
        AutoCompletar(tallad, e)
    End Sub
    Private Sub cajad_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cajad.SelectedIndexChanged
        llena_combos(tallad, "SELECT DISTINCT TALLA FROM CAJAS01 WHERE CAJA = '" & cajad.Text & "' ORDER BY TALLA", "TALLA")
        llena_combos(tipod, "SELECT DISTINCT TIPO FROM CAJAS01 WHERE CAJA = '" & cajad.Text & "' ORDER BY TIPO", "TIPO")
    End Sub

    Private Sub tallad_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles tallad.SelectedIndexChanged
        unid.Text = saldo_caja(co, cajad.Text, tipod.Text, tallad.Text, orden.Text, fechad.Text)
    End Sub
    Private Sub unidn_TextChanged(sender As System.Object, e As System.EventArgs) Handles unidn.TextChanged
        Dim selStart As Integer = unidn.SelectionStart
        Dim selMoveLeft As Integer = 0
        Dim newStr As String = "" 'Build a new string by copying each valid character from the existing string. The new string starts as blank and valid characters are added 1 at a time.

        For i As Integer = 0 To unidn.Text.Length - 1

            If "0123456789".IndexOf(unidn.Text(i)) <> -1 Then 'Characters that are in the allowed set will be added to the new string.
                newStr = newStr & unidn.Text(i)

            ElseIf i < selStart Then 'Characters that are not valid are removed - if these characters are before the cursor, we need to move the cursor left to account for their removal.
                selMoveLeft = selMoveLeft + 1

            End If
        Next

        unidn.Text = newStr 'Place the new text into the textbox.
        unidn.SelectionStart = selStart - selMoveLeft 'Move the cursor to the appropriate location.
    End Sub

    Private Sub sigue_Click(sender As System.Object, e As System.EventArgs) Handles sigue.Click
        Dim ok As Boolean
        If CDec(unidn.Text) = 0 Then
            MsgBox("El número de unidades no puede ser 0. ", MsgBoxStyle.Critical, "Por favor revise !!!")
            Exit Sub
        End If
        If CDec(unidn.Text) > CDec(unid.Text) Then
            MsgBox("El número Maximo de unidades a trasladar es de " & Trim(unid.Text), MsgBoxStyle.Critical, "Por favor revise !!!")
            Exit Sub
        End If
        agrega_fg(ok)
        If Not ok Then
            MsgBox("No pueden haber dos registros dentro una misma Caja con Talla y Tipo de Segunda repetidos.", MsgBoxStyle.Critical, "Por favor revise !!!")
            Exit Sub
        End If
        ok = modifica_caja(co, cajad.Text, tipod.Text, tallad.Text, unidn.Text, orden.Text, unid.Text)
        unidn.Text = 0
        cajad.Focus()
    End Sub

    Private Sub agrega_fg(ByRef ok As Boolean)
        Dim l As Integer
        Dim uni As Integer
        ok = True
        l = fg.FindRow(cajad.Text + tallad.Text + tipon.Text, 1, 0, False)
        If l > -1 Then
            ok = False
            Exit Sub
        End If
        l = fg.Rows.Count
        fg.Rows.Count = l + 1
        fg(l, 0) = cajad.Text + tallad.Text + tipon.Text
        uni = revisa_caja_talla(cajad.Text, tallad.Text, CInt(unid.Text))
        fg(l, 10) = uni

        fg(l, 1) = cajad.Text
        fg(l, 2) = tipod.Text
        fg(l, 3) = tallad.Text
        fg(l, 4) = unid.Text
        fg(l, 5) = tipon.Text
        fg(l, 6) = unidn.Text
        fg(l, 7) = orden.Text
        fg(l, 8) = fechad.Text
        fg(l, 9) = cajad.Text + tallad.Text

    End Sub
    Private Function revisa_caja_talla(ByVal caja As String, ByVal talla As String, ByVal uni As Integer) As Integer
        Dim unidades As Integer
        Dim x As Integer
        x = fg.FindRow(caja + talla, 1, 9, False)
        If x = -1 Then
            unidades = uni
        Else
            unidades = fg(x, 10)
        End If
        Return unidades
    End Function

    Private Sub tipod_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles tipod.SelectedIndexChanged
        llena_combos(tallad, "SELECT DISTINCT TALLA FROM CAJAS01 WHERE CAJA = '" & cajad.Text & "' AND TIPO = '" & tipod.Text & "' ORDER BY TALLA", "TALLA")
    End Sub

    Private Sub Cancela_Click(sender As System.Object, e As System.EventArgs) Handles Cancela.Click
        limpia_variables()
        corte.Focus()
    End Sub

    Private Sub fg_Click(sender As System.Object, e As System.EventArgs) Handles fg.DoubleClick
        elimina_linea()
    End Sub
    Private Sub elimina_linea()
        Dim ok As Boolean
        Dim l As Integer = fg.RowSel
        'Try
        ok = modifica_caja(co, fg(l, 1), fg(l, 2), fg(l, 3), fg(l, 6) * -1, orden.Text, unid.Text)
        fg.Rows.Remove(l)
        unidn.Text = 0
        cajad.Focus()
        'Catch
        'End Try
    End Sub
    '============================= Actualiza la Base de Datos =============================
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
        Dim tipo As String
        Dim busca As String
        Dim terceras As String = "38"
        Dim ts As Integer
        Dim fechac As String
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
            caja_nueva = pre + Format(corre + 1, h)

            strsql = "UPDATE CAJAS02 SET CORRELATIVO = CORRELATIVO + 1"
            comando.CommandText = strsql
            afectados = comando.ExecuteNonQuery()

            For i = 1 To fg.Rows.Count - 1
                fechac = Format(fg(i, 8), "yyyy-MM-dd HH:mm:ss")
                ts = 0
                busca = ts
                tipo = "P"
                strsql = "UPDATE CAJAS01 SET UNIDADES = UNIDADES + '" & fg(i, 6) & "' WHERE CAJA = '" & pre + Format(corre + 1, h) & "' AND CORTE = '" & corte.Text & "' AND TALLA = '" & fg(i, 3) & "' AND TIPO = '" & tipo & "'"
                comando.CommandText = strsql
                afectados = comando.ExecuteNonQuery()
                strsql = "UPDATE CAJAS04 SET UNIDADES = UNIDADES + '" & fg(i, 6) & "' WHERE CAJA = '" & pre + Format(corre + 1, h) & "' AND CORTE = '" & corte.Text & "' AND TALLA = '" & fg(i, 3) & "' AND TIPO = '" & tipo & "'"
                comando.CommandText = strsql
                afectados = comando.ExecuteNonQuery()
                If afectados = 0 Then
                    strsql = "INSERT INTO CAJAS01 (CAJA,CORTE,TALLA,TIPO,UNIDADES,CLIENTE,UBICACION,FECHA,ESTADO,ESCALA,ORDEN,IMPRESO,TIPO_SEG,SECCION) VALUES ('" &
                                                   pre + Format(corre + 1, h) & "','" & corte.Text & "','" &
                                                   fg(i, 3) & "','" & tipo & "','" &
                                                   fg(i, 6) & "','" & cliente & "','00','" & fg(i, 8) & "','P','" &
                                                   escala & "','" & fg(i, 7) & "','" & obj.usuario & "','" & ts & "','" & seccion & "')"
                    comando.CommandText = strsql
                    afectados = comando.ExecuteNonQuery()

                    Try
                        strsql = "INSERT INTO CAJAS04 (CAJA,CORTE,TALLA,TIPO,UNIDADES,FECHA ,QUIEN) " &
                                     "VALUES ( '" & pre + Format(corre + 1, h) & "','" & corte.Text & "','" &
                                                    fg(i, 3) & "','" & tipo & "','" &
                                                    fg(i, 6) & "','" & fg(i, 8) & "','" &
                                                    obj.usuario & "')"
                        comando.CommandText = strsql
                        afectados = afectados + comando.ExecuteNonQuery()
                    Catch
                    End Try
                End If
                strsql = "UPDATE CAJAS01 SET UNIDADES = UNIDADES - '" & fg(i, 6) & "' WHERE CAJA = '" & fg(i, 1) & "' AND CORTE = '" & corte.Text & "' AND TALLA = '" & fg(i, 3) & "' AND TIPO = '" & fg(i, 2) & "'"
                comando.CommandText = strsql
                afectados = comando.ExecuteNonQuery()

                strsql = "UPDATE CAJAS04 SET UNIDADES = UNIDADES - '" & fg(i, 6) & "' WHERE CAJA = '" & fg(i, 1) & "' AND CORTE = '" & corte.Text & "' AND TALLA = '" & fg(i, 3) & "' AND TIPO = '" & fg(i, 2) & "'"
                comando.CommandText = strsql
                afectados = comando.ExecuteNonQuery()

                strsql = "INSERT INTO CAJAS03 (TIPO,CAJA_D,CORTE_D,TALLA_D,TIPO_D,CAJA_A,CORTE_A,TALLA_A,TIPO_A,TIPO_SEG,UNIDADES,FECHA,QUIEN,ORDEN) VALUES ('" & _
                                              "S','" & fg(i, 1) & "','" & corte.Text & "','" & fg(i, 3) & "','" & fg(i, 2) & "','" & pre + Format(corre + 1, h) & "','" & corte.Text & "','" & fg(i, 3) & "','" & tipo & "','" & ts & "','" & _
                                               fg(i, 6) & "',GETDATE(),'" & _
                                              obj.usuario & "','" & fg(i, 7) & "')"
                comando.CommandText = strsql
                afectados = comando.ExecuteNonQuery()

                If InStr(fg(i, 1), caja_nueva) = 0 Then
                    caja_nueva = caja_nueva + fg(i, 1)
                End If

            Next
            transaccion.Commit()
            MsgBox("Grabacion Exitosa", MsgBoxStyle.Exclamation, "Grabaciones")
        Catch e As Exception
            Try
                MsgBox("Inconsistencia en Datos" + vbLf + e.Message, MsgBoxStyle.Critical, "Por favor avise a Jesús Acosta !!!!")
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

    Private Function talla() As ComboBox
        Throw New NotImplementedException
    End Function

End Class

