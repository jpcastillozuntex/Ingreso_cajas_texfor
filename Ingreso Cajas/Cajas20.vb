Imports C1.Win.C1FlexGrid

Public Class Cajas20
    Inherits System.Windows.Forms.Form
    Dim cnn As New SqlClient.SqlConnection()
    Dim dt As New DataTable
    Dim co As New DataTable
    Dim pp As New DataTable
    Friend WithEvents fg As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents corte As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Cancela As System.Windows.Forms.Button
    Dim dr As DataRow
    Dim caja_nueva As String
    Dim cliente As String
    Dim escala As String
    Dim col As Integer
    Dim tp As New DataTable
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cpo As System.Windows.Forms.Label
    Friend WithEvents estilo As System.Windows.Forms.Label
    Friend WithEvents colo As System.Windows.Forms.Label
    Friend WithEvents t_uni As System.Windows.Forms.Label
    Dim dj As DataRow
    Dim ta As String = "|XS|S|M|L|XL|XL2|XL3|XL4|XL5|XL6"
    Friend WithEvents seccion As System.Windows.Forms.ComboBox
    Friend WithEvents UPC As System.Windows.Forms.Label
    Friend WithEvents talla As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents mixto As System.Windows.Forms.CheckBox
    Friend WithEvents Automatica As System.Windows.Forms.CheckBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents up As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents codigo As System.Windows.Forms.TextBox
    Friend WithEvents ucaja As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents imprime As System.Windows.Forms.Label
    Friend WithEvents si As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Dim ts(10) As String
    Dim obj As New empresas
    Dim user_sec As String
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Cajas20))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.corte = New System.Windows.Forms.ComboBox()
        Me.Cancela = New System.Windows.Forms.Button()
        Me.seccion = New System.Windows.Forms.ComboBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.si = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.fg = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cpo = New System.Windows.Forms.Label()
        Me.estilo = New System.Windows.Forms.Label()
        Me.colo = New System.Windows.Forms.Label()
        Me.t_uni = New System.Windows.Forms.Label()
        Me.UPC = New System.Windows.Forms.Label()
        Me.talla = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.mixto = New System.Windows.Forms.CheckBox()
        Me.Automatica = New System.Windows.Forms.CheckBox()
        Me.up = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.codigo = New System.Windows.Forms.TextBox()
        Me.ucaja = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.imprime = New System.Windows.Forms.Label()
        CType(Me.fg, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.up, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'corte
        '
        Me.corte.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.corte.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.corte.Items.AddRange(New Object() {"JT", "TRECENTO", "ZUNTEX"})
        Me.corte.Location = New System.Drawing.Point(230, 63)
        Me.corte.Name = "corte"
        Me.corte.Size = New System.Drawing.Size(296, 33)
        Me.corte.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.corte, "No. Corte.")
        '
        'Cancela
        '
        Me.Cancela.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.Cancela.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cancela.ForeColor = System.Drawing.Color.Black
        Me.Cancela.Image = CType(resources.GetObject("Cancela.Image"), System.Drawing.Image)
        Me.Cancela.Location = New System.Drawing.Point(1117, 24)
        Me.Cancela.Name = "Cancela"
        Me.Cancela.Size = New System.Drawing.Size(68, 51)
        Me.Cancela.TabIndex = 104
        Me.Cancela.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.Cancela, "Otro Corte")
        Me.Cancela.UseVisualStyleBackColor = False
        '
        'seccion
        '
        Me.seccion.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.seccion.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.seccion.Items.AddRange(New Object() {"JT", "TRECENTO", "ZUNTEX"})
        Me.seccion.Location = New System.Drawing.Point(231, 15)
        Me.seccion.Name = "seccion"
        Me.seccion.Size = New System.Drawing.Size(296, 33)
        Me.seccion.TabIndex = 117
        Me.ToolTip1.SetToolTip(Me.seccion, "No. Sección.")
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.Button2.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Location = New System.Drawing.Point(1041, 761)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(69, 51)
        Me.Button2.TabIndex = 128
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.Button2, "Imprime la Caja")
        Me.Button2.UseVisualStyleBackColor = False
        Me.Button2.Visible = False
        '
        'si
        '
        Me.si.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.si.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.si.ForeColor = System.Drawing.Color.Black
        Me.si.Image = CType(resources.GetObject("si.Image"), System.Drawing.Image)
        Me.si.Location = New System.Drawing.Point(1041, 24)
        Me.si.Name = "si"
        Me.si.Size = New System.Drawing.Size(69, 51)
        Me.si.TabIndex = 1
        Me.si.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.si, "Elige corte.")
        Me.si.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.Button1.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Location = New System.Drawing.Point(1117, 761)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(68, 51)
        Me.Button1.TabIndex = 135
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.Button1, "Limpia el contenido de la Caja")
        Me.Button1.UseVisualStyleBackColor = False
        '
        'fg
        '
        Me.fg.AllowEditing = False
        Me.fg.AllowFiltering = True
        Me.fg.ColumnInfo = resources.GetString("fg.ColumnInfo")
        Me.fg.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.fg.HighLight = C1.Win.C1FlexGrid.HighLightEnum.Never
        Me.fg.Location = New System.Drawing.Point(14, 260)
        Me.fg.Name = "fg"
        Me.fg.Rows.DefaultSize = 21
        Me.fg.Size = New System.Drawing.Size(1197, 123)
        Me.fg.StyleInfo = resources.GetString("fg.StyleInfo")
        Me.fg.TabIndex = 92
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.SteelBlue
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(14, 61)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(197, 40)
        Me.Label3.TabIndex = 94
        Me.Label3.Text = "Corte:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.SteelBlue
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(14, 111)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(197, 41)
        Me.Label1.TabIndex = 107
        Me.Label1.Text = "CPO:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.SteelBlue
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(14, 158)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(197, 41)
        Me.Label2.TabIndex = 108
        Me.Label2.Text = "Estilo:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.SteelBlue
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(14, 205)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(197, 41)
        Me.Label4.TabIndex = 109
        Me.Label4.Text = "Color:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.SteelBlue
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(14, 11)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(197, 41)
        Me.Label5.TabIndex = 110
        Me.Label5.Text = "Seccion:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cpo
        '
        Me.cpo.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.cpo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.cpo.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.cpo.Location = New System.Drawing.Point(230, 111)
        Me.cpo.Name = "cpo"
        Me.cpo.Size = New System.Drawing.Size(295, 41)
        Me.cpo.TabIndex = 111
        Me.cpo.Text = " "
        Me.cpo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'estilo
        '
        Me.estilo.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.estilo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.estilo.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.estilo.Location = New System.Drawing.Point(231, 158)
        Me.estilo.Name = "estilo"
        Me.estilo.Size = New System.Drawing.Size(295, 41)
        Me.estilo.TabIndex = 112
        Me.estilo.Text = " "
        Me.estilo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'colo
        '
        Me.colo.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.colo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.colo.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.colo.Location = New System.Drawing.Point(231, 205)
        Me.colo.Name = "colo"
        Me.colo.Size = New System.Drawing.Size(295, 41)
        Me.colo.TabIndex = 113
        Me.colo.Text = " "
        Me.colo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        't_uni
        '
        Me.t_uni.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.t_uni.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.t_uni.Font = New System.Drawing.Font("Microsoft Sans Serif", 120.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.t_uni.ForeColor = System.Drawing.Color.Black
        Me.t_uni.Location = New System.Drawing.Point(535, 398)
        Me.t_uni.Name = "t_uni"
        Me.t_uni.Size = New System.Drawing.Size(676, 334)
        Me.t_uni.TabIndex = 115
        Me.t_uni.Text = "0"
        Me.t_uni.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'UPC
        '
        Me.UPC.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.UPC.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.UPC.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.UPC.Location = New System.Drawing.Point(231, 498)
        Me.UPC.Name = "UPC"
        Me.UPC.Size = New System.Drawing.Size(211, 40)
        Me.UPC.TabIndex = 123
        Me.UPC.Text = " "
        Me.UPC.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'talla
        '
        Me.talla.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.talla.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.talla.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.talla.Location = New System.Drawing.Point(230, 451)
        Me.talla.Name = "talla"
        Me.talla.Size = New System.Drawing.Size(211, 40)
        Me.talla.TabIndex = 122
        Me.talla.Text = " "
        Me.talla.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.SteelBlue
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.White
        Me.Label10.Location = New System.Drawing.Point(14, 498)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(197, 40)
        Me.Label10.TabIndex = 121
        Me.Label10.Text = "UPC"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.SteelBlue
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.White
        Me.Label11.Location = New System.Drawing.Point(14, 451)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(197, 40)
        Me.Label11.TabIndex = 120
        Me.Label11.Text = "Talla:"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.SteelBlue
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(9, 685)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(198, 41)
        Me.Label8.TabIndex = 125
        Me.Label8.Text = "Impresión Auto:"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.SteelBlue
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.White
        Me.Label9.Location = New System.Drawing.Point(9, 638)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(198, 41)
        Me.Label9.TabIndex = 124
        Me.Label9.Text = "Tallas Mixtas:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'mixto
        '
        Me.mixto.AutoSize = True
        Me.mixto.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.mixto.Location = New System.Drawing.Point(251, 649)
        Me.mixto.Name = "mixto"
        Me.mixto.Size = New System.Drawing.Size(18, 17)
        Me.mixto.TabIndex = 126
        Me.mixto.UseVisualStyleBackColor = True
        '
        'Automatica
        '
        Me.Automatica.AutoSize = True
        Me.Automatica.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Automatica.Location = New System.Drawing.Point(251, 698)
        Me.Automatica.Name = "Automatica"
        Me.Automatica.Size = New System.Drawing.Size(18, 17)
        Me.Automatica.TabIndex = 127
        Me.Automatica.UseVisualStyleBackColor = True
        '
        'up
        '
        Me.up.AllowEditing = False
        Me.up.AllowFiltering = True
        Me.up.ColumnInfo = resources.GetString("up.ColumnInfo")
        Me.up.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.up.HighLight = C1.Win.C1FlexGrid.HighLightEnum.Never
        Me.up.Location = New System.Drawing.Point(3, 597)
        Me.up.Name = "up"
        Me.up.Rows.Count = 4
        Me.up.Rows.DefaultSize = 21
        Me.up.Size = New System.Drawing.Size(1198, 122)
        Me.up.StyleInfo = resources.GetString("up.StyleInfo")
        Me.up.TabIndex = 129
        Me.up.Visible = False
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.SteelBlue
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(14, 402)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(197, 40)
        Me.Label6.TabIndex = 130
        Me.Label6.Text = "Codigo:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'codigo
        '
        Me.codigo.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.codigo.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.codigo.Location = New System.Drawing.Point(230, 404)
        Me.codigo.Name = "codigo"
        Me.codigo.Size = New System.Drawing.Size(211, 34)
        Me.codigo.TabIndex = 131
        '
        'ucaja
        '
        Me.ucaja.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ucaja.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.ucaja.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.ucaja.Location = New System.Drawing.Point(231, 546)
        Me.ucaja.Name = "ucaja"
        Me.ucaja.Size = New System.Drawing.Size(211, 40)
        Me.ucaja.TabIndex = 133
        Me.ucaja.Text = " "
        Me.ucaja.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.SteelBlue
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.White
        Me.Label12.Location = New System.Drawing.Point(14, 546)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(197, 40)
        Me.Label12.TabIndex = 132
        Me.Label12.Text = "U.x Caja:"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'imprime
        '
        Me.imprime.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.imprime.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.imprime.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.imprime.ForeColor = System.Drawing.Color.Black
        Me.imprime.Location = New System.Drawing.Point(535, 398)
        Me.imprime.Name = "imprime"
        Me.imprime.Size = New System.Drawing.Size(650, 334)
        Me.imprime.TabIndex = 134
        Me.imprime.Text = "Imprimiendo Etiqueta."
        Me.imprime.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cajas20
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 19)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1097, 653)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ucaja)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.codigo)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.t_uni)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Automatica)
        Me.Controls.Add(Me.mixto)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.UPC)
        Me.Controls.Add(Me.talla)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.seccion)
        Me.Controls.Add(Me.colo)
        Me.Controls.Add(Me.estilo)
        Me.Controls.Add(Me.cpo)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Cancela)
        Me.Controls.Add(Me.si)
        Me.Controls.Add(Me.corte)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.fg)
        Me.Controls.Add(Me.imprime)
        Me.Controls.Add(Me.up)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Cajas20"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Ingreso de Unidades"
        CType(Me.fg, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.up, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    Private Sub Cajas05_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        AddHandler seccion.KeyPress, AddressOf keypressed1
        AddHandler corte.KeyPress, AddressOf keypressed2
        AddHandler codigo.KeyPress, AddressOf keypressed3
        Dim sec As String = obj.seccion
        If Trim(obj.seccion) <> "" Then
            llena_combos(seccion, "SELECT  SECCION FROM SECCIONES WHERE SECCION = '" & sec & "'  AND ACTIVA = 'S'", "SECCION")
        Else
            llena_combos(seccion, "SELECT  SECCION FROM SECCIONES WHERE SECCION LIKE 'TEX%' AND ACTIVA = 'S'", "SECCION")
        End If
        ts = ta.Split("|")
        limpia_variables()
    End Sub
    Private Sub limpia_variables()
        setea_grid()
        seccion.Enabled = True
        corte.Enabled = True
        cpo.Text = ""
        estilo.Text = ""
        colo.Text = ""
        talla.Text = ""
        UPC.Text = ""
        codigo.Text = ""
        ucaja.Text = ""
        t_uni.Text = "0"
        mixto.Checked = False
        Automatica.Checked = True
        UPC.Enabled = False
        mixto.Enabled = False
        codigo.Enabled = False
        si.Visible = True
        seccion.Focus()
    End Sub
    Private Sub habilita()
        seccion.Enabled = False
        corte.Enabled = False
        UPC.Enabled = True
        mixto.Enabled = True
        codigo.Enabled = True
        si.Visible = False
    End Sub
    Private Sub setea_grid()
        fg.Rows.Count = 1
        fg.Rows.Fixed = 1
        fg.Rows.Count = 3
        fg.Rows(0).Height = 30
        fg.Rows(1).Height = 30
        up.Rows.Count = 1
        up.Rows.Fixed = 1
        up.Rows.Count = 4
        limpia_Grids()
    End Sub
    Private Sub si_Click(sender As System.Object, e As System.EventArgs) Handles si.Click
        llena_corte()
        limpia_codigo()
    End Sub

    Private Sub llena_corte()
        Dim ok As Boolean
        Dim i As Integer
        llena_tablas(co, "SELECT CORTES.*,CPO_D.ESCALA, E_TALLAS.* FROM CORTES LEFT JOIN CPO_D ON CPO_D.CPO = CORTES.CPO AND CPO_D.ESTILO = CORTES.ESTILO AND CPO_D.COLOR = CORTES.COLOR LEFT JOIN E_TALLAS ON CPO_D.ESCALA = E_TALLAS.ESCALA WHERE CORTE = '" & corte.Text & "'", cnn)
        If co.Rows.Count > 0 Then
            dr = co.Rows(0)
            dj = dr
            cpo.Text = dr("CPO")
            estilo.Text = dr("ESTILO")
            colo.Text = dr("COLOR")
            seccion.Text = dr("SECCION")
            cliente = dr("CLIENTE")
            escala = dr("ESCALA")
            fg(1, 1) = "Cortado"
        End If
        For i = 1 To 10
            'If escala = "00" Then
            '    dj(i + 24) = ts(i)
            'End If
            fg(0, i + 1) = dj(i + 24)
            fg(1, i + 1) = dj(i + 5)
            up(0, i) = dj(i + 24)
        Next
        llena_empaque(estilo.Text, ok)
        If Not ok Then
            limpia_variables()
        End If
        producido()
        busca_upc(ok)
        If ok Then
            habilita()
        Else
            limpia_variables()
        End If
    End Sub
    Private Sub producido()
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim strsql As String = "SELECT ORDEN, SUM(CAJAS01.UNIDADES) AS PROD FROM CAJAS01 LEFT JOIN CAJAS04 ON CAJAS01.CAJA = CAJAS04.CAJA AND CAJAS01.TALLA = CAJAS04.TALLA AND CAJAS01.TIPO = CAJAS04.TIPO WHERE CAJAS01.CORTE = '" & corte.Text & "' GROUP BY ORDEN"
        llena_tablas(dt, strsql, cnn)
        For Each dr In dt.Rows
            fg(2, CInt(dr("ORDEN")) + 2) = dr("PROD")
        Next
    End Sub
    Private Sub busca_upc(ByRef ok As Boolean)
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim i As Integer
        ok = False
        llena_tablas(dt, "SELECT * FROM UPC WHERE CLIENTE = '" & cliente & "' AND ESTILO = '" & estilo.Text & "' AND COLOR = '" & colo.Text & "'", cnn)
        If dt.Rows.Count = 0 Then
            MsgBox("Aún no existen registrados UPC para este CORTE !!!!", MsgBoxStyle.Critical, "Por favor revise.")
            Exit Sub
        Else
            dr = dt.Rows(0)
        End If
        For i = 1 To 10
            up(1, i) = Trim(dr(i + 4))
        Next
        ok = True
    End Sub
    Private Sub limpia_codigo()
        codigo.Text = ""
        codigo.Focus()
    End Sub
    Private Sub llena_empaque(ByRef estilo As String, ByRef ok As Boolean)
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim strsql As String = "SELECT * FROM ESTILO_EMP WHERE ESTILO = '" & estilo & "'"
        Dim i As Integer
        llena_tablas(dt, strsql, cnn)
        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            For i = 1 To 10
                Try
                    up(2, i) = dr(i)
                Catch
                End Try
            Next
            ok = True
        Else
            MsgBox("Aún no ha ingresado el Número de unidades por Caja.", MsgBoxStyle.Critical, "Por favor Revise !!!!")
            ok = False
        End If
    End Sub
    '================================== HANDLERS ================================
    Private Sub keypressed1(ByVal o As [Object], ByVal e As KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            If Trim(seccion.Text) <> "" Then
                corte.Focus()
            End If
        End If
    End Sub 'keypressed
    '================================== HANDLERS ================================
    Private Sub keypressed2(ByVal o As [Object], ByVal e As KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            If Trim(corte.Text) <> "" Then
                si.Focus()
            End If
        End If
    End Sub 'keypressed

    '================================== HANDLERS ================================
    Private Sub keypressed3(ByVal o As [Object], ByVal e As KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            If Trim(codigo.Text) <> "" Then
                Try
                    My.Computer.Audio.Play("c:\scan\beep.wav")
                Catch
                End Try
                procesa()
            End If
        End If
    End Sub 'keypressed

    Private Sub procesa()
        If CDec(t_uni.Text) = 0 Then
            llena_talla()
            If ucaja.Text = "0" Then
                MsgBox("Upc de Prenda y UPC registerado no coinciden." & codigo.Text, MsgBoxStyle.Critical, "Por favor revisar !!!!")
                limpia_codigo()
                corte.Focus()
                Exit Sub
            End If
            UPC.Text = codigo.Text
            mixto.Enabled = False
            Automatica.Enabled = False
        Else
            If codigo.Text <> UPC.Text And mixto.Checked = False Then
                MsgBox("No puede ingresar distntas tallas en la misma caja.", MsgBoxStyle.Critical, "Por favor revise.")
                limpia_codigo()
                Exit Sub
            End If
        End If
        suma_unidades()
    End Sub
    Private Sub suma_unidades()
        llena_talla()
        up(3, col) = up(3, col) + 1
        t_uni.Text = t_uni.Text + 1
        If Automatica.Checked = True Then
            If CInt(t_uni.Text) = CInt(ucaja.Text) Then
                imprime_Etiqueta()
            End If
        End If
        limpia_codigo()
    End Sub
    Private Sub llena_talla()
        Dim i As Integer
        For i = 1 To 10
            If up(1, i) = codigo.Text Then
                talla.Text = up(0, i)
                ucaja.Text = up(2, i)
                col = i
                Exit For
            Else
                ucaja.Text = "0"
            End If
        Next
     
    End Sub
    Private Sub imprime_Etiqueta()
        Dim ok As Boolean
        Dim caja As String = ""
        Dim pr As New C1cajas.prt
        verifica_Tallas(ok)
        If ok Then
            t_uni.Visible = False
            imprime.Visible = True
            graba_datos(caja, ok)
            If ok Then
                ok = pr.imprime_cajas_s(caja, caja, obj.seccion, obj.numero, obj.constr)
                otra_caja()
            End If
        End If

    End Sub
    Private Sub verifica_Tallas(ByRef ok As Boolean)
        Dim i As Integer
        ok = False
        For i = 1 To 10
            If (up(3, i) + fg(2, i + 1)) > fg(1, i + 1) Then
                MsgBox("Está tratando de ingresar " + CStr(up(3, i)) + " unidades en la talla " + CStr(fg(0, i + 1)) + " y solo se cortaron " + CStr(fg(1, i + 1)), MsgBoxStyle.Critical, "Por favor revise y corrija.")
                Exit Sub
            End If
        Next
        ok = True
    End Sub
    Private Sub otra_caja()
        Dim i As Integer
        t_uni.Visible = True
        t_uni.Text = "0"
        talla.Text = ""
        UPC.Text = ""
        ucaja.Text = ""
        mixto.Enabled = True
        Automatica.Enabled = True
        mixto.Checked = False
        Automatica.Checked = True
        For i = 1 To 10
            up(3, i) = 0
        Next
        llena_corte()
        limpia_codigo()
    End Sub
    Private Sub seccion_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles seccion.KeyPress
        AutoCompletar(seccion, e)
    End Sub
    Private Sub seccion_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles seccion.KeyDown
        If e.KeyCode = Keys.Delete Then
            e.Handled = True
        End If
    End Sub
    Private Sub corteS_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles corte.KeyPress
        AutoCompletar(corte, e)
    End Sub

    Private Sub cortes_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles corte.KeyDown
        If e.KeyCode = Keys.Delete Then
            e.Handled = True
        End If
    End Sub

    Private Sub Cancela_Click(sender As System.Object, e As System.EventArgs) Handles Cancela.Click
        limpia_variables()
        corte.Focus()
    End Sub
    '============================= Actualiza la Base de Datos =============================
    Private Sub graba_datos(ByRef caja As String, ByRef ok As Boolean)
        Dim strsql As String
        Dim afectados As Integer
        Dim corre As Integer
        Dim pre As String = ""
        Dim i As Integer
        Dim obj As New empresas
        Dim h As String = "0000000"
        If obj.numero = "1" Then
            pre = "JT"
        ElseIf obj.numero = "3" Then
            pre = "ZU"
        End If
        Dim j As Object
        ok = False
        caja = ""
        cnn.Open()
        ' Comienza la transacción
        Dim transaccion As SqlClient.SqlTransaction = cnn.BeginTransaction()

        ' Crea el comando para la transacción
        Dim comando As SqlClient.SqlCommand = cnn.CreateCommand()
        comando.Transaction = transaccion

        Try

            strsql = "UPDATE CAJAS02 SET CORRELATIVO = CORRELATIVO + 1"
            comando.CommandText = strsql
            afectados = comando.ExecuteNonQuery()

            strsql = "SELECT CORRELATIVO FROM CAJAS02"
            comando.CommandText = strsql
            j = comando.ExecuteScalar()
            corre = j.ToString
            caja = pre + Format(corre, h)

            For i = 1 To 10
                If up(3, i) > 0 Then
                    strsql = "INSERT INTO CAJAS01 (CAJA,CORTE,TALLA,TIPO,UNIDADES,CLIENTE,UBICACION,FECHA,ESTADO,ESCALA,ORDEN,IMPRESO,TIPO_SEG,SECCION) VALUES ('" & _
                                                   caja & "','" & corte.Text & "','" & up(0, i) & "','P','" & up(3, i) & "','" & cliente & "','00',GETDATE(),'P','" & escala & "','" & CStr(i - 1) & "','" & _
                                                   obj.usuario & "','0','" & seccion.Text & "')"

                    comando.CommandText = strsql
                    afectados = comando.ExecuteNonQuery()

                    strsql = "INSERT INTO CAJAS04 (CAJA,CORTE,TALLA,TIPO,UNIDADES,FECHA,QUIEN) VALUES ('" & _
                                                 caja & "','" & corte.Text & "','" & up(0, i) & "','P','" & up(3, i) & "',GETDATE(),'" & obj.usuario & "')"

                    comando.CommandText = strsql
                    afectados = comando.ExecuteNonQuery()

                End If

            Next

            transaccion.Commit()

            ok = True
        Catch e As Exception
            Try
                transaccion.Rollback()
            Catch ex As SqlClient.SqlException
              
            End Try
        Finally
            cnn.Close()
        End Try
        If Not ok Then
            MsgBox("Inconsistencia en Datos", MsgBoxStyle.Critical, "Por favor revise !!!!")
        End If
    End Sub

    Private Sub seccion_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles seccion.SelectedIndexChanged
        Dim strsql As String = "SELECT CORTE FROM CORTES WHERE CORTE NOT IN (SELECT DISTINCT CORTE FROM PROD_DIARIA) AND SECCION = '" & seccion.Text & "' AND TOTAL > 0 AND EXPORTADO <> 'S'"
        llena_combos(corte, strsql, "CORTE")
    End Sub

    Private Sub Automatica_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles Automatica.CheckedChanged
        If Automatica.Checked Then
            Button2.Visible = False
        Else
            Button2.Visible = True
        End If
    End Sub

    Private Sub mixto_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles mixto.CheckedChanged
        Automatica.Checked = False
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        otra_caja()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        imprime_Etiqueta()
    End Sub

    Private Sub limpia_Grids()
        Dim i As Integer
        For i = 0 To 10
            Try
                up(0, i) = ""
                fg(0, i + 1) = ""
            Catch
            End Try
        Next
        fg(1, 1) = "Cortado"
        fg(2, 1) = "Producido"
    End Sub
End Class

