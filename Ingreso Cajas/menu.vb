Imports System.IO
Imports System.Reflection
Public Class menu
    Inherits System.Windows.Forms.Form
    Private HasConnected As Boolean = False
    Dim empresa As String
    Dim nombre As String
    Dim dt As DataTable
    Dim cnn As New SqlClient.SqlConnection()
    Dim contador As Integer
    Dim bien As Boolean
    Dim tipo As String
    Dim usua As String = "TEVOC"
    Public dia_hoy As Date = Today
    Dim retval
    Dim menus As String
    Dim empres As String
    Dim men As New System.Windows.Forms.MenuItem()
    Dim a As Integer = Screen.PrimaryScreen.Bounds.Height - 50
    Dim l As Integer = Screen.PrimaryScreen.Bounds.Width - 5
    Friend WithEvents foto As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Dim obj As empresas
    Dim cs As New C1cajasLib_SIF.util
    Dim usuario As String
    Dim ip As String
    Dim i1 As Integer
    Friend WithEvents menudo As MenuStrip
    Dim i2 As Integer

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
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents grupo1 As System.Windows.Forms.GroupBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Empre As System.Windows.Forms.ListBox

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(menu))
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.grupo1 = New System.Windows.Forms.GroupBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Empre = New System.Windows.Forms.ListBox()
        Me.foto = New System.Windows.Forms.PictureBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.menudo = New System.Windows.Forms.MenuStrip()
        Me.grupo1.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.foto, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.ForeColor = System.Drawing.Color.Black
        Me.TextBox1.Location = New System.Drawing.Point(152, 64)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TextBox1.Size = New System.Drawing.Size(255, 26)
        Me.TextBox1.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.TextBox1, "Por favor ingrese la palabra Clave que le asigno el Adminstrador del Sistema.   S" &
        "in esa palabra no podrá accesar .")
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(142, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(265, 32)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Por favor Ingrese su Password:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(36, Byte), Integer), CType(CType(47, Byte), Integer), CType(CType(106, Byte), Integer))
        Me.Label2.Location = New System.Drawing.Point(6, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(208, 40)
        Me.Label2.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(36, Byte), Integer), CType(CType(47, Byte), Integer), CType(CType(106, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(288, 624)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(472, 64)
        Me.Label3.TabIndex = 8
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ToolTip1.SetToolTip(Me.Label3, "Nombre del Usuario")
        '
        'grupo1
        '
        Me.grupo1.BackColor = System.Drawing.Color.FromArgb(CType(CType(164, Byte), Integer), CType(CType(197, Byte), Integer), CType(CType(223, Byte), Integer))
        Me.grupo1.Controls.Add(Me.PictureBox2)
        Me.grupo1.Controls.Add(Me.Label1)
        Me.grupo1.Controls.Add(Me.TextBox1)
        Me.grupo1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.grupo1.ForeColor = System.Drawing.Color.Black
        Me.grupo1.Location = New System.Drawing.Point(288, 488)
        Me.grupo1.Name = "grupo1"
        Me.grupo1.Size = New System.Drawing.Size(472, 128)
        Me.grupo1.TabIndex = 9
        Me.grupo1.TabStop = False
        Me.grupo1.Text = "Password"
        '
        'PictureBox2
        '
        Me.PictureBox2.Location = New System.Drawing.Point(72, 32)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(64, 56)
        Me.PictureBox2.TabIndex = 3
        Me.PictureBox2.TabStop = False
        '
        'Empre
        '
        Me.Empre.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Empre.ItemHeight = 20
        Me.Empre.Items.AddRange(New Object() {"1.- JT. TRADING", "2.- TRECENTO", "3.- ZUNTEX"})
        Me.Empre.Location = New System.Drawing.Point(876, 43)
        Me.Empre.Name = "Empre"
        Me.Empre.Size = New System.Drawing.Size(154, 24)
        Me.Empre.TabIndex = 10
        Me.Empre.Visible = False
        '
        'foto
        '
        Me.foto.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.foto.Location = New System.Drawing.Point(825, 524)
        Me.foto.Name = "foto"
        Me.foto.Size = New System.Drawing.Size(128, 123)
        Me.foto.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.foto.TabIndex = 12
        Me.foto.TabStop = False
        Me.foto.Visible = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.PictureBox3.ErrorImage = CType(resources.GetObject("PictureBox3.ErrorImage"), System.Drawing.Image)
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.InitialImage = CType(resources.GetObject("PictureBox3.InitialImage"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(433, 84)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(405, 366)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 13
        Me.PictureBox3.TabStop = False
        Me.PictureBox3.Visible = False
        '
        'menudo
        '
        Me.menudo.AutoSize = False
        Me.menudo.BackColor = System.Drawing.Color.DarkCyan
        Me.menudo.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.menudo.Location = New System.Drawing.Point(0, 0)
        Me.menudo.Name = "menudo"
        Me.menudo.Size = New System.Drawing.Size(1362, 40)
        Me.menudo.Stretch = False
        Me.menudo.TabIndex = 14
        '
        'menu
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1362, 681)
        Me.Controls.Add(Me.menudo)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.foto)
        Me.Controls.Add(Me.grupo1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Empre)
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "menu"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.grupo1.ResumeLayout(False)
        Me.grupo1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.foto, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub Menu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AddHandler TextBox1.KeyPress, AddressOf keypressed1
        posiciona_objetos()
        Empre.SelectedIndex = 0
        setea_empresa()
        TextBox1.Focus()
    End Sub
    Private Sub posiciona_objetos()
        Dim r As Point
        r.X = l - 180
        r.Y = Empre.Location.Y
        Empre.Location = r
        r.X = CInt((l / 2) - 170)
        r.Y = CInt((a / 2) - 182)
        PictureBox3.Location = r
        r.X = CInt((l / 2) - 220)
        r.Y = CInt(a - 200)
        grupo1.Location = r
        r.X = CInt((l / 2) - 190)
        r.Y = a - 150
        Label3.Location = r
        r.X = l - 150
        r.Y = 70
        r.X = CInt((l / 2) - 220)
        r.Y = CInt(a - 200)
        foto.Location = r
    End Sub

    '============================================================================
    '                           password
    '============================================================================
    Sub keypressed1(ByVal o As [Object], ByVal e As KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            e.Handled = True
            If Trim(TextBox1.Text) <> "" Then
                bien = True
                busca_clave()
                If bien Then
                    procesa()
                End If
            End If
        End If
    End Sub
    Private Sub busca_clave()
        Dim obj As New empresas()
        Dim mn As New C1cajasLib_SIF.men
        Dim existe As Boolean = True
        Dim ok As Boolean
        Dim strSQL As String = ""
        Dim ft As New DataTable
        Dim dr As DataRow
        Dim sec As String = ""
        contador = contador + 1
        If contador > 6 Then
            End
        End If
        mn.menu(TextBox1.Text, usuario, ip, sec, dt, ok)
        Try
            If Not ok Then
                existe = False
                bien = False
                Label3.Text = "Clave Incorrecta !!!!!!.      " + CStr(contador) + "  Intentos "
                TextBox1.Text = ""
            Else
                Label3.Text = usuario
                nombre = usuario
                obj.usuario = Label3.Text
                obj.clave = TextBox1.Text
                obj.seccion = sec
                foto.Visible = True
                strSQL = "SELECT * FROM MEN_JAP_F WHERE PASSWORD = '" & TextBox1.Text & "'"
                llena_tablas(ft, strSQL, cnn)
                If ft.Rows.Count > 0 Then
                    dr = ft.Rows(0)

                    Try
                        Dim fotogra() As Byte = dr("FOTO")
                        Dim stmBLOBData As New MemoryStream(fotogra)
                        foto.Image = Image.FromStream(stmBLOBData)
                    Catch
                    End Try
                End If
                bien = True
            End If
        Catch
            MsgBox("No he podido conectarme al servidor.  Por favor verifique su conección", MsgBoxStyle.Critical, "Conección Perdida.")
            Close()
        End Try
    End Sub
    Private Sub setea_empresa()
        Dim selectedIndex As Integer
        selectedIndex = Empre.SelectedIndex
        Dim selectedItem As Object
        selectedItem = Empre.SelectedItem
        If selectedIndex > -1 Then
            empresa = Mid(selectedItem.ToString(), 1, 1)
            nombre = Mid(selectedItem.ToString(), 5)
            obj = New empresas()
            obj.numero = empresa
            obj.nombre = nombre
            obj.constr = cs.con_string(empresa - 1)
            obj.conole = cs.con_ole(empresa - 1)
            Label2.Text = nombre
        End If
    End Sub
    Private Sub Empre_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Empre.SelectedIndexChanged
        setea_empresa()
    End Sub
    Private Sub procesa()
        Dim ms As New ToolStripMenuItem
        Dim mj As New ToolStripMenuItem
        Dim mb As New ToolStripMenuItem
        Dim dr As DataRow
        Dim i As Integer = -1
        Dim menu As String
        Dim descripcion As String
        Dim col As String
        Dim rojo As Integer
        Dim verde As Integer
        Dim azul As Integer
        Dim colo As Color
        For Each dr In dt.Rows
            menu = dr("MENU")
            descripcion = dr("DESCRIPCION")
            col = dr("COLOR")
            rojo = CInt(Mid(col, 1, 3))
            verde = CInt(Mid(col, 5, 3))
            azul = CInt(Mid(col, 9, 3))
            colo = Color.FromArgb(rojo, verde, azul)
            If CDec(Mid(menu, 4)) = 0 Then
                i = i + 1
                i1 = -1
                menudo.Items.Add(menu)
                menudo.Items(i).Text = descripcion
                menudo.Items(i).Name = menu
                menudo.Items(i).Visible = True
                menudo.Items(i).BackColor = colo
                ms = menudo.Items(i)
            ElseIf CDec(Mid(menu, 6)) = 0 Then
                menusg(ms, menu, descripcion, colo)
            ElseIf CDec(Mid(menu, 8)) = 0 Then
                mj = ms.DropDownItems(i1)
                menutr(mj, dr("MENU"), dr("DESCRIPCION"), colo)
            Else
                mb = mj.DropDownItems(i2)
                menubr(mb, dr("MENU"), dr("DESCRIPCION"), colo)
            End If
        Next
        TextBox1.Visible = False
        grupo1.Visible = False
        Empre.Visible = True
        PictureBox3.Visible = True
        Label3.Visible = True
    End Sub

    Private Sub menusg(ByVal ms As ToolStripMenuItem, ByVal m As String, ByVal descripcion As String, ByVal colo As Color)
        Dim i As Integer = CInt(Mid(m, 4, 2)) - 1
        Dim mj As New ToolStripMenuItem
        Dim MenuHijo As New ToolStripMenuItem(m, Nothing, New EventHandler(AddressOf MenuItem_Click), m)
        MenuHijo.Text = descripcion
        MenuHijo.BackColor = colo
        ms.DropDownItems.Add(MenuHijo)
        i1 = i1 + 1
        i2 = -1
    End Sub
    Private Sub menutr(ByVal ms As ToolStripMenuItem, ByVal m As String, ByVal descripcion As String, ByVal colo As Color)
        Dim mj As New ToolStripMenuItem
        Dim MenuHijo As New ToolStripMenuItem(m, Nothing, New EventHandler(AddressOf MenuItem_Click), m)
        MenuHijo.Text = descripcion
        MenuHijo.BackColor = colo
        ms.DropDownItems.Add(MenuHijo)
        i2 = i2 + 1
    End Sub
    Private Sub menubr(ByVal ms As ToolStripMenuItem, ByVal m As String, ByVal descripcion As String, ByVal colo As Color)
        Dim mj As New ToolStripMenuItem
        Dim MenuHijo As New ToolStripMenuItem(m, Nothing, New EventHandler(AddressOf MenuItem_Click), m)
        MenuHijo.Text = descripcion
        MenuHijo.BackColor = colo
        ms.DropDownItems.Add(MenuHijo)
    End Sub

    Private Sub MenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim prg As String = ""
        Dim ancho As Integer = 0
        Dim alto As Integer = 0
        Dim cont As Integer = 0
        prg = busca_prg(dt, sender.name, prg, ancho, alto, cont)
        If prg.Length > 1 Then
            Dim frm As Form = CreateForm(prg)
            If ancho > 0 Or alto > 0 Then
                Try
                    frm.Size = New Size(l, a)
                Catch
                End Try

            End If
            Try
                frm.Text = frm.Text + "  (" + obj.nombre + ")"
                frm.ShowDialog()
            Catch
            End Try
        End If

    End Sub
    Public Function busca_prg(ByVal db As DataTable, ByVal nombre As String, ByRef prg As String, ByRef ancho As Integer, ByRef alto As Integer, ByRef cont As String) As String
        Dim dd As DataRow()
        Dim dr As DataRow
        dd = db.Select("MENU = '" & nombre & "'")
        If dd.Length > 0 Then
            dr = dd(0)
            prg = dr("PROGRAMA")
            ancho = dr("ANCHO")
            alto = dr("ALTO")
            cont = dr("CONTROL")
        End If
        Return prg
    End Function

    Public Shared Function CreateObjectInstance(ByVal objectName As String) As Object
        Dim obj As Object
        Try
            If objectName.LastIndexOf(".") = -1 Then
                objectName = [Assembly].GetEntryAssembly.GetName.Name & "." & objectName
            End If
            obj = [Assembly].GetEntryAssembly.CreateInstance(objectName)
        Catch ex As Exception
            obj = Nothing
        End Try
        Return obj

    End Function

    Public Shared Function CreateForm(ByVal formName As String) As Form
        ' Regresa instancia de la forma
        Return DirectCast(CreateObjectInstance(formName), Form)
    End Function

End Class