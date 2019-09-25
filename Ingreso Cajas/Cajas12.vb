Imports C1.Win.C1FlexGrid
Imports System.Drawing.Printing
Imports System.IO

Public Class Cajas12
    Inherits System.Windows.Forms.Form
    Dim cnn As New SqlClient.SqlConnection()
    Dim dt As New DataTable()
    Dim dr As DataRow
    Public strsql As String
    Dim con(3) As String
    Dim pr As New DataTable
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
    Friend WithEvents fg As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Cajas12))
        Me.fg = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        CType(Me.fg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fg
        '
        Me.fg.AllowEditing = False
        Me.fg.AllowFiltering = True
        Me.fg.AllowMerging = C1.Win.C1FlexGrid.AllowMergingEnum.Free
        Me.fg.AllowSorting = C1.Win.C1FlexGrid.AllowSortingEnum.None
        Me.fg.BackColor = System.Drawing.Color.White
        Me.fg.ColumnInfo = resources.GetString("fg.ColumnInfo")
        Me.fg.FocusRect = C1.Win.C1FlexGrid.FocusRectEnum.None
        Me.fg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.fg.ForeColor = System.Drawing.Color.Black
        Me.fg.HighLight = C1.Win.C1FlexGrid.HighLightEnum.Never
        Me.fg.Location = New System.Drawing.Point(12, 21)
        Me.fg.Name = "fg"
        Me.fg.Rows.DefaultSize = 19
        Me.fg.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row
        Me.fg.Size = New System.Drawing.Size(991, 659)
        Me.fg.StyleInfo = resources.GetString("fg.StyleInfo")
        Me.fg.SubtotalPosition = C1.Win.C1FlexGrid.SubtotalPositionEnum.BelowData
        Me.fg.TabIndex = 22
        Me.ToolTip1.SetToolTip(Me.fg, "Area de Datos.")
        '
        'Cajas12
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1015, 696)
        Me.Controls.Add(Me.fg)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Red
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Cajas12"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Prendas producidas por Sección"
        Me.ToolTip1.SetToolTip(Me, " ")
        CType(Me.fg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub FPO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        setea_grid()
        llena_grid()
    End Sub

    Private Sub setea_grid()
        fg.Rows.Count = 1
        fg.Rows(0).Height = 30
    End Sub

    Private Sub llena_grid()
        Dim l As Integer = 1
        Dim dr As DataRow
        conexiones(con)
        For i = 0 To 2
            Try
                dt = New DataTable
                cnn.ConnectionString = con(i)
                llena_tablas_e(dt, strsql, cnn)
                pr.Merge(dt)
            Catch
            End Try
        Next
        fg.Rows.Count = pr.Rows.Count + 1
        For Each dr In pr.Rows
            fg(l, 1) = dr("CAJA")
            fg(l, 2) = dr("CORTE")
            fg(l, 3) = dr("CPO")
            fg(l, 4) = dr("ESTILO")
            fg(l, 5) = dr("COLOR")
            fg(l, 6) = dr("TIPO")
            fg(l, 7) = dr("TALLA")
            fg(l, 8) = dr("UNIDADES")
            fg(l, 9) = dr("FECHA")
            l = l + 1
        Next
        sub_total()
    End Sub
    Private Sub sub_total()
        fg.Subtotal(AggregateEnum.Sum, 0, 0, 8, "Gran Total -->")
    End Sub

    Private Sub fg_Click(sender As System.Object, e As System.EventArgs) Handles fg.AfterFilter
        sub_total()
    End Sub
End Class



