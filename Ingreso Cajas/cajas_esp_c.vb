Imports C1.Win.C1FlexGrid

Public Class cajas_esp_c
    Inherits System.Windows.Forms.Form
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
    Friend WithEvents fg As C1.Win.C1FlexGrid.C1FlexGrid

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(cajas_esp_c))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fg = New C1.Win.C1FlexGrid.C1FlexGrid()
        CType(Me.fg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fg
        '
        Me.fg.AllowDragging = C1.Win.C1FlexGrid.AllowDraggingEnum.None
        Me.fg.AllowEditing = False
        Me.fg.AllowFiltering = True
        Me.fg.AllowResizing = C1.Win.C1FlexGrid.AllowResizingEnum.None
        Me.fg.ColumnInfo = resources.GetString("fg.ColumnInfo")
        Me.fg.FocusRect = C1.Win.C1FlexGrid.FocusRectEnum.None
        Me.fg.Font = New System.Drawing.Font("Tahoma", 9.75!)
        Me.fg.ForeColor = System.Drawing.Color.Black
        Me.fg.HighLight = C1.Win.C1FlexGrid.HighLightEnum.WithFocus
        Me.fg.KeyActionEnter = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross
        Me.fg.KeyActionTab = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross
        Me.fg.Location = New System.Drawing.Point(8, 34)
        Me.fg.Name = "fg"
        Me.fg.Rows.Count = 100
        Me.fg.Rows.DefaultSize = 19
        Me.fg.Size = New System.Drawing.Size(1000, 662)
        Me.fg.StyleInfo = resources.GetString("fg.StyleInfo")
        Me.fg.SubtotalPosition = C1.Win.C1FlexGrid.SubtotalPositionEnum.BelowData
        Me.fg.TabIndex = 76
        '
        'cajas_esp_c
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1014, 707)
        Me.Controls.Add(Me.fg)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Red
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "cajas_esp_c"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Utilización de Cajas"
        CType(Me.fg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim cnn As New SqlClient.SqlConnection()
    Dim dr As DataRow
    Dim mt As New DataTable()
    Dim dt As New DataTable()
    Dim strsql As String
    Private Sub cajas_esp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        setea_grid()
        llena_grid()
    End Sub
    Private Sub llena_grid()
        Dim dt As New DataTable()
        Dim dr As DataRow
        Dim l As Integer = 1
        Dim i As Integer
        llena_tablas(dt, "SELECT * FROM MAT_EST_COL_ES LEFT JOIN MATERIALES ON MAT_EST_COL_ES.CODIGO = MATERIALES.CODIGO ORDER BY ESTILO", cnn)
        fg.Rows.Count = dt.Rows.Count + 1
        For Each dr In dt.Rows
            fg(l, 1) = dr("ESTILO")
            fg(l, 2) = dr("CODIGO")
            fg(l, 3) = dr("DESCRIPCION")
            For i = 4 To 13
                fg(l, i) = dr(i - 2)
            Next
            l = l + 1
        Next
    End Sub

    Private Sub setea_grid()
        fg.Rows.Count = 1
        fg.Rows(0).Height = 30
    End Sub
End Class
