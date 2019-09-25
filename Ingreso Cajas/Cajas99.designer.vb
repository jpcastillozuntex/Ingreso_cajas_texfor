<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Cajas99
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Cajas99))
        Me.fg = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.abrear = New System.Windows.Forms.OpenFileDialog()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.caja = New System.Windows.Forms.TextBox()
        Me.jg = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.fg, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.jg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fg
        '
        Me.fg.AllowDragging = C1.Win.C1FlexGrid.AllowDraggingEnum.None
        Me.fg.AllowEditing = False
        Me.fg.AllowFiltering = True
        Me.fg.AllowMergingFixed = C1.Win.C1FlexGrid.AllowMergingEnum.None
        Me.fg.AllowSorting = C1.Win.C1FlexGrid.AllowSortingEnum.None
        Me.fg.ColumnInfo = resources.GetString("fg.ColumnInfo")
        Me.fg.Font = New System.Drawing.Font("Courier New", 8.25!)
        Me.fg.HighLight = C1.Win.C1FlexGrid.HighLightEnum.Never
        Me.fg.KeyActionTab = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross
        Me.fg.Location = New System.Drawing.Point(12, 73)
        Me.fg.Name = "fg"
        Me.fg.Rows.DefaultSize = 19
        Me.fg.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Cell
        Me.fg.Size = New System.Drawing.Size(833, 594)
        Me.fg.StyleInfo = resources.GetString("fg.StyleInfo")
        Me.fg.TabIndex = 0
        '
        'abrear
        '
        Me.abrear.FileName = "OpenFileDialog1"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(186, 33)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Caja:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'caja
        '
        Me.caja.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.caja.Location = New System.Drawing.Point(85, 28)
        Me.caja.MaxLength = 9
        Me.caja.Name = "caja"
        Me.caja.Size = New System.Drawing.Size(229, 26)
        Me.caja.TabIndex = 0
        '
        'jg
        '
        Me.jg.AllowDragging = C1.Win.C1FlexGrid.AllowDraggingEnum.None
        Me.jg.AllowEditing = False
        Me.jg.AllowFiltering = True
        Me.jg.AllowMergingFixed = C1.Win.C1FlexGrid.AllowMergingEnum.None
        Me.jg.AllowSorting = C1.Win.C1FlexGrid.AllowSortingEnum.None
        Me.jg.BackColor = System.Drawing.Color.LightGreen
        Me.jg.ColumnInfo = resources.GetString("jg.ColumnInfo")
        Me.jg.Font = New System.Drawing.Font("Courier New", 8.25!)
        Me.jg.HighLight = C1.Win.C1FlexGrid.HighLightEnum.Never
        Me.jg.KeyActionTab = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross
        Me.jg.Location = New System.Drawing.Point(851, 73)
        Me.jg.Name = "jg"
        Me.jg.Rows.Count = 1
        Me.jg.Rows.DefaultSize = 19
        Me.jg.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Cell
        Me.jg.Size = New System.Drawing.Size(226, 594)
        Me.jg.StyleInfo = resources.GetString("jg.StyleInfo")
        Me.jg.TabIndex = 12
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(851, 37)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(226, 33)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "TEXSUN J"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cajas99
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1093, 690)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.jg)
        Me.Controls.Add(Me.caja)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.fg)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Cajas99"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Ingreso de Cajas"
        CType(Me.fg, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.jg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents fg As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents abrear As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents caja As System.Windows.Forms.TextBox
    Friend WithEvents jg As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
