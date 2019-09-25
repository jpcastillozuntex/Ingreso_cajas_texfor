Imports C1.Win.C1FlexGrid
Imports System.Drawing.Printing
Imports System.IO

Public Class Cajas10
    Inherits System.Windows.Forms.Form
    Dim h As String = "#######0.00"
    Dim ok As Boolean
    Dim si As Integer
    Dim cnn As New SqlClient.SqlConnection
    Dim dt As New DataTable
    Dim lineas As Integer
    Dim fecha As String
    Dim fecha1 As String
    Dim corte As String
    Dim path As String = "c:\estado_cliente.xls"
    Dim obj As New empresas
    Friend WithEvents fg As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents excel As System.Windows.Forms.Button
    Dim constr As String = obj.constr
    Dim cnstr As C1cajasLib_SIF.util

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
    Friend WithEvents S1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents fecha_i As System.Windows.Forms.DateTimePicker
    Friend WithEvents Cancela As System.Windows.Forms.Button
    Friend WithEvents Imprime As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Cajas10))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fecha_i = New System.Windows.Forms.DateTimePicker()
        Me.Imprime = New System.Windows.Forms.Button()
        Me.Cancela = New System.Windows.Forms.Button()
        Me.S1 = New System.Windows.Forms.Button()
        Me.fg = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.excel = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.fg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fecha_i
        '
        Me.fecha_i.CustomFormat = "MM/yyyy"
        Me.fecha_i.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fecha_i.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.fecha_i.Location = New System.Drawing.Point(72, 9)
        Me.fecha_i.Name = "fecha_i"
        Me.fecha_i.ShowUpDown = True
        Me.fecha_i.Size = New System.Drawing.Size(96, 26)
        Me.fecha_i.TabIndex = 61
        Me.ToolTip1.SetToolTip(Me.fecha_i, "Mes")
        '
        'Imprime
        '
        Me.Imprime.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.Imprime.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Imprime.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Imprime.ForeColor = System.Drawing.Color.Black
        Me.Imprime.Image = CType(resources.GetObject("Imprime.Image"), System.Drawing.Image)
        Me.Imprime.Location = New System.Drawing.Point(632, 8)
        Me.Imprime.Name = "Imprime"
        Me.Imprime.Size = New System.Drawing.Size(35, 35)
        Me.Imprime.TabIndex = 69
        Me.Imprime.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.Imprime, "Presione este Boton si Desea Imprimir los Datos mostrados en la Pantalla.")
        Me.Imprime.UseVisualStyleBackColor = False
        Me.Imprime.Visible = False
        '
        'Cancela
        '
        Me.Cancela.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.Cancela.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Cancela.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cancela.ForeColor = System.Drawing.Color.Black
        Me.Cancela.Image = CType(resources.GetObject("Cancela.Image"), System.Drawing.Image)
        Me.Cancela.Location = New System.Drawing.Point(584, 8)
        Me.Cancela.Name = "Cancela"
        Me.Cancela.Size = New System.Drawing.Size(35, 35)
        Me.Cancela.TabIndex = 66
        Me.Cancela.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.Cancela, "Presione este Boton para Cancelar toda la operación y limpiar todos los datos sin" & _
        " Grabar.")
        Me.Cancela.UseVisualStyleBackColor = False
        Me.Cancela.Visible = False
        '
        'S1
        '
        Me.S1.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.S1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.S1.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.S1.ForeColor = System.Drawing.Color.Black
        Me.S1.Image = CType(resources.GetObject("S1.Image"), System.Drawing.Image)
        Me.S1.Location = New System.Drawing.Point(536, 8)
        Me.S1.Name = "S1"
        Me.S1.Size = New System.Drawing.Size(35, 35)
        Me.S1.TabIndex = 57
        Me.S1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.S1, "Chequeo de Datos")
        Me.S1.UseVisualStyleBackColor = False
        '
        'fg
        '
        Me.fg.AllowEditing = False
        Me.fg.AllowMerging = C1.Win.C1FlexGrid.AllowMergingEnum.Free
        Me.fg.AllowSorting = C1.Win.C1FlexGrid.AllowSortingEnum.None
        Me.fg.BackColor = System.Drawing.Color.White
        Me.fg.ColumnInfo = resources.GetString("fg.ColumnInfo")
        Me.fg.FocusRect = C1.Win.C1FlexGrid.FocusRectEnum.None
        Me.fg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.fg.ForeColor = System.Drawing.Color.Black
        Me.fg.HighLight = C1.Win.C1FlexGrid.HighLightEnum.Never
        Me.fg.Location = New System.Drawing.Point(12, 49)
        Me.fg.Name = "fg"
        Me.fg.Rows.DefaultSize = 19
        Me.fg.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row
        Me.fg.Size = New System.Drawing.Size(986, 635)
        Me.fg.StyleInfo = resources.GetString("fg.StyleInfo")
        Me.fg.SubtotalPosition = C1.Win.C1FlexGrid.SubtotalPositionEnum.BelowData
        Me.fg.TabIndex = 70
        Me.ToolTip1.SetToolTip(Me.fg, "Area de Datos.")
        '
        'excel
        '
        Me.excel.BackColor = System.Drawing.Color.FromArgb(CType(CType(46, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(245, Byte), Integer))
        Me.excel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.excel.Font = New System.Drawing.Font("Comic Sans MS", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.excel.ForeColor = System.Drawing.Color.Black
        Me.excel.Image = CType(resources.GetObject("excel.Image"), System.Drawing.Image)
        Me.excel.Location = New System.Drawing.Point(679, 8)
        Me.excel.Name = "excel"
        Me.excel.Size = New System.Drawing.Size(35, 35)
        Me.excel.TabIndex = 71
        Me.excel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.excel, "Presione este Boton si Desea  convertir a Excel los datos. ")
        Me.excel.UseVisualStyleBackColor = False
        Me.excel.Visible = False
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 32)
        Me.Label1.TabIndex = 62
        Me.Label1.Text = "Mes:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Cajas10
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1010, 696)
        Me.Controls.Add(Me.excel)
        Me.Controls.Add(Me.fg)
        Me.Controls.Add(Me.Imprime)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.fecha_i)
        Me.Controls.Add(Me.Cancela)
        Me.Controls.Add(Me.S1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Red
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Cajas10"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Produccion Diaria por Seccion"
        CType(Me.fg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub Prod_sec(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fg.Height = Me.Height - 120
        fg.Width = Me.Width - 40
        fg.Rows.Count = 0
    End Sub

    Private Sub setea_fg()
        fg.Clear()
        fg.Rows.Count = 1
        fg.Rows.Fixed = 1
        fg.Cols.Fixed = 0
        fg.Cols.Count = 2
        fg.Cols(0).Width = 5
        fg.Rows(0).Height = 30
        fg.Cols(0).Name = "DIA"
        fg(0, 0) = fg.Cols(0).Name
        fg.Cols(0).Width = 40
        fg.Cols(1).Width = 80
    End Sub

    Private Sub llena_mes()
        Dim i As Integer
        Dim linea As Integer
        Dim fecha As Date
        Dim dia As String
        For i = 1 To 31
            Try
                fecha = Mid(fecha_i.Text, 4) + "-" + Mid(fecha_i.Text, 1, 2) + "-" + CStr(i)
                dia = UCase(Format(fecha, "dddd"))
                linea = fg.Rows.Count
                fg.Rows.Count = linea + 1
                fg(linea, 0) = Format(i, "00")
                fg(linea, 1) = dia
            Catch
            End Try
        Next
    End Sub

    Private Sub produccion(ByVal cnn As SqlClient.SqlConnection)
        Dim mes As String = Mid(fecha_i.Text, 1, 2)
        Dim ano As String = Mid(fecha_i.Text, 4)
        Dim indice As Integer
        Dim fechas As Date
        Dim fecha As String
        Dim dia As Integer
        Dim seccion As String
        Dim prendas As Integer
        Dim strSQL As String
        strSQL = "SELECT CONVERT (date,FECHA) AS FECHA , SUM(UNIDADES) AS PRENDA, SECCION FROM CAJAS04 LEFT JOIN CORTES ON CAJAS04.CORTE = CORTES.CORTE WHERE MONTH(FECHA) = '" & mes & "' AND YEAR(FECHA) ='" & ano & "' AND SECCION LIKE 'TEX%'  GROUP BY SECCION,CONVERT (date,FECHA) ORDER BY SECCION"
        dt = New DataTable
        Dim dr As DataRow
        llena_tablas_e(dt, strSQL, cnn)
        For Each dr In dt.Rows
            fechas = dr("FECHA")
            fecha = Format(fechas, "dd/MM/yyyy")
            dia = CInt(Mid(fecha, 1, 2))
            seccion = dr("SECCION")
            prendas = dr("PRENDA")
            indice = fg.Cols.IndexOf(seccion)
            If indice < 1 Then
                crea_columna(seccion)
                indice = fg.Cols.Count - 1
                End If
            fg(dia, indice) = fg(dia, indice) + prendas
        Next
    End Sub

    Private Sub crea_columna(ByVal seccion As String)
        Dim col As Integer
        col = fg.Cols.Count
        fg.Cols.Count = col + 1
        fg.Cols(col).Name = seccion
        fg(0, col) = fg.Cols(col).Name
        fg.Cols(col).Width = 90
        fg.Cols(col).DataType = GetType(Integer)
        fg.Cols(col).Format = "####,##0"
        fg.Cols(col).TextAlign = TextAlignEnum.RightCenter
        fg.Cols(col).TextAlignFixed = TextAlignEnum.RightCenter
    End Sub

    Private Sub S1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles S1.Click
        setea_fg()
        llena_mes()
        llena_produccion()
        S1.Visible = False
        Cancela.Visible = True
        fecha_i.Enabled = False
        Imprime.Visible = True
        excel.Visible = True
        Dim obj As New empresas
        obj.constr = constr
    End Sub
    Private Sub llena_produccion()
        Dim i As Integer
        Dim str As New C1cajasLib_SIF.util
        For i = 0 To 2
            cnn.ConnectionString = str.con_string(i)
            produccion(cnn)
        Next
        crea_columna("TOTALES")
        totales()
    End Sub
    Private Sub totales()
        Dim i As Integer
        Dim j As Integer
        Dim cols As Integer = fg.Cols.Count - 1
        Dim filas As Integer = fg.Rows.Count
        fg.Rows.Count = fg.Rows.Count + 1
        fg(filas, 1) = "TOTAL MES"
        fg.Rows(filas).Height = 30
        For j = 1 To filas - 1
            For i = 2 To cols - 1
                fg(j, cols) = fg(j, cols) + fg(j, i)
                fg(filas, i) = fg(filas, i) + fg(j, i)
            Next
            fg(filas, cols) = fg(filas, cols) + fg(j, cols)
        Next
        fg.Rows(filas).Style = fg.Styles("amarillo")
        fg.Cols(cols).Style = fg.Styles("amarillo")
    End Sub

    Private Sub CANCELA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancela.Click
        Cancela.Visible = False
        excel.Visible = False
        S1.Visible = True
        fecha_i.Enabled = True
        Imprime.Visible = False
        fg.Rows.Count = 0
    End Sub

    Private Sub Imprime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Imprime.Click
        Print_fg()
    End Sub

    Private Sub Print_fg()
        Dim pd As PrintDocument
        Dim fecha As String
        fecha = Format(Today, "dd/MM/yyyy")
        pd = fg.PrintParameters.PrintDocument()
        With pd.DefaultPageSettings
            .Landscape = True
            .Margins.Left = 20
            .Margins.Right = 50
            .Margins.Top = 50
            .Margins.Bottom = 50
        End With
        fg.PrintParameters.HeaderFont = New Font("Arial Narrow", 14, FontStyle.Bold)
        fg.PrintParameters.FooterFont = New Font("Arial Narrow", 12, FontStyle.Italic)
        fg.PrintGrid("Telas", PrintGridFlags.FitToPageWidth + PrintGridFlags.ShowPrintDialog, "Reporte de prendas producidas. del mes:  " + fecha_i.Text + Chr(9) + Chr(9) + "Pagina {0}", "")
    End Sub

    Private Sub excel_Click(sender As System.Object, e As System.EventArgs) Handles excel.Click
        Dim ok As Boolean
        a_excel(fg, "c:\reportes\prod_diaria_sec.xls", ok)
        If ok Then
            Close()
        End If
    End Sub

    Private Sub fg_Click(sender As System.Object, e As System.EventArgs) Handles fg.Click
        Dim seccion As String = fg(0, fg.ColSel)
        Dim fecha As String = Format(fecha_i.Value, "yyyy") + "-" + Format(fecha_i.Value, "MM") & "-" & Format(CInt(fg(fg.RowSel, 0)), "00")
        Dim strsql As String = " "
        If fg.ColSel > 1 And fg.ColSel < fg.Cols.Count Then
            If fg(fg.RowSel, fg.ColSel) > 0 Then
                strsql = "SELECT * FROM CAJAS04 LEFT JOIN CORTES ON CAJAS04.CORTE = CORTES.CORTE LEFT JOIN CAJAS01 ON CAJAS01.CAJA = CAJAS04.CAJA WHERE CONVERT(date,CAJAS04.FECHA) = '" & fecha & "' AND CAJAS01.SECCION = '" & seccion & "'"
                Dim forma As New Cajas12
                forma.strsql = strsql
                forma.ShowDialog()
            End If
        End If

    End Sub
End Class



