Public Class ImprimirDataGrid

    Private DataGrid As DataGridView
    Private TituloInforme As String

    Private ContadorFilas As Integer = 0
    Private NumeroPagina As Integer = 1
    Private PaginaNueva As Boolean = True

    Public Shared Function Imprimir(ByRef DataGrid As DataGridView, Optional ByVal TituloInforme As String = "", Optional ByVal VistaPrevia As Boolean = False) As DialogResult

        If DataGrid.RowCount = 0 Then
            MiMessageBox.ShowWinMessage("El DataGrid NO tiene Filas.", "Imprimir DataGrid", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            Return DialogResult.None
        Else
            Dim ObjImpDGV As ImprimirDataGrid = New ImprimirDataGrid
            ObjImpDGV.PrtDocumento_Imprimir(DataGrid, TituloInforme, VistaPrevia)
            Return DialogResult.OK
        End If

    End Function

    Private Sub PrtDocumento_Imprimir(ByRef DataGrid As DataGridView, Optional ByVal TituloInforme As String = "", Optional ByVal VistaPrevia As Boolean = False)

        Me.DataGrid = DataGrid
        Me.TituloInforme = TituloInforme

        Dim PrtDialogo As New PrintDialog
        Dim PrtSettings As PrinterSettings = New PrinterSettings

        Dim PrtDocumento As New PrintDocument
        AddHandler PrtDocumento.PrintPage, AddressOf PrtDocumento_PrintPage     'Método de Evento para cada Página a imprimir
        AddHandler PrtDocumento.BeginPrint, AddressOf PrtDocumento_BeginPrint   'Metodo antes de que se imprima la primera Página del Documento

        If Not VistaPrevia Then
            With PrtDialogo
                .AllowCurrentPage = False
                .AllowPrintToFile = False
                .AllowSelection = False
                .AllowSomePages = False
                .PrintToFile = False
                .ShowHelp = False
                .ShowNetwork = False
                .PrinterSettings = PrtSettings
                If .ShowDialog() = DialogResult.OK Then
                    PrtSettings = .PrinterSettings
                Else
                    Exit Sub
                End If
            End With
        End If

        PrtDocumento.PrinterSettings = PrtSettings
        'PrtDocumento.PrintController = New StandardPrintController() 'Para que NO muestre la ventana "Imprimiendo ...". El controlador por defecto es "PrintControllerWithStatusDialog"
        PrtDocumento.PrinterSettings.DefaultPageSettings.Margins = New Margins(15, 15, 60, 35)
        If GetColumnasVisiblesTotalWidth(DataGrid) > 1000 Then
            PrtDocumento.PrinterSettings.DefaultPageSettings.Landscape = True
        Else
            PrtDocumento.PrinterSettings.DefaultPageSettings.Landscape = False
        End If
        If VistaPrevia Then
            Dim PrtPrev As New PrintPreviewDialog
            Dim Screen As Screen = Screen.PrimaryScreen
            PrtPrev.Document = PrtDocumento
            PrtPrev.StartPosition = FormStartPosition.CenterParent
            PrtPrev.PrintPreviewControl.Zoom = 1
            PrtPrev.Width = GetColumnasVisiblesTotalWidth(DataGrid) / 1.2! + 50 : PrtPrev.Height = Screen.Bounds.Height / 1.5!
            PrtPrev.Text = IIf(String.IsNullOrWhiteSpace(TituloInforme), DataGrid.Name, TituloInforme) + " - Vista Previa de Impresión"
            PrtPrev.ShowDialog()
        Else
            PrtDocumento.Print()
        End If

    End Sub

    Private Sub PrtDocumento_PrintPage(ByVal sender As Object, ByVal e As PrintPageEventArgs) 'Handles PrtDocumento.PrintPage

        Dim Formato As StringFormat = New StringFormat(StringFormatFlags.LineLimit) : Formato.LineAlignment = StringAlignment.Center : Formato.Trimming = StringTrimming.EllipsisCharacter
        Dim Rectangulo As Rectangle

        Dim FuenteTitulo As Font = New Font("Microsoft Sans Serif", 12.0!, FontStyle.Bold, GraphicsUnit.Point, CType(0, Byte))
        Dim FuenteHeader As Font = New Font("Microsoft Sans Serif", 9.75!, FontStyle.Regular, GraphicsUnit.Point, CType(0, Byte))
        Dim FuenteCeldas As Font = New Font("Microsoft Sans Serif", 8.0!, FontStyle.Regular, GraphicsUnit.Point, CType(0, Byte))

        Dim y As Int32 = e.MarginBounds.Top
        Dim x As Int32
        Dim h As Int32 = 0
        Dim Fila As DataGridViewRow

        'Imprime Titulo y Cabecera para cada Página Nueva ...
        If PaginaNueva Then
            Fila = DataGrid.Rows(ContadorFilas)
            x = e.MarginBounds.Left

            'Titulo
            Rectangulo = New Rectangle(x, e.MarginBounds.Top / 3, GetColumnasVisiblesTotalWidth(DataGrid) / 1.25F, e.MarginBounds.Top / 3 + 5) : Formato.Alignment = StringAlignment.Center
            e.Graphics.FillRectangle(Brushes.CornflowerBlue, Rectangulo) : e.Graphics.DrawRectangle(Pens.Black, Rectangulo)
            If String.IsNullOrWhiteSpace(TituloInforme) Then
                e.Graphics.DrawString("DataGridView " + DataGrid.Name + " - " + Now.ToLongDateString + " - " + Now.ToShortTimeString, FuenteTitulo, Brushes.Black, Rectangulo, Formato)
            Else
                e.Graphics.DrawString(TituloInforme, FuenteTitulo, Brushes.Black, Rectangulo, Formato)
            End If

            'Cabecera
            For Each Celda As DataGridViewCell In Fila.Cells
                If Celda.Visible Then
                    Rectangulo = New Rectangle(x, y, Celda.Size.Width / 1.25F, Celda.Size.Height / 1.25F)
                    e.Graphics.FillRectangle(Brushes.RoyalBlue, Rectangulo) : e.Graphics.DrawRectangle(Pens.Black, Rectangulo)
                    Select Case GetAlineacionHorizontal(DataGrid.Columns(Celda.ColumnIndex).DefaultCellStyle.Alignment.ToString)
                        Case AlineacionHorizontal.RIGHT
                            Formato.Alignment = StringAlignment.Far
                            Rectangulo.Offset(-1, 0)
                        Case AlineacionHorizontal.CENTER
                            Formato.Alignment = StringAlignment.Center
                        Case Else
                            Formato.Alignment = StringAlignment.Near
                            Rectangulo.Offset(2, 0)
                    End Select
                    e.Graphics.DrawString(DataGrid.Columns(Celda.ColumnIndex).HeaderText, FuenteHeader, Brushes.White, Rectangulo, Formato)
                    x += Rectangulo.Width
                    h = Math.Max(h, Rectangulo.Height)
                End If
            Next
            y += h
            'Imprime Número de Página
            Rectangulo = New Rectangle(e.MarginBounds.Left, e.PageBounds.Bottom - 25, (e.MarginBounds.Right - e.MarginBounds.Left) / 2, 15) : Formato.Alignment = StringAlignment.Near
            e.Graphics.DrawString(Now.ToLongDateString + " - " + Now.ToShortTimeString, FuenteCeldas, Brushes.Black, Rectangulo, Formato)
            Rectangulo = New Rectangle((e.MarginBounds.Right - e.MarginBounds.Left) / 2, e.PageBounds.Bottom - 25, (e.MarginBounds.Right - e.MarginBounds.Left) / 2, 15) : Formato.Alignment = StringAlignment.Far
            e.Graphics.DrawString("Página : " + NumeroPagina.ToString, FuenteCeldas, Brushes.Black, Rectangulo, Formato)
        End If
        PaginaNueva = False

        'Imprime cada Fila
        For i As Int32 = ContadorFilas To DataGrid.RowCount - 1
            Fila = DataGrid.Rows(i)
            h = 0
            x = e.MarginBounds.Left 'Reset x para las lineas de Datos
            For Each Celda As DataGridViewCell In Fila.Cells
                If Celda.Visible Then
                    Rectangulo = New Rectangle(x, y, Celda.Size.Width / 1.25, Celda.Size.Height / 1.25F)
                    e.Graphics.DrawRectangle(Pens.Black, Rectangulo)
                    Select Case GetAlineacionHorizontal(DataGrid.Columns(Celda.ColumnIndex).DefaultCellStyle.Alignment.ToString)
                        Case AlineacionHorizontal.RIGHT
                            Formato.Alignment = StringAlignment.Far
                            Rectangulo.Offset(-1, 0)
                        Case AlineacionHorizontal.CENTER
                            Formato.Alignment = StringAlignment.Center
                        Case Else
                            Formato.Alignment = StringAlignment.Near
                            Rectangulo.Offset(2, 0)
                    End Select

                    If Celda.ValueType = System.Type.GetType("System.Boolean") Then
                        e.Graphics.DrawString(IIf(Celda.Value, "V", "F"), FuenteCeldas, Brushes.Black, Rectangulo, Formato) 'Para que escriba 'V' o 'F' en vez de 'True' o 'False'.
                    Else
                        e.Graphics.DrawString(Celda.FormattedValue.ToString(), FuenteCeldas, Brushes.Black, Rectangulo, Formato)
                    End If
                    x += Rectangulo.Width
                    h = Math.Max(h, Rectangulo.Height)
                    End If
            Next
            y += h
            ContadorFilas = i + 1

            If y + h > e.MarginBounds.Bottom Then
                e.HasMorePages = True
                PaginaNueva = True : NumeroPagina += 1
                Return
            End If
        Next

    End Sub

    Private Sub PrtDocumento_QueryPageSettings(ByVal sender As Object, ByVal e As System.Drawing.Printing.QueryPageSettingsEventArgs) 'Handles PrtDocumento.QueryPageSettings

        'Before it prints each page, the PrintDocument object raises its QueryPageSettings event. The program's event handler sets e.PageSettings.Landscape

    End Sub

    Private Sub PrtDocumento_BeginPrint(sender As Object, e As PrintEventArgs) 'Handles PrtDocumento.BeginPrint

        'Se restablecen valores, pues, en caso de PreView, cuando se clicka el boton imprimir se vuelve a generar el documento.
        ContadorFilas = 0
        NumeroPagina = 1
        PaginaNueva = True
    End Sub

End Class
