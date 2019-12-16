Imports iText.Layout
Imports iText.Kernel.Pdf
Imports iText.Kernel.Geom
Imports iText.Layout.Element
Imports iText.Layout.Properties
Imports iText.Kernel.Font
Imports iText.IO.Font.Constants
Imports iText.Kernel.Colors
Imports iText.Kernel.Events
Imports iText.Kernel.Pdf.Canvas

Public Module itext7

    Public Function ExportarDataGridaPDF(ByRef DataGrid As DataGridView, Optional ByVal NombreFichero As String = "", Optional ByVal TituloInforme As String = "", Optional ByVal CrearEnTemp As Boolean = False, Optional ByVal MensajeProcesando As String = "") As DialogResult

        '===============================================================================================
        'Exporta un DataGridView a archivo PDF utilizando los servicios "iText7" de los paquetes "NuGet".
        '===============================================================================================

        If DataGrid.RowCount = 0 Then
            MiMessageBox.ShowWinMessage("El DataGrid NO tiene Filas.", "Exportar DataGrid a PDF", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            Return DialogResult.None
        Else

            If String.IsNullOrWhiteSpace(NombreFichero) Then
                NombreFichero = DataGrid.Name + ".pdf"
            End If
            If System.IO.Path.GetExtension(NombreFichero) = String.Empty Then
                NombreFichero += ".pdf"
            End If

            Dim FicheroDestino As String
            If CrearEnTemp Then
                FicheroDestino = System.IO.Path.GetTempPath() + NombreFichero
            Else
                Using OFD As New OpenFileDialog()
                    OFD.Filter = "Archivo PDF | *.pdf|Todos los archivos (*.*)|*.*"
                    OFD.CheckFileExists = False
                    OFD.AddExtension = True
                    OFD.DefaultExt = "pdf"
                    OFD.Multiselect = False
                    OFD.RestoreDirectory = True
                    OFD.FileName = NombreFichero
                    If OFD.ShowDialog() = DialogResult.Cancel Then
                        MiMessageBox.ShowWinMessage("Exportación CANCELDA por el Usuario.", "Exportar DataGrid a PDF", MsgBoxStyle.Information, MsgBoxStyle.OkOnly)
                        Return DialogResult.Cancel
                    End If
                    FicheroDestino = OFD.FileName
                End Using
            End If

            Dim VMensaje As New MiVentanaMensaje(200, 1, IIf(MensajeProcesando = String.Empty, "Exportando DataGrid a PDF...", MensajeProcesando))

            Try

                Dim PdfWriter As PdfWriter = New PdfWriter(FicheroDestino)                       'Escritor del archivo pdf
                Dim PdfDoc As PdfDocument = New PdfDocument(PdfWriter)                           'Documento PDF que se almacenara via el escritor
                Dim Doc As Document = New Document(PdfDoc, IIf(GetColumnasVisiblesTotalWidth(DataGrid) > 1000, PageSize.A4.Rotate, PageSize.A4)) 'Documento PDF en si, con pagina tamaño letra

                Doc.SetFontSize(8.0F)
                Doc.SetMargins(40.0F, 10.0F, 35.0F, 15.0F)                                       'Margenes del documento

                Dim FuenteHeader As PdfFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA) 'Fuente a usar en las Cabeceras de la tabla
                Dim FuenteCeldas As PdfFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA) 'Fuente a usar en las Celdas de la tabla

                Dim Evento As EventoPagina = New EventoPagina(Doc, DataGrid.Name, TituloInforme) 'Manejador de evento de pagina, el cual agregara el encabezado y pie de pagina
                PdfDoc.AddEventHandler(PdfDocumentEvent.END_PAGE, Evento)                        'Indicamos que el manejador se encargara del evento END_PAGE

                Dim Tabla As New Table(UnitValue.CreatePercentArray(GetColumnasVisiblesWidth(DataGrid)))    'Crea un objeto Table con el numero de columnas del DataGridView, con ancho de celdas de la Tabla proporcional al ancho de las columnas del DGV. 
                Tabla.SetFixedLayout()
                Tabla.SetPaddingTop(0)
                Tabla.SetPaddingBottom(0)
                Tabla.SetPaddingLeft(3)
                Tabla.SetPaddingRight(3)

                'Cabecera
                For i As Integer = 0 To DataGrid.ColumnCount - 1
                    If DataGrid.Columns(i).Visible Then
                        Dim Celda As Cell = New Cell()
                        Celda.Add(New Paragraph(DataGrid.Columns(i).HeaderText))
                        Celda.SetFont(FuenteHeader)
                        Celda.SetFontSize(10.0F)
                        Celda.SetFontColor(iText.Kernel.Colors.ColorConstants.WHITE)
                        Celda.SetBackgroundColor(New DeviceRgb(0, 0, 192))
                        If GetAlineacionHorizontal(DataGrid.Columns(i).DefaultCellStyle.Alignment.ToString) = AlineacionHorizontal.LEFT Then
                            Celda.SetTextAlignment(TextAlignment.LEFT)
                        ElseIf GetAlineacionHorizontal(DataGrid.Columns(i).DefaultCellStyle.Alignment.ToString) = AlineacionHorizontal.CENTER Then
                            Celda.SetTextAlignment(TextAlignment.CENTER)
                        ElseIf GetAlineacionHorizontal(DataGrid.Columns(i).DefaultCellStyle.Alignment.ToString) = AlineacionHorizontal.RIGHT Then
                            Celda.SetTextAlignment(TextAlignment.RIGHT)
                        Else
                            Celda.SetTextAlignment(TextAlignment.LEFT)
                        End If
                        Tabla.AddHeaderCell(Celda)
                    End If
                Next

                'Tabla
                For i As Integer = 0 To DataGrid.RowCount - 1
                    For j As Integer = 0 To DataGrid.ColumnCount - 1
                        If DataGrid.Columns(j).Visible Then
                            Dim Celda As Cell = New Cell()
                            Celda.Add(New Paragraph(Convert.ToString(DataGrid.Rows(i).Cells(j).FormattedValue)))
                            Celda.SetFont(FuenteCeldas)
                            'Celda.SetFontSize(10)
                            If GetAlineacionHorizontal(DataGrid.Columns(j).DefaultCellStyle.Alignment.ToString) = AlineacionHorizontal.LEFT Then
                                Celda.SetTextAlignment(TextAlignment.LEFT)
                            ElseIf GetAlineacionHorizontal(DataGrid.Columns(j).DefaultCellStyle.Alignment.ToString) = AlineacionHorizontal.CENTER Then
                                Celda.SetTextAlignment(TextAlignment.CENTER)
                            ElseIf GetAlineacionHorizontal(DataGrid.Columns(j).DefaultCellStyle.Alignment.ToString) = AlineacionHorizontal.RIGHT Then
                                Celda.SetTextAlignment(TextAlignment.RIGHT)
                            Else
                                Celda.SetTextAlignment(TextAlignment.LEFT)
                            End If
                            Tabla.AddCell(Celda)
                        End If
                    Next
                Next
                Doc.Add(Tabla)
                Doc.Close()
                VMensaje.Close()
                Return DialogResult.OK
            Catch ex As Exception
                VMensaje.Close()
                MiMessageBox.ShowWinMessage("Se ha producido una Excepción en el proceso de Exportación del DataGrid;" + vbCrLf + vbCrLf + ex.Message, "Exportar DataGrid a PDF", MsgBoxStyle.Critical, MsgBoxStyle.OkOnly)
                Return DialogResult.Abort
            End Try

        End If

    End Function

    Public Class EventoPagina

        Implements IEventHandler

        Private Documento As Document
        Private NombreDatagrid As String
        Private TituloInforme As String

        Public Sub New(Documento As Document, NombreDatagrid As String, TituloInforme As String)

            Me.Documento = Documento
            Me.NombreDatagrid = NombreDatagrid
            Me.TituloInforme = TituloInforme

        End Sub

        Public Sub HandleEvent([event] As [Event]) Implements IEventHandler.HandleEvent     'Manejador del Evento de Cambio de Página, agrega el Encabezado y Pie de Página

            Dim Evento As PdfDocumentEvent = [event]

            Dim PdfDoc As PdfDocument = Evento.GetDocument()
            Dim Pagina As PdfPage = Evento.GetPage()
            Dim Canvas As PdfCanvas = New PdfCanvas(Pagina.NewContentStreamBefore(), Pagina.GetResources(), PdfDoc)

            'Cabecera
            Dim TablaEncabezado As Table = New Table({1.0F}) : TablaEncabezado.SetWidth(Pagina.GetPageSize().GetWidth() - (Documento.GetLeftMargin() + Documento.GetRightMargin()))
            Dim CeldaC As Cell = New Cell() : CeldaC.SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA_OBLIQUE)) : CeldaC.SetBold() : CeldaC.SetBackgroundColor(New DeviceRgb(192, 192, 255))
            Dim Aux As String = IIf(String.IsNullOrWhiteSpace(TituloInforme), "DataGridView " + NombreDatagrid, TituloInforme)
            CeldaC.Add(New Paragraph(Aux)) : CeldaC.SetTextAlignment(TextAlignment.CENTER)
            TablaEncabezado.AddCell(CeldaC)
            Dim RectanguloEncabezado As Rectangle = New Rectangle(PdfDoc.GetDefaultPageSize().GetX() + Documento.GetLeftMargin(), PdfDoc.GetDefaultPageSize().GetTop() - Documento.GetTopMargin() * 1.5F, Pagina.GetPageSize().GetWidth() - (Documento.GetLeftMargin() + Documento.GetRightMargin()), 50.0F)

            Dim CanvasEncabezado As Canvas = New Canvas(Canvas, PdfDoc, RectanguloEncabezado)
            CanvasEncabezado.Add(TablaEncabezado)

            'Pie
            Dim TablaPie As Table = New Table(UnitValue.CreatePercentArray({1.0F, 1.0F})) : TablaPie.SetWidth(Pagina.GetPageSize().GetWidth() - (Documento.GetLeftMargin() + Documento.GetRightMargin()))
            Dim CeldaFecha As Cell = New Cell()
            CeldaFecha.Add(New Paragraph(Now.ToLongDateString + " - " + Now.ToShortTimeString)) : CeldaFecha.SetFontSize(8) : CeldaFecha.SetTextAlignment(TextAlignment.LEFT) : CeldaFecha.SetPaddingLeft(5)
            TablaPie.AddCell(CeldaFecha)
            Dim CeldaNumPag As Cell = New Cell()
            CeldaNumPag.Add(New Paragraph("Página. " + PdfDoc.GetPageNumber(Pagina).ToString)) : CeldaNumPag.SetFontSize(8) : CeldaNumPag.SetTextAlignment(TextAlignment.RIGHT) : CeldaNumPag.SetPaddingRight(5)
            TablaPie.AddCell(CeldaNumPag)
            Dim RectanguloPie As Rectangle = New Rectangle(PdfDoc.GetDefaultPageSize().GetX() + Documento.GetLeftMargin(), PdfDoc.GetDefaultPageSize().GetBottom(), Pagina.GetPageSize().GetWidth() - (Documento.GetLeftMargin() + Documento.GetRightMargin()), 30.0F)

            Dim CanvasPie As Canvas = New Canvas(Canvas, PdfDoc, RectanguloPie)
            CanvasPie.Add(TablaPie)

        End Sub

    End Class

End Module
