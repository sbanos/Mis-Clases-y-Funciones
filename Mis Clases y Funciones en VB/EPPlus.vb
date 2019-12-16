Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Public Module EPPlus

    Public Function ExportarDataGridaEXCEL(ByRef DataGrid As DataGridView, Optional ByVal NombreFichero As String = "", Optional ByVal TituloInforme As String = "", Optional ByVal CrearEnTemp As Boolean = False, Optional ByVal MensajeProcesando As String = "") As DialogResult

        '===============================================================================================
        'Exporta un DataGridView a Hoja Excel utilizando los servicios "EPPlus" de los paquetes "NuGet".
        'Muy RAPIDA
        '===============================================================================================

        If DataGrid.RowCount = 0 Then
            MiMessageBox.ShowWinMessage("El DataGrid NO tiene Filas.", "Exportar DataGrid a EXCEL", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            Return DialogResult.None
        Else

            If String.IsNullOrWhiteSpace(NombreFichero) Then
                NombreFichero = DataGrid.Name + ".xlsx"
            End If
            If System.IO.Path.GetExtension(NombreFichero) = String.Empty Then
                NombreFichero += ".xlsx"
            End If

            Dim FicheroDestino As String
            If CrearEnTemp Then
                FicheroDestino = System.IO.Path.GetTempPath() + NombreFichero
            Else
                Using OFD As New OpenFileDialog()
                    OFD.Filter = "Archivo Excel | *.xlsx|Todos los archivos (*.*)|*.*"
                    OFD.CheckFileExists = False
                    OFD.AddExtension = True
                    OFD.DefaultExt = "xlsx"
                    OFD.Multiselect = False
                    OFD.RestoreDirectory = True
                    OFD.FileName = NombreFichero
                    If OFD.ShowDialog() = DialogResult.Cancel Then
                        MiMessageBox.ShowWinMessage("Exportación CANCELDA por el Usuario.", "Exportar DataGrid a EXCEL", MsgBoxStyle.Information, MsgBoxStyle.OkOnly)
                        Return DialogResult.Cancel
                    End If
                    FicheroDestino = OFD.FileName
                End Using
            End If

            Dim VMensaje As New MiVentanaMensaje(200, 1, IIf(MensajeProcesando = String.Empty, "Exportando DataGrid a EXCEL...", MensajeProcesando))

            Using ExPackage As ExcelPackage = New ExcelPackage

                ExPackage.Workbook.Properties.Author = "Santos Baños - Funciones y Utilidades"
                ExPackage.Workbook.Properties.Title = DataGrid.Name
                ExPackage.Workbook.Properties.Subject = DataGrid.Name + " Exportado (EPPlus)"
                ExPackage.Workbook.Properties.Created = DateTime.Now

                Dim ExHoja As ExcelWorksheet = ExPackage.Workbook.Worksheets.Add(DataGrid.Name)

                Try
                    Dim k As Int16 = 1
                    For i As Integer = 0 To DataGrid.ColumnCount - 1
                        If DataGrid.Columns(i).Visible Then
                            ExHoja.Cells(4, k).Value = DataGrid.Columns(i).HeaderText
                            ExHoja.Cells(4, k).Style.Font.Bold = True
                            ExHoja.Cells(4, k).Style.Font.Size = 12
                            ExHoja.Cells(4, k).Style.Font.Color.SetColor(System.Drawing.Color.White)
                            ExHoja.Cells(4, k).Style.Fill.PatternType = ExcelFillStyle.Solid
                            ExHoja.Cells(4, k).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Navy)
                            k += 1
                        End If
                    Next
                    For i As Integer = 0 To DataGrid.RowCount - 1
                        k = 1
                        For j As Integer = 0 To DataGrid.ColumnCount - 1
                            If DataGrid.Rows(i).Cells(j).Visible Then
                                If String.IsNullOrEmpty(DataGrid.Columns(j).DefaultCellStyle.Format) Then
                                    ExHoja.Cells.Item(i + 5, k).Value = Convert.ToString(DataGrid.Rows(i).Cells(j).Value)
                                ElseIf DataGrid.Columns(j).DefaultCellStyle.Format.Substring(0, 1) = "N" Then
                                    ExHoja.Cells.Item(i + 5, k).Value = DataGrid.Rows(i).Cells(j).Value
                                    ExHoja.Cells.Item(i + 5, k).Style.Numberformat.Format = "#,##0" + IIf(DataGrid.Columns(j).DefaultCellStyle.Format.Substring(1, 1) <> "0", "." + StrDup(CInt(DataGrid.Columns(j).DefaultCellStyle.Format.Substring(1, 1)), "0"), String.Empty)
                                ElseIf DataGrid.Columns(j).DefaultCellStyle.Format = "d" Then
                                    ExHoja.Cells.Item(i + 5, k).Value = DataGrid.Rows(i).Cells(j).Value
                                    ExHoja.Cells.Item(i + 5, k).Style.Numberformat.Format = "dd/mm/yyyy"
                                ElseIf DataGrid.Columns(j).DefaultCellStyle.Format.Substring(0, 1) = "C" Then
                                    ExHoja.Cells.Item(i + 5, k).Value = DataGrid.Rows(i).Cells(j).Value
                                    ExHoja.Cells.Item(i + 5, k).Style.Numberformat.Format = "#,##0.00 €;[Red]-#,##0.00 €"
                                ElseIf Right(DataGrid.Columns(j).DefaultCellStyle.Format.Trim, 3) = "'%'" Then
                                    ExHoja.Cells.Item(i + 5, k).Value = Convert.ToDecimal(DataGrid.Rows(i).Cells(j).Value / 100)
                                    ExHoja.Cells.Item(i + 5, k).Style.Numberformat.Format = "0%"
                                Else
                                    ExHoja.Cells.Item(i + 5, k).Value = Convert.ToString(DataGrid.Rows(i).Cells(j).FormattedValue)
                                End If
                                k += 1
                            End If
                        Next
                    Next
                    ExHoja.Cells(ExHoja.Dimension.Address).AutoFitColumns() 'Ajustar Ancho Columnas Automaticamente
                    Using Rango As ExcelRange = ExHoja.Cells(2, 3, 2, 3)    'Inserto Titulo despues de "AutoFitColumns" para que no afecte al ancho General de las Columnas.
                        Rango.Value = IIf(String.IsNullOrWhiteSpace(TituloInforme), "DataGridView " + DataGrid.Name, TituloInforme) + " - " + Now.ToLongDateString + " - " + Now.ToShortTimeString
                        Rango.Style.Font.Size = 16
                        Rango.Style.Font.Bold = True
                        Rango.Style.Font.Italic = True
                    End Using
                    ExHoja.Protection.IsProtected = False
                    'ExHoja.Protection.AllowSelectLockedCells = False
                    ExPackage.SaveAs(New FileInfo(FicheroDestino))
                    VMensaje.Close()
                    Return DialogResult.OK
                Catch Ex As Exception
                    VMensaje.Close()
                    MiMessageBox.ShowWinMessage("Se ha producido una Excepción en el proceso de Exportación del DataGrid;" + vbCrLf + vbCrLf + Ex.Message, "Exportar DataGrid a EXCEL", MsgBoxStyle.Critical, MsgBoxStyle.OkOnly)
                    Return DialogResult.Abort
                End Try

            End Using

        End If

    End Function

End Module
