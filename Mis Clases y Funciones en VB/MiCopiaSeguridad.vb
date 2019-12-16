Public NotInheritable Class MiCopiaSeguridad

#Region "Variables para la Clase ..."

    Private Shared ObjMiCopiaSeguridad As MiCopiaSeguridad
    Private Shared WithEvents Ventana As Form

    Private Shared Componentes As ComponentModel.Container

    Private Shared WithEvents BtCancelar As Button
    Private Shared WithEvents SlEsc As ToolStripStatusLabel
    Private Shared WithEvents SsCopiaSeguridad As StatusStrip
    Private Shared WithEvents SlBuscarCarpeta As ToolStripStatusLabel
    Private Shared WithEvents Label3 As Label
    Private Shared WithEvents CarpetaDestino As MiTextBox
    Private Shared WithEvents BtCopiaSeguridad As Button
    Private Shared WithEvents PictureBox3 As PictureBox
    Private Shared WithEvents FolderBrowserDialog As FolderBrowserDialog
    Private Shared WithEvents SlEntrar As ToolStripStatusLabel

    Private Shared CarpetaCopiaSeguridad As String = String.Empty

    Public Const Cancelado As String = "#@$Ñ5jy&j<"             'Constante devuelta por la funcion, para indicar se se pulsó el botón CANCELAR

#End Region

#Region "Constructor de la clase"
    Private Sub New()

        Ventana = New Form()

    End Sub

#End Region

    Public Shared Function ShowDialogo(Optional ByVal CarpetaPreSeleccionada As String = "C:\") As String

        CrearVentana()
        InicializarVariables(CarpetaPreSeleccionada)

        Ventana.ActiveControl = BtCopiaSeguridad
        BtCopiaSeguridad.Focus()
        Ventana.ShowDialog()

        Try
            If Componentes IsNot Nothing Then
                Componentes.Dispose()
            End If
        Finally
            Ventana.Dispose() : Ventana = Nothing
        End Try

        Return CarpetaCopiaSeguridad

    End Function

    Private Shared Sub InicializarVariables(Optional ByVal CarpetaPreSeleccionada As String = "C:\")

        If CarpetaPreSeleccionada = String.Empty Then
            CarpetaPreSeleccionada = "C:\"
        End If
        CarpetaDestino.Text = CarpetaPreSeleccionada

    End Sub

    Private Shared Sub BtCopiaSeguridad_Click(sender As Object, e As EventArgs) Handles BtCopiaSeguridad.Click

        If CarpetaDestino.Text.Trim = String.Empty Then
            MiMessageBox.ShowWinMessage("Especificar una Carpeta Local ...", "Destino Copia Seguridad", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            CarpetaDestino.Focus()
        ElseIf Not Directory.Exists(CarpetaDestino.Text.Trim) Then
            MiMessageBox.ShowWinMessage("La Carpeta Local especificada NO existe." + vbCrLf + vbCrLf + "Especificar una Carpeta Local existente ...", "Destino Copia Seguridad", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            CarpetaDestino.Focus()
        Else
            CarpetaCopiaSeguridad = CarpetaDestino.Text.Trim
            Ventana.DialogResult = DialogResult.OK
        End If

    End Sub

    Private Shared Sub CarpetaDestino_DoubleClick(sender As Object, e As EventArgs) Handles CarpetaDestino.DoubleClick, SlBuscarCarpeta.Click

        FolderBrowserDialog.SelectedPath = CarpetaDestino.Text
        If FolderBrowserDialog.ShowDialog() = DialogResult.OK Then
            CarpetaDestino.Text = FolderBrowserDialog.SelectedPath
        End If

    End Sub
    Private Shared Sub CarpetaDestino_Enter(sender As Object, e As EventArgs) Handles CarpetaDestino.Enter

        SlBuscarCarpeta.Visible = True

    End Sub

    Private Shared Sub CarpetaDestino_Leave(sender As Object, e As EventArgs) Handles CarpetaDestino.Leave

        SlBuscarCarpeta.Visible = False

    End Sub

    Private Shared Sub BtCancelar_Click(sender As Object, e As EventArgs) Handles BtCancelar.Click, SlEsc.Click

        CarpetaCopiaSeguridad = Cancelado 'DialogResult.Cancel.ToString.ToUpper
        Ventana.DialogResult = DialogResult.Cancel

    End Sub

    Private Shared Sub CrearVentana()

        ObjMiCopiaSeguridad = Nothing
        ObjMiCopiaSeguridad = New MiCopiaSeguridad()

        Componentes = New System.ComponentModel.Container()

        BtCancelar = New System.Windows.Forms.Button()
        SlEsc = New System.Windows.Forms.ToolStripStatusLabel()
        SsCopiaSeguridad = New System.Windows.Forms.StatusStrip()
        SlEntrar = New System.Windows.Forms.ToolStripStatusLabel()
        SlBuscarCarpeta = New System.Windows.Forms.ToolStripStatusLabel()
        Label3 = New System.Windows.Forms.Label()
        CarpetaDestino = New MisClasesFuncionesC.MiTextBox(Componentes)
        BtCopiaSeguridad = New System.Windows.Forms.Button()
        FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog()
        PictureBox3 = New System.Windows.Forms.PictureBox()

        SsCopiaSeguridad.SuspendLayout()
        CType(PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Ventana.SuspendLayout()
        '
        'BtCancelar
        '
        BtCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        BtCancelar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        BtCancelar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        BtCancelar.Location = New System.Drawing.Point(492, 33)
        BtCancelar.Name = "BtCancelar"
        BtCancelar.Size = New System.Drawing.Size(90, 23)
        BtCancelar.TabIndex = 3
        BtCancelar.Text = "Cancelar"
        BtCancelar.UseVisualStyleBackColor = True
        '
        'SlEsc
        '
        SlEsc.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        SlEsc.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        SlEsc.Name = "SlEsc"
        SlEsc.Size = New System.Drawing.Size(88, 17)
        SlEsc.Text = "Esc=Cancelar"
        '
        'SsCopiaSeguridad
        '
        SsCopiaSeguridad.BackColor = System.Drawing.SystemColors.ControlLight
        SsCopiaSeguridad.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        SsCopiaSeguridad.GripStyle = System.Windows.Forms.ToolStripGripStyle.Visible
        SsCopiaSeguridad.Items.AddRange(New System.Windows.Forms.ToolStripItem() {SlEntrar, SlEsc, SlBuscarCarpeta})
        SsCopiaSeguridad.Location = New System.Drawing.Point(0, 60)
        SsCopiaSeguridad.Name = "SsCopiaSeguridad"
        SsCopiaSeguridad.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        SsCopiaSeguridad.Size = New System.Drawing.Size(585, 22)
        SsCopiaSeguridad.TabIndex = 0
        SsCopiaSeguridad.Text = "StatusStrip1"
        '
        'SlEntrar
        '
        SlEntrar.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        SlEntrar.ForeColor = System.Drawing.Color.Green
        SlEntrar.Name = "SlEntrar"
        SlEntrar.Size = New System.Drawing.Size(137, 17)
        SlEntrar.Text = "Entrar=RealizarCopia"
        '
        'SlBuscarCarpeta
        '
        SlBuscarCarpeta.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        SlBuscarCarpeta.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        SlBuscarCarpeta.Name = "SlBuscarCarpeta"
        SlBuscarCarpeta.Size = New System.Drawing.Size(146, 17)
        SlBuscarCarpeta.Text = "D.Click=BuscarCarpeta"
        SlBuscarCarpeta.Visible = False
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(215, 12)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(63, 13)
        Label3.TabIndex = 0
        Label3.Text = "En Carpeta:"
        Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CarpetaDestino
        '
        CarpetaDestino.BackColor = System.Drawing.SystemColors.Window
        CarpetaDestino.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        CarpetaDestino.Formato = ""
        CarpetaDestino.Location = New System.Drawing.Point(281, 8)
        CarpetaDestino.Margin = New System.Windows.Forms.Padding(4)
        CarpetaDestino.Name = "CarpetaDestino"
        CarpetaDestino.NumericoSolo = False
        CarpetaDestino.NumeroSolo = False
        CarpetaDestino.SeleccionarTodoCuandoClick = True
        CarpetaDestino.Size = New System.Drawing.Size(300, 22)
        CarpetaDestino.TabIndex = 2
        CarpetaDestino.TextoVacioNumPermitido = True
        CarpetaDestino.UsarEnterComoTab = True
        '
        'BtCopiaSeguridad
        '
        BtCopiaSeguridad.BackColor = System.Drawing.Color.Beige
        BtCopiaSeguridad.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        BtCopiaSeguridad.ForeColor = System.Drawing.Color.Green
        BtCopiaSeguridad.Location = New System.Drawing.Point(38, 7)
        BtCopiaSeguridad.Name = "BtCopiaSeguridad"
        BtCopiaSeguridad.Size = New System.Drawing.Size(175, 32)
        BtCopiaSeguridad.TabIndex = 1
        BtCopiaSeguridad.Text = "Copia de Seguridad"
        BtCopiaSeguridad.UseVisualStyleBackColor = False
        '
        'FolderBrowserDialog
        '
        FolderBrowserDialog.Description = "Carpeta Destino para Copia de Seguridad"
        '
        'PictureBox3
        '
        PictureBox3.Image = Global.MisClasesFuncionesVB.My.Resources.Resources.CopiasSeguridad_1_32
        PictureBox3.Location = New System.Drawing.Point(3, 7)
        PictureBox3.Name = "PictureBox3"
        PictureBox3.Size = New System.Drawing.Size(32, 32)
        PictureBox3.TabIndex = 42
        PictureBox3.TabStop = False
        '
        'MiCopiaSeguridad
        '
        Ventana.AcceptButton = BtCopiaSeguridad
        Ventana.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Ventana.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Ventana.CancelButton = BtCancelar
        Ventana.ClientSize = New System.Drawing.Size(585, 82)
        Ventana.ControlBox = False
        Ventana.Controls.Add(BtCancelar)
        Ventana.Controls.Add(SsCopiaSeguridad)
        Ventana.Controls.Add(Label3)
        Ventana.Controls.Add(CarpetaDestino)
        Ventana.Controls.Add(BtCopiaSeguridad)
        Ventana.Controls.Add(PictureBox3)
        Ventana.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Ventana.KeyPreview = True
        Ventana.Name = "MiCopiaSeguridad"
        Ventana.ShowInTaskbar = False
        Ventana.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        SsCopiaSeguridad.ResumeLayout(False)
        SsCopiaSeguridad.PerformLayout()
        CType(PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()

        Ventana.ResumeLayout(False)
        Ventana.PerformLayout()

    End Sub

End Class
