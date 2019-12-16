Public NotInheritable Class MiDestinoImpresion

#Region "Variables para la Clase ..."

    Private Shared ObjMiDestinoImpresion As MiDestinoImpresion
    Private Shared WithEvents Ventana As Form

    Private Shared Componentes As ComponentModel.Container

    Private Shared WithEvents FolderBrowserDialog As FolderBrowserDialog
    Private Shared WithEvents EMailDestino As MiTextBox
    Private Shared WithEvents BtCancelar As Button
    Private Shared WithEvents CarpetaDestino As MiTextBox
    Private Shared WithEvents GbFormatoCarpetaLocal As GroupBox
    Private Shared WithEvents PictureBox3 As PictureBox
    Private Shared WithEvents BtCarpeta As Button
    Private Shared WithEvents PictureBox2 As PictureBox
    Private Shared WithEvents PictureBox1 As PictureBox
    Private Shared WithEvents CbImpresoraDestino As ComboBox
    Private Shared WithEvents Label2 As Label
    Private Shared WithEvents BtEMail As Button
    Private Shared WithEvents BtImpresoraLocal As Button
    Private Shared WithEvents RtbTexto As RichTextBox
    Private Shared WithEvents Label1 As Label
    Private Shared WithEvents Label3 As Label
    Private Shared WithEvents RbExcel As RadioButton
    Private Shared WithEvents RbWord As RadioButton
    Private Shared WithEvents RbPDF As RadioButton
    Private Shared WithEvents SsDestinoImpresion As StatusStrip
    Private Shared WithEvents SlEsc As ToolStripStatusLabel
    Private Shared WithEvents SlImpresora As ToolStripStatusLabel
    Private Shared WithEvents SleMail As ToolStripStatusLabel
    Private Shared WithEvents SlCarpeta As ToolStripStatusLabel
    Private Shared WithEvents SlBuscarCarpeta As ToolStripStatusLabel
    Private Shared WithEvents CbPrevisualizar As CheckBox
    Private Shared WithEvents SlPrevisualizar As ToolStripStatusLabel

    Private Shared ParametrosDestinoImpresion As String = String.Empty

#End Region

#Region "Constructor de la clase"
    Private Sub New()

        Ventana = New Form()

    End Sub

#End Region

    Public Shared Function ShowDialogo(Optional ByVal ImpresoraPreSeleccionada As String = "", Optional ByVal eMailPreSeleccionado As String = "", Optional ByVal CarpetaPreSeleccionada As String = "C:\", Optional ByVal TextoAdicionalCabecera As String = "", Optional ByVal ArticuloTextoAdicionalCabecera As String = "") As String

        CrearVentana()
        InicializarVariables(ImpresoraPreSeleccionada, eMailPreSeleccionado, CarpetaPreSeleccionada, TextoAdicionalCabecera, ArticuloTextoAdicionalCabecera)

        Ventana.ActiveControl = BtImpresoraLocal
        BtImpresoraLocal.Focus()
        Ventana.ShowDialog()

        Try
            If Componentes IsNot Nothing Then
                Componentes.Dispose()
            End If
        Finally
            Ventana.Dispose() : Ventana = Nothing
        End Try

        Return ParametrosDestinoImpresion

    End Function

    Private Shared Sub InicializarVariables(Optional ByVal ImpresoraPreSeleccionada As String = "", Optional ByVal eMailPreSeleccionado As String = "", Optional ByVal CarpetaPreSeleccionada As String = "C:\", Optional ByVal TextoAdicionalCabecera As String = "", Optional ByVal ArticuloTextoAdicionalCabecera As String = "")

        For Each PName As String In System.Drawing.Printing.PrinterSettings.InstalledPrinters
            CbImpresoraDestino.Items.Add(PName)
        Next

        If CarpetaPreSeleccionada = String.Empty Then
            CarpetaPreSeleccionada = "C:\"
        End If

        If TextoAdicionalCabecera <> String.Empty Then
            RtbTexto.Text = RtbTexto.Text + " " + IIf(ArticuloTextoAdicionalCabecera = String.Empty, String.Empty, ArticuloTextoAdicionalCabecera + " ") + TextoAdicionalCabecera + ";"
        Else
            RtbTexto.Text = RtbTexto.Text + ";"
        End If

        RtbResaltarTexto(RtbTexto, "Seleccionar", True, New Font("Microsoft Sans Serif", 12.0!, FontStyle.Bold), SystemColors.WindowText, Color.LightGray)
        RtbResaltarTexto(RtbTexto, TextoAdicionalCabecera, True, New Font("Microsoft Sans Serif", 12.0!, FontStyle.Bold), Color.FromArgb(0, 0, 192), Color.LightGray)
        RtbResaltarTexto(RtbTexto, "Dirección de Envio", True, New Font("Microsoft Sans Serif", 12.0!, FontStyle.Italic), SystemColors.WindowText, Color.LightGray)

        CbImpresoraDestino.SelectedIndex = CbImpresoraDestino.Items.IndexOf(ImpresoraPreSeleccionada)
        If CbImpresoraDestino.SelectedIndex = -1 Then
            Dim ImpresoraPredeterminada As System.Drawing.Printing.PrintDocument = New System.Drawing.Printing.PrintDocument()
            CbImpresoraDestino.SelectedIndex = CbImpresoraDestino.Items.IndexOf(ImpresoraPredeterminada.PrinterSettings.PrinterName)
        End If
        EMailDestino.Text = eMailPreSeleccionado
        CarpetaDestino.Text = CarpetaPreSeleccionada
        RbPDF.Checked = True
        CbPrevisualizar.Checked = False

    End Sub

    Private Shared Sub MiDestinoImpresion_KeyDown(sender As Object, e As KeyEventArgs) Handles Ventana.KeyDown 'La propiedad Ventana.KeyPreviw ha de estar a TRUE

        Select Case e.KeyData
            Case Keys.Escape
                e.Handled = True    'Evento Controlado
                BtCancelar_Click(BtCancelar, EventArgs.Empty)
            Case Keys.Control + Keys.I
                e.Handled = True    'Evento Controlado
                BtImpresoraLocal_Click(BtImpresoraLocal, EventArgs.Empty)
            Case Keys.Control + Keys.E
                e.Handled = True    'Evento Controlado
                BtEMail_Click(BtEMail, EventArgs.Empty)
            Case Keys.Control + Keys.C
                e.Handled = True    'Evento Controlado
                BtCarpeta_Click(BtCarpeta, EventArgs.Empty)
            Case Keys.Control + Keys.V
                e.Handled = True    'Evento Controlado
                CbPrevisualizar.Checked = Not CbPrevisualizar.Checked
            Case Else
                e.Handled = False   'Asegura que el proceso es pasado al control que tiene el foco
        End Select

    End Sub

    'Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean
    '
    'Select Case keyData
    'Case Keys.Escape
    'BtCancelar_Click(BtCancelar, EventArgs.Empty)
    'Case Keys.Control + Keys.I
    'BtImpresoraLocal_Click(BtImpresoraLocal, EventArgs.Empty)
    'Case Keys.Control + Keys.E
    'BtEMail_Click(BtEMail, EventArgs.Empty)
    'Case Keys.Control + Keys.C
    'BtCarpeta_Click(BtCarpeta, EventArgs.Empty)
    'Case Keys.Control + Keys.V
    'CbPrevisualizar.Checked = Not CbPrevisualizar.Checked
    'Case Else
    'Return MyBase.ProcessCmdKey(msg, keyData)
    'End Select
    '
    'Return True
    '
    'End Function

    Private Shared Sub BtImpresoraLocal_Click(sender As Object, e As EventArgs) Handles BtImpresoraLocal.Click, SlImpresora.Click

        ParametrosDestinoImpresion = DialogResult.OK.ToString + "|I|" + CbImpresoraDestino.Text + "|" + CbPrevisualizar.Checked.ToString
        Ventana.DialogResult = DialogResult.OK

    End Sub

    Private Shared Sub BtEMail_Click(sender As Object, e As EventArgs) Handles BtEMail.Click, SleMail.Click

        If EMailDestino.Text.Trim = String.Empty Then
            MiMessageBox.ShowWinMessage("Especificar una Dirección eMail ...", "Destino Impresión eMail", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            EMailDestino.Focus()
        ElseIf Not ValidarEmail(EMailDestino.Text.Trim) Then
            MiMessageBox.ShowWinMessage("La Dirección eMail especificada no es una Dirección eMail Válida." + vbCrLf + vbCrLf + "Especificar una Dirección eMail válida ...", "Destino Impresión eMail", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            EMailDestino.Focus()
        Else
            ParametrosDestinoImpresion = DialogResult.OK.ToString + "|E|" + EMailDestino.Text.Trim + "|" + CbPrevisualizar.Checked.ToString
            Ventana.DialogResult = DialogResult.OK
        End If

    End Sub

    Private Shared Sub BtCarpeta_Click(sender As Object, e As EventArgs) Handles BtCarpeta.Click, SlCarpeta.Click

        If CarpetaDestino.Text.Trim = String.Empty Then
            MiMessageBox.ShowWinMessage("Especificar una Carpeta Local ...", "Destino Carpeta Local", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            CarpetaDestino.Focus()
        ElseIf Not Directory.Exists(CarpetaDestino.Text.Trim) Then
            MiMessageBox.ShowWinMessage("La Carpeta Local especificada NO existe." + vbCrLf + vbCrLf + "Especificar una Carpeta Local existente ...", "Destino Carpeta Local", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            CarpetaDestino.Focus()
        Else
            ParametrosDestinoImpresion = DialogResult.OK.ToString + "|D|" + CarpetaDestino.Text.Trim + "|" + IIf(RbPDF.Checked, "PDF", IIf(RbWord.Checked, "WORD", IIf(RbExcel.Checked, "EXCEL", "XXX"))) + "|" + CbPrevisualizar.Checked.ToString
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

        ParametrosDestinoImpresion = DialogResult.Cancel.ToString.ToUpper
        Ventana.DialogResult = DialogResult.Cancel

    End Sub

    Private Shared Sub CrearVentana()

        ObjMiDestinoImpresion = Nothing
        ObjMiDestinoImpresion = New MiDestinoImpresion()

        Componentes = New System.ComponentModel.Container()

        FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog()
        EMailDestino = New MisClasesFuncionesC.MiTextBox(Componentes)
        BtCancelar = New System.Windows.Forms.Button()
        CarpetaDestino = New MisClasesFuncionesC.MiTextBox(Componentes)
        GbFormatoCarpetaLocal = New System.Windows.Forms.GroupBox()
        RbExcel = New System.Windows.Forms.RadioButton()
        RbWord = New System.Windows.Forms.RadioButton()
        RbPDF = New System.Windows.Forms.RadioButton()
        PictureBox3 = New System.Windows.Forms.PictureBox()
        BtCarpeta = New System.Windows.Forms.Button()
        PictureBox2 = New System.Windows.Forms.PictureBox()
        PictureBox1 = New System.Windows.Forms.PictureBox()
        CbImpresoraDestino = New System.Windows.Forms.ComboBox()
        Label2 = New System.Windows.Forms.Label()
        BtEMail = New System.Windows.Forms.Button()
        BtImpresoraLocal = New System.Windows.Forms.Button()
        RtbTexto = New System.Windows.Forms.RichTextBox()
        Label1 = New System.Windows.Forms.Label()
        Label3 = New System.Windows.Forms.Label()
        SsDestinoImpresion = New System.Windows.Forms.StatusStrip()
        SlEsc = New System.Windows.Forms.ToolStripStatusLabel()
        SlImpresora = New System.Windows.Forms.ToolStripStatusLabel()
        SleMail = New System.Windows.Forms.ToolStripStatusLabel()
        SlCarpeta = New System.Windows.Forms.ToolStripStatusLabel()
        SlBuscarCarpeta = New System.Windows.Forms.ToolStripStatusLabel()
        CbPrevisualizar = New System.Windows.Forms.CheckBox()
        SlPrevisualizar = New System.Windows.Forms.ToolStripStatusLabel()
        GbFormatoCarpetaLocal.SuspendLayout()
        CType(PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()

        SsDestinoImpresion.SuspendLayout()
        Ventana.SuspendLayout()
        '
        'FolderBrowserDialog
        '
        FolderBrowserDialog.Description = "Carpeta Destino para Ficheros Exportados (PDF, Word, ó Excel)"
        '
        'EMailDestino
        '
        EMailDestino.BackColor = System.Drawing.SystemColors.Window
        EMailDestino.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        EMailDestino.Formato = ""
        EMailDestino.Location = New System.Drawing.Point(319, 68)
        EMailDestino.Margin = New System.Windows.Forms.Padding(4)
        EMailDestino.Name = "EMailDestino"
        EMailDestino.NumericoSolo = False
        EMailDestino.NumeroSolo = False
        EMailDestino.SeleccionarTodoCuandoClick = True
        EMailDestino.Size = New System.Drawing.Size(300, 22)
        EMailDestino.TabIndex = 4
        EMailDestino.TextoVacioNumPermitido = True
        EMailDestino.UsarEnterComoTab = True
        '
        'BtCancelar
        '
        BtCancelar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        BtCancelar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        BtCancelar.Location = New System.Drawing.Point(530, 138)
        BtCancelar.Name = "BtCancelar"
        BtCancelar.Size = New System.Drawing.Size(90, 23)
        BtCancelar.TabIndex = 8
        BtCancelar.Text = "Cancelar"
        BtCancelar.UseVisualStyleBackColor = True
        '
        'CarpetaDestino
        '
        CarpetaDestino.BackColor = System.Drawing.SystemColors.Window
        CarpetaDestino.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        CarpetaDestino.Formato = ""
        CarpetaDestino.Location = New System.Drawing.Point(319, 103)
        CarpetaDestino.Margin = New System.Windows.Forms.Padding(4)
        CarpetaDestino.Name = "CarpetaDestino"
        CarpetaDestino.NumericoSolo = False
        CarpetaDestino.NumeroSolo = False
        CarpetaDestino.SeleccionarTodoCuandoClick = True
        CarpetaDestino.Size = New System.Drawing.Size(300, 22)
        CarpetaDestino.TabIndex = 7
        CarpetaDestino.TextoVacioNumPermitido = True
        CarpetaDestino.UsarEnterComoTab = True
        '
        'GbFormatoCarpetaLocal
        '
        GbFormatoCarpetaLocal.Controls.Add(RbExcel)
        GbFormatoCarpetaLocal.Controls.Add(RbWord)
        GbFormatoCarpetaLocal.Controls.Add(RbPDF)
        GbFormatoCarpetaLocal.Location = New System.Drawing.Point(191, 92)
        GbFormatoCarpetaLocal.Name = "GbFormatoCarpetaLocal"
        GbFormatoCarpetaLocal.Size = New System.Drawing.Size(65, 74)
        GbFormatoCarpetaLocal.TabIndex = 6
        GbFormatoCarpetaLocal.TabStop = False
        '
        'RbExcel
        '
        RbExcel.AutoSize = True
        RbExcel.Location = New System.Drawing.Point(3, 47)
        RbExcel.Name = "RbExcel"
        RbExcel.Size = New System.Drawing.Size(59, 20)
        RbExcel.TabIndex = 3
        RbExcel.TabStop = True
        RbExcel.Text = "Excel"
        RbExcel.UseVisualStyleBackColor = True
        '
        'RbWord
        '
        RbWord.AutoSize = True
        RbWord.Location = New System.Drawing.Point(3, 27)
        RbWord.Name = "RbWord"
        RbWord.Size = New System.Drawing.Size(59, 20)
        RbWord.TabIndex = 2
        RbWord.TabStop = True
        RbWord.Text = "Word"
        RbWord.UseVisualStyleBackColor = True
        '
        'RbPDF
        '
        RbPDF.AutoSize = True
        RbPDF.Location = New System.Drawing.Point(3, 7)
        RbPDF.Name = "RbPDF"
        RbPDF.Size = New System.Drawing.Size(53, 20)
        RbPDF.TabIndex = 1
        RbPDF.TabStop = True
        RbPDF.Text = "PDF"
        RbPDF.UseVisualStyleBackColor = True
        '
        'PictureBox3
        '
        PictureBox3.Image = Global.MisClasesFuncionesVB.My.Resources.Resources.CarpetaDisco_1_32
        PictureBox3.Location = New System.Drawing.Point(4, 98)
        PictureBox3.Name = "PictureBox3"
        PictureBox3.Size = New System.Drawing.Size(32, 32)
        PictureBox3.TabIndex = 25
        PictureBox3.TabStop = False
        '
        'BtCarpeta
        '
        BtCarpeta.BackColor = System.Drawing.Color.Beige
        BtCarpeta.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        BtCarpeta.Location = New System.Drawing.Point(40, 98)
        BtCarpeta.Name = "BtCarpeta"
        BtCarpeta.Size = New System.Drawing.Size(150, 32)
        BtCarpeta.TabIndex = 5
        BtCarpeta.Text = "Carpeta Local"
        BtCarpeta.UseVisualStyleBackColor = False
        '
        'PictureBox2
        '
        PictureBox2.Image = Global.MisClasesFuncionesVB.My.Resources.Resources.email_1_32
        PictureBox2.Location = New System.Drawing.Point(4, 63)
        PictureBox2.Name = "PictureBox2"
        PictureBox2.Size = New System.Drawing.Size(32, 32)
        PictureBox2.TabIndex = 24
        PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        PictureBox1.Image = Global.MisClasesFuncionesVB.My.Resources.Resources.Impresora_4_32
        PictureBox1.Location = New System.Drawing.Point(4, 28)
        PictureBox1.Name = "PictureBox1"
        PictureBox1.Size = New System.Drawing.Size(32, 32)
        PictureBox1.TabIndex = 22
        PictureBox1.TabStop = False
        '
        'CbImpresoraDestino
        '
        CbImpresoraDestino.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        CbImpresoraDestino.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        CbImpresoraDestino.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        CbImpresoraDestino.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        CbImpresoraDestino.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        CbImpresoraDestino.Location = New System.Drawing.Point(319, 32)
        CbImpresoraDestino.Margin = New System.Windows.Forms.Padding(4)
        CbImpresoraDestino.Name = "CbImpresoraDestino"
        CbImpresoraDestino.Size = New System.Drawing.Size(300, 24)
        CbImpresoraDestino.Sorted = True
        CbImpresoraDestino.TabIndex = 2
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Location = New System.Drawing.Point(208, 34)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(108, 16)
        Label2.TabIndex = 0
        Label2.Text = "Impresora Local:"
        Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'BtEMail
        '
        BtEMail.BackColor = System.Drawing.Color.AliceBlue
        BtEMail.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        BtEMail.Location = New System.Drawing.Point(40, 63)
        BtEMail.Name = "BtEMail"
        BtEMail.Size = New System.Drawing.Size(150, 32)
        BtEMail.TabIndex = 3
        BtEMail.Text = "eMail"
        BtEMail.UseVisualStyleBackColor = False
        '
        'BtImpresoraLocal
        '
        BtImpresoraLocal.BackColor = System.Drawing.Color.Linen
        BtImpresoraLocal.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        BtImpresoraLocal.Location = New System.Drawing.Point(40, 28)
        BtImpresoraLocal.Name = "BtImpresoraLocal"
        BtImpresoraLocal.Size = New System.Drawing.Size(150, 32)
        BtImpresoraLocal.TabIndex = 1
        BtImpresoraLocal.Text = "Impresora Local"
        BtImpresoraLocal.UseVisualStyleBackColor = False
        '
        'RtbTexto
        '
        RtbTexto.BackColor = System.Drawing.Color.LightGray
        RtbTexto.BorderStyle = System.Windows.Forms.BorderStyle.None
        RtbTexto.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        RtbTexto.Location = New System.Drawing.Point(4, 4)
        RtbTexto.Multiline = False
        RtbTexto.Name = "RtbTexto"
        RtbTexto.ReadOnly = True
        RtbTexto.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None
        RtbTexto.Size = New System.Drawing.Size(615, 24)
        RtbTexto.TabIndex = 0
        RtbTexto.TabStop = False
        RtbTexto.Text = "Seleccionar Dirección de Envio"
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(213, 70)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(104, 16)
        Label1.TabIndex = 0
        Label1.Text = "Dirección eMail:"
        Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Location = New System.Drawing.Point(257, 105)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(59, 16)
        Label3.TabIndex = 0
        Label3.Text = "Carpeta:"
        Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'SsDestinoImpresion
        '
        SsDestinoImpresion.BackColor = System.Drawing.SystemColors.ControlLight
        SsDestinoImpresion.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        SsDestinoImpresion.GripStyle = System.Windows.Forms.ToolStripGripStyle.Visible
        SsDestinoImpresion.Items.AddRange(New System.Windows.Forms.ToolStripItem() {SlEsc, SlImpresora, SleMail, SlCarpeta, SlPrevisualizar, SlBuscarCarpeta})
        SsDestinoImpresion.Location = New System.Drawing.Point(0, 166)
        SsDestinoImpresion.Name = "SsDestinoImpresion"
        SsDestinoImpresion.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        SsDestinoImpresion.Size = New System.Drawing.Size(623, 22)
        SsDestinoImpresion.TabIndex = 0
        SsDestinoImpresion.Text = "StatusStrip1"
        '
        'SlEsc
        '
        SlEsc.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        SlEsc.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        SlEsc.Name = "SlEsc"
        SlEsc.Size = New System.Drawing.Size(88, 17)
        SlEsc.Text = "Esc=Cancelar"
        '
        'SlImpresora
        '
        SlImpresora.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        SlImpresora.ForeColor = System.Drawing.SystemColors.ControlText
        SlImpresora.Name = "SlImpresora"
        SlImpresora.Size = New System.Drawing.Size(92, 17)
        SlImpresora.Text = "^I=Impresora"
        '
        'SleMail
        '
        SleMail.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        SleMail.ForeColor = System.Drawing.SystemColors.ControlText
        SleMail.Name = "SleMail"
        SleMail.Size = New System.Drawing.Size(67, 17)
        SleMail.Text = "^E=eMail"
        '
        'SlCarpeta
        '
        SlCarpeta.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        SlCarpeta.ForeColor = System.Drawing.SystemColors.ControlText
        SlCarpeta.Name = "SlCarpeta"
        SlCarpeta.Size = New System.Drawing.Size(81, 17)
        SlCarpeta.Text = "^C=Carpeta"
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
        'CbPrevisualizar
        '
        CbPrevisualizar.AutoSize = True
        CbPrevisualizar.Location = New System.Drawing.Point(41, 137)
        CbPrevisualizar.Name = "CbPrevisualizar"
        CbPrevisualizar.Size = New System.Drawing.Size(104, 20)
        CbPrevisualizar.TabIndex = 0
        CbPrevisualizar.TabStop = False
        CbPrevisualizar.Text = "Previsualizar"
        CbPrevisualizar.UseVisualStyleBackColor = True
        '
        'SlPrevisualizar
        '
        SlPrevisualizar.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        SlPrevisualizar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        SlPrevisualizar.Name = "SlPrevisualizar"
        SlPrevisualizar.Size = New System.Drawing.Size(115, 17)
        SlPrevisualizar.Text = "^V=PreVisualizar"
        '
        'MiDestinoImpresion
        '
        Ventana.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Ventana.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Ventana.BackColor = System.Drawing.Color.LightGray
        Ventana.ClientSize = New System.Drawing.Size(623, 188)
        Ventana.ControlBox = False
        Ventana.Controls.Add(CbPrevisualizar)
        Ventana.Controls.Add(SsDestinoImpresion)
        Ventana.Controls.Add(EMailDestino)
        Ventana.Controls.Add(BtCancelar)
        Ventana.Controls.Add(Label3)
        Ventana.Controls.Add(CarpetaDestino)
        Ventana.Controls.Add(GbFormatoCarpetaLocal)
        Ventana.Controls.Add(PictureBox3)
        Ventana.Controls.Add(BtCarpeta)
        Ventana.Controls.Add(PictureBox2)
        Ventana.Controls.Add(PictureBox1)
        Ventana.Controls.Add(CbImpresoraDestino)
        Ventana.Controls.Add(Label2)
        Ventana.Controls.Add(BtEMail)
        Ventana.Controls.Add(BtImpresoraLocal)
        Ventana.Controls.Add(RtbTexto)
        Ventana.Controls.Add(Label1)
        Ventana.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Ventana.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Ventana.Margin = New System.Windows.Forms.Padding(4)
        Ventana.Name = "MiDestinoImpresion"
        Ventana.ShowIcon = False
        Ventana.ShowInTaskbar = False
        Ventana.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Ventana.KeyPreview = True
        GbFormatoCarpetaLocal.ResumeLayout(False)
        GbFormatoCarpetaLocal.PerformLayout()
        CType(PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()

        SsDestinoImpresion.ResumeLayout(False)
        SsDestinoImpresion.PerformLayout()
        Ventana.ResumeLayout(False)
        Ventana.PerformLayout()


    End Sub

End Class
