Public NotInheritable Class MiInputBox

#Region "Variables para la Clase ..."

    Private Shared ObjMiInputBox As MiInputBox
    Private Shared WithEvents Ventana As Form

    Private Shared Componentes As ComponentModel.Container

    Private Shared WithEvents BtACEPTAR As Button
    Private Shared WithEvents BtCANCELAR As Button
    Private Shared WithEvents SsEstado As StatusStrip
    Private Shared WithEvents SlEsc As System.Windows.Forms.ToolStripStatusLabel
    Private Shared WithEvents SlF1 As System.Windows.Forms.ToolStripStatusLabel
    Private Shared WithEvents SlSepara1 As System.Windows.Forms.ToolStripStatusLabel
    Private Shared WithEvents SlMensaje As System.Windows.Forms.ToolStripStatusLabel
    Private Shared WithEvents TbPrompt As System.Windows.Forms.TextBox
    Private Shared WithEvents TbNota As System.Windows.Forms.TextBox
    Private Shared WithEvents TbTexto As System.Windows.Forms.TextBox
    Private Shared WithEvents Timer As System.Windows.Forms.Timer

    Private Shared SwF1Objeto As Boolean = False
    Private Shared WithEvents ObF1Objeto As Object
    Private Shared SwMsgProcesoCancelado As Boolean = False
    Private Shared SwEntradaObligatoria As Boolean = False
    Private Shared TipoDato0 As String

    Public Const Cancelado As String = "#@$Ñ5jy&j<"             'Constante devuelta por la funcion, para indicar se se pulsó el botón CANCELAR
    Public Const TipoDatoNumerico As String = "N"               'Numerico
    Public Const TipoDatoNumericoDecimal As String = "ND"       'Numerico con Decimales
    Public Const TipoDatoString As String = "S"                 'String

#End Region

#Region "Constructor de la clase"
    Private Sub New()
        Ventana = New Form()
        AddHandler Ventana.Shown, AddressOf ParaHacerUnaVezMostrado
    End Sub

    Private Sub ParaHacerUnaVezMostrado(sender As Object, e As EventArgs)

        Cursor.Position = Ventana.Location + Ventana.ActiveControl.Location + New Point(40, 12)

    End Sub
#End Region

    Public Shared Function ShowDialogo(ByVal Prompt As String, Optional ByVal Nota As String = "", Optional ByVal Titulo As String = "", Optional ByVal ValorDefecto As String = "", Optional ByVal TipoDato As String = "S", Optional EntradaObligatoria As Boolean = False, Optional ByVal MsgProcesoCancelado As Boolean = True, Optional F1Objecto As Object = Nothing, Optional ByVal F1StatusLabel As String = "") As String

        CrearVentana()
        InicializarVariables(Prompt, Nota, Titulo, ValorDefecto, TipoDato, EntradaObligatoria, MsgProcesoCancelado, F1Objecto, F1StatusLabel)

        Dim Valordevuelto As String = String.Empty

        TbTexto.Focus()
        Ventana.ShowDialog()
        If Ventana.DialogResult = Windows.Forms.DialogResult.OK Then
            Valordevuelto = TbTexto.Text.Trim
        Else
            Valordevuelto = Cancelado
        End If

        Try
            If Componentes IsNot Nothing Then
                Componentes.Dispose()
            End If
        Finally
            Ventana.Dispose() : Ventana = Nothing
        End Try

        Return Valordevuelto

    End Function

    Private Shared Sub InicializarVariables(ByVal Prompt As String, Optional ByVal Nota As String = "", Optional ByVal Titulo As String = "", Optional ByVal ValorDefecto As String = "", Optional ByVal TipoDato As String = "S", Optional EntradaObligatoria As Boolean = False, Optional ByVal MsgProcesoCancelado As Boolean = True, Optional F1Objecto As Object = Nothing, Optional ByVal F1StatusLabel As String = "")

        If F1Objecto Is Nothing Then
            SwF1Objeto = False
            SlF1.Text = String.Empty
            SlF1.Visible = False
        Else
            ObF1Objeto = F1Objecto
            SlF1.Text = F1StatusLabel
            SwF1Objeto = True
            SlF1.Visible = True
            'Nota: El objeto ha de dejar un valor en una Variable Publica tipo String con Nombre "DialogValor"
        End If

        TbPrompt.Text = Prompt
        TbNota.Text = Nota
        If Titulo Is Nothing Then
            Ventana.Text = String.Empty
        Else
            Ventana.Text = Titulo.Trim
        End If
        If ValorDefecto Is Nothing Then
            TbTexto.Text = String.Empty
        Else
            TbTexto.Text = ValorDefecto.Trim
        End If
        TipoDato0 = TipoDato
        SwMsgProcesoCancelado = MsgProcesoCancelado
        SwEntradaObligatoria = EntradaObligatoria

    End Sub

    Private Shared Sub BtACEPTAR_Click(sender As Object, e As EventArgs) Handles BtACEPTAR.Click

        If SwEntradaObligatoria Then
            If TbTexto.Text.Trim = String.Empty Then
                SlMensaje.ForeColor = Color.FromArgb(192, 0, 0)
                SlMensaje.Text = "Entrada OBLIGATORIA"
                Beep()
                Timer.Enabled = True
                TbTexto.Focus()
            Else
                Ventana.DialogResult = Windows.Forms.DialogResult.OK
            End If
        Else
            Ventana.DialogResult = Windows.Forms.DialogResult.OK

        End If
    End Sub

    Private Shared Sub BtCANCELAR_Click(sender As Object, e As EventArgs) Handles BtCANCELAR.Click

        If SwMsgProcesoCancelado Then
            MiMessageBox.ShowWinMessage("Proceso CANCELADO ..." + vbCrLf + vbCrLf + "... A Petición del Usuario.", "Entrada de Datos", MessageBoxIcon.Information, MessageBoxButtons.OK)
        End If
        Ventana.DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Shared Sub TbTexto_KeyDown(sender As Object, e As KeyEventArgs) Handles TbTexto.KeyDown

        ' Si no numerico, nos vamos
        If TipoDato0 <> TipoDatoNumerico And TipoDato0 <> TipoDatoNumericoDecimal Then
            Return
        End If

        ' Borramos mensaje que pudiera haber ....
        SlMensaje.Text = String.Empty
        Timer.Enabled = False

        ' Teclas de Control/Navegacion permitidas
        Select Case e.KeyData
            Case Keys.F1
                If SwF1Objeto Then
                    ObF1Objeto.ShowDialog()
                    If ObF1Objeto.DialogResult = DialogResult.OK Then
                        TbTexto.Text = Trim(ObF1Objeto.DialogValue)
                        TbTexto.SelectionStart = 999
                    End If
                End If
                Return
            Case Keys.Control, Keys.Shift, Keys.LControlKey, Keys.LControlKey, Keys.Home, Keys.End
                Return
            Case Keys.Delete, Keys.Insert
                Return
            Case Keys.Back
                Return
            Case Keys.Left, Keys.Right, Keys.Up, Keys.Down, Keys.PageDown, Keys.PageUp
                Return
            Case Keys.Control + Keys.V
                Return
            Case Keys.Control + Keys.C
                Return
            Case Keys.Shift + Keys.Home
                Return
            Case Keys.Shift + Keys.End
                Return
            Case Keys.CapsLock, Keys.NumLock
                Return
            Case 131089, 65552 ' Control y Shift ... que parece son codigos distintos de Keys.Control y Keys.Shift
                Return
        End Select

        ' Teclas de números permitidas.
        Dim keyValue() As Integer = {48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105}

        ' Verificar si tecla permitida
        Dim TeclaPermitida As Boolean = False
        Dim swPorPunto As Boolean = False
        If TipoDato0 = TipoDatoNumericoDecimal And (e.KeyData = Keys.Decimal Or e.KeyData = 190) Then
            If InStr(TbTexto.Text, ".") = 0 Then
                TeclaPermitida = True
            Else
                TeclaPermitida = False
                swPorPunto = True
            End If
        Else
            Dim i As Integer
            For i = 0 To keyValue.Length - 1
                If keyValue(i) = e.KeyValue Then
                    TeclaPermitida = True
                    Exit For
                End If
            Next
        End If

        ' Si no es una tecla permitida, mensaje, y evitamos que se desencadene el evento KeyPress.
        If Not (TeclaPermitida) Then
            SlMensaje.ForeColor = Color.FromArgb(192, 0, 192)
            If TipoDato0 = TipoDatoNumerico Then
                SlMensaje.Text = "Solo Número Permitido"
            ElseIf TipoDato0 = TipoDatoNumericoDecimal Then
                If swPorPunto Then
                    SlMensaje.Text = "Decimal YA Introducido"
                Else
                    SlMensaje.Text = "Solo Numérico Permitido"
                End If
            Else
                SlMensaje.Text = "?"
            End If
            Beep()
            Timer.Enabled = True
            e.SuppressKeyPress = True
        End If

    End Sub

    Private Shared Sub TbTexto_Enter(sender As Object, e As EventArgs) Handles TbTexto.Enter

        TbTexto.BackColor = Color.Gold

    End Sub

    Private Shared Sub TbTexto_Leave(sender As Object, e As EventArgs) Handles TbTexto.Leave

        TbTexto.BackColor = SystemColors.Window

    End Sub

    Private Shared Sub Timer_Tick(sender As Object, e As EventArgs) Handles Timer.Tick

        If SlMensaje.Visible Then
            SlMensaje.Visible = False
        Else
            SlMensaje.Visible = True
        End If

    End Sub

    Private Shared Sub CrearVentana()

        ObjMiInputBox = Nothing
        ObjMiInputBox = New MiInputBox()

        Componentes = New System.ComponentModel.Container()

        BtACEPTAR = New System.Windows.Forms.Button()
        BtCANCELAR = New System.Windows.Forms.Button()
        SsEstado = New System.Windows.Forms.StatusStrip()
        SlEsc = New System.Windows.Forms.ToolStripStatusLabel()
        SlF1 = New System.Windows.Forms.ToolStripStatusLabel()
        SlSepara1 = New System.Windows.Forms.ToolStripStatusLabel()
        SlMensaje = New System.Windows.Forms.ToolStripStatusLabel()
        TbPrompt = New System.Windows.Forms.TextBox()
        TbNota = New System.Windows.Forms.TextBox()
        TbTexto = New System.Windows.Forms.TextBox()
        Timer = New System.Windows.Forms.Timer(Componentes)

        SsEstado.SuspendLayout()
        Ventana.SuspendLayout()
        '
        'BtACEPTAR
        '
        BtACEPTAR.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        BtACEPTAR.ForeColor = System.Drawing.Color.Green
        BtACEPTAR.Location = New System.Drawing.Point(320, 7)
        BtACEPTAR.Name = "BtACEPTAR"
        BtACEPTAR.Size = New System.Drawing.Size(90, 42)
        BtACEPTAR.TabIndex = 2
        BtACEPTAR.Text = "Aceptar"
        BtACEPTAR.UseVisualStyleBackColor = True
        '
        'BtCANCELAR
        '
        BtCANCELAR.DialogResult = System.Windows.Forms.DialogResult.Cancel
        BtCANCELAR.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        BtCANCELAR.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        BtCANCELAR.Location = New System.Drawing.Point(320, 54)
        BtCANCELAR.Name = "BtCANCELAR"
        BtCANCELAR.Size = New System.Drawing.Size(90, 42)
        BtCANCELAR.TabIndex = 3
        BtCANCELAR.Text = "Cancelar"
        BtCANCELAR.UseVisualStyleBackColor = True
        '
        'SsEstado
        '
        SsEstado.Items.AddRange(New System.Windows.Forms.ToolStripItem() {SlEsc, SlF1, SlSepara1, SlMensaje})
        SsEstado.Location = New System.Drawing.Point(0, 101)
        SsEstado.Name = "SsEstado"
        SsEstado.Size = New System.Drawing.Size(415, 24)
        SsEstado.TabIndex = 0
        SsEstado.Text = "StatusStrip1"
        '
        'SlEsc
        '
        SlEsc.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        SlEsc.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        SlEsc.Name = "SlEsc"
        SlEsc.Size = New System.Drawing.Size(83, 19)
        SlEsc.Text = "Esc=Cancel"
        '
        'SlF1
        '
        SlF1.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        SlF1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        SlF1.Name = "SlF1"
        SlF1.Size = New System.Drawing.Size(58, 19)
        SlF1.Text = "F1=xxx"
        '
        'SlSepara1
        '
        SlSepara1.Name = "SlSepara1"
        SlSepara1.Size = New System.Drawing.Size(10, 19)
        SlSepara1.Text = "|"
        '
        'SlMensaje
        '
        SlMensaje.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        SlMensaje.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        SlMensaje.Name = "SlMensaje"
        SlMensaje.Size = New System.Drawing.Size(0, 19)
        '
        'TbPrompt
        '
        TbPrompt.BackColor = System.Drawing.SystemColors.Control
        TbPrompt.BorderStyle = System.Windows.Forms.BorderStyle.None
        TbPrompt.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        TbPrompt.Location = New System.Drawing.Point(5, 5)
        TbPrompt.Multiline = True
        TbPrompt.Name = "TbPrompt"
        TbPrompt.Size = New System.Drawing.Size(310, 50)
        TbPrompt.TabIndex = 0
        TbPrompt.TabStop = False
        '
        'TbNota
        '
        TbNota.BackColor = System.Drawing.SystemColors.Control
        TbNota.BorderStyle = System.Windows.Forms.BorderStyle.None
        TbNota.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        TbNota.Location = New System.Drawing.Point(5, 55)
        TbNota.Name = "TbNota"
        TbNota.Size = New System.Drawing.Size(310, 18)
        TbNota.TabIndex = 0
        TbNota.TabStop = False
        '
        'TbTexto
        '
        TbTexto.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold)
        TbTexto.Location = New System.Drawing.Point(5, 72)
        TbTexto.Name = "TbTexto"
        TbTexto.Size = New System.Drawing.Size(310, 23)
        TbTexto.TabIndex = 1
        '
        'Timer
        '
        Timer.Interval = 1000
        '
        'MiInputBox
        '
        Ventana.AcceptButton = BtACEPTAR
        Ventana.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Ventana.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Ventana.CancelButton = BtCANCELAR
        Ventana.ClientSize = New System.Drawing.Size(415, 125)
        Ventana.ControlBox = False
        Ventana.Controls.Add(TbTexto)
        Ventana.Controls.Add(TbNota)
        Ventana.Controls.Add(TbPrompt)
        Ventana.Controls.Add(SsEstado)
        Ventana.Controls.Add(BtCANCELAR)
        Ventana.Controls.Add(BtACEPTAR)
        Ventana.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Ventana.Name = "MiInputBoxExt"
        Ventana.ShowInTaskbar = False
        Ventana.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Ventana.Text = "Input Box"

        SsEstado.ResumeLayout(False)
        SsEstado.PerformLayout()
        Ventana.ResumeLayout(False)
        Ventana.PerformLayout()

    End Sub

End Class
