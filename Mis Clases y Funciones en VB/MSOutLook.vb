Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Public NotInheritable Class MSOutlook

    Implements IDisposable

    Private m_streams As IList(Of Stream)
    Private WithEvents AplicacionOutlook As Global.Microsoft.Office.Interop.Outlook.Application
    Private ItemEmail As Microsoft.Office.Interop.Outlook.MailItem
    'Private AttachsEmail As Outlook.Attachments
    'Private AttachEmail As Outlook.Attachment
    Private InspectorOutlook As Global.Microsoft.Office.Interop.Outlook.Inspector

    Public Sub EnviarEMail(ByVal Path_FileName As String, ByVal EmailAsunto As String, ByVal EmailCuerpo As String, Optional ByVal EmailDestino As String = "", Optional ByVal EditarEmailAntesEnviar As Boolean = False, Optional ByVal HTMLBody As Boolean = False)

        If Not EstaOutlookInstalado() Then
            MiMessageBox.ShowWinMessage("Parece que OUTLOOK NO está INSTALADO en el Equipo." + vbCr + vbCr + "eMail NO puede ser ENVIADO.", "Enviar Email", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        Dim HayDestinatario As Boolean = True   ' De entrada, se supone que se ha especidficado destinatario

        If String.IsNullOrWhiteSpace(EmailDestino) Then
            Using GetEMailDestino As GetEMailDestino = New GetEMailDestino(, EditarEmailAntesEnviar)
                GetEMailDestino.ShowDialog()
                If GetEMailDestino.DialogResult = DialogResult.OK Then
                    EmailDestino = GetEMailDestino.EMailDestino
                    EditarEmailAntesEnviar = GetEMailDestino.EditarEmailAntesEnviar
                Else
                    MiMessageBox.ShowWinMessage("Envío eMail CANCELADO ..." + vbCrLf + vbCrLf + "... A petición del Usuario.", "Enviar Email", MsgBoxStyle.Information, MsgBoxStyle.OkOnly)
                    DisposeEmail(ItemEmail, True)
                    Exit Sub
                End If
            End Using
        End If

        Try
            AplicacionOutlook = DirectCast(Marshal.GetActiveObject("Outlook.Application"), Outlook.Application) 'Hay un proceso OUTLOOK en ejecucion. Se utiliza el metodo GetActiveObject para obtener el proceso y lanzarlo como una Aplicacion objecto
        Catch Ex As System.Runtime.InteropServices.COMException
            MiMessageBox.ShowWinMessage("OUTLOOK está CERRADO...," + vbCrLf + vbCrLf + "El eMail quedará depositado en la Bandeja de Salida, ..." + vbCrLf + vbCrLf + "... el mismo será enviado una vez Outlook sea abierto.", "Enviar Email", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            AplicacionOutlook = New Outlook.Application ' Creo una nueva instancia de Aplicacion Outlook
        End Try

        Try
            ItemEmail = AplicacionOutlook.CreateItem(Outlook.OlItemType.olMailItem)
            InspectorOutlook = ItemEmail.GetInspector()

            ItemEmail.Subject = EmailAsunto
            If HTMLBody Then
                ItemEmail.HTMLBody = EmailCuerpo
            Else
                ItemEmail.Body = EmailCuerpo
            End If
            ItemEmail.To = EmailDestino
            ItemEmail.CC = String.Empty
            ItemEmail.BCC = String.Empty

            Dim AttachsEmail As Outlook.Attachments = ItemEmail.Attachments
            Dim AttachEmail As Outlook.Attachment = AttachsEmail.Add(Path_FileName, , ItemEmail.Body.Length + 1, Path.GetFileName(Path_FileName))

            If String.IsNullOrWhiteSpace(EmailDestino) AndAlso Not EditarEmailAntesEnviar Then
                HayDestinatario = False
                Dim ListasDirecciones As Outlook.AddressLists
                Dim CarpetaContactos As Outlook.Folder = CType(AplicacionOutlook.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts), Outlook.Folder)
                ListasDirecciones = AplicacionOutlook.Session.AddressLists
                For Each ListaDirecciones As Outlook.AddressList In ListasDirecciones
                    Dim CarpetaTest As Outlook.Folder = CType(ListaDirecciones.GetContactsFolder(), Outlook.Folder)
                    If Not (CarpetaTest Is Nothing) Then
                        ' Test to determine if Folder returned by GetCarpetaContactos has same EntryID as default Contacts folder.
                        If (AplicacionOutlook.Session.CompareEntryIDs(CarpetaContactos.EntryID, CarpetaTest.EntryID)) Then
                            Dim DialogoSeleccionarNombres As Outlook.SelectNamesDialog = AplicacionOutlook.Session.GetSelectNamesDialog()
                            DialogoSeleccionarNombres.InitialAddressList = ListaDirecciones
                            DialogoSeleccionarNombres.NumberOfRecipientSelectors = 3 ' Recipientes "Para" + "CC" + "CCO" 
                            DialogoSeleccionarNombres.AllowMultipleSelection = False
                            Dim SalvaInspectorOutlookleft As Integer = InspectorOutlook.Left
                            'InspectorOutlook.Left = -9999   ' Set the Inspector off screen.
                            InspectorOutlook.Activate()
                            DialogoSeleccionarNombres.Display()
                            If DialogoSeleccionarNombres.Recipients.Count > 0 Then
                                Dim i As Integer
                                For i = 1 To DialogoSeleccionarNombres.Recipients.Count
                                    If DialogoSeleccionarNombres.Recipients.Item(i).Type = 1 Then
                                        ItemEmail.To += ";" + DialogoSeleccionarNombres.Recipients.Item(i).Address
                                    ElseIf DialogoSeleccionarNombres.Recipients.Item(i).Type = 2 Then
                                        ItemEmail.CC += ";" + DialogoSeleccionarNombres.Recipients.Item(i).Address
                                    ElseIf DialogoSeleccionarNombres.Recipients.Item(i).Type = 3 Then
                                        ItemEmail.BCC += ";" + DialogoSeleccionarNombres.Recipients.Item(i).Address
                                    End If
                                Next
                                HayDestinatario = True
                            End If
                            InspectorOutlook.Left = SalvaInspectorOutlookleft
                            DialogoSeleccionarNombres = Nothing
                        End If
                    End If
                Next
                CarpetaContactos = Nothing
                ListasDirecciones = Nothing
            End If

            If EditarEmailAntesEnviar Then
                InspectorOutlook.Activate()
            Else
                If HayDestinatario Then
                    ItemEmail.Send()
                Else
                    InspectorOutlook.Close(Outlook.OlInspectorClose.olDiscard)
                    MiMessageBox.ShowWinMessage("No se ha especificado ningún Destinatario para el eMail." + vbCr + vbCr + "Envio Email CANCELADO", "Enviar Email", MsgBoxStyle.Critical, MsgBoxStyle.OkOnly)
                End If
            End If
        Catch Ex As System.Runtime.InteropServices.COMException
            MiMessageBox.ShowWinMessage("Se ha producido un ERROR en el proceso de envio del eMail;" + vbCrLf + vbCrLf + Ex.Message + vbCrLf + vbCrLf + "Envio eMail CANCELADO", "Enviar Email", MsgBoxStyle.Critical, MsgBoxStyle.OkOnly)
            DisposeEmail(ItemEmail, True)
            Exit Sub
        End Try

    End Sub

    Private Sub DisposeEmail(ByVal ItemEmail2 As Object, ByRef Cancel As Boolean) Handles AplicacionOutlook.ItemSend

        If Object.ReferenceEquals(ItemEmail2, Me.ItemEmail) Then
            ItemEmail = Nothing
            'AttachsEmail = Nothing
            'AttachEmail = Nothing
            InspectorOutlook = Nothing
            AplicacionOutlook = Nothing
        End If

    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose

        If m_streams IsNot Nothing Then
            For Each stream As Stream In m_streams
                stream.Close()
            Next
            m_streams = Nothing
        End If

    End Sub

    Shared Function EstaOutlookInstalado() As Boolean

        Return (Not CreateObject("Outlook.Application") Is Nothing)

    End Function

    Shared Function SelDesdeLibretaDirecciones() As String

        Dim EMailDestino As String = String.Empty

        Dim AplicacionOutlook As New Global.Microsoft.Office.Interop.Outlook.Application
        Dim ItemEmail As Microsoft.Office.Interop.Outlook.MailItem = AplicacionOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
        Dim InspectorOutlook As Global.Microsoft.Office.Interop.Outlook.Inspector
        Dim ListasDirecciones As Outlook.AddressLists
        Dim CarpetaContactos As Outlook.Folder = CType(AplicacionOutlook.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts), Outlook.Folder)
        InspectorOutlook = ItemEmail.GetInspector()

        ListasDirecciones = AplicacionOutlook.Session.AddressLists
        For Each ListaDirecciones As Outlook.AddressList In ListasDirecciones
            Dim CarpetaTest As Outlook.Folder = CType(ListaDirecciones.GetContactsFolder(), Outlook.Folder)
            If Not (CarpetaTest Is Nothing) Then
                ' Test to determine if Folder returned by GetContactsFolder has same EntryID as default Contacts folder.
                If (AplicacionOutlook.Session.CompareEntryIDs(CarpetaContactos.EntryID, CarpetaTest.EntryID)) Then
                    Dim DialogoSeleccionarNombres As Outlook.SelectNamesDialog = AplicacionOutlook.Session.GetSelectNamesDialog()
                    DialogoSeleccionarNombres.InitialAddressList = ListaDirecciones
                    DialogoSeleccionarNombres.NumberOfRecipientSelectors = 1 ' Recipientes "Para" (solo)
                    DialogoSeleccionarNombres.AllowMultipleSelection = True
                    DialogoSeleccionarNombres.Caption = "Libreta de Direcciones de OutLook"
                    Dim SalvaInspectorOutllokLeft As Integer = InspectorOutlook.Left
                    InspectorOutlook.Left = -9999   ' Set the Inspector off screen.
                    InspectorOutlook.Activate()
                    If DialogoSeleccionarNombres.Display() Then
                        EMailDestino = String.Empty
                        Dim i As Integer
                        For i = 1 To DialogoSeleccionarNombres.Recipients.Count
                            If i = 1 Then
                                EMailDestino = DialogoSeleccionarNombres.Recipients.Item(i).Address
                            Else
                                EMailDestino += ";" + DialogoSeleccionarNombres.Recipients.Item(i).Address
                            End If
                        Next
                    End If
                    InspectorOutlook.Left = SalvaInspectorOutllokLeft
                    InspectorOutlook.Close(Microsoft.Office.Interop.Outlook.OlInspectorClose.olDiscard)
                    DialogoSeleccionarNombres = Nothing
                End If
            End If
        Next

        CarpetaContactos = Nothing
        ListasDirecciones = Nothing
        ItemEmail = Nothing
        InspectorOutlook = Nothing
        AplicacionOutlook = Nothing

        Return EMailDestino

    End Function

End Class

Public Class GetEMailDestino

    Inherits System.Windows.Forms.Form

    Private WithEvents CbEditartEmailAntesEnviar As CheckBox
    Private WithEvents PbEmail As PictureBox
    Private WithEvents RtbTexto As RichTextBox
    Private WithEvents LbDireccionEMail As Label
    Private WithEvents BtAceptar As Button
    Private WithEvents BtCancelar As Button
    Private WithEvents TbEMailDestino As MiTextBox
    Private WithEvents SsGetEMailDestino As StatusStrip
    Private WithEvents SlEntrar As ToolStripStatusLabel
    Private WithEvents SlEsc As ToolStripStatusLabel
    Private WithEvents SlEditarEmailAntesEnviar As ToolStripStatusLabel
    Private WithEvents SlDClick As ToolStripStatusLabel
    Private WithEvents LbNota As Label

    Public EMailDestino As String
    Public EditarEmailAntesEnviar As Boolean

    Public Sub New(Optional ByVal eMailPreSeleccionado As String = "", Optional ByVal EditarEmailAntesEnviar As Boolean = False)

        MyBase.New()

        InitializarComponentes()

        TbEMailDestino.Text = eMailPreSeleccionado
        CbEditartEmailAntesEnviar.Checked = EditarEmailAntesEnviar

        RtbResaltarTexto(RtbTexto, "Introducir", True, New Font("Microsoft Sans Serif", 12.0!, FontStyle.Bold), SystemColors.WindowText, Color.LightGray)
        RtbResaltarTexto(RtbTexto, "Dirección eMail", True, New Font("Microsoft Sans Serif", 12.0!, FontStyle.Italic), Color.FromArgb(0, 0, 192), Color.LightGray)

    End Sub

    Private Sub GetEMailDestino_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ActiveControl = TbEMailDestino : TbEMailDestino.Focus()

    End Sub

    Private Sub GetEMailDestino_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown 'La propiedad Ventana.KeyPreviw ha de estar a TRUE

        Select Case e.KeyData
            Case Keys.Escape
                e.Handled = True    'Evento Controlado
                BtCancelar_Click(BtCancelar, EventArgs.Empty)
            Case Keys.Control + Keys.Enter
                e.Handled = True    'Evento Controlado
                BtAceptar_Click(BtAceptar, EventArgs.Empty)
            Case Keys.Control + Keys.E
                e.Handled = True    'Evento Controlado
                CbEditartEmailAntesEnviar.Checked = Not CbEditartEmailAntesEnviar.Checked
            Case Else
                e.Handled = False   'Asegura que el proceso es pasado al control que tiene el foco
        End Select

    End Sub

    Private Sub BtAceptar_Click(sender As Object, e As EventArgs) Handles BtAceptar.Click, SlEntrar.Click

        If Not String.IsNullOrWhiteSpace(TbEMailDestino.Text) AndAlso Not ValidarEmail(TbEMailDestino.Text.Trim) Then
            MiMessageBox.ShowWinMessage("La Dirección eMail especificada no es una Dirección eMail Válida." + vbCrLf + vbCrLf + "Especificar una Dirección eMail válida ...", "Destino Impresión eMail", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            TbEMailDestino.Focus()
            Exit Sub
        End If

        EMailDestino = TbEMailDestino.Text
        EditarEmailAntesEnviar = CbEditartEmailAntesEnviar.Checked
        DialogResult = DialogResult.OK
        Close()

    End Sub

    Private Sub BtCancelar_Click(sender As Object, e As EventArgs) Handles BtCancelar.Click, SlEsc.Click

        EMailDestino = String.Empty
        EditarEmailAntesEnviar = False
        DialogResult = DialogResult.Cancel
        Close()

    End Sub

    Private Sub TbEmailDestino_DoubleClick(sender As System.Object, e As System.EventArgs) Handles TbEMailDestino.DoubleClick

        TbEMailDestino.Text = MSOutlook.SelDesdeLibretaDirecciones()

    End Sub

    Private Sub TbEmailDestino_Enter(sender As System.Object, e As System.EventArgs) Handles TbEMailDestino.Enter

        SlDClick.Visible = True

    End Sub

    Private Sub TbEmailDestino_Leave(sender As System.Object, e As System.EventArgs) Handles TbEMailDestino.Leave

        SlDClick.Visible = False

    End Sub

    Private Sub InitializarComponentes()

        Me.components = New System.ComponentModel.Container()
        Me.CbEditartEmailAntesEnviar = New System.Windows.Forms.CheckBox()
        Me.RtbTexto = New System.Windows.Forms.RichTextBox()
        Me.LbDireccionEMail = New System.Windows.Forms.Label()
        Me.BtAceptar = New System.Windows.Forms.Button()
        Me.BtCancelar = New System.Windows.Forms.Button()
        Me.TbEMailDestino = New MisClasesFuncionesC.MiTextBox(Me.components)
        Me.PbEmail = New System.Windows.Forms.PictureBox()
        Me.SsGetEMailDestino = New System.Windows.Forms.StatusStrip()
        Me.SlEntrar = New System.Windows.Forms.ToolStripStatusLabel()
        Me.SlEsc = New System.Windows.Forms.ToolStripStatusLabel()
        Me.SlEditarEmailAntesEnviar = New System.Windows.Forms.ToolStripStatusLabel()
        Me.SlDClick = New System.Windows.Forms.ToolStripStatusLabel()
        Me.LbNota = New System.Windows.Forms.Label()
        CType(Me.PbEmail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SsGetEMailDestino.SuspendLayout()
        Me.SuspendLayout()
        '
        'CbEditartEmailAntesEnviar
        '
        Me.CbEditartEmailAntesEnviar.AutoSize = True
        Me.CbEditartEmailAntesEnviar.Location = New System.Drawing.Point(4, 75)
        Me.CbEditartEmailAntesEnviar.Name = "CbEditartEmailAntesEnviar"
        Me.CbEditartEmailAntesEnviar.Size = New System.Drawing.Size(195, 20)
        Me.CbEditartEmailAntesEnviar.TabIndex = 0
        Me.CbEditartEmailAntesEnviar.TabStop = False
        Me.CbEditartEmailAntesEnviar.Text = "Editar eMail Antes de Enviar"
        Me.CbEditartEmailAntesEnviar.UseVisualStyleBackColor = True
        '
        'RtbTexto
        '
        Me.RtbTexto.BackColor = System.Drawing.Color.LightGray
        Me.RtbTexto.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RtbTexto.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RtbTexto.Location = New System.Drawing.Point(4, 3)
        Me.RtbTexto.Margin = New System.Windows.Forms.Padding(4)
        Me.RtbTexto.Multiline = False
        Me.RtbTexto.Name = "RtbTexto"
        Me.RtbTexto.ReadOnly = True
        Me.RtbTexto.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None
        Me.RtbTexto.Size = New System.Drawing.Size(444, 22)
        Me.RtbTexto.TabIndex = 0
        Me.RtbTexto.TabStop = False
        Me.RtbTexto.Text = "Introducir Dirección eMail de Envio;"
        '
        'LbDireccionEMail
        '
        Me.LbDireccionEMail.AutoSize = True
        Me.LbDireccionEMail.Location = New System.Drawing.Point(38, 31)
        Me.LbDireccionEMail.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LbDireccionEMail.Name = "LbDireccionEMail"
        Me.LbDireccionEMail.Size = New System.Drawing.Size(104, 16)
        Me.LbDireccionEMail.TabIndex = 0
        Me.LbDireccionEMail.Text = "Dirección eMail:"
        Me.LbDireccionEMail.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'BtAceptar
        '
        Me.BtAceptar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtAceptar.ForeColor = System.Drawing.Color.Green
        Me.BtAceptar.Location = New System.Drawing.Point(270, 71)
        Me.BtAceptar.Name = "BtAceptar"
        Me.BtAceptar.Size = New System.Drawing.Size(85, 26)
        Me.BtAceptar.TabIndex = 2
        Me.BtAceptar.Text = "Aceptar"
        Me.BtAceptar.UseVisualStyleBackColor = True
        '
        'BtCancelar
        '
        Me.BtCancelar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtCancelar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.BtCancelar.Location = New System.Drawing.Point(360, 71)
        Me.BtCancelar.Name = "BtCancelar"
        Me.BtCancelar.Size = New System.Drawing.Size(85, 26)
        Me.BtCancelar.TabIndex = 3
        Me.BtCancelar.Text = "Cancelar"
        Me.BtCancelar.UseVisualStyleBackColor = True
        '
        'TbEMailDestino
        '
        Me.TbEMailDestino.BackColor = System.Drawing.SystemColors.Window
        Me.TbEMailDestino.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TbEMailDestino.Formato = ""
        Me.TbEMailDestino.Location = New System.Drawing.Point(145, 28)
        Me.TbEMailDestino.Margin = New System.Windows.Forms.Padding(5)
        Me.TbEMailDestino.Name = "TbEMailDestino"
        Me.TbEMailDestino.NumericoSolo = False
        Me.TbEMailDestino.NumeroSolo = False
        Me.TbEMailDestino.SeleccionarTodoCuandoClick = True
        Me.TbEMailDestino.Size = New System.Drawing.Size(300, 22)
        Me.TbEMailDestino.TabIndex = 1
        Me.TbEMailDestino.TextoVacioNumPermitido = True
        Me.TbEMailDestino.UsarEnterComoTab = True
        '
        'PbEmail
        '
        Me.PbEmail.Image = Global.MisClasesFuncionesVB.My.Resources.Resources.email_1_32
        Me.PbEmail.Location = New System.Drawing.Point(4, 24)
        Me.PbEmail.Margin = New System.Windows.Forms.Padding(4)
        Me.PbEmail.Name = "PbEmail"
        Me.PbEmail.Size = New System.Drawing.Size(32, 32)
        Me.PbEmail.TabIndex = 0
        Me.PbEmail.TabStop = False
        '
        'SsGetEMailDestino
        '
        Me.SsGetEMailDestino.BackColor = System.Drawing.SystemColors.ControlLight
        Me.SsGetEMailDestino.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SlEntrar, Me.SlEsc, Me.SlEditarEmailAntesEnviar, Me.SlDClick})
        Me.SsGetEMailDestino.Location = New System.Drawing.Point(0, 98)
        Me.SsGetEMailDestino.Name = "SsGetEMailDestino"
        Me.SsGetEMailDestino.Size = New System.Drawing.Size(448, 22)
        Me.SsGetEMailDestino.TabIndex = 4
        '
        'SlEntrar
        '
        Me.SlEntrar.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SlEntrar.ForeColor = System.Drawing.Color.Green
        Me.SlEntrar.Name = "SlEntrar"
        Me.SlEntrar.Size = New System.Drawing.Size(110, 17)
        Me.SlEntrar.Text = "^Entrar=Aceptar"
        '
        'SlEsc
        '
        Me.SlEsc.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SlEsc.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.SlEsc.Name = "SlEsc"
        Me.SlEsc.Size = New System.Drawing.Size(88, 17)
        Me.SlEsc.Text = "Esc=Cancelar"
        '
        'SlEditarEmailAntesEnviar
        '
        Me.SlEditarEmailAntesEnviar.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SlEditarEmailAntesEnviar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.SlEditarEmailAntesEnviar.Name = "SlEditarEmailAntesEnviar"
        Me.SlEditarEmailAntesEnviar.Size = New System.Drawing.Size(103, 17)
        Me.SlEditarEmailAntesEnviar.Text = "^E=EditarEmail"
        '
        'SlDClick
        '
        Me.SlDClick.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SlDClick.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SlDClick.Name = "SlDClick"
        Me.SlDClick.Size = New System.Drawing.Size(103, 17)
        Me.SlDClick.Text = "D.Click=LibDirecc"
        '
        'LbNota
        '
        Me.LbNota.AutoSize = True
        Me.LbNota.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LbNota.Location = New System.Drawing.Point(144, 54)
        Me.LbNota.Name = "LbNota"
        Me.LbNota.Size = New System.Drawing.Size(297, 13)
        Me.LbNota.TabIndex = 5
        Me.LbNota.Text = "Nota: Si NO especifcado, se muestra  Libreta de Direcciones."
        '
        'GetEMailDestino
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(449, 124)
        Me.ControlBox = False
        Me.Controls.Add(Me.LbNota)
        Me.Controls.Add(Me.PbEmail)
        Me.Controls.Add(Me.SsGetEMailDestino)
        Me.Controls.Add(Me.TbEMailDestino)
        Me.Controls.Add(Me.BtCancelar)
        Me.Controls.Add(Me.BtAceptar)
        Me.Controls.Add(Me.LbDireccionEMail)
        Me.Controls.Add(Me.RtbTexto)
        Me.Controls.Add(Me.CbEditartEmailAntesEnviar)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "GetEMailDestino"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        CType(Me.PbEmail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SsGetEMailDestino.ResumeLayout(False)
        Me.SsGetEMailDestino.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private components As System.ComponentModel.IContainer

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

End Class
