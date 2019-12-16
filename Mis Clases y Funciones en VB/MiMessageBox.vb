Public NotInheritable Class MiMessageBox

#Region "Variables para el Mensaje ..."

    Private Shared ObjWinMessage As MiMessageBox

    Private Shared WithEvents CmdOk As Button                                   'Para el Boton Aceptar en los MessageBox
    Private Shared WithEvents CmdYes As Button                                  'Para el Boton Si en los MessageBox
    Private Shared WithEvents CmdNo As Button                                   'Para el Boton No en los MessageBox
    Private Shared WithEvents CmdCancel As Button                               'Para el Boton Cancelar en los MessageBox
    Private Shared WithEvents CmdAbort As Button                                'Para el Boton Abortar en los MessageBox
    Private Shared WithEvents CmdRetry As Button                                'Para el Boton Reintentar en los MessageBox
    Private Shared WithEvents CmdIgnore As Button                               'Para el Boton Ignorar en los MessageBox

    Private Shared WithEvents WinMsg As Form                                    'Para crear la forma del MessageBox
    Private Shared WithEvents LblHeader As Label                                'Label para el Titulo del MessageBox
    Private Shared PicIcon As PictureBox                                        'Para el icono del Mensaje del MessageBox
    Private Shared WithEvents WinMessage As Label                               'Para el Contenido del MessageBox

    Private Shared MaxWidth As Integer = SystemInformation.WorkingArea.Width    'Maximo acho del MessageBox
    Private Shared MaxHeight As Integer = SystemInformation.WorkingArea.Height  'Maximo alto del MessageBox

    'Inicializacion de la variables
    Private Shared FormWidth As Integer = 0
    Private Shared FormHeight As Integer = 0

    Private Shared HeaderWidth As Integer = 0

    Private Shared MsgWidth As Integer = 0
    Private Shared MsgHeight As Integer = 0

    Private Shared WinReturn As Integer = 0

    'Tipo de Letra, Tamaño y Estilo de la Fuente Default
    Private Shared ReadOnly WinFnt As New Font("Arial", 10, FontStyle.Regular)
    Private Shared WinLocation As Point

    Private Shared FormSize As Size 'Tamaño del MessageBox
    Private Shared MessageSize As Size 'Tamaño del Contenido del MessageBox

    Private Shared WinResult As DialogResult 'Almacenara el resultado del MessageBox, que seria por ejemplo cuando se de clic en algun boton
    Private Shared WinMake As DialogResult

    'Se inicializa la variable del contenido del Mensaje y el Titulo
    Private Shared MessageText As String = ""
    Private Shared HeaderText As String = ""

    'Colores de Fondo
    Private Shared MessageColorFondo As Color
    Private Shared HeaderColorFondo As Color
    Private Shared MarcoColor As Color

#End Region

#Region "Constructor de la clase"
    Private Sub New()
        WinMsg = New Form()
        AddHandler WinMsg.Shown, AddressOf ParaHacerUnaVezMostrado
    End Sub

    Private Sub ParaHacerUnaVezMostrado(sender As Object, e As EventArgs)

        Cursor.Position = WinMsg.Location + WinMsg.ActiveControl.Location + New Point(40, 12)

    End Sub
#End Region

    'Funcion que crea o emite el MessageBox
    Public Shared Function ShowWinMessage(ByVal WinText As String,
                                          Optional ByVal WinHeader As String = "",
                                          Optional ByVal WinIcon As MessageBoxIcon = MessageBoxIcon.None,
                                          Optional ByVal WinButtons As MessageBoxButtons = MessageBoxButtons.OK,
                                          Optional ByVal WinDefault As MessageBoxDefaultButton = MessageBoxDefaultButton.Button1,
                                          Optional ByVal WinPosition As FormStartPosition = FormStartPosition.CenterParent,
                                          Optional ByVal WinLocationInitialX As Integer = 0,
                                          Optional ByVal WinLocationInitialY As Integer = 0) As DialogResult

        If WinText Is Nothing Then
            Throw New ArgumentNullException("WinText")
        End If

        ' Si longitud > 75 y no tiene saltos de linea, añado saltos de linea ...
        Dim l As Integer = Len(WinText)
        If l > 75 And WinText.IndexOf(vbCrLf, StringComparison.Ordinal) = -1 Then     'Si el string es mayor de 75, y no tiene salto de linea .... 
            Dim i As Integer = 75
            Do While i < l
                WinText = WinText.Substring(0, i - 1) + Replace(WinText.Substring(i), " ", vbCrLf, 1, 1)
                i += 75
            Loop
        End If

        Beep()
        WinResult = MakeMessage(WinText, WinHeader, WinIcon, WinButtons, WinDefault, WinPosition, WinLocationInitialX, WinLocationInitialY)
        Return WinResult

    End Function

    Private Shared Function MakeMessage(ByVal WinText As String,
                                        Optional ByVal WinHeader As String = "",
                                        Optional ByVal WinIcon As MessageBoxIcon = MessageBoxIcon.None,
                                        Optional ByVal WinButtons As MessageBoxButtons = MessageBoxButtons.OK,
                                        Optional ByVal WinDefault As MessageBoxDefaultButton = MessageBoxDefaultButton.Button1,
                                        Optional ByVal WinPosition As FormStartPosition = FormStartPosition.CenterParent,
                                        Optional ByVal WinLocationInitialX As Integer = 0,
                                        Optional ByVal WinLocationInitialY As Integer = 0) As DialogResult

        ObjWinMessage = Nothing
        ObjWinMessage = New MiMessageBox()

        MessageText = "" : MessageText = WinText                                'Se establece el Mensaje del MessageBox
        HeaderText = "" : HeaderText = WinHeader                                'Se establece el Titulo del MessageBox

        MaxWidth -= 60                                                          'Al Maximo Ancho del MessageBox le restamos el margen que dejaremos en el Formulario
        MaxHeight -= 120                                                        'Al Maximo Alto del MessageBox le restamos el margen que dejaremos en el Formulario

        'Se inicializan la variables
        FormWidth = 0 : FormHeight = 0
        HeaderWidth = 0
        MsgWidth = 0 : MsgHeight = 0
        WinReturn = 0

        'Comprueba si la longitud del Texto Mensaje y Titulo sea igual a Cero
        If MessageText.Trim().Length = 0 And HeaderText.Trim().Length = 0 Then
            FormSize = New Size(305, 135)                                       'Se establece el tamaño del formulario
            FormWidth = 305 : FormHeight = 135                                  'Se establece el ancho y alto del formulario
            WinMsg.Size = New Size(FormSize.Width, FormSize.Height)             'Se establece el tamaño de la forma base
            GoTo Mess                                                           'Se sale de esta desicion y se va asta la linea Mess
        End If

        HeaderWidth = TamañoCadena(HeaderText.Trim(), MaxWidth, WinFnt).Width   'Se establece el ancho del TITULO
        MessageSize = TamañoCadena(MessageText.Trim(), MaxWidth, WinFnt)        'Se establece el tamaño (Alto y Ancho) del Mensaje
        MsgWidth = MessageSize.Width : MsgHeight = MessageSize.Height           'Se establece el ancho del Mensaje

        'MsgBox(New Size(MsgWidth, MsgHeight).ToString)

        'Si el Tamaño o Longitud del Titulo es mayor de 0 y el Mensaje es Igual a 0
        If HeaderText.Trim().Length > 0 And MessageText.Trim().Length = 0 Then
            HeaderWidth = HeaderWidth + 80                                      'Se establece el ancho del Titulo
            WinReturn = Math.Max(HeaderWidth, 305)                              'Se establece el ancho que tendra el formulario
            FormSize = New Size(WinReturn, 135)                                 'Se establece el tamaño (ancho y alto) del formulario
            FormWidth = FormSize.Width : FormHeight = FormSize.Height           'Se establece el Ancho y Alto del formulario
            GoTo Mess                                                           'Se sale de esta desicion y se va hasta la linea Mess
        End If

        HeaderWidth = HeaderWidth + 80                                          'Se establece el ancho del Label (Titulo)
        FormWidth = MsgWidth + 60                                               'Se establece el ancho del formulario
        FormHeight = MsgHeight + 120                                            'Se establece el alto del formulario

        'Si el Tamaño o Longitud del Titulo es igual a 0 y el Mensaje es Mayor a 0
        If HeaderText.Trim().Length = 0 And MessageText.Trim().Length > 0 Then
            FormWidth = Math.Max(FormWidth, 305)                                'Se establece el Ancho del Formulario
            FormHeight = Math.Max(FormHeight, 135)                              'Se establece el Alto del Formulario
            FormSize = New Size(FormWidth, FormHeight)                          'Se establece el Tamaño (Ancho y Alto) del Formulario
            FormWidth = FormSize.Width : FormHeight = FormSize.Height           'Se establece el Ancho y Alto del Formulario
            GoTo Mess                                                           'Se sale de esta decision y se va hasta la linea Mess
        End If

        'Si el Tamaño o Longitud del Titulo es mayor de 0 y el Mensaje es Mayor de 0
        If HeaderText.Trim().Length > 0 And MessageText.Trim().Length > 0 Then
            WinReturn = Math.Max(HeaderWidth, FormWidth)                        'Retorna el Ancho del Formulario
            FormWidth = Math.Max(WinReturn, 305)                                'Se establece el Ancho del Formulario
            FormHeight = Math.Max(FormHeight, 135)                              'Se establece el Alto del Formulario
            FormSize = New Size(FormWidth, FormHeight)                          'Se establece el Tamaño (Ancho y Alto) del Formulario
            FormWidth = FormSize.Width : FormHeight = FormSize.Height           'Se establece el Ancho y Alto del Formulario
            GoTo Mess                                                           'Se sale de esta desicion y se va hasta la linea Mess
        End If

Mess:
        Call CrearBasePantalla(WinPosition, WinLocationInitialX, WinLocationInitialY) 'Se llama al procedimiento que crea la Base del Formulario del MessageBox

        Call CrearTitulo()                                                      'Se llama al procedimiento que crea el Titulo del MessageBox
        Call CrearIconoMensaje()                                                'Se llama al procedimiento que crea el Icono del Mesaje
        Call CrearMensaje()                                                     'Se llama al procedimiento que crea el Cuerpo o el Mensaje a mostrar del MessageBox

        Call AñadiraPantallaBase()                                              'Se añaden al formulario base el Titulo, Mensaje e Icono creados anteriormente.

        WinMessage.Text = WinText                                               'Se establece el Mensaje
        LblHeader.Text = WinHeader                                              'Se establece el Titulo
        'Se selecciona el Icono a mostrar en el MessageBox (Aca se puede cambiar por una imagen de nuestra preferencia)
        'Tambien el Color de Fondo, en funcion del Icono seleccionado.
        Select Case WinIcon
            'Case MessageBoxIcon.Asterisk 'En Caso que se seleccione Asterisco
            'PicIcon.Image = Drawing.SystemIcons.Asterisk.ToBitmap() 'Se selecciona los iconos default del MessageBox Normal que estan guardados en Drawing.SystemIcons
            'PicIcon.Image = La imagen de tu preferencia asi se puede cambiar en cada uno de los iconos
            'MessageColorFondo = Color.FromArgb(166, 197, 227)
            'HeaderColorFondo = Color.FromArgb(166, 197, 227)
            'MarcoColor = Color.FromArgb(96, 155, 173)
            Case MessageBoxIcon.Error
                PicIcon.Image = Drawing.SystemIcons.Error.ToBitmap()
                MessageColorFondo = Color.FromArgb(255, 128, 128)
                HeaderColorFondo = Color.Tomato
                MarcoColor = Color.FromArgb(192, 0, 0)
            Case MessageBoxIcon.Exclamation
                PicIcon.Image = Drawing.SystemIcons.Exclamation.ToBitmap()
                MessageColorFondo = Color.Gold
                HeaderColorFondo = Color.Goldenrod
                MarcoColor = Color.DarkGoldenrod
            Case MessageBoxIcon.Hand
                PicIcon.Image = Drawing.SystemIcons.Hand.ToBitmap()
                MessageColorFondo = Color.FromArgb(255, 128, 128)
                HeaderColorFondo = Color.Tomato
                MarcoColor = Color.FromArgb(192, 0, 0)
            Case MessageBoxIcon.Information
                PicIcon.Image = Drawing.SystemIcons.Information.ToBitmap()
                MessageColorFondo = Color.CornflowerBlue ' FromArgb(166, 197, 227)
                HeaderColorFondo = Color.RoyalBlue
                MarcoColor = Color.FromArgb(96, 155, 173)
            Case MessageBoxIcon.None
                PicIcon.Image = Nothing
                MessageColorFondo = Color.Gainsboro
                HeaderColorFondo = Color.Silver
                MarcoColor = Color.Gray
            Case MessageBoxIcon.Question
                PicIcon.Image = Drawing.SystemIcons.Question.ToBitmap()
                MessageColorFondo = Color.Violet
                HeaderColorFondo = Color.Orchid
                MarcoColor = Color.MediumVioletRed
            Case MessageBoxIcon.Stop
                PicIcon.Image = Drawing.SystemIcons.Error.ToBitmap()
                MessageColorFondo = Color.FromArgb(255, 128, 128)
                HeaderColorFondo = Color.Tomato
                MarcoColor = Color.FromArgb(192, 0, 0)
            Case MessageBoxIcon.Warning
                PicIcon.Image = Drawing.SystemIcons.Warning.ToBitmap()
                MessageColorFondo = Color.Gold
                HeaderColorFondo = Color.Goldenrod
                MarcoColor = Color.DarkGoldenrod
            Case Else
                MessageColorFondo = Color.FromArgb(166, 197, 227)
                HeaderColorFondo = Color.FromArgb(166, 197, 227)
                MarcoColor = Color.FromArgb(96, 155, 173)
        End Select

        Call CrearBotonesMensajes(WinButtons, WinDefault)

        WinMsg.ShowDialog()

        Return WinMake

    End Function

    'Se crea la Forma Basica del MessageBox
    Private Shared Sub CrearBasePantalla(Optional ByVal WinPosition As FormStartPosition = FormStartPosition.CenterParent,
                                         Optional ByVal WinLocationInitialX As Integer = 0,
                                         Optional ByVal WinLocationInitialY As Integer = 0)
        With WinMsg
            .Text = ""                                          'Se vacia el Titulo de la forma del MessageBox
            .Size = New Size(FormWidth, FormHeight)             'Se establece el tamaño del formulario
            If WinPosition = FormStartPosition.Manual Then      'Se establece la posicion inicial del formulario en tiempo de ejecucion
                .Location = New Point(WinLocationInitialX, WinLocationInitialY)
            Else
                .StartPosition = WinPosition
            End If
            .FormBorderStyle = FormBorderStyle.None             'Se establece el estilo del borde del formulario
            .ShowInTaskbar = False : .ShowIcon = False          'Se establece que el formulario no se mostrara en la barra de tareas y que el Icono no aparecera en el Titulo
            .Opacity = 1                                        'Se establece el nivel de opacidad del formulario.
            .Font = WinFnt                                      'Se establece la fuente del formulario
        End With

    End Sub

    'Procedimiento que crea el Label para el Titulo del MessageBox
    Private Shared Sub CrearTitulo()

        LblHeader = New Label()

        With LblHeader
            .Text = ""                                  'Se establece el Texto del Titulo en vacio.
            .AutoSize = False                           'Se establece que el control no cambie automaticamente el tamaño para mostrar el contenido.
            .Dock = DockStyle.Top                       'Se establece que los bordes del Label (Titulo) se acoplaran a su contrl principal (Formulario Base)
            .BackColor = Color.Transparent              'Se establece el color de fondo del Label (Titulo)
            .TextAlign = ContentAlignment.MiddleCenter  'Se establce la alineacion del texto del Label (Titulo) en la izquierda.
            .Height = 24                                'Se establece el alto del Label (Titulo)
            .Font = WinFnt                              'Se establece la fuente del Label (Titulo)
            .SendToBack()                               'Envia el Label (Titulo) al final del orden Z
            .Visible = True                             'Se pone visible el Label (Titulo)
        End With

    End Sub

    'Procedimiento que crea el Icono del MessageBox
    Private Shared Sub CrearIconoMensaje()

        PicIcon = New PictureBox()

        With PicIcon
            .Size = New Size(35, 35)                    'Se establece el tamaño alto y ancho del icono del MessageBox
            .Location = New Point(8, 32)                'Establece la posicion donde estara ubicado el icono del MessageBox
            .BackColor = Color.Transparent              'Establece el color de fondo del icono del MessageBox
            .BorderStyle = BorderStyle.None             'Establece el estilo del borde para el icono del MessageBox
            .SendToBack()                               'Envia el icono o imagen al final del orden Z
            .Visible = True                             'Se pone visible el icono o imagen del MessageBox
        End With

    End Sub

    'Procedimiento que crea el Cuerpo o el Mensaje a mostrar en el MessageBox
    Private Shared Sub CrearMensaje()

        WinMessage = New Label()

        With WinMessage
            .Text = ""                                  'Se establece el Texto del Mensaje en vacio.
            .Size = New Size(MsgWidth, MsgHeight)       'Se establece el tamaño alto y ancho del Label (Mensaje) del MessageBox
            .Location = New Point(48, 32)               'Establece la posicion donde estara ubicado el Label (Mensaje)
            .AutoSize = True 'False                     'Se establece que el Label (Mensaje) no cambie automaticamente el tamaño para mostrar el contenido.
            .Font = WinFnt                              'Se establece la fuente del Label (Mensaje)
            .TextAlign = ContentAlignment.TopLeft       'Se establce la alineacion del texto del Label (Titulo) en la izquierda.
            .BackColor = Color.Transparent              'Se establece el color de fondo del Label (Mensaje)
            .SendToBack()                               'Envia el Label (Mensaje) al final del orden Z
            .Visible = True                             'Se pone visible el Label (Mensaje)
            '.BorderStyle = BorderStyle.FixedSingle
        End With

    End Sub

    'Procedimiento que añade al formulario base el Titulo, Icono y Mensaje.
    Private Shared Sub AñadiraPantallaBase()

        With WinMsg
            .Controls.Add(LblHeader)                    'Se añade el Label Titulo
            .Controls.Add(PicIcon)                      'Se añade el Icono
            .Controls.Add(WinMessage)                   'Se añade el Mensaje
            .Refresh()
        End With

    End Sub

    'Funcion donde se crear, se ubican y se les da el tamaño correcto a los botones. Segun la opcion que se seleccione
    Private Shared Sub CrearBotonesMensajes(ByVal WinButtons As MessageBoxButtons, ByVal WinDefault As MessageBoxDefaultButton)

        Select Case WinButtons
            Case MessageBoxButtons.AbortRetryIgnore
                CmdAbort = New Button()
                Call BotonPropiedades(CmdAbort, "&Abortar", New Size(80, 24), New Point(CInt(FormWidth - (103 + 80 * 2 + 5 * 2)), FormHeight - 70))
                CmdRetry = New Button()
                Call BotonPropiedades(CmdRetry, "&Reintentar", New Size(80, 24), New Point(CInt(FormWidth - (103 + 80 + 5)), FormHeight - 70))
                CmdIgnore = New Button()
                Call BotonPropiedades(CmdIgnore, "&Ignorar", New Size(80, 24), New Point(FormWidth - 103, FormHeight - 70))
                If WinDefault = MessageBoxDefaultButton.Button1 Then
                    WinMsg.ActiveControl = CmdAbort
                ElseIf WinDefault = MessageBoxDefaultButton.Button2 Then
                    WinMsg.ActiveControl = CmdRetry
                Else
                    WinMsg.ActiveControl = CmdIgnore
                End If
            Case MessageBoxButtons.OK
                CmdOk = New Button()
                Call BotonPropiedades(CmdOk, "&Aceptar", New Size(80, 24), New Point(FormWidth - 103, FormHeight - 70))
                WinMsg.ActiveControl = CmdOk
            Case MessageBoxButtons.OKCancel
                CmdOk = New Button()
                Call BotonPropiedades(CmdOk, "&Aceptar", New Size(80, 24), New Point(CInt(FormWidth - (103 + 80 + 5)), FormHeight - 70))
                CmdCancel = New Button()
                Call BotonPropiedades(CmdCancel, "&Cancelar", New Size(80, 24), New Point(FormWidth - 103, FormHeight - 70))
                If WinDefault = MessageBoxDefaultButton.Button1 Then
                    WinMsg.ActiveControl = CmdOk
                Else
                    WinMsg.ActiveControl = CmdCancel
                End If
            Case MessageBoxButtons.RetryCancel
                CmdRetry = New Button()
                Call BotonPropiedades(CmdRetry, "&Reintentar", New Size(80, 24), New Point(CInt(FormWidth - (103 + 80 + 5)), FormHeight - 70))
                CmdCancel = New Button()
                Call BotonPropiedades(CmdCancel, "&Cancelar", New Size(80, 24), New Point(FormWidth - 103, FormHeight - 70))
                If WinDefault = MessageBoxDefaultButton.Button1 Then
                    WinMsg.ActiveControl = CmdRetry
                Else
                    WinMsg.ActiveControl = CmdCancel
                End If
            Case MessageBoxButtons.YesNo
                CmdYes = New Button()
                Call BotonPropiedades(CmdYes, "&Si", New Size(80, 24), New Point(CInt(FormWidth - (103 + 80 + 5)), FormHeight - 70))
                CmdNo = New Button()
                Call BotonPropiedades(CmdNo, "&No", New Size(80, 24), New Point(FormWidth - 103, FormHeight - 70))
                If WinDefault = MessageBoxDefaultButton.Button1 Then
                    WinMsg.ActiveControl = CmdYes
                Else
                    WinMsg.ActiveControl = CmdNo
                End If
            Case MessageBoxButtons.YesNoCancel
                CmdYes = New Button()
                Call BotonPropiedades(CmdYes, "&Si", New Size(80, 24), New Point(CInt(FormWidth - (103 + 80 * 2 + 5 * 2)), FormHeight - 70))
                CmdNo = New Button()
                Call BotonPropiedades(CmdNo, "&No", New Size(80, 24), New Point(CInt(FormWidth - (103 + 80 + 5)), FormHeight - 70))
                CmdCancel = New Button()
                Call BotonPropiedades(CmdCancel, "&Cancelar", New Size(80, 24), New Point(FormWidth - 103, FormHeight - 70))
                If WinDefault = MessageBoxDefaultButton.Button1 Then
                    WinMsg.ActiveControl = CmdYes
                ElseIf WinDefault = MessageBoxDefaultButton.Button2 Then
                    WinMsg.ActiveControl = CmdNo
                Else
                    WinMsg.ActiveControl = CmdCancel
                End If
        End Select

    End Sub

    'Procedimiento con las Propiedades del Boton
    Private Shared Sub BotonPropiedades(ByVal Btn As Button, ByVal Txt As String, ByVal Sz As Size, ByVal Lc As Point)

        With Btn
            .BringToFront()                                                             'Coloca el boton al principio del orden Z
            .Size = Sz                                                                  'Se establece el alto y ancho del Boton
            .Text = Txt                                                                 'Se coloca el texto que tendra el Boton
            .BackColor = HeaderColorFondo                                               'Se establece el Color de fondo del Boton
            .FlatAppearance.BorderSize = 0                                              'Establece el valor que especifica el tamaño, en pixeles, del borde alrededor del boton
            .FlatStyle = FlatStyle.Standard                                             'Establece la apariencia del estilo plano del boton
            .Location = Lc                                                              'Establece la posicion del Boton
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right    'Se establece el borde del contenedor al que esta enlazado el boton
            .TextAlign = ContentAlignment.MiddleCenter                                  'Se establece la alineacion del texto del boton en el Centro
            .Font = New Font("Arial", 10, FontStyle.Bold)                               'Se establece el tamaño, nombre Fuente y Estilo del Texto del Boton
            .Visible = True                                                             'Se pone visible el Boton
        End With

        WinMsg.Controls.Add(Btn) 'Se agregar el boton en el formulario del Mensaje

    End Sub

    'Funcion que devuelve el tamaño de la cadena
    Private Shared Function TamañoCadena(ByVal WinMsgText As String, ByVal WinWdth As Integer, ByVal WinFnt As Font) As Size

        Dim GRA As Graphics = WinMsg.CreateGraphics()
        Dim SZF As SizeF = GRA.MeasureString(WinMsgText, WinFnt, WinWdth)   'Mide la cadena especificada al dibujarla
        GRA.Dispose()                                                       'Libera los recursos usados por GRA

        Dim SZ As New Size(Convert.ToInt16(SZF.Width) + 50,
                           Convert.ToInt16(SZF.Height))                     'Establece el ancho y alto de la cadena

        Return SZ                                                           'Devuelve el tamaño (ancho y alto) de la cadena

    End Function

    'Procedimiento para dibujar el area donde ira el Mensaje y el icono del MessageBox
    Private Shared Sub WinMsg_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles WinMsg.Paint

        Dim MGraphics As Graphics = e.Graphics                              'Variable donde se almacenara el grafico
        Dim MPen As New Pen(MarcoColor, 1)                                  'Variable donde se almacenara el grosor y el color de la linea

        Dim Area As New Rectangle(0, 0, WinMsg.Width - 1, WinMsg.Height - 1) 'Variable que almacena un rectangulo
        Dim LGradient As New _
            LinearGradientBrush(Area, MessageColorFondo,
            Color.FromArgb(245, 251, 251),
            LinearGradientMode.BackwardDiagonal)                            'Variable que almacenara los colores que se usaran en el Gradiente, como ba a ser ese gradiente y a que se va a aplicar ese gradiente
        MGraphics.FillRectangle(LGradient, Area)                            'Se rellena el rectangulo con el Gradiente
        MGraphics.DrawRectangle(MPen, Area)                                 'Se dibuja el rectangulo ya con el gradiente aplicado y los bordes que se han configurado en la varibale MPen

    End Sub

    Private Shared Sub LblHeader_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles LblHeader.MouseDown

        WinLocation = e.Location                                            'Variable donde se almacenara la Localizacion del Label cuando se pase el mouse por encima

    End Sub

    'Cuando se mueve el MessageBox con el Mouse a otra posicion
    Private Shared Sub LblHeader_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles LblHeader.MouseMove

        'Si el boton presionado es el Derecho sobre el Titulo del MessageBox
        If String.Compare(Control.MouseButtons.ToString(), "Left", StringComparison.CurrentCultureIgnoreCase) = 0 Then
            Dim MSize As New Size(WinLocation) With {
                .Width = e.X - WinLocation.X,                               'Se almacena en la posicion X donde esta actualmente el formulario
                .Height = e.Y - WinLocation.Y                               'Se almacena en la posicion Y donde esta actualmente el formulario
                }
            WinMsg.Location = Point.Add(WinMsg.Location, MSize)             'Se ubica el formulario en la nueva posicion
        End If

    End Sub

    'Procedimiento para dibujar el area donde ira el Titulo del MessageBox
    Private Shared Sub LblHeader_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles LblHeader.Paint

        Dim MGraphics As Graphics = e.Graphics                          'Variable donde se almacenara el grafico final
        Dim MPen As New Pen(MarcoColor, 1)                              'Variable donde se almacenara el grosor y el color de la linea

        Dim Area As New Rectangle(0, 0, LblHeader.Width - 1, LblHeader.Height - 1) 'Variable que almacena un rectangulo
        Dim LGradient As New _
            LinearGradientBrush(Area, HeaderColorFondo,
                                Color.FromArgb(245, 251, 251),
                                LinearGradientMode.BackwardDiagonal)    'Variable que almacenara los colores que se usaran en el Gradiente, como ba a ser ese gradiente y a que se va a aplicar ese gradiente
        MGraphics.FillRectangle(LGradient, Area)                        'Se rellena el rectangulo con el Gradiente
        MGraphics.DrawRectangle(MPen, Area)                             'Se dibuja el rectangulo ya con el gradiente aplicado y los bordes que se han configurado en la varibale MPen

        Dim DrawFont As New Font("Arial", 10, FontStyle.Bold)           'Se almacena el nombre, el tamaño y el estilo de la fuente para el Titulo
        Dim DrawBrush As New SolidBrush(Color.Black)                    'Se almacena el color de la fuente del Titulo
        Dim DrawPoint As New PointF(2.0F, 3.0F)                         'Se almacena la posicion donde estara el Titulo

        Dim DrawGradientBrush As New _
            LinearGradientBrush(e.Graphics.ClipBounds,
                                Color.White, Color.FromArgb(122, 158, 226),
                                LinearGradientMode.ForwardDiagonal)     'Se crea un Gradiente Lineal


        e.Graphics.DrawString(HeaderText.Trim.ToString(), DrawFont, DrawBrush, DrawPoint) 'Se dibuja el Titulo con los datos anteriormente configurados

    End Sub

    Private Shared Sub CmdAbort_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdAbort.Click

        WinMsg.Dispose() : WinMsg = Nothing
        WinMake = DialogResult.Abort

    End Sub

    Private Shared Sub CmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdCancel.Click

        WinMsg.Dispose() : WinMsg = Nothing
        WinMake = DialogResult.Cancel

    End Sub

    Private Shared Sub CmdIgnore_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdIgnore.Click

        WinMsg.Dispose() : WinMsg = Nothing
        WinMake = DialogResult.Ignore

    End Sub

    Private Shared Sub CmdNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdNo.Click

        WinMsg.Dispose() : WinMsg = Nothing
        WinMake = DialogResult.No

    End Sub

    Private Shared Sub CmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdOk.Click

        WinMsg.Dispose() : WinMsg = Nothing
        WinMake = DialogResult.OK

    End Sub

    Private Shared Sub CmdRetry_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdRetry.Click

        WinMsg.Dispose() : WinMsg = Nothing
        WinMake = DialogResult.Retry

    End Sub

    Private Shared Sub CmdYes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdYes.Click

        WinMsg.Dispose() : WinMsg = Nothing
        WinMake = DialogResult.Yes

    End Sub


End Class
