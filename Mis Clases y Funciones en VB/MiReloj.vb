
Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public NotInheritable Class MiReloj

#Region "Declaraciones y Variables ..."

    Private Shared ObjReloj As MiReloj

    Private Shared WithEvents Reloj As Form
    Private Shared WithEvents Panel As System.Windows.Forms.Panel
    Private Shared WithEvents Esfera As System.Windows.Forms.PictureBox
    Private Shared WithEvents Temporizador As System.Windows.Forms.Timer
    Private Shared ToolTip1 As System.Windows.Forms.ToolTip

    Private Shared E03 As System.Windows.Forms.Label
    Private Shared E06 As System.Windows.Forms.Label
    Private Shared E09 As System.Windows.Forms.Label
    Private Shared E12 As System.Windows.Forms.Label
    Private Shared M05 As Microsoft.VisualBasic.PowerPacks.OvalShape
    Private Shared M10 As Microsoft.VisualBasic.PowerPacks.OvalShape
    Private Shared M20 As Microsoft.VisualBasic.PowerPacks.OvalShape
    Private Shared M25 As Microsoft.VisualBasic.PowerPacks.OvalShape
    Private Shared M35 As Microsoft.VisualBasic.PowerPacks.OvalShape
    Private Shared M40 As Microsoft.VisualBasic.PowerPacks.OvalShape
    Private Shared M50 As Microsoft.VisualBasic.PowerPacks.OvalShape
    Private Shared M55 As Microsoft.VisualBasic.PowerPacks.OvalShape
    Private Shared ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer

    Private Shared Formulario As New GraphicsPath()
    Private Shared PanelReloj As New GraphicsPath()
    Private Shared EsferaReloj As New GraphicsPath()

    Private Shared AnguloSegundos As Integer
    Private Shared AnguloMinutos As Integer
    Private Shared AnguloHoras As Integer

    Private Shared swMouseDown As Boolean = False
    Private Shared PosicionReloj As Point

#End Region

#Region "Constructor de la clase"
    Private Sub New()
        Reloj = New Form()
    End Sub
#End Region

    Public Shared Sub ShowReloj(Optional ByVal FormatoReloj As Boolean = False, Optional PosicionInicialX As Int16 = 0, Optional PosicionInicialY As Int16 = 0)

        If Not IsNothing(Reloj) Then    ' El Reloj ya está arrancado
            Exit Sub
        End If

        'Reloj = New Form()

        Call MakeReloj(FormatoReloj)  ' True = Circulo - False = Cuadrado

        Reloj.Show()
        If Not (PosicionInicialX = 0 And PosicionInicialY = 0) Then
            Reloj.Location = New Point(PosicionInicialX, PosicionInicialY)
        End If

        ' Pone en marcha el Reloj, con intervalo de 1 Segundo
        Temporizador.Interval = 1000
        Temporizador.Enabled = True

    End Sub

    Public Shared Sub ReShowReloj(ByVal PosicionX As Int16, ByVal PosicionY As Int16)

        If Not IsNothing(Reloj) Then
            Reloj.Location = New Point(PosicionX, PosicionY)
        End If

    End Sub

    Public Shared Sub CloseReloj()

        Reloj.Close()
        Reloj.Dispose() : Reloj = Nothing

    End Sub

    Private Shared Sub MakeReloj(Optional ByVal FormatoReloj As Boolean = False)

        ObjReloj = Nothing
        ObjReloj = New MiReloj

        Panel = New System.Windows.Forms.Panel()
        Esfera = New System.Windows.Forms.PictureBox()
        Temporizador = New System.Windows.Forms.Timer()
        ToolTip1 = New System.Windows.Forms.ToolTip()
        ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()

        E03 = New System.Windows.Forms.Label()
        E06 = New System.Windows.Forms.Label()
        E09 = New System.Windows.Forms.Label()
        E12 = New System.Windows.Forms.Label()
        M05 = New Microsoft.VisualBasic.PowerPacks.OvalShape()
        M10 = New Microsoft.VisualBasic.PowerPacks.OvalShape()
        M20 = New Microsoft.VisualBasic.PowerPacks.OvalShape()
        M25 = New Microsoft.VisualBasic.PowerPacks.OvalShape()
        M35 = New Microsoft.VisualBasic.PowerPacks.OvalShape()
        M40 = New Microsoft.VisualBasic.PowerPacks.OvalShape()
        M50 = New Microsoft.VisualBasic.PowerPacks.OvalShape()
        M55 = New Microsoft.VisualBasic.PowerPacks.OvalShape()

        Reloj.BackColor = System.Drawing.Color.DarkGray
        Reloj.ClientSize = New System.Drawing.Size(156, 156)
        Reloj.ControlBox = False
        Reloj.Controls.Add(Panel)
        Reloj.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Reloj.ShowInTaskbar = False
        Reloj.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Reloj.TopMost = True

        Panel.BackColor = System.Drawing.Color.Transparent
        Panel.Controls.Add(E12)
        Panel.Controls.Add(E03)
        Panel.Controls.Add(E09)
        Panel.Controls.Add(E06)
        Panel.Controls.Add(Esfera)
        Panel.Controls.Add(ShapeContainer1)
        Panel.Location = New System.Drawing.Point(3, 3)
        Panel.Margin = New System.Windows.Forms.Padding(2)
        Panel.Size = New System.Drawing.Size(150, 150)

        Esfera.Location = New System.Drawing.Point(16, 16)
        Esfera.Margin = New System.Windows.Forms.Padding(2)
        Esfera.Size = New System.Drawing.Size(118, 118)
        ToolTip1.SetToolTip(Esfera, "Doble Click para Cerrar")
        ToolTip1.IsBalloon = True

        ShapeContainer1.Location = New System.Drawing.Point(0, 0)
        ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {M55, M50, M40, M35, M25, M20, M10, M05})
        ShapeContainer1.Size = New System.Drawing.Size(112, 122)

        E03.BackColor = System.Drawing.Color.Transparent
        E03.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        E03.ForeColor = System.Drawing.Color.Gold
        E03.Location = New System.Drawing.Point(132, 66)
        E03.Size = New System.Drawing.Size(16, 16)
        E03.Text = "3"

        E06.BackColor = System.Drawing.Color.Transparent
        E06.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        E06.ForeColor = System.Drawing.Color.Gold
        E06.Location = New System.Drawing.Point(68, 132)
        E06.Size = New System.Drawing.Size(16, 16)
        E06.Text = "6"

        E09.BackColor = System.Drawing.Color.Transparent
        E09.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        E09.ForeColor = System.Drawing.Color.Gold
        E09.Location = New System.Drawing.Point(2, 66)
        E09.Size = New System.Drawing.Size(16, 16)
        E09.Text = "9"

        E12.BackColor = System.Drawing.Color.Transparent
        E12.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        E12.ForeColor = System.Drawing.Color.Gold
        E12.Location = New System.Drawing.Point(63, 2)
        E12.Size = New System.Drawing.Size(24, 16)
        E12.Text = "12"

        M05.FillColor = System.Drawing.Color.Gold
        M05.FillStyle = Microsoft.VisualBasic.PowerPacks.FillStyle.Solid
        M05.Location = New System.Drawing.Point(104, 18)
        M05.Size = New System.Drawing.Size(5, 5)

        M10.FillColor = System.Drawing.Color.Gold
        M10.FillStyle = Microsoft.VisualBasic.PowerPacks.FillStyle.Solid
        M10.Location = New System.Drawing.Point(127, 41)
        M10.Size = New System.Drawing.Size(5, 5)

        M20.FillColor = System.Drawing.Color.Gold
        M20.FillStyle = Microsoft.VisualBasic.PowerPacks.FillStyle.Solid
        M20.Location = New System.Drawing.Point(127, 104)
        M20.Size = New System.Drawing.Size(5, 5)

        M25.FillColor = System.Drawing.Color.Gold
        M25.FillStyle = Microsoft.VisualBasic.PowerPacks.FillStyle.Solid
        M25.Location = New System.Drawing.Point(104, 127)
        M25.Size = New System.Drawing.Size(5, 5)

        M35.FillColor = System.Drawing.Color.Gold
        M35.FillStyle = Microsoft.VisualBasic.PowerPacks.FillStyle.Solid
        M35.Location = New System.Drawing.Point(41, 127)
        M35.Size = New System.Drawing.Size(5, 5)

        M40.FillColor = System.Drawing.Color.Gold
        M40.FillStyle = Microsoft.VisualBasic.PowerPacks.FillStyle.Solid
        M40.Location = New System.Drawing.Point(18, 104)
        M40.Size = New System.Drawing.Size(5, 5)

        M50.FillColor = System.Drawing.Color.Gold
        M50.FillStyle = Microsoft.VisualBasic.PowerPacks.FillStyle.Solid
        M50.Location = New System.Drawing.Point(18, 41)
        M50.Size = New System.Drawing.Size(5, 5)

        M55.FillColor = System.Drawing.Color.Gold
        M55.FillStyle = Microsoft.VisualBasic.PowerPacks.FillStyle.Solid
        M55.Location = New System.Drawing.Point(41, 18)
        M55.Size = New System.Drawing.Size(5, 5)

        'Forma del Formulario
        Dim RFormulario As New Rectangle(0, 0, Reloj.Height, Reloj.Height)
        If FormatoReloj Then
            'Forma Circular del Formulario
            Formulario.AddEllipse(RFormulario)
            Reloj.Region = New Region(Formulario)
        Else
            'Forma Cuadrtada del Formulario
            Formulario.AddRectangle(RFormulario)
        End If

        'Forma del Panel donde se aloja la Esfera
        Dim RPanel As New Rectangle(0, 0, Panel.Height, Panel.Height)
        If FormatoReloj Then
            'Forma Circular del Panel donde se aloja la Esfera
            PanelReloj.AddEllipse(RPanel)
            Panel.Region = New Region(PanelReloj)
        Else
            'Forma Cuadrada del Panel donde se aloja la Esfera
            PanelReloj.AddRectangle(RPanel)
        End If

        'Forma Circular de la Esfera
        Dim REsfera As New Rectangle(0, 0, Esfera.Height, Esfera.Height)
        EsferaReloj.AddEllipse(REsfera)
        Esfera.Region = New Region(EsferaReloj)

        DibujarManecillas()

    End Sub

    Private Shared Sub Form_Paint(sender As Object, e As PaintEventArgs) Handles Reloj.Paint

        e.Graphics.DrawPath(New Pen(Color.Black, 3), Formulario)

    End Sub

    Private Shared Sub Panel_Paint(sender As Object, e As PaintEventArgs) Handles Panel.Paint

        Dim Rectangulo1 As Rectangle = New Rectangle(0, 0, Panel.Width, Panel.Height / 2 + 2)
        Dim Rectangulo2 As Rectangle = New Rectangle(0, Panel.Height / 2 + 2, Panel.Width, Panel.Height)
        Dim myBrush1 As Brush = New Drawing.Drawing2D.LinearGradientBrush(Rectangulo1, Color.DarkBlue, Color.SteelBlue, Drawing.Drawing2D.LinearGradientMode.Vertical)
        Dim myBrush2 As Brush = New Drawing.Drawing2D.LinearGradientBrush(Rectangulo1, Color.SteelBlue, Color.DarkBlue, Drawing.Drawing2D.LinearGradientMode.Vertical)
        e.Graphics.FillRectangle(myBrush1, Rectangulo1)
        e.Graphics.FillRectangle(myBrush2, Rectangulo2)

        e.Graphics.DrawPath(New Pen(Color.Black, 3), PanelReloj)

        'Dim Rectangulo3 As Rectangle = New Rectangle(12, 12, Me.Panel.Width - 24, Me.Panel.Height - 24)
        'PanelReloj.AddRectangle(Rectangulo3)
        'PanelReloj.AddEllipse(Rectangulo3)
        'e.Graphics.DrawPath(New Pen(Color.Black, 1), PanelReloj)

    End Sub

    Private Shared Sub DibujarManecillas() Handles Temporizador.Tick

        ' Define las Variables del Grafico
        Dim MapaBits As Bitmap = New Bitmap(Esfera.Width, Esfera.Height)
        Dim DibujoManecillas As Graphics = Graphics.FromImage(MapaBits)

        ' Lineas Suaves
        DibujoManecillas.SmoothingMode = SmoothingMode.HighQuality

        ' Establece los Angulos 
        AnguloSegundos = 225 + (Date.Now.Second * 6)                        ' Ajuste + 360º/60 Segundos
        AnguloMinutos = 225 + (Date.Now.Minute * 6)                         ' Ajuste + 360º/60 Minutos
        AnguloHoras = 225 + ((Date.Now.Hour + Date.Now.Minute / 60) * 30)   ' Ajuste + 360º/12 Horas

        Dim AlineacionCentrada As New StringFormat With {.LineAlignment = StringAlignment.Near, .Alignment = StringAlignment.Center}
        Dim Fuente As Font = New Font("Microsoft Sans Serif", 9.75, FontStyle.Bold)
        ' Hora del Dia
        Dim rCajaHora As Rectangle
        With rCajaHora
            .X = 25
            .Y = Esfera.Height / 10 * 2 - 4
            .Width = Esfera.Width - 50
            .Height = Fuente.Height
        End With
        DibujoManecillas.DrawString(Now.ToLongTimeString, Fuente, Brushes.Gold, rCajaHora, AlineacionCentrada)
        ' Fecha del Dia
        Dim rCajaFecha As Rectangle
        With rCajaFecha
            .X = 15
            .Y = Esfera.Height / 10 * 7 + 1
            .Width = Esfera.Width - 30
            .Height = Fuente.Height
        End With
        DibujoManecillas.DrawString(Date.Today.ToShortDateString, Fuente, Brushes.Silver, rCajaFecha, AlineacionCentrada)

        'Dibuja la Manecilla de las Horas
        Dim ManecillaHoras As New Pen(Brushes.Black, 3)
        ManecillaHoras.SetLineCap(0, LineCap.ArrowAnchor, DashCap.Flat)
        DibujoManecillas.TranslateTransform(Esfera.Width / 2, Esfera.Height / 2) ' Origen del Sistemas de Coordenadas respecto al BitMap donde se va a pintar 
        DibujoManecillas.RotateTransform(AnguloHoras)
        DibujoManecillas.DrawLine(ManecillaHoras, 0, 0, CInt(Esfera.Width / 4), CInt(Esfera.Width / 4))
        DibujoManecillas.ResetTransform()

        'Dibuja la Manecilla de los Minutos
        Dim ManecillaMinutos As New Pen(Brushes.Black, 2)
        ManecillaMinutos.SetLineCap(0, LineCap.ArrowAnchor, DashCap.Flat)
        DibujoManecillas.TranslateTransform(Esfera.Width / 2, Esfera.Height / 2) ' Origen del Sistemas de Coordenadas respecto al BitMap donde se va a pintar 
        DibujoManecillas.RotateTransform(AnguloMinutos)
        DibujoManecillas.DrawLine(ManecillaMinutos, 0, 0, CInt(Esfera.Width / 3), CInt(Esfera.Width / 3))
        DibujoManecillas.ResetTransform()

        'Dibuja la Manecilla de los Segundos
        Dim ManecillaSegundos As New Pen(Brushes.Red, 2)
        ManecillaSegundos.SetLineCap(0, LineCap.Flat, DashCap.Flat)
        DibujoManecillas.TranslateTransform(Esfera.Width / 2, Esfera.Height / 2) ' Origen del Sistemas de Coordenadas respecto al BitMap donde se va a pintar 
        DibujoManecillas.RotateTransform(AnguloSegundos)
        DibujoManecillas.DrawLine(ManecillaSegundos, 0, 0, CInt(Esfera.Width / 2), CInt(Esfera.Width / 2))
        DibujoManecillas.ResetTransform()

        ' Tapar la Unión de las 3 Manecillas
        DibujoManecillas.DrawEllipse(Pens.Black, CInt(Esfera.Width / 2) - 4, CInt(Esfera.Height / 2) - 4, 8, 8)
        DibujoManecillas.FillEllipse(Brushes.DarkRed, CInt(Esfera.Width / 2) - 4, CInt(Esfera.Height / 2) - 4, 8, 8)

        ' Imagen Pintada en PictureBox
        Esfera.Image = MapaBits

    End Sub

    Private Shared Sub Esfera_DoubleClick(sender As Object, e As EventArgs) Handles Esfera.DoubleClick
        Reloj.Close()
        Reloj.Dispose() : Reloj = Nothing
    End Sub

    'Desplazar el Reloj por la Pantalla - Inicio
    Private Shared Sub Esfera_MouseDown(sender As Object, e As MouseEventArgs) Handles Panel.MouseDown, Esfera.MouseDown
        PosicionReloj = New Point(Cursor.Position.X - Reloj.Location.X, Cursor.Position.Y - Reloj.Location.Y)
        swMouseDown = True
    End Sub
    Private Shared Sub Esfera_MouseMove(sender As Object, e As MouseEventArgs) Handles Esfera.MouseMove ', Panel.MouseMove
        If swMouseDown Then
            'Reloj.Location = New Point(Reloj.Location.X + e.X, Reloj.Location.Y + e.Y)
            Reloj.Location = New Point(Cursor.Position.X - PosicionReloj.X, Cursor.Position.Y - PosicionReloj.Y)
        End If
    End Sub
    Private Shared Sub Esfera_MouseUp(sender As Object, e As MouseEventArgs) Handles Panel.MouseUp, Esfera.MouseUp
        swMouseDown = False
    End Sub
    'Desplazar el Reloj por la Pantalla - Fin

End Class

