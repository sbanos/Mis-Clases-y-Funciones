Public Class MiVentanaMensaje

    'Inherits System.Windows.Forms.Form
    Private Shared WithEvents VentanaMensaje As Form

    Public LineaMensaje1 As System.Windows.Forms.Label
    Public LineaMensaje2 As System.Windows.Forms.Label
    Public LineaMensaje3 As System.Windows.Forms.Label
    Public LineaMensaje4 As System.Windows.Forms.Label
    Public LineaMensaje5 As System.Windows.Forms.Label
    Public LineaMensaje6 As System.Windows.Forms.Label
    Public LineaMensaje7 As System.Windows.Forms.Label
    Public LineaMensaje8 As System.Windows.Forms.Label
    Public LineaMensaje9 As System.Windows.Forms.Label
    Public LineaMensaje10 As System.Windows.Forms.Label

    Public Sub New(Optional AnchoVentana As Int16 = 350, Optional LineasVentana As Int16 = 5, Optional ByVal TxtLineaMensaje1 As String = "", Optional ByVal TxtLineaMensaje2 As String = "", Optional ByVal TxtLineaMensaje3 As String = "", Optional ByVal TxtLineaMensaje4 As String = "", Optional ByVal TxtLineaMensaje5 As String = "", Optional ByVal TxtLineaMensaje6 As String = "", Optional ByVal TxtLineaMensaje7 As String = "", Optional ByVal TxtLineaMensaje8 As String = "", Optional ByVal TxtLineaMensaje9 As String = "", Optional ByVal TxtLineaMensaje10 As String = "", Optional ByVal ColorFondo As System.Drawing.Color = Nothing)

        If LineasVentana < 1 Then
            LineasVentana = 1
        ElseIf LineasVentana > 10 Then
            LineasVentana = 10
        End If

        VentanaMensaje = New Form()

        InitializarComponente(AnchoVentana, LineasVentana, ColorFondo)

        If LineasVentana > 0 Then
            LineaMensaje1.Text = TxtLineaMensaje1
        End If
        If LineasVentana > 1 Then
            LineaMensaje2.Text = TxtLineaMensaje2
        End If
        If LineasVentana > 2 Then
            LineaMensaje3.Text = TxtLineaMensaje3
        End If
        If LineasVentana > 3 Then
            LineaMensaje4.Text = TxtLineaMensaje4
        End If
        If LineasVentana > 4 Then
            LineaMensaje5.Text = TxtLineaMensaje5
        End If
        If LineasVentana > 5 Then
            LineaMensaje6.Text = TxtLineaMensaje6
        End If
        If LineasVentana > 6 Then
            LineaMensaje7.Text = TxtLineaMensaje7
        End If
        If LineasVentana > 7 Then
            LineaMensaje8.Text = TxtLineaMensaje8
        End If
        If LineasVentana > 8 Then
            LineaMensaje9.Text = TxtLineaMensaje9
        End If
        If LineasVentana > 9 Then
            LineaMensaje10.Text = TxtLineaMensaje10
        End If

        VentanaMensaje.Show()
        VentanaMensaje.Refresh()

    End Sub

    Private Sub InitializarComponente(Optional AnchoVentana As Int16 = 350, Optional LineasVentana As Int16 = 5, Optional ByVal ColorFondo As System.Drawing.Color = Nothing)
        '
        'Lineas Mensaje
        '
        If LineasVentana > 0 Then
            LineaMensaje1 = New System.Windows.Forms.Label With {
                .Location = New System.Drawing.Point(0, 0 * 25),
                .Size = New System.Drawing.Size(AnchoVentana, 25),
                .TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                .Font = New Font("Arial", 10, FontStyle.Regular)
            }
        End If
        If LineasVentana > 1 Then
            LineaMensaje2 = New System.Windows.Forms.Label With {
                .Location = New System.Drawing.Point(0, 1 * 25),
                .Size = New System.Drawing.Size(AnchoVentana, 25),
                .TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                .Font = New Font("Arial", 10, FontStyle.Regular)
            }
        End If
        If LineasVentana > 2 Then
            LineaMensaje3 = New System.Windows.Forms.Label With {
                .Location = New System.Drawing.Point(0, 2 * 25),
                .Size = New System.Drawing.Size(AnchoVentana, 25),
                .TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                .Font = New Font("Arial", 10, FontStyle.Regular)
            }
        End If
        If LineasVentana > 3 Then
            LineaMensaje4 = New System.Windows.Forms.Label With {
                .Location = New System.Drawing.Point(0, 3 * 25),
                .Size = New System.Drawing.Size(AnchoVentana, 25),
                .TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                .Font = New Font("Arial", 10, FontStyle.Regular)
            }
        End If
        If LineasVentana > 4 Then
            LineaMensaje5 = New System.Windows.Forms.Label With {
                .Location = New System.Drawing.Point(0, 4 * 25),
                .Size = New System.Drawing.Size(AnchoVentana, 25),
                .TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                .Font = New Font("Arial", 10, FontStyle.Regular)
            }
        End If
        If LineasVentana > 5 Then
            LineaMensaje6 = New System.Windows.Forms.Label With {
                .Location = New System.Drawing.Point(0, 5 * 25),
                .Size = New System.Drawing.Size(AnchoVentana, 25),
                .TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                .Font = New Font("Arial", 10, FontStyle.Regular)
            }
        End If
        If LineasVentana > 6 Then
            LineaMensaje7 = New System.Windows.Forms.Label With {
                .Location = New System.Drawing.Point(0, 6 * 25),
                .Size = New System.Drawing.Size(AnchoVentana, 25),
                .TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                .Font = New Font("Arial", 10, FontStyle.Regular)
            }
        End If
        If LineasVentana > 7 Then
            LineaMensaje8 = New System.Windows.Forms.Label With {
                .Location = New System.Drawing.Point(0, 7 * 25),
                .Size = New System.Drawing.Size(AnchoVentana, 25),
                .TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                .Font = New Font("Arial", 10, FontStyle.Regular)
            }
        End If
        If LineasVentana > 8 Then
            LineaMensaje9 = New System.Windows.Forms.Label With {
                .Location = New System.Drawing.Point(0, 8 * 25),
                .Size = New System.Drawing.Size(AnchoVentana, 25),
                .TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                .Font = New Font("Arial", 10, FontStyle.Regular)
            }
        End If
        If LineasVentana > 9 Then
            LineaMensaje10 = New System.Windows.Forms.Label With {
                .Location = New System.Drawing.Point(0, 9 * 25),
                .Size = New System.Drawing.Size(AnchoVentana, 25),
                .TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                .Font = New Font("Arial", 10, FontStyle.Regular)
            }
        End If
        '
        'Mensaje
        '
        VentanaMensaje.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        VentanaMensaje.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        If ColorFondo = Nothing Then
            VentanaMensaje.BackColor = System.Drawing.SystemColors.ControlDark
        Else
            VentanaMensaje.BackColor = ColorFondo
        End If
        VentanaMensaje.ClientSize = New System.Drawing.Size(AnchoVentana, LineasVentana * 25)
        VentanaMensaje.ControlBox = False
        VentanaMensaje.Controls.Add(LineaMensaje1)
        VentanaMensaje.Controls.Add(LineaMensaje2)
        VentanaMensaje.Controls.Add(LineaMensaje3)
        VentanaMensaje.Controls.Add(LineaMensaje4)
        VentanaMensaje.Controls.Add(LineaMensaje5)
        VentanaMensaje.Controls.Add(LineaMensaje6)
        VentanaMensaje.Controls.Add(LineaMensaje7)
        VentanaMensaje.Controls.Add(LineaMensaje8)
        VentanaMensaje.Controls.Add(LineaMensaje9)
        VentanaMensaje.Controls.Add(LineaMensaje10)
        VentanaMensaje.MaximizeBox = False
        VentanaMensaje.MinimizeBox = False
        VentanaMensaje.Name = "VentanaMensaje"
        VentanaMensaje.Opacity = 0.87R
        VentanaMensaje.ShowInTaskbar = False
        VentanaMensaje.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        VentanaMensaje.Text = "Procesando ..."
        VentanaMensaje.Font = New Font("Arial", 10, FontStyle.Regular)
        'TopMost = True

        VentanaMensaje.Cursor = Cursors.WaitCursor

    End Sub

    Public Sub Refresh()
        VentanaMensaje.Refresh()
    End Sub

    Public Sub Close()
        VentanaMensaje.Cursor = Cursors.Default
        VentanaMensaje.Close()
    End Sub

    ' Reemplaza a Dispose para limpiar la lista de componentes.
    'Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    'Try
    'If disposing AndAlso components IsNot Nothing Then
    '           components.Dispose()
    'End If
    'Finally
    'MyBase.Dispose(disposing)
    'End Try
    'End Sub
    'Private components As System.ComponentModel.IContainer

End Class
