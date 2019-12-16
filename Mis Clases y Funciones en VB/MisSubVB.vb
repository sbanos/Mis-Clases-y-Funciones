Public Module MisSubVB

    Public Sub SoloNumerico(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
            Case Else
                If InStr("1234567890,", e.KeyChar) = 0 Then
                    Beep()
                    e.Handled = True
                End If
        End Select
    End Sub

    Public Sub SoloNumeros(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Select Case Asc(e.KeyChar)
            Case 8
            Case 13
            Case Else
                If InStr("1234567890", e.KeyChar) = 0 Then
                    Beep()
                    e.Handled = True
                End If
        End Select
    End Sub

End Module
