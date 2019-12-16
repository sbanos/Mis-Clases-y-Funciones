Public Class MiGestorConnectionString

    Shared Function GetConnectionString(ByVal ConnectionStringName As String) As String

        Dim Appconfig As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        Dim ConnStringSettings As ConnectionStringSettings = Appconfig.ConnectionStrings.ConnectionStrings(ConnectionStringName)

        If ConnStringSettings IsNot Nothing Then
            Return ConnStringSettings.ConnectionString
        Else
            Return String.Empty
        End If

    End Function

    Shared Sub SaveConnectionString(ByVal ConnectionStringName As String, ByVal ConnectionString As String)

        Dim Appconfig As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

        Appconfig.ConnectionStrings.ConnectionStrings(ConnectionStringName).ConnectionString = ConnectionString

        Appconfig.Save()

    End Sub

    Shared Function GetConnectionStringNames() As List(Of String)

        Dim cns As List(Of String) = New List(Of String)
        Dim Appconfig As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

        For Each cn As ConnectionStringSettings In Appconfig.ConnectionStrings.ConnectionStrings
            cns.Add(cn.Name)
        Next

        Return cns

    End Function

    Shared Function GetFirstConnectionStringName() As String

        Return GetConnectionStringNames().FirstOrDefault()

    End Function

    Shared Function GetFirstConnectionString() As String

        Return GetConnectionString(GetFirstConnectionStringName())

    End Function

    Shared Function GetOleDBProviderName(ByVal ConnectionString As String) As String

        Dim Builder As OleDbConnectionStringBuilder = New OleDbConnectionStringBuilder(ConnectionString)

        Return Builder.Provider

    End Function

    Shared Function SetOleDBProviderName(ByVal ConnectionString As String, ByVal Provider As String) As String

        Dim Builder As OleDbConnectionStringBuilder = New OleDbConnectionStringBuilder(ConnectionString) With {
            .Provider = Provider
        }

        Return Builder.ConnectionString

    End Function

    Shared Function GetOleDBDataSource(ByVal ConnectionString As String) As String

        Dim Builder As OleDbConnectionStringBuilder = New OleDbConnectionStringBuilder(ConnectionString)

        Return Builder.DataSource

    End Function

    Shared Function SetOleDBDataSource(ByVal ConnectionString As String, ByVal DataSource As String) As String

        Dim Builder As OleDbConnectionStringBuilder = New OleDbConnectionStringBuilder(ConnectionString) With {
            .DataSource = DataSource
        }

        Return Builder.ConnectionString

    End Function

    Shared Function GetSQLDataSource(ByVal ConnectionString As String) As String

        Dim Builder As SqlConnectionStringBuilder = New SqlConnectionStringBuilder(ConnectionString)

        Return Builder.DataSource

    End Function

    Shared Function SetSQLDataSource(ByVal ConnectionString As String, ByVal DataSource As String) As String

        Dim Builder As SqlConnectionStringBuilder = New SqlConnectionStringBuilder(ConnectionString) With {
            .DataSource = DataSource
        }

        Return Builder.ConnectionString

    End Function

    Shared Function GetSQLInitialCatalog(ByVal ConnectionString As String) As String

        Dim Builder As SqlConnectionStringBuilder = New SqlConnectionStringBuilder(ConnectionString)

        Return Builder.InitialCatalog

    End Function

    Shared Function SetSQLInitialCatalog(ByVal ConnectionString As String, ByVal InitialCatalog As String) As String

        Dim Builder As SqlConnectionStringBuilder = New SqlConnectionStringBuilder(ConnectionString) With {
            .InitialCatalog = InitialCatalog
        }

        Return Builder.ConnectionString

    End Function

    Shared Function GetSQLIntegratedSecurity(ByVal ConnectionString As String) As String

        Dim Builder As SqlConnectionStringBuilder = New SqlConnectionStringBuilder(ConnectionString)

        Return Builder.IntegratedSecurity

    End Function

    Shared Function SetSQLIntegratedSecurity(ByVal ConnectionString As String, ByVal IntegratedSecurity As Boolean) As String

        Dim Builder As SqlConnectionStringBuilder = New SqlConnectionStringBuilder(ConnectionString) With {
            .IntegratedSecurity = IntegratedSecurity
        }

        Return Builder.ConnectionString

    End Function

    Shared Function GetSQLUserID(ByVal ConnectionString As String) As String

        Dim Builder As SqlConnectionStringBuilder = New SqlConnectionStringBuilder(ConnectionString)

        Return Builder.UserID

    End Function

    Shared Function SetSQLUserID(ByVal ConnectionString As String, ByVal UserID As String) As String

        Dim Builder As SqlConnectionStringBuilder = New SqlConnectionStringBuilder(ConnectionString) With {
            .UserID = UserID
        }

        Return Builder.ConnectionString

    End Function

    Shared Function GetSQLPassword(ByVal ConnectionString As String) As String

        Dim Builder As SqlConnectionStringBuilder = New SqlConnectionStringBuilder(ConnectionString)

        Return Builder.Password

    End Function

    Shared Function SetSQLPassword(ByVal ConnectionString As String, ByVal Password As String) As String

        Dim Builder As SqlConnectionStringBuilder = New SqlConnectionStringBuilder(ConnectionString) With {
            .Password = Password
        }

        Return Builder.ConnectionString

    End Function

    Shared Function TestSQLConnectionString(ByVal StringConexion As String, Optional ByVal EmitirMensajeOK As Boolean = True, Optional ByVal EmitirMensajeKO As Boolean = True) As Boolean

        Dim Sqlcon As SqlConnection = New SqlConnection()

        Dim Respuesta As String = String.Empty

        Dim VMensaje As New MiVentanaMensaje(300, 1, "Verificando Conexión a SQL Server ...", ,  , , , , , , , , Color.GreenYellow)

        Try

            Sqlcon.ConnectionString = StringConexion
            Sqlcon.Open()

        Catch ex As Exception

            Respuesta = ex.Message

        End Try

        VMensaje.close()

        If Sqlcon.State = ConnectionState.Open Then
            Sqlcon.Close()
            If EmitirMensajeOK Then
                MiMessageBox.ShowWinMessage("OK", "Test Conexion SQL", MsgBoxStyle.Information, MsgBoxStyle.OkOnly)
            End If
            Return True
        Else
            Sqlcon.Close()
            If EmitirMensajeKO Then
                MiMessageBox.ShowWinMessage("ConnStr = <" + StringConexion + ">" + vbCrLf + vbCrLf + Respuesta, "Test Conexion SQL", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            End If
            Return False
        End If

    End Function

End Class
