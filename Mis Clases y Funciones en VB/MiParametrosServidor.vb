Public Class MiParametrosServidor

    Private StringParametros As String     'IdConexion=xxx;NombreServidor=xxx;Puerto=xxx;InstanciaBaseDatos=xxx;BaseDatos=xxx;TipoAutenticacion=xxx;ID=xxx;Password=xxx
    Private IdDeConexion As String = String.Empty
    Private NombreDeServidor As String = String.Empty
    Private InstanciaDeBaseDatos As String = String.Empty
    Private NumeroDePuerto As String = String.Empty
    Private BaseDeDatos As String = String.Empty
    Private TipoDeAutenticacion As String = String.Empty
    Private IdDeUsuario As String = String.Empty
    Private PalabraDeClave As String = String.Empty

    Sub New(Optional ByVal StringParametros As String = "")

        Me.StringParametros = StringParametros

        If Me.StringParametros <> String.Empty Then
            Replace(StringParametros, " ", "") ' Quita los posibles espacios que pudieran haber ...
            ObtenerPropiedades()
        End If

    End Sub

    Private Sub ObtenerPropiedades()

        Dim Parametros() As String = StringParametros.Split(";")

        For Each Parametro As String In Parametros
            If Parametro.IndexOf("IdConexion=") >= 0 Then
                IdDeConexion = Mid(Parametro, 12)
            End If
            If Parametro.IndexOf("NombreServidor=") >= 0 Then
                NombreDeServidor = Mid(Parametro, 16)
            End If
            If Parametro.IndexOf("InstanciaBaseDatos=") >= 0 Then
                InstanciaDeBaseDatos = Mid(Parametro, 20)
            End If
            If Parametro.IndexOf("Puerto=") >= 0 Then
                NumeroDePuerto = Mid(Parametro, 8)
            End If
            If Parametro.IndexOf("BaseDatos=") >= 0 Then
                BaseDeDatos = Mid(Parametro, 11)
            End If
            If Parametro.IndexOf("TipoAutenticacion=") >= 0 Then
                TipoDeAutenticacion = Mid(Parametro, 19)
            End If
            If Parametro.IndexOf("IdUsuario=") >= 0 Then
                IdDeUsuario = Mid(Parametro, 11)
            End If
            If Parametro.IndexOf("PalabraClave=") >= 0 Then
                PalabraDeClave = Mid(Parametro, 14)
            End If
        Next

    End Sub

    Private Sub ReconstruirCadena()

        StringParametros = String.Empty
        If IdDeConexion <> String.Empty Then
            StringParametros = StringParametros + "IdConexion=" + IdDeConexion + ";"
        End If
        If NombreDeServidor <> String.Empty Then
            StringParametros = StringParametros + "NombreServidor=" + NombreDeServidor + ";"
        End If
        If InstanciaDeBaseDatos <> String.Empty Then
            StringParametros = StringParametros + "InstanciaBaseDatos=" + InstanciaDeBaseDatos + ";"
        End If
        If NumeroDePuerto <> String.Empty Then
            StringParametros = StringParametros + "Puerto=" + NumeroDePuerto + ";"
        Else
            StringParametros = StringParametros + "Puerto=Default;" 'Valor por defecto
        End If
        If BaseDeDatos <> String.Empty Then
            StringParametros = StringParametros + "BaseDatos=" + BaseDeDatos + ";"
        End If
        If TipoDeAutenticacion = String.Empty OrElse (TipoDeAutenticacion <> "Windows" AndAlso TipoDeAutenticacion <> "SQL Server") Then
            StringParametros = StringParametros + "TipoAutenticacion=Windows;" ' Valor por defecto y Corrección de errores de especificacion
        Else
            StringParametros = StringParametros + "TipoAutenticacion=" + TipoDeAutenticacion + ";"
        End If
        If IdDeUsuario <> String.Empty Then
            StringParametros = StringParametros + "IdUsuario=" + IdDeUsuario + ";"
        End If
        If PalabraDeClave <> String.Empty Then
            StringParametros = StringParametros + "PalabraClave=" + PalabraDeClave + ";"
        End If
        StringParametros = Mid(StringParametros, 1, Len(StringParametros) - 1)  'Quita el ultimo ";"

    End Sub

    Public Property IdConexion() As String

        Get
            Return IdDeConexion
        End Get

        Set(ByVal Valor As String)
            IdDeConexion = Valor
            ReconstruirCadena()
        End Set

    End Property

    Public Property NombreServidor() As String

        Get
            Return NombreDeServidor
        End Get

        Set(ByVal Valor As String)
            NombreDeServidor = Valor
            ReconstruirCadena()
        End Set

    End Property

    Public Property InstanciaDaseDatos() As String

        Get
            Return InstanciaDeBaseDatos
        End Get

        Set(ByVal Valor As String)
            InstanciaDeBaseDatos = Valor
            ReconstruirCadena()
        End Set

    End Property

    Public Property Puerto() As String

        Get
            If NumeroDePuerto = String.Empty Then
                NumeroDePuerto = "Default"
                ReconstruirCadena()
            End If
            Return NumeroDePuerto
        End Get

        Set(ByVal Valor As String)
            NumeroDePuerto = Valor
            ReconstruirCadena()
        End Set

    End Property

    Public Property DaseDatos() As String

        Get
            Return BaseDeDatos
        End Get

        Set(ByVal Valor As String)
            BaseDeDatos = Valor
            ReconstruirCadena()
        End Set

    End Property

    Public Property TipoAutenticacion() As String

        Get
            If TipoDeAutenticacion = String.Empty OrElse (TipoDeAutenticacion <> "Windows" AndAlso TipoDeAutenticacion <> "SQL Server") Then
                TipoDeAutenticacion = "Windows"
                ReconstruirCadena()
            End If
            Return TipoDeAutenticacion
        End Get

        Set(ByVal Valor As String)
            TipoDeAutenticacion = Valor
            ReconstruirCadena()
        End Set

    End Property

    Public Property IdUsuario() As String

        Get
            Return IdDeUsuario
        End Get

        Set(ByVal Valor As String)
            IdDeUsuario = Valor
            ReconstruirCadena()
        End Set

    End Property

    Public Property PalabraClave() As String

        Get
            Return PalabraDeClave
        End Get

        Set(ByVal Valor As String)
            PalabraDeClave = Valor
            ReconstruirCadena()
        End Set

    End Property


    Public Property CadenaParametros() As String

        Get
            ReconstruirCadena()
            Return StringParametros
        End Get

        Set(ByVal Valor As String)
            StringParametros = Valor
        End Set



    End Property

End Class

