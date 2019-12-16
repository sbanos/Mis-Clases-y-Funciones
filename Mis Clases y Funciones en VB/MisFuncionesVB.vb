Imports System.Text.RegularExpressions
Imports System.Globalization

Public Module MisFuncionesVB
    Public Function MiInputBox(ByVal Prompt As String, Optional ByVal Titulo As String = "", Optional ByVal ValorDefecto As String = "", Optional ByVal TipoDato As String = "S") As String
        Dim Valor As String = String.Empty
        Do While True
            Valor = InputBox(Prompt, Titulo, ValorDefecto)
            If Valor = String.Empty Then
                MiMessageBox.ShowWinMessage("Proceso Cancelado", Titulo, MessageBoxIcon.Information, MessageBoxButtons.OK)
                Exit Do
            ElseIf TipoDato = "N" And Not IsNumeric(Valor) Then
                MiMessageBox.ShowWinMessage("Error en Tipo de Dato Introducido = " + Valor + vbCrLf + vbCrLf + "El Dato ha de ser Numérico" + vbCrLf + vbCrLf + "Corregir ...", Titulo, MessageBoxIcon.Exclamation, MessageBoxButtons.OK)
            ElseIf TipoDato = "F" And Not IsDate(Valor) Then
                MiMessageBox.ShowWinMessage("Error en Tipo de Dato Introducido = " + Valor + vbCrLf + vbCrLf + "El Dato ha de ser Formato Fecha" + vbCrLf + vbCrLf + "Corregir ...", Titulo, MessageBoxIcon.Exclamation, MessageBoxButtons.OK)
            Else
                Exit Do
            End If
        Loop
        Return Valor
    End Function

    Public Function ValidarEmail(ByVal email As String) As Boolean

        Dim emailRegex As New System.Text.RegularExpressions.Regex("^(?<user>[^@]+)@(?<host>.+)$")
        Dim emailMatch As System.Text.RegularExpressions.Match = emailRegex.Match(email)
        Return emailMatch.Success

    End Function

    Public Function VacioOBasura(ByVal TextoAValidar As String) As Boolean

        If TextoAValidar Is Nothing Then
            Throw New ArgumentNullException("TextoAValidar")
        End If

        TextoAValidar = TextoAValidar.Trim

        If Len(TextoAValidar) > 0 AndAlso Asc(TextoAValidar) = 0 Then  'Caracter Nulo
            Return True
        ElseIf TextoAValidar = String.Empty OrElse
            TextoAValidar = "." OrElse
            TextoAValidar = ". ." OrElse
            TextoAValidar = "," OrElse
            TextoAValidar = ", ," OrElse
            TextoAValidar = "- -" OrElse
            TextoAValidar = "-" Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function ValidarNif(ByRef Nif As String) As Boolean

        '*******************************************************************
        ' Nombre:     ValidateNif
        '             por Enrique Martínez Montejo
        '
        ' Finalidad:  Validar el NIF/NIE pasado a la función.
        '
        ' Entradas:
        '     NIF:    String. El NIF/NIE que se desea verificar. El número
        '             será devuelto formateado y con el NIF/NIE correcto.
        ' Resultados:
        ' Boolean:    True/False
        '*******************************************************************

        If Nif Is Nothing Then
            Throw New ArgumentNullException("Nif")
        End If

        Dim nifTemp As String = Nif.Trim().ToUpper(CultureInfo.CurrentCulture)

        If (nifTemp.Length > 9) Then Return False

        ' Guardamos el dígito de control.
        Dim dcTemp As Char = nifTemp.Chars(Nif.Length - 1)

        ' Compruebo si el dígito de control es un número.
        If (Char.IsDigit(dcTemp)) Then Return Nothing

        ' Nos quedamos con los caracteres, sin el dígito de control.
        nifTemp = nifTemp.Substring(0, Nif.Length - 1)

        If (nifTemp.Length < 8) Then
            Dim paddingChar As String = New String("0"c, 8 - nifTemp.Length)
            nifTemp = nifTemp.Insert(nifTemp.Length, paddingChar)
        End If

        ' Obtengo el dígito de control correspondiente, utilizando
        ' para ello una llamada a la función GetDcCif.
        '
        Dim dc As Char = GetDcNif(Nif)

        If (Not (dc = Nothing)) Then
            Nif = nifTemp & dc
        End If

        Return (dc = dcTemp)

    End Function

    Public Function GetDcNif(ByVal Nif As String) As Char

        '*******************************************************************
        ' Nombre:     GetDcNif
        '             por Enrique Martínez Montejo
        '
        ' Finalidad:  Devuelve la letra correspondiente al NIF o al NIE
        '             (Número de Identificación de Extranjero)
        '
        ' Entradas:
        '     NIF:    String. La cadena del NIF cuya letra final se desea
        '             obtener.
        '
        ' Resultados:
        '     String: La letra del NIF/NIE.
        '*******************************************************************

        If Nif Is Nothing Then
            Throw New ArgumentNullException("Nif")
        End If

        ' Pasamos el NIF a mayúscula a la vez que eliminamos los
        ' espacios en blanco al comienzo y al final de la cadena.
        '
        Nif = Nif.Trim().ToUpper(CultureInfo.CurrentCulture)

        ' El NIF está formado de uno a nueve números seguido de una letra.
        '
        ' El NIF de otros colectivos de personas físicas, está
        ' formato por una letra (K, L, M), seguido de 7 números
        ' y de una letra final.
        '
        ' El NIE está formado de una letra inicial (X, Y, Z),
        ' seguido de 7 números y de una letra final.
        ' 
        ' En el patrón de la expresión regular, defino cuatro grupos en el
        ' siguiente orden:
        '
        ' 1º) 1 a 8 dígitos.
        ' 2º) 1 a 8 dígitos + 1 letra.
        ' 3º) 1 letra + 1 a 7 dígitos.
        ' 4º) 1 letra + 1 a 7 dígitos + 1 letra.
        '
        Dim re As New Regex("(^\d{1,8}$)|(^\d{1,8}[A-Z]$)|(^[K-MX-Z]\d{1,7}$)|(^[K-MX-Z]\d{1,7}[A-Z]$)", RegexOptions.IgnoreCase)

        If (Not (re.IsMatch(Nif))) Then Return Nothing

        ' Nos quedamos únicamente con los números del NIF, y
        ' los formateamos con ceros a la izquierda si su
        ' longitud es inferior a siete caracteres.
        '
        re = New Regex("(\d{1,8})")

        Dim numeros As String = re.Match(Nif).Value.PadLeft(7, "0"c)

        ' Primer carácter del NIF.
        '
        Dim firstChar As Char = Nif.Chars(0)

        ' Si procede, reemplazamos la letra del NIE por el peso que le corresponde.
        '
        If (firstChar = "X"c) Then
            numeros = "0" & numeros
        ElseIf (firstChar = "Y"c) Then
            numeros = "1" & numeros
        ElseIf (firstChar = "Z"c) Then
            numeros = "2" & numeros
        End If

        ' Tabla del NIF
        '
        '  0T  1R  2W  3A  4G  5M  6Y  7F  8P  9D
        ' 10X 11B 12N 13J 14Z 15S 16Q 17V 18H 19L
        ' 20C 21K 22E 23T
        '
        ' Procedo a calcular el NIF/NIE
        '
        Dim dni As Integer = CInt(numeros)

        ' La operación consiste en calcular el resto de dividir el DNI
        ' entre 23 (sin decimales). Dicho resto (que estará entre 0 y 22),
        ' se busca en la tabla y nos da la letra del NIF.
        '
        ' Obtenemos el resto de la división.
        '
        Dim r As Integer = dni Mod 23

        ' Obtenemos el dígito de control del NIF
        '
        Dim dc As Char = CChar("TRWAGMYFPDXBNJZSQVHLCKE".Substring(r, 1))

        Return dc

    End Function

    Public Function ValidarCif(ByRef Cif As String) As Boolean

        '*******************************************************************
        ' Nombre:     ValidateCif
        '             por Enrique Martínez Montejo
        '
        ' Finalidad:  Validar el NIF pasado a la función.
        '
        ' Entradas:
        '     nif:    String. El NIF que se desea verificar. El número
        '             será devuelto formateado con el NIF correcto.
        ' Resultados:
        '     Boolean: True/False
        '*******************************************************************

        If Cif Is Nothing Then
            Throw New ArgumentNullException("Cif")
        End If

        Dim CifTemp As String = Cif.Trim().ToUpper(CultureInfo.CurrentCulture)

        If (CifTemp.Length < 9) Then Return False

        ' Guardamos el noveno carácter.
        Dim dcTemp As Char = CifTemp.Chars(8)

        ' Nos quedamos con los primeros ocho caracteres.
        '
        CifTemp = CifTemp.Substring(0, 8)

        ' Obtengo el dígito de control correspondiente, utilizando
        ' para ello una llamada a la función GetDcCif
        '
        Dim dc As Char = GetDcCif(Cif)

        If (Not (dc = Nothing)) Then
            Cif = CifTemp & dc
        End If

        Return (dc = dcTemp)

    End Function

    Public Function GetDcCif(ByVal nif As String) As Char

        '*******************************************************************
        ' Nombre:     GetDcCif
        '             por Enrique Martínez Montejo
        '
        ' Finalidad:  Obtener el Dígito de Control de un NIF correspondiente
        '             a personas jurídicas y otras entidades con o sin
        '             personalidad jurídica.
        '
        ' Entradas:
        '     nif:    String. El NIF cuyo dígito de control se desea obtener.
        '
        ' Resultados:
        '     String: La letra o el número correspondiente al NIF.
        '*******************************************************************

        ' Pasamos el NIF a mayúscula a la vez que eliminamos todos los
        ' carácteres que no sean números o letras. Note que no incluyo
        ' la letra I, porque si bien no puede aparecer como carácter
        ' inicial de un NIF, sí puede ser un dígito de control válido,
        ' lo que no sucede con las letras O y T.
        '
        Dim re As New Regex("([^A-W0-9]|[OT]|[^\w])", RegexOptions.IgnoreCase)

        nif = re.Replace(nif, String.Empty).ToUpper(CultureInfo.CurrentCulture)

        ' Para calcular el CIF, se debe de haber pasado un máximo
        ' de nueve caracteres a la función: una letra válida (que
        ' necesariamente deberá de estar comprendida en el intervalo
        ' ABCDEFGHJKLMNPQRSUVW), siete números, y el dígito de control,
        ' que puede ser un número o una letra, dependiendo de la entidad
        ' a la que pertenezca el NIF.
        '
        ' En el patrón de la expresión regular, defino dos grupos en el
        ' siguiente orden:
        ' 1º) 1 letra + 7 u 8 dígitos.
        ' 2º) 1 letra + 7 dígitos + 1 letra.
        '
        ' Note que en el siguiente patrón, no incluyo la letra I como
        ' carácter de inicio válido del NIF.
        '
        re = New Regex("(^[A-HJ-W]\d{7,8}$)|(^[A-HJ-W]\d{7}[A-Z]$)")

        If (Not (re.IsMatch(nif))) Then Return Nothing

        ' Guardo el último carácter pasado, siempre que
        ' el NIF disponga de nueve caracteres.
        '
        Dim lastChar As Char = Nothing
        If (nif.Length = 9) Then lastChar = nif.Chars(8)

        Dim sumaPar, sumaImpar As Int32

        ' A continuación, la cadena debe tener 7 dígitos.
        '
        Dim digits As String = nif.Substring(1, 7)

        For n As Int32 = 0 To digits.Length - 1 Step 2

            If (n < 6) Then
                ' Sumo las cifras pares del número que se corresponderá
                ' con los caracteres 1, 3 y 5 de la variable «digits».
                '
                sumaImpar += CInt(CStr(digits.Chars(n + 1)))
            End If

            ' Multiplico por dos cada cifra impar (caracteres 0, 2, 4 y 6).
            '
            Dim dobleImpar As Int32 = 2 * CInt(CStr(digits.Chars(n)))

            ' Acumulo la suma del doble de números impares.
            '
            sumaPar += (dobleImpar Mod 10) + (dobleImpar \ 10)

        Next

        ' Sumo las cifras pares e impares.
        '
        Dim sumaTotal As Int32 = sumaPar + sumaImpar

        ' Me quedo con la cifra de las unidades y se la resto a 10, siempre
        ' y cuando la cifra de las unidades sea distinta de cero.
        '
        sumaTotal = (10 - (sumaTotal Mod 10)) Mod 10

        Dim characters() As Char = {"J"c, "A"c, "B"c, "C"c, "D"c, "E"c, "F"c, "G"c, "H"c, "I"c}

        Dim dc As Char = characters(sumaTotal)

        ' Devuelvo el Dígito de Control dependiendo del primer carácter
        ' del NIF pasado a la función.
        '
        Dim firstChar As Char = nif.Chars(0)

        Select Case firstChar
            Case "N"c, "P"c, "Q"c, "R"c, "S"c, "W"c, "K"c, "L"c, "M"c
                ' NIF de entidades cuyo dígito de control se corresponde
                ' con una letra. Se incluyen las letras K, L y M porque
                ' el cálculo del dígito de control es el mismo que para
                ' el CIF.
                '
                ' Al estar los índices de los arrays en base cero, el primer
                ' elemento del array se corresponderá con la unidad del número
                ' 10, es decir, el número cero.
                '
                Return characters(sumaTotal)

            Case "C"c
                If ((lastChar = CStr(sumaTotal)) OrElse (lastChar = dc)) Then
                    ' Devuelvo el dígito de control pasado, que
                    ' puede ser un número o una letra.
                    Return lastChar
                Else
                    ' Devuelvo la letra correspondiente al cálculo
                    ' del dígito de control del NIF.
                    Return dc
                End If
            Case Else
                ' NIF de las restantes entidades, cuyo dígito de control es un número.
                '
                Return CChar(CStr(sumaTotal))
        End Select

    End Function

    Public Function TipoDatoOleDBGenerico(ByVal StringConexion As String, ByVal Tabla As String, ByVal Campo As String) As Int16

        ' Esta Funcion devuelve un entero identificador del Tipo de Datos Generico del Campo+Tabla a consultar;
        '       0 = Vacio o Indeterminado o No Tratado
        '       1 = Numerico
        '       2 = String de Caracteres
        '       3 = Logico
        '       4 = DateTime
        '       5 = Time

        Dim TipoOleDB As OleDbType = TipoDatoOleDB(StringConexion, Tabla, Campo)
        Select Case TipoOleDB
            Case OleDbType.BigInt, OleDbType.Currency, OleDbType.Decimal, OleDbType.Integer, OleDbType.Double, OleDbType.Numeric, OleDbType.Single, OleDbType.SmallInt, OleDbType.TinyInt, OleDbType.UnsignedBigInt, OleDbType.UnsignedInt, OleDbType.UnsignedSmallInt, OleDbType.UnsignedTinyInt, OleDbType.VarNumeric
                Return 1
            Case OleDbType.BSTR, OleDbType.Char, OleDbType.LongVarChar, OleDbType.LongVarWChar, OleDbType.VarWChar, OleDbType.WChar
                Return 2
            Case OleDbType.Boolean
                Return 3
            Case OleDbType.Date, OleDbType.DBDate, OleDbType.DBTimeStamp
                Return 4
            Case OleDbType.DBTime
                Return 5
            Case Else
                Return 0
        End Select

    End Function

    Public Function TipoDatoOleDB(ByVal StringConexion As String, ByVal Tabla As String, ByVal Campo As String) As OleDbType

        If StringConexion Is Nothing Then
            Throw New ArgumentNullException("StringConexion")
        End If
        If Campo Is Nothing Then
            Throw New ArgumentNullException("Campo")
        End If

        'The numbers are defined in Metadata. You could cast it to (OleDbType). 

        'using System;
        'namespace System.Data.OleDb 
        'public enum OleDbType 

        'Specifies the data type of a field, a property, for use in an System.Data.OleDb.OleDbParameter.

        'No value (DBTYPE_EMPTY).
        'Empty = 0,
        'A 16-bit signed integer (DBTYPE_I2). This maps to System.Int16.
        'SmallInt = 2,
        'A 32-bit signed integer (DBTYPE_I4). This maps to System.Int32.
        'Integer = 3,
        'A floating-point number within the range of -3.40E +38 through 3.40E +38 (DBTYPE_R4). This maps to System.Single.
        'Single = 4,
        'A floating-point number within the range of -1.79E +308 through 1.79E +308 (DBTYPE_R8). This maps to System.Double.
        'Double = 5,
        'A currency value ranging from -2 63 (or -922,337,203,685,477.5808) to 2 63 -1 (or +922,337,203,685,477.5807) with an accuracy to a ten-thousandth of a currency unit (DBTYPE_CY). This maps to System.Decimal.
        'Currency = 6,
        'Date data, stored as a double (DBTYPE_DATE). The whole portion is the number of days since December 30, 1899, and the fractional portion is a fraction of a day. This maps to System.DateTime.
        'Date = 7,
        'A null-terminated character string of Unicode characters (DBTYPE_BSTR). This maps to System.String.
        'BSTR = 8,
        'A pointer to an IDispatch interface (DBTYPE_IDISPATCH). This maps to System.Object.
        'IDispatch = 9,
        'A 32-bit error code (DBTYPE_ERROR). This maps to System.Exception.
        'Error = 10,
        'A Boolean value (DBTYPE_BOOL). This maps to System.Boolean.
        'Boolean = 11,
        'A special data type that can contain numeric, string, binary, or date data, and also the special values Empty and Null (DBTYPE_VARIANT). This type is assumed if no other is specified. This maps to System.Object.
        'Variant = 12,
        'A pointer to an IUnknown interface (DBTYPE_UNKNOWN). This maps to System.Object.
        'Unknown = 13,
        'A fixed precision and scale numeric value between -10 38 -1 and 10 38 -1 (DBTYPE_DECIMAL). This maps to System.Decimal.
        'Decimal = 14,
        'A 8-bit signed integer (DBTYPE_I1). This maps to System.SByte.
        'TinyInt = 16,
        'A 8-bit unsigned integer (DBTYPE_UI1). This maps to System.Byte.
        'UnsignedTinyInt = 17,
        'A 16-bit unsigned integer (DBTYPE_UI2). This maps to System.UInt16.
        'UnsignedSmallInt = 18,
        'A 32-bit unsigned integer (DBTYPE_UI4). This maps to System.UInt32.
        'UnsignedInt = 19,
        'BigInt = 20,
        '64-bit unsigned integer (DBTYPE_UI8). This maps to System.UInt64.
        'UnsignedBigInt = 21,
        'A 64-bit unsigned integer representing the number of 100-nanosecond intervals since January 1, 1601 (DBTYPE_FILETIME). This maps to System.DateTime.
        'Filetime = 64,
        'A globally unique identifier (or GUID) (DBTYPE_GUID). This maps to System.Guid.
        'Guid = 72,
        'A stream of binary data (DBTYPE_BYTES). This maps to an System.Array of type System.Byte.
        'Binary = 128,
        'A character string (DBTYPE_STR). This maps to System.String.
        'Char = 129,
        'A null-terminated stream of Unicode characters (DBTYPE_WSTR). This maps to System.String.
        'WChar = 130,
        '// An exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC). This maps to System.Decimal.
        'Numeric = 131,
        'Date data in the format yyyymmdd (DBTYPE_DBDATE). This maps to System.DateTime.
        'DBDate = 133,
        'Time data in the format hhmmss (DBTYPE_DBTIME). This maps to System.TimeSpan.
        'DBTime = 134,
        'Data and time data in the format yyyymmddhhmmss (DBTYPE_DBTIMESTAMP). This maps to System.DateTime.
        'DBTimeStamp = 135,
        'An automation PROPVARIANT (DBTYPE_PROP_VARIANT). This maps to System.Object.
        'PropVariant = 138,
        'A variable-length numeric value (System.Data.OleDb.OleDbParameter only). This maps to System.Decimal.
        'VarNumeric = 139,
        'A variable-length stream of non-Unicode characters (System.Data.OleDb.OleDbParameter only). This maps to System.String.
        'VarChar = 200,
        'A long string value (System.Data.OleDb.OleDbParameter only). This maps to System.String.
        'LongVarChar = 201,
        'A variable-length, null-terminated stream of Unicode characters (System.Data.OleDb.OleDbParameter only). This maps to System.String.
        'VarWChar = 202,
        'A long null-terminated Unicode string value (System.Data.OleDb.OleDbParameter only). This maps to System.String.
        'LongVarWChar = 203,
        'A variable-length stream of binary data (System.Data.OleDb.OleDbParameter only). This maps to an System.Array of type System.Byte.
        'VarBinary = 204,
        'A long binary value (System.Data.OleDb.OleDbParameter only). This maps to an System.Array of type System.Byte.
        'LongVarBinary = 205,

        Dim TablaEsquema As DataTable
        Dim TipoOleDB As OleDbType

        Using Conexion As New OleDbConnection
            Conexion.ConnectionString = StringConexion
            Conexion.Open()
            TablaEsquema = Conexion.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, Tabla, Nothing})
            'Conexion.Close()
        End Using

        If TablaEsquema.Rows.Count = 0 Then
            MiMessageBox.ShowWinMessage("La Tabla '" + Tabla + "', cuyo Esquema a consultar ..." + vbCrLf + vbCrLf + "NO Existe", "Llamada a Función TipoDatoOleDB(...)", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            TipoOleDB = 0
        Else
            Dim i As Integer
            Dim SwOK As Boolean = False
            For i = 0 To TablaEsquema.Rows.Count - 1
                If TablaEsquema.Rows(i).Item("COLUMN_NAME").ToString = Campo.Trim Then
                    SwOK = True
                    Exit For
                End If
            Next i
            If SwOK Then
                TipoOleDB = TablaEsquema.Rows(i).Item("DATA_TYPE")
            Else
                MiMessageBox.ShowWinMessage("El Campo '" + Campo + "', cuyo Tipo de Datos a consultar ..." + vbCrLf + vbCrLf + "NO Existe en la Tabla '" + Tabla + "'", "Llamada a Función TipoDatoOleDB(...)", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
                TipoOleDB = 999
            End If
        End If

        Return TipoOleDB

    End Function

    Public Function TipoDatoSqlGenerico(ByVal StringConexion As String, ByVal Tabla As String, ByVal Campo As String) As Int16

        ' Esta Funcion devuelve un entero identificador del Tipo de Datos Generico del Campo+Tabla a consultar;
        '       0 = Vacio o Indeterminado o No Tratado
        '       1 = Numerico
        '       2 = String de Caracteres o Char
        '       3 = Logico
        '       4 = DateTime
        '       5 = Time
        '       6 = Binario o Bit

        Dim TipoSql As SqlDbType = TipoDatoSql(StringConexion, Tabla, Campo)

        Select Case TipoSql
            Case SqlDbType.BigInt, SqlDbType.Decimal, SqlDbType.Float, SqlDbType.Int, SqlDbType.Money, SqlDbType.Real, SqlDbType.SmallInt, SqlDbType.SmallMoney, SqlDbType.TinyInt
                Return 1
            Case SqlDbType.Char, SqlDbType.NChar, SqlDbType.NText, SqlDbType.NVarChar, SqlDbType.Text, SqlDbType.VarChar
                Return 2
            Case OleDbType.Boolean
                Return 3
            Case SqlDbType.Date, SqlDbType.DateTime, SqlDbType.DateTime2, SqlDbType.DateTimeOffset, SqlDbType.SmallDateTime, SqlDbType.Timestamp
                Return 4
            Case SqlDbType.Time
                Return 5
            Case SqlDbType.Binary, SqlDbType.Bit, SqlDbType.VarBinary
                Return 6
            Case Else
                Return 0
        End Select

    End Function

    Public Function TipoDatoSql(ByVal StringConexion As String, ByVal Tabla As String, ByVal Campo As String) As SqlDbType

        If StringConexion Is Nothing Then
            Throw New ArgumentNullException("StringConexion")
        End If
        If Campo Is Nothing Then
            Throw New ArgumentNullException("Campo")
        End If

        Dim TablaEsquema As DataTable
        Dim TipoSql As SqlDbType

        Using Conexion As New SqlConnection
            Conexion.ConnectionString = StringConexion
            Conexion.Open()
            TablaEsquema = Conexion.GetSchema("Columns", New String() {Nothing, Nothing, Tabla, Nothing})
            Conexion.Close()
        End Using

        If TablaEsquema.Rows.Count = 0 Then
            MiMessageBox.ShowWinMessage("La Tabla '" + Tabla + "', cuyo Esquema a consultar ..." + vbCrLf + vbCrLf + "NO Existe", "Llamada a Función TipoDatoSql(...)", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            TipoSql = 0
        Else
            Dim i As Integer
            Dim SwOK As Boolean = False
            For i = 0 To TablaEsquema.Rows.Count - 1
                If TablaEsquema.Rows(i).Item("COLUMN_NAME").ToString = Campo.Trim Then
                    SwOK = True
                    Exit For
                End If
            Next i
            If SwOK Then
                TipoSql = [Enum].Parse(GetType(SqlDbType), TablaEsquema.Rows(i).Item("DATA_TYPE"), True)
            Else
                MiMessageBox.ShowWinMessage("El Campo '" + Campo + "', cuyo Tipo de Datos a consultar ..." + vbCrLf + vbCrLf + "NO Existe en la Tabla '" + Tabla + "'", "Llamada a Función TipoDatoSql(...)", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
                TipoSql = 999
            End If
        End If

        Return TipoSql

    End Function

    Public Function MesesDiferencia(ByVal FechaFin As DateTime, ByVal FechaInicio As DateTime) As Int16

        'Return Math.Abs((FechaFin.Month - FechaInicio.Month) + 12 * (FechaFin.Year - FechaInicio.Year))

        Return DateDiff(DateInterval.Month, FechaInicio, FechaFin)

    End Function

    Public Function DiasDiferencia(ByVal FechaFin As DateTime, ByVal FechaInicio As DateTime) As Int16

        'Difference in days, hours, And minutes.
        Dim DifFecha As TimeSpan = FechaFin - FechaInicio

        Return DateDiff(DateInterval.Day, FechaInicio, FechaFin)
        'Return DifFecha.Days

    End Function

    Public Function RtbResaltarTexto(ByVal Rtb As RichTextBox,
                                     ByVal TextoBusqueda As String,
                                     ByVal IgnorarCase As Boolean,
                                     ByVal Font As System.Drawing.Font,
                                     ByVal ForeColor As System.Drawing.Color,
                                     ByVal BackColor As System.Drawing.Color) As Integer

        If String.IsNullOrEmpty(TextoBusqueda) Then
            Return 0
        End If

        'Establece criterio letter-case.
        Dim RichTextBoxFinds As RichTextBoxFinds = If(IgnorarCase, RichTextBoxFinds.None, RichTextBoxFinds.MatchCase)
        Dim StringComparison As StringComparison = If(IgnorarCase, StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)

        'Salva la actual 'Caret Position' para restaurarla al final.
        Dim CaretPosition As Integer = Rtb.SelectionStart

        Dim CuentaOcurrencias As Integer = 0
        Dim LongitudTexto As Integer = Rtb.TextLength
        Dim PrimerIndice As Integer = 0
        Dim UltimoIndice As Integer = Rtb.Text.LastIndexOf(TextoBusqueda, StringComparison)

        While (PrimerIndice <= UltimoIndice)

            Dim findIndex As Integer = Rtb.Find(TextoBusqueda, PrimerIndice, LongitudTexto, RichTextBoxFinds)
            If (findIndex <> -1) Then
                CuentaOcurrencias += 1
            Else
                Continue While
            End If

            Rtb.SelectionColor = ForeColor
            Rtb.SelectionBackColor = BackColor
            Rtb.SelectionFont = Font

            PrimerIndice = (Rtb.Text.IndexOf(TextoBusqueda, findIndex, StringComparison) + 1)

        End While ' (PrimerIndice <= UltimoIndice)

        'Restaura la 'Caret Position'. Reset 'Selection Length' a cero.
        Rtb.Select(CaretPosition, length:=0)

        Return CuentaOcurrencias

    End Function

    Function GetAlineacionHorizontal(ByRef Alineacion As String) As AlineacionHorizontal

        Select Case Alineacion.Trim.ToUpper
            Case "LEFT", "TOPLEFT", "MIDDLELEFT", "BOTTOMLEFT"
                Return AlineacionHorizontal.LEFT
            Case "CENTER", "TOPCENTER", "MIDDLECENTER", "BOTTOMCENTER"
                Return AlineacionHorizontal.CENTER
            Case "RIGHT", "TOPRIGHT", "MIDDLERIGHT", "BOTTOMRIGHT"
                Return AlineacionHorizontal.RIGHT
            Case Else
                Return AlineacionHorizontal.NOTSET
        End Select

    End Function

    Public Function GetBufferedString(StringOriginal As String, LongitudFijada As Int16, AlineacionHorizontal As AlineacionHorizontal) As String

        ' ===================================================================================
        ' Devuelve un String Ajustado, tal como indicado, y fijado a la longitud especificada.
        ' Util para alinear concatenaciones de strings de Salida
        ' ===================================================================================

        If (StringOriginal.Length < LongitudFijada) Then
            Dim BufString As String = Space(LongitudFijada - StringOriginal.Length)
            Select Case AlineacionHorizontal
                Case AlineacionHorizontal.LEFT
                    Return StringOriginal + BufString
                Case AlineacionHorizontal.RIGHT
                    Return BufString + StringOriginal
                Case AlineacionHorizontal.CENTER
                    Dim HalfString As String = BufString.Substring(BufString.Length / 2)
                    StringOriginal = HalfString + StringOriginal
                    BufString = Space(LongitudFijada - StringOriginal.Length)
                    Return StringOriginal + BufString
                Case Else
                    Return StringOriginal + BufString
            End Select
        Else
            Return StringOriginal.Substring(0, LongitudFijada)
        End If

    End Function

    Public Function GetColumnasVisiblesWidth(ByVal DataGrid As DataGridView) As Single()

        '=======================================================================================
        'Devuelve una Tabla con el Ancho de cada una de las columnas Visibles de un DatagridView
        '=======================================================================================

        Dim Valores() As Single
        Dim n As Int16 = 0
        For i As Integer = 0 To DataGrid.ColumnCount - 1
            If DataGrid.Columns(i).Visible Then
                ReDim Preserve Valores(n)
                Valores(n) = DataGrid.Columns(i).Width
                n += 1
            End If
        Next
        Return Valores

    End Function

    Public Function GetColumnasVisiblesTotalWidth(ByVal DataGrid As DataGridView) As Integer

        '=======================================================================================
        'Devuelve la Suma Total de todos los Anchoa de las columnas Visibles de un DatagridView
        '=======================================================================================

        Dim AnchoTotal As Integer = 0
        For i As Integer = 0 To DataGrid.ColumnCount - 1
            If DataGrid.Columns(i).Visible Then
                AnchoTotal += DataGrid.Columns(i).Width
            End If
        Next
        Return AnchoTotal

    End Function

    Public Function GetColumnasVisiblesNum(ByVal DataGrid As DataGridView) As Int16

        '==========================================================
        'Devuelve el Numero de Columnas Visibles de un DatagridView
        '==========================================================

        Dim Contador As Int16 = 0
        For i As Integer = 0 To DataGrid.ColumnCount - 1
            If DataGrid.Columns(i).Visible Then
                Contador += 1
            End If
        Next
        Return Contador

    End Function

    Public Function ExportarDataGridaTXT(ByRef DataGrid As DataGridView, Optional ByVal NombreFichero As String = "", Optional ByVal CrearEnTemp As Boolean = False, Optional ByVal MensajeProcesando As String = "") As DialogResult

        '===============================================================================================
        'Exporta un DataGridView a Fichero de TEXTO.
        '===============================================================================================

        If DataGrid.RowCount = 0 Then
            MiMessageBox.ShowWinMessage("El DataGrid NO tiene Filas.", "Exportar DataGrid a .TXT", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            Return DialogResult.None
        Else

            If String.IsNullOrWhiteSpace(NombreFichero) Then
                NombreFichero = DataGrid.Name + ".txt"
            End If
            If Path.GetExtension(NombreFichero) = String.Empty Then
                NombreFichero += ".txt"
            End If

                Dim FicheroDestino As String
            If CrearEnTemp Then
                FicheroDestino = System.IO.Path.GetTempPath() + NombreFichero
            Else
                Using OFD As New OpenFileDialog()
                    OFD.Filter = "Archivos de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
                    OFD.CheckFileExists = False
                    OFD.AddExtension = True
                    OFD.DefaultExt = "txt"
                    OFD.Multiselect = False
                    OFD.RestoreDirectory = True
                    OFD.FileName = NombreFichero
                    If OFD.ShowDialog() = DialogResult.Cancel Then
                        MiMessageBox.ShowWinMessage("Exportación CANCELDA por el Usuario.", "Exportar DataGrid a TXT", MsgBoxStyle.Information, MsgBoxStyle.OkOnly)
                        Return DialogResult.Cancel
                    End If
                    FicheroDestino = OFD.FileName
                End Using
            End If

            Dim VMensaje As New MiVentanaMensaje(200, 1, IIf(MensajeProcesando = String.Empty, "Exportando DataGrid a TXT...", MensajeProcesando))

            Try
                Using TW As TextWriter = New StreamWriter(FicheroDestino)
                    Dim TextoSalida As String
                    For i As Integer = 0 To DataGrid.ColumnCount - 1
                        If DataGrid.Columns(i).Visible Then
                            TextoSalida = GetBufferedString(DataGrid.Columns(i).HeaderText, DataGrid.Columns(i).Width / 8, GetAlineacionHorizontal(DataGrid.Columns(i).DefaultCellStyle.Alignment.ToString)) + Space(1)
                            TW.Write(TextoSalida)
                        End If
                    Next
                    TW.WriteLine()
                    For i As Integer = 0 To DataGrid.ColumnCount - 1
                        If DataGrid.Columns(i).Visible Then
                            TextoSalida = StrDup(CInt(DataGrid.Columns(i).Width / 8), "-") + Space(1)
                            TW.Write(TextoSalida)
                        End If
                    Next
                    TW.WriteLine()
                    For i As Integer = 0 To DataGrid.RowCount - 1
                        For j As Integer = 0 To DataGrid.ColumnCount - 1
                            If DataGrid.Rows(i).Cells(j).Visible Then
                                If DataGrid.Rows(i).Cells(j).ValueType = System.Type.GetType("System.Boolean") Then
                                    TextoSalida = GetBufferedString(IIf(DataGrid.Rows(i).Cells(j).Value, "V", "F"), DataGrid.Columns(j).Width / 8, GetAlineacionHorizontal(DataGrid.Columns(j).DefaultCellStyle.Alignment.ToString)) + Space(1)
                                Else
                                    TextoSalida = GetBufferedString(Convert.ToString(DataGrid.Rows(i).Cells(j).FormattedValue), DataGrid.Columns(j).Width / 8, GetAlineacionHorizontal(DataGrid.Columns(j).DefaultCellStyle.Alignment.ToString)) + Space(1)
                                End If
                                TW.Write(TextoSalida)
                                End If
                        Next
                        If i < (DataGrid.RowCount - 1) Then
                            TW.WriteLine()
                        End If
                    Next
                    TW.Close()
                    VMensaje.Close()
                    Return DialogResult.OK
                End Using
            Catch Ex As Exception
                VMensaje.Close()
                MiMessageBox.ShowWinMessage("Se ha producido una Excepción en el proceso de Exportación del DataGrid;" + vbCrLf + vbCrLf + Ex.Message, "Exportar DataGrid a TXT", MsgBoxStyle.Critical, MsgBoxStyle.OkOnly)
                Return DialogResult.Abort
            End Try

        End If

    End Function

    Public Function ExportarDataGridaEXCEL_OfficeInterop(ByRef DataGrid As DataGridView, Optional ByVal NombreFichero As String = "", Optional ByVal CrearEnTemp As Boolean = False, Optional ByVal MensajeProcesando As String = "") As DialogResult

        '===============================================================================================
        'Exporta un DataGridView a Hoja Excel utilizando los servicios "Microsoft.Office.Interop.Excel".
        '* Aún, No tiene codificado la aplicacion de formatos de salida (Celdas de la Hoja de Calculo).
        'Es muy LENTO. Se recomienda utilizar la funcion  "ExportarDataGridaEXCEL" que es mucho mas RAPIDA.
        '===============================================================================================


        If DataGrid.RowCount = 0 Then
            MiMessageBox.ShowWinMessage("El DataGrid NO tiene Filas.", "Exportar DataGrid a EXCEL", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            Return DialogResult.None
        Else

            If String.IsNullOrWhiteSpace(NombreFichero) Then
                NombreFichero = DataGrid.Name + ".xlsx"
            End If
            If Path.GetExtension(NombreFichero) = String.Empty Then
                NombreFichero += ".xlsx"
            End If

            Dim ExApp As New Microsoft.Office.Interop.Excel.Application
            Dim ExLibro As Microsoft.Office.Interop.Excel.Workbook = ExApp.Workbooks.Add
            Dim ExHoja As Microsoft.Office.Interop.Excel.Worksheet

            ExApp.Visible = False 'Para que no se muestre mientras se crea
            ExHoja = ExLibro.ActiveSheet

            Dim FicheroDestino As String
            If CrearEnTemp Then
                FicheroDestino = System.IO.Path.GetTempPath() + NombreFichero
            Else
                Using OFD As New OpenFileDialog()
                    OFD.Filter = "Archivo Excel | *.xlsx|Todos los archivos (*.*)|*.*"
                    OFD.CheckFileExists = False
                    OFD.AddExtension = True
                    OFD.DefaultExt = "xlsx"
                    OFD.Multiselect = False
                    OFD.RestoreDirectory = True
                    OFD.FileName = NombreFichero
                    If OFD.ShowDialog() = DialogResult.Cancel Then
                        MiMessageBox.ShowWinMessage("Exportación CANCELDA por el Usuario.", "Exportar DataGrid a EXCEL", MsgBoxStyle.Information, MsgBoxStyle.OkOnly)
                        Return DialogResult.Cancel
                    End If
                    FicheroDestino = OFD.FileName
                End Using
            End If

            Dim VMensaje As New MiVentanaMensaje(200, 1, IIf(MensajeProcesando = String.Empty, "Exportando DataGrid a EXCEL...", MensajeProcesando))

            Try
                Dim k As Int16 = 1
                For i As Integer = 0 To DataGrid.ColumnCount - 1
                    If DataGrid.Columns(i).Visible Then
                        ExHoja.Cells.Item(1, k).Value = DataGrid.Columns(i).HeaderText
                        ExHoja.Cells.Item(1, k).Font.Bold = True
                        ExHoja.Cells.Item(1, k).Font.Size = 12
                        ExHoja.Cells.Item(1, k).Font.Color = System.Drawing.Color.White
                        ExHoja.Cells.Item(1, k).Interior.Color = System.Drawing.Color.Navy
                        k += 1
                    End If
                Next
                For i As Integer = 0 To DataGrid.RowCount - 1
                    k = 1
                    For j As Integer = 0 To DataGrid.ColumnCount - 1
                        If DataGrid.Rows(i).Cells(j).Visible Then
                            ExHoja.Cells.Item(i + 2, k).Value = Convert.ToString(DataGrid.Rows(i).Cells(j).FormattedValue)
                            k += 1
                        End If
                    Next
                Next
                ExLibro.SaveAs(FicheroDestino)
                ExApp.Workbooks.Close()
                ExApp.Quit()
                VMensaje.Close()
                Return DialogResult.OK
            Catch Ex As Exception
                VMensaje.Close()
                MiMessageBox.ShowWinMessage("Se ha producido una Excepción en el proceso de Exportación del DataGrid;" + vbCrLf + vbCrLf + Ex.Message, "Exportar DataGrid a EXCEL", MsgBoxStyle.Critical, MsgBoxStyle.OkOnly)
                Return DialogResult.Abort
            End Try

        End If

    End Function

    Public Function EnviaEmailDataGridEXCEL(ByRef DataGrid As DataGridView, Optional ByVal NombreFichero As String = "", Optional ByVal TituloInforme As String = "") As DialogResult

        If Not MSOutlook.EstaOutlookInstalado() Then
            MiMessageBox.ShowWinMessage("Parece que OUTLOOK NO está INSTALADO en el Equipo." + vbCr + vbCr + "eMail NO puede ser ENVIADO.", "Enviar DataGrid por Email", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            Return DialogResult.Abort
        End If

        If String.IsNullOrWhiteSpace(NombreFichero) Then
            NombreFichero = DataGrid.Name + ".xlsx"
        End If
        If Path.GetExtension(NombreFichero) = String.Empty Then
            NombreFichero += ".xlsx"
        End If

        'ExportarDataGridaXXX(), cuando el Parametro CrearEnTemp=True, crea el fichero en "System.IO.Path.GetTempPath()"
        If ExportarDataGridaEXCEL(DataGrid, NombreFichero, TituloInforme, True, "Enviando DataGrid por eMail...") = DialogResult.OK Then 'El fichero se crea en "&UserConfigPath\Temp" con del nombre indicado o DataGridView.Name
            Dim EMailCuerpo As String = "<p Class=MsoNormal>Hola, <o:p></o:p></p>" +
                                        "<p Class=MsoNormal><span style='mso-tab-count:1'>" + Chr(CInt("&H0A")) + "</span>Adjunto, enviado el <b><i>Fichero de referencia</i></b>.<o:p></o:p></p>" +
                                        "<p Class=MsoNormal>Atentamente,<o:p></o:p></p>"
            Try
                Using Enviar As New MSOutlook()
                    Enviar.EnviarEMail(System.IO.Path.GetTempPath() + NombreFichero, IIf(Not String.IsNullOrWhiteSpace(TituloInforme), TituloInforme, IIf(NombreFichero = DataGrid.Name + ".xlsx", "DataGrid " + DataGrid.Name + " en Formato EXCEL ...", NombreFichero)), EMailCuerpo, , False, True)
                End Using
                File.Delete(System.IO.Path.GetTempPath() + NombreFichero)
                Return DialogResult.OK
            Catch ex As Exception
                MiMessageBox.ShowWinMessage("Se ha producido un ERROR en el proceso de enviar DataGrid por eMail;" + vbCrLf + vbCrLf + ex.Message, "Enviar DataGrid por Email", MsgBoxStyle.Critical, MsgBoxStyle.OkOnly)
                Return DialogResult.Abort
            End Try
        Else
            Return DialogResult.Abort
        End If

    End Function

    Public Function EnviaEmailDataGridPDF(ByRef DataGrid As DataGridView, Optional ByVal NombreFichero As String = "", Optional ByVal TituloInforme As String = "") As DialogResult

        If Not MSOutlook.EstaOutlookInstalado() Then
            MiMessageBox.ShowWinMessage("Parece que OUTLOOK NO está INSTALADO en el Equipo." + vbCr + vbCr + "eMail NO puede ser ENVIADO.", "Enviar DataGrid por Email", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            Return DialogResult.Abort
        End If

        If String.IsNullOrWhiteSpace(NombreFichero) Then
            NombreFichero = DataGrid.Name + ".pdf"
        End If
        If Path.GetExtension(NombreFichero) = String.Empty Then
            NombreFichero += ".pdf"
        End If

        'ExportarDataGridaXXX(), cuando el Parametro CrearEnTemp=True, crea el fichero en "System.IO.Path.GetTempPath()"
        If ExportarDataGridaPDF(DataGrid, NombreFichero, TituloInforme, True, "Enviando DataGrid por eMail...") = DialogResult.OK Then 'El fichero se crea en "&UserConfigPath\Temp" con del nombre indicado o DataGridView.Name
            Dim EMailCuerpo As String = "<p Class=MsoNormal>Hola, <o:p></o:p></p>" +
                                        "<p Class=MsoNormal><span style='mso-tab-count:1'>" + Chr(CInt("&H0A")) + "</span>Adjunto, enviado el <b><i>Fichero de referencia</i></b>.<o:p></o:p></p>" +
                                        "<p Class=MsoNormal>Atentamente,<o:p></o:p></p>"
            Try
                Using Enviar As New MSOutlook()
                    Enviar.EnviarEMail(System.IO.Path.GetTempPath() + NombreFichero, IIf(Not String.IsNullOrWhiteSpace(TituloInforme), TituloInforme, IIf(NombreFichero = DataGrid.Name + ".pdf", "DataGrid " + DataGrid.Name + " en Formato PDF ...", NombreFichero)), EMailCuerpo, , False, True)
                End Using
                File.Delete(System.IO.Path.GetTempPath() + NombreFichero)
                Return DialogResult.OK
            Catch ex As Exception
                MiMessageBox.ShowWinMessage("Se ha producido un ERROR en el proceso de enviar DataGrid por eMail;" + vbCrLf + vbCrLf + ex.Message, "Enviar DataGrid por Email", MsgBoxStyle.Critical, MsgBoxStyle.OkOnly)
                Return DialogResult.Abort
            End Try
        Else
            Return DialogResult.Abort
        End If

    End Function

    Public Function EnviaEmailDataGridTXT(ByRef DataGrid As DataGridView, Optional ByVal NombreFichero As String = "") As DialogResult

        If Not MSOutlook.EstaOutlookInstalado() Then
            MiMessageBox.ShowWinMessage("Parece que OUTLOOK NO está INSTALADO en el Equipo." + vbCr + vbCr + "eMail NO puede ser ENVIADO.", "Enviar DataGrid por Email", MsgBoxStyle.Exclamation, MsgBoxStyle.OkOnly)
            Return DialogResult.Abort
        End If

        If String.IsNullOrWhiteSpace(NombreFichero) Then
            NombreFichero = DataGrid.Name + ".txt"
        End If
        If Path.GetExtension(NombreFichero) = String.Empty Then
            NombreFichero += ".txt"
        End If

        'ExportarDataGridaXXX(), cuando el Parametro CrearEnTemp=True, crea el fichero en "System.IO.Path.GetTempPath()"
        If ExportarDataGridaTXT(DataGrid, NombreFichero, True, "Enviando DataGrid por eMail...") = DialogResult.OK Then 'El fichero se crea en "&UserConfigPath\Temp" con del nombre indicado o DataGridView.Name
            Dim EMailCuerpo As String = "<p Class=MsoNormal>Hola, <o:p></o:p></p>" +
                                        "<p Class=MsoNormal><span style='mso-tab-count:1'>" + Chr(CInt("&H0A")) + "</span>Adjunto, enviado el <b><i>Fichero de referencia</i></b>.<o:p></o:p></p>" +
                                        "<p Class=MsoNormal>Atentamente,<o:p></o:p></p>"
            Try
                Using Enviar As New MSOutlook()
                    Enviar.EnviarEMail(System.IO.Path.GetTempPath() + NombreFichero, IIf(NombreFichero = DataGrid.Name + ".txt", "DataGrid " + DataGrid.Name + " en Formato TXT ...", NombreFichero), EMailCuerpo, , False, True)
                End Using
                File.Delete(System.IO.Path.GetTempPath() + NombreFichero)
                Return DialogResult.OK
            Catch ex As Exception
                MiMessageBox.ShowWinMessage("Se ha producido un ERROR en el proceso de enviar DataGrid por eMail;" + vbCrLf + vbCrLf + ex.Message, "Enviar DataGrid por Email", MsgBoxStyle.Critical, MsgBoxStyle.OkOnly)
                Return DialogResult.Abort
            End Try
        Else
            Return DialogResult.Abort
        End If

    End Function

End Module
