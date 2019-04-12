Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class Form1

    ' Esta es tu conexion global
    Dim sqlConexion As SqlConnection

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        AbrirConexion()


    End Sub

    ' Abrea la conexion a sql Server
    Private Sub AbrirConexion()

        Try

            Dim sCadenaConexion As String = "Tu cadena de conexion al servidor"
            sqlConexion = New SqlConnection(sCadenaConexion)
            sqlConexion.Open()
            ' DataTable con los datos que se enviaran al sp
            Dim dtDatos As DataTable = New DataTable
            ' Columnas y tipo de dato que tiene la tabla de tipo type que es type_empleados
            dtDatos.Columns.Add("id", System.Type.GetType("System.String"))
            dtDatos.Columns.Add("nombre", System.Type.GetType("System.String"))
            dtDatos.Columns.Add("direccion", System.Type.GetType("System.String"))
            ' Datos que en este caso se van a insertar
            dtDatos.Rows.Add("444", "ZZZ", "AAA")

            Dim sqlComando = New SqlCommand()
            ' Ya esto te lo sabes :p
            sqlComando.Connection = sqlConexion
            sqlComando.CommandText = "spRegistrarEmpleado"
            sqlComando.CommandType = CommandType.StoredProcedure
            sqlComando.Parameters.AddWithValue("@datos_empleados", dtDatos)
            sqlComando.ExecuteNonQuery()

            sqlConexion.Close()


        Catch exc As Exception

        End Try

    End Sub

    ' Esta es la funcion que manda llamar otra funcion
    Private Sub FuncionUno()

        Dim sqlTransaccion As SqlTransaction

        Try
            sqlTransaccion = sqlConexion.BeginTransaction("NombreTransaccion")

            ' Aqui va un ciclo que manda llamar la FuncionDos()
            If FuncionDos() = False Then
                ' Al regresar false significa que algo esta mal entoces se hara el rollback
                sqlTransaccion.Rollback()
                ' Se hace el break de VB para salir del ciclo
            End If


        Catch exc As Exception

            sqlTransaccion.Rollback()

        End Try

    End Sub

    ' Esta es la funcion booleana de si todo salio bien regresa True, False si es lo contrario
    Private Function FuncionDos() As Boolean

        Dim resultado As Boolean

        Return resultado

    End Function

    'Public Function GuardarGestion(ByRef dsDatos As Data.DataSet, ByVal bolEsFinado As Boolean, ByVal intCodigoCobrador As Integer, ByVal intCodigoCobranza As Integer) As Boolean
    '    ' String
    '    Dim strSQL As String
    '    ' SqlConnection
    '    Dim scnConexion As New SqlConnection
    '    ' SqlCommand
    '    Dim scmdComando As New SqlCommand
    '    'drInformacion
    '    Dim drInformacion As Data.DataRow
    '    ' SqlTransacion
    '    Dim sTrans As SqlTransaction
    '    'Boolean
    '    Dim bolTodoBien As Boolean = True
    '    'DataTable
    '    Dim dtPrestamosVales As New DataTable

    '    drInformacion = dsDatos.Tables("Gestion").Rows(0)
    '    Try
    '        ' Aquí se establece y abre la conexión hacia la base de datos
    '        scnConexion.ConnectionString = MyBase.CadenaDeConexion
    '        scnConexion.Open()
    '        sTrans = scnConexion.BeginTransaction()
    '    Catch ex As Exception
    '        MyBase.OcurrioError = True
    '        MyBase.DescripcionDelError = "Ocurrió un error al conectarse a la Base de Datos. " & ex.Message
    '    End Try
    '    If Not MyBase.OcurrioError Then
    '        If drInformacion("Codigo") = -1 Then
    '            strSQL = "INSERT INTO [dbo].[GestionesCobranzaVales]([CodigoDistribuidor],[CodigoCliente],[CodigoPlaza],[CodigoTipo],[CodigoTitularDistribuidorGestionado],[CodigoTitularClienteValeGestionado],[CodigoAvalGestionado],[CodigoReferenciaGestionado],[CodigoPersonaViveGestionado],[CodigoCompromiso] ,[Comentario],[FechaGestion],[CodigoGestorEmpleado],[CodigoGestorCobrador],[Mensaje],[Fecha],[CodigoUsuario],[PersonaGestionada], [CodigoResultadoGestiones]) VALUES ("
    '            strSQL &= "" & drInformacion("CodigoDistribuidor") & ", "
    '            strSQL &= "" & IIf(Integer.Parse(drInformacion("CodigoCliente").ToString()) = -1, "Null", drInformacion("CodigoCliente")) & ", "
    '            strSQL &= "'" & drInformacion("CodigoPlaza") & "', "
    '            strSQL &= "'" & drInformacion("CodigoTipo") & "', "
    '            strSQL &= "" & IIf(Integer.Parse(drInformacion("CodigoTitularDistribuidorGestionado").ToString()) = -1, "Null", drInformacion("CodigoTitularDistribuidorGestionado")) & ", "
    '            strSQL &= "" & IIf(Integer.Parse(drInformacion("CodigoTitularClienteValeGestionado").ToString()) = -1, "Null", drInformacion("CodigoTitularClienteValeGestionado")) & ", "
    '            strSQL &= "" & IIf(Integer.Parse(drInformacion("CodigoAvalGestionado").ToString()) = -1, "Null", drInformacion("CodigoAvalGestionado")) & ", "
    '            strSQL &= "" & IIf(Integer.Parse(drInformacion("CodigoReferenciaGestionado").ToString()) = -1, "Null", drInformacion("CodigoReferenciaGestionado")) & ", "
    '            strSQL &= "" & IIf(Integer.Parse(drInformacion("CodigoPersonaViveGestionado").ToString()) = -1, "Null", drInformacion("CodigoPersonaViveGestionado")) & ", "
    '            'strSQL &= "'" & drInformacion("Resultado") & "', "
    '            strSQL &= "" & IIf(Integer.Parse(drInformacion("CodigoCompromiso").ToString()) = -1, "Null", drInformacion("CodigoCompromiso")) & ", "
    '            strSQL &= "'" & drInformacion("Comentario") & "', "
    '            strSQL &= "'" & drInformacion("FechaGestion") & "', "
    '            strSQL &= "" & IIf(Integer.Parse(drInformacion("CodigoGestorEmpleado").ToString()) = -1, "Null", drInformacion("CodigoGestorEmpleado")) & ", "
    '            strSQL &= "" & IIf(Integer.Parse(drInformacion("CodigoGestorCobrador").ToString()) = -1, "Null", drInformacion("CodigoGestorCobrador")) & ", "
    '            strSQL &= "'" & drInformacion("Mensaje") & "', "
    '            strSQL &= "'" & drInformacion("Fecha") & "', "
    '            strSQL &= "" & drInformacion("CodigoUsuario") & ", "
    '            strSQL &= "'" & drInformacion("PersonaGestionada") & "', "
    '            strSQL &= "" & drInformacion("CodigoResultadoGestiones") & ")"
    '        Else
    '            strSQL = "UPDATE [dbo].[GestionesCobranzaVales]"
    '            strSQL &= " SET [CodigoDistribuidor] = " & drInformacion("CodigoDistribuidor") & ", "
    '            strSQL &= "[CodigoCliente] = " & IIf(Integer.Parse(drInformacion("CodigoCliente").ToString()) = -1, "Null", drInformacion("CodigoCliente")) & ", "
    '            strSQL &= "[CodigoPlaza] = '" & drInformacion("CodigoPlaza") & "', "
    '            strSQL &= "[CodigoTipo] = '" & drInformacion("CodigoTipo") & "' ,"
    '            strSQL &= "[CodigoTitularDistribuidorGestionado] = " & IIf(Integer.Parse(drInformacion("CodigoTitularDistribuidorGestionado").ToString()) = -1, "Null", drInformacion("CodigoTitularDistribuidorGestionado")) & ", "
    '            strSQL &= "[CodigoTitularClienteValeGestionado] = " & IIf(Integer.Parse(drInformacion("CodigoTitularClienteValeGestionado").ToString()) = -1, "Null", drInformacion("CodigoTitularClienteValeGestionado")) & ", "
    '            strSQL &= "[CodigoAvalGestionado] = " & IIf(Integer.Parse(drInformacion("CodigoAvalGestionado").ToString()) = -1, "Null", drInformacion("CodigoAvalGestionado")) & ", "
    '            strSQL &= "[CodigoReferenciaGestionado] = " & IIf(Integer.Parse(drInformacion("CodigoReferenciaGestionado").ToString()) = -1, "Null", drInformacion("CodigoReferenciaGestionado")) & ", "
    '            strSQL &= "[CodigoPersonaViveGestionado] = " & IIf(Integer.Parse(drInformacion("CodigoPersonaViveGestionado").ToString()) = -1, "Null", drInformacion("CodigoPersonaViveGestionado")) & ", "
    '            'strSQL &= "[Resultado] = '" & drInformacion("Resultado") & "', "
    '            strSQL &= "[CodigoCompromiso] = " & IIf(Integer.Parse(drInformacion("CodigoCompromiso").ToString()) = -1, "Null", drInformacion("CodigoCompromiso")) & ", "
    '            strSQL &= "[Comentario] = '" & drInformacion("Comentario") & "', "
    '            strSQL &= "[FechaGestion] = '" & drInformacion("FechaGestion") & "', "
    '            strSQL &= "[CodigoGestorEmpleado] = " & IIf(Integer.Parse(drInformacion("CodigoGestorEmpleado").ToString()) = -1, "Null", drInformacion("CodigoGestorEmpleado")) & ", "
    '            strSQL &= "[CodigoGestorCobrador] = " & IIf(Integer.Parse(drInformacion("CodigoGestorCobrador").ToString()) = -1, "Null", drInformacion("CodigoGestorCobrador")) & ", "
    '            strSQL &= "[Mensaje] = '" & drInformacion("Mensaje") & "', "
    '            strSQL &= "[PersonaGestionada] = '" & drInformacion("PersonaGestionada") & "', "
    '            strSQL &= "[CodigoResultadoGestiones] = " & drInformacion("CodigoResultadoGestiones") & " "
    '            'strSQL &= "[Fecha] = '" & drInformacion("Fecha") & "', "
    '            'strSQL &= "[CodigoUsuario] = " & drInformacion("CodigoUsuario") & ""
    '            strSQL &= "Where Codigo = " & drInformacion("Codigo")
    '        End If
    '        Try
    '            ' Aquí se establecen los parámetros y se ejecuta el comando
    '            scmdComando.CommandText = strSQL
    '            scmdComando.Connection = scnConexion
    '            scmdComando.Transaction = sTrans
    '            scmdComando.ExecuteNonQuery()
    '            If drInformacion("Codigo") = -1 Then
    '                strSQL = "Select @@Identity From GestionesCobranzaVales (NoLock)"
    '                scmdComando.CommandText = strSQL
    '                drInformacion("Codigo") = scmdComando.ExecuteScalar
    '                scmdComando.Transaction = sTrans
    '            End If
    '            If ObtenerPrestamosValesDeLaGestion(Integer.Parse(drInformacion("CodigoDistribuidor").ToString), intCodigoCobrador, intCodigoCobranza, Integer.Parse(drInformacion("Codigo").ToString), dtPrestamosVales, IIf(Integer.Parse(drInformacion("CodigoCliente").ToString) > -1, False, True), Integer.Parse(drInformacion("CodigoCliente").ToString)) Then
    '                If Not GuardarPrestamosValesDeLaGestion(sTrans, scnConexion, scmdComando, Integer.Parse(drInformacion("Codigo").ToString), dtPrestamosVales) Then
    '                    sTrans.Rollback()
    '                    bolTodoBien = False
    '                Else
    '                    If Not GuardarDistribuidorFinado(sTrans, scnConexion, scmdComando, Integer.Parse(drInformacion("CodigoDistribuidor")), bolEsFinado) Then
    '                        sTrans.Rollback()
    '                        bolTodoBien = False
    '                    Else
    '                        sTrans.Commit()
    '                    End If
    '                End If
    '            End If
    '        Catch ex As Exception
    '            If bolTodoBien Then
    '                sTrans.Rollback()
    '            End If
    '            MyBase.OcurrioError = True
    '            MyBase.DescripcionDelError = "Ocurrió un error al acceder a la tabla Distribuidores. " & ex.Message
    '        End Try
    '        scmdComando.Connection.Close()
    '        scmdComando.Connection.Dispose()
    '        scnConexion.Close()
    '    End If
    '    scmdComando.Dispose()
    '    scnConexion.Dispose()
    '    Return Not MyBase.OcurrioError
    'End Function

End Class
