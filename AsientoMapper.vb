Imports System.Data.Common
Imports SkY.RutinasDeBase
Imports SkY.RutinasDeBase.SQL

Public Class AsientoMapper
    Inherits cMapperGenerico

    Public Sub Cargar(ByVal xAsiento As Asiento)
        Dim sDR As DbDataReader
        sDR = Conn.ExecuteDataReader("SELECT * FROM Asientos WHERE asieje_Cod = " & SqlStr(xAsiento.Ejercicio) & " AND asi_Id = " & SqlNum(xAsiento.Id))

        Try
            If sDR.Read() Then
                xAsiento.Numero = sDR("asi_Nro")
                xAsiento.Fecha = sDR("asi_Fecha")
                xAsiento.Descripcion = sDR("asi_Desc")
            End If
            xAsiento.Cuentas = New List(Of AsientoDetalle)
            sDR.Close()
            sDR = Conn.ExecuteDataReader("SELECT * FROM AsientosDetalle WHERE adeeje_Cod = " & SqlStr(xAsiento.Ejercicio) & " AND adeasi_Id = " & SqlNum(xAsiento.Id) & " ORDER BY ade_Columna")
            While sDR.Read
                Dim sDetalle As New AsientoDetalle()
                sDetalle.Ejercicio = xAsiento.Ejercicio
                sDetalle.IdAsiento = xAsiento.Id
                sDetalle.Id = sDR("ade_Id")
                sDetalle.Cuenta = sDR("adecue_Cod")
                sDetalle.Columna = IIf(sDR("ade_Columna") = False, EnumColumna.eDebe, EnumColumna.eHaber)
                sDetalle.Importe = sDR("ade_Importe")
                xAsiento.Cuentas.Add(sDetalle)
            End While
        Catch ex As Exception
            Throw New ApplicationException("Error al Cargar el Asiento:  " & ex.Message & ".")
        Finally
            'cierro conexiones
            sDR.Close()
            sDR = Nothing

        End Try
    End Sub

    Public Sub Grabar(ByVal xAsiento As Asiento, Optional xVentas As List(Of String) = Nothing, Optional xCompras As List(Of String) = Nothing, Optional xFinanzas As List(Of String) = Nothing)
        Dim sValues As String = ""
        Dim sFields As String = ""
        Try
            With xAsiento
                sValues &= SqlStr(.Ejercicio)
                sFields &= "asieje_Cod"

                sValues &= ", " & SqlNum(.Numero)
                sFields &= ", asi_Nro"

                sValues &= ", " & SqlStr(.Descripcion)
                sFields &= ", asi_Desc"

                sValues &= ", " & SqlDate(.Fecha)
                sFields &= ", asi_Fecha"

                Conn.BeginTransaction()
                .Id = Conn.ExecuteScalar("INSERT INTO Asientos (" & sFields & ") VALUES (" & sValues & ") SELECT SCOPE_IDENTITY()")
            End With

            For Each sCuenta As AsientoDetalle In xAsiento.Cuentas
                With sCuenta
                    .Ejercicio = xAsiento.Ejercicio
                    .IdAsiento = xAsiento.Id

                    sValues = SqlStr(.Ejercicio)
                    sFields = "adeeje_Cod"

                    sValues &= ", " & SqlNum(.IdAsiento)
                    sFields &= ", adeasi_Id"

                    sValues &= ", " & SqlNum(.Id)
                    sFields &= ", ade_Id"

                    sValues &= ", " & SqlStr(.Cuenta)
                    sFields &= ", adecue_Cod"

                    sValues &= ", " & SqlBool(.Columna = EnumColumna.eHaber)
                    sFields &= ", ade_Columna"

                    sValues &= ", " & SqlNum(.Importe)
                    sFields &= ", ade_Importe"

                    Conn.ExecuteNonQuery("INSERT INTO AsientosDetalle (" & sFields & ") VALUES (" & sValues & ")")
                End With
            Next

            If xVentas IsNot Nothing AndAlso xVentas.Count > 0 Then
                For Each sVenta As String In xVentas
                    sValues = SqlStr(xAsiento.Ejercicio)
                    sFields = "raveje_Cod"

                    sValues &= ", " & SqlNum(xAsiento.Id)
                    sFields &= ", ravasi_Id"

                    sValues &= ", " & SqlStr(sVenta.Split("|")(0))
                    sFields &= ", ravemp_Cod"

                    sValues &= ", " & SqlStr(sVenta.Split("|")(1))
                    sFields &= ", ravsuc_Cod"

                    sValues &= ", " & SqlNum(sVenta.Split("|")(2))
                    sFields &= ", ravcve_Id"

                    Conn.ExecuteNonQuery("INSERT INTO RelAsientosVentas (" & sFields & ") VALUES (" & sValues & ")")
                Next

                Conn.ExecuteNonQuery("UPDATE CabVenta SET cve_PasadoCG = 'S' FROM CabVenta INNER JOIN RelAsientosVentas ON ravemp_Cod = cveemp_Cod AND ravsuc_Cod = cvesuc_Cod AND ravcve_Id = cve_Id WHERE raveje_Cod = " & SqlStr(xAsiento.Ejercicio) & " AND ravasi_Id = " & SqlNum(xAsiento.Id))

            End If

            If xCompras IsNot Nothing AndAlso xCompras.Count > 0 Then
                For Each sCompra As String In xCompras
                    sValues = SqlStr(xAsiento.Ejercicio)
                    sFields = "raceje_Cod"

                    sValues &= ", " & SqlNum(xAsiento.Id)
                    sFields &= ", racasi_Id"

                    sValues &= ", " & SqlStr(sCompra.Split("|")(0))
                    sFields &= ", racemp_Cod"

                    sValues &= ", " & SqlStr(sCompra.Split("|")(1))
                    sFields &= ", racsuc_Cod"

                    sValues &= ", " & SqlNum(sCompra.Split("|")(2))
                    sFields &= ", raccco_Id"

                    Conn.ExecuteNonQuery("INSERT INTO RelAsientosCompras (" & sFields & ") VALUES (" & sValues & ")")
                Next

                Conn.ExecuteNonQuery("UPDATE CabCompra SET cco_PasadoCG = 'S' FROM CabCompra INNER JOIN RelAsientosCompras ON racemp_Cod = ccoemp_Cod AND racsuc_Cod = ccosuc_Cod AND raccco_Id = cco_Id WHERE raceje_Cod = " & SqlStr(xAsiento.Ejercicio) & " AND racasi_Id = " & SqlNum(xAsiento.Id))

            End If

            If xFinanzas IsNot Nothing AndAlso xFinanzas.Count > 0 Then
                For Each sFinanza As String In xFinanzas
                    sValues = SqlStr(xAsiento.Ejercicio)
                    sFields = "rafeje_Cod"

                    sValues &= ", " & SqlNum(xAsiento.Id)
                    sFields &= ", rafasi_Id"

                    sValues &= ", " & SqlStr(sFinanza.Split("|")(0))
                    sFields &= ", rafemp_Cod"

                    sValues &= ", " & SqlStr(sFinanza.Split("|")(1))
                    sFields &= ", rafsuc_Cod"

                    sValues &= ", " & SqlNum(sFinanza.Split("|")(2))
                    sFields &= ", rafcmf_Id"

                    Conn.ExecuteNonQuery("INSERT INTO RelAsientosFondos (" & sFields & ") VALUES (" & sValues & ")")
                Next

                Conn.ExecuteNonQuery("UPDATE CabMovF SET cmf_PasadoCG = 'S' FROM CabMovF INNER JOIN RelAsientosFondos ON rafemp_Cod = cmfemp_Cod AND rafsuc_Cod = cmfsuc_Cod AND rafcmf_Id = cmf_Id WHERE rafeje_Cod = " & SqlStr(xAsiento.Ejercicio) & " AND rafasi_Id = " & SqlNum(xAsiento.Id))

            End If

            Conn.CommitTransaction()
        Catch ex As Exception
            Conn.RollbackTransaction()
            Throw New ApplicationException("Error al Grabar el Asiento: " & ex.Message & ".")
        End Try
    End Sub

    Public Sub Modificar(ByVal xAsiento As Asiento)
        Dim sQuery As String = ""
        Try
            With xAsiento
                sQuery &= "asi_Desc = " & SqlStr(.Descripcion)
                sQuery &= ", asi_Nro = " & SqlNum(.Numero)
                sQuery &= ", asi_Fecha = " & SqlDate(.Fecha)
                Conn.BeginTransaction()
                Conn.ExecuteNonQuery("UPDATE Asientos SET " & sQuery & " WHERE asieje_Cod = " & SqlStrNull(xAsiento.Ejercicio) & " AND asi_Id = " & SqlNum(xAsiento.Id))

                Conn.ExecuteNonQuery("DELETE FROM AsientosDetalle WHERE adeeje_Cod = " & SqlStrNull(xAsiento.Ejercicio) & " AND adeasi_Id = " & SqlNum(xAsiento.Id))

                For Each sCuenta As AsientoDetalle In xAsiento.Cuentas
                    Dim sValues As String = ""
                    Dim sFields As String = ""
                    With sCuenta
                        .Ejercicio = xAsiento.Ejercicio
                        .IdAsiento = xAsiento.Id

                        sValues &= SqlStr(.Ejercicio)
                        sFields &= "adeeje_Cod"

                        sValues &= ", " & SqlNum(.IdAsiento)
                        sFields &= ", adeasi_Id"

                        sValues &= ", " & SqlNum(.Id)
                        sFields &= ", ade_Id"

                        sValues &= ", " & SqlStr(.Cuenta)
                        sFields &= ", adecue_Cod"

                        sValues &= ", " & SqlBool(.Columna = EnumColumna.eHaber)
                        sFields &= ", ade_Columna"

                        sValues &= ", " & SqlNum(.Importe)
                        sFields &= ", ade_Importe"

                        Conn.ExecuteNonQuery("INSERT INTO AsientosDetalle (" & sFields & ") VALUES (" & sValues & ")")
                    End With
                Next

                Conn.CommitTransaction()
            End With
        Catch ex As Exception
            Conn.RollbackTransaction()
            Throw New ApplicationException("Error al Modificar el Asiento: " & ex.Message & ".")
        End Try
    End Sub

    Public Sub Eliminar(ByVal xAsiento As Asiento)
        Try
            Conn.BeginTransaction()
            Conn.ExecuteNonQuery("UPDATE CabVenta SET cve_PasadoCG = 'N' FROM CabVenta INNER JOIN RelAsientosVentas ON ravemp_Cod = cveemp_Cod AND ravsuc_Cod = cvesuc_Cod AND ravcve_Id = cve_Id WHERE raveje_Cod = " & SqlStr(xAsiento.Ejercicio) & " AND ravasi_Id = " & SqlNum(xAsiento.Id))
            Conn.ExecuteNonQuery("UPDATE CabCompra SET cco_PasadoCG = 'N' FROM CabCompra INNER JOIN RelAsientosCompras ON racemp_Cod = ccoemp_Cod AND racsuc_Cod = ccosuc_Cod AND raccco_Id = cco_Id WHERE raceje_Cod = " & SqlStr(xAsiento.Ejercicio) & " AND racasi_Id = " & SqlNum(xAsiento.Id))
            Conn.ExecuteNonQuery("UPDATE CabMovF SET cmf_PasadoCG = 'N' FROM CabMovF INNER JOIN RelAsientosFondos ON rafemp_Cod = cmfemp_Cod AND rafsuc_Cod = cmfsuc_Cod AND rafcmf_Id = cmf_Id WHERE rafeje_Cod = " & SqlStr(xAsiento.Ejercicio) & " AND rafasi_Id = " & SqlNum(xAsiento.Id))
            Conn.ExecuteNonQuery("DELETE FROM RelAsientosVentas WHERE raveje_Cod = " & SqlStrNull(xAsiento.Ejercicio) & " AND ravasi_Id = " & SqlNum(xAsiento.Id))
            Conn.ExecuteNonQuery("DELETE FROM RelAsientosCompras WHERE raceje_Cod = " & SqlStrNull(xAsiento.Ejercicio) & " AND racasi_Id = " & SqlNum(xAsiento.Id))
            Conn.ExecuteNonQuery("DELETE FROM RelAsientosFondos WHERE rafeje_Cod = " & SqlStrNull(xAsiento.Ejercicio) & " AND rafasi_Id = " & SqlNum(xAsiento.Id))
            Conn.ExecuteNonQuery("DELETE FROM AsientosDetalle WHERE adeeje_Cod = " & SqlStrNull(xAsiento.Ejercicio) & " AND adeasi_Id = " & SqlNum(xAsiento.Id))
            Conn.ExecuteNonQuery("DELETE FROM Asientos WHERE asieje_Cod = " & SqlStrNull(xAsiento.Ejercicio) & " AND asi_Id = " & SqlNum(xAsiento.Id))
            Conn.CommitTransaction()
        Catch ex As Exception
            Conn.RollbackTransaction()
            Throw New ApplicationException("Error al Eliminar el Asiento: " & ex.Message & ".")
        End Try
    End Sub

    Public Function Grilla(xEjercicio As String) As DataSet
        Return Conn.ExecuteDataSet("SELECT asieje_Cod, asi_Id, asi_Fecha AS [Fecha], asi_Nro AS [Número], asi_Desc AS [Descripción], SUM(ade_Importe) AS Importe FROM Asientos INNER JOIN AsientosDetalle ON asieje_Cod = adeeje_Cod AND asi_Id = adeasi_Id WHERE asieje_Cod = " & SqlStr(xEjercicio) & " AND ade_Columna = 0 GROUP BY asieje_Cod, asi_Id, asi_Fecha, asi_Nro, asi_Desc ORDER BY asi_Fecha DESC, asi_Nro DESC")
    End Function

    Public Function GrillaDetalleDummy() As DataSet
        Return Conn.ExecuteDataSet("SELECT adeeje_Cod, adeasi_Id, ade_Id, adecue_Cod AS [Código], cue_Desc AS [Descripción], CASE WHEN ade_Columna = 0 THEN ade_Importe ELSE 0 END AS [Debe], CASE WHEN ade_Columna = 1 THEN ade_Importe ELSE 0 END AS [Haber] FROM AsientosDetalle INNER JOIN Cuentas ON adeeje_Cod = cueeje_Cod AND adecue_Cod = cue_Cod WHERE 1=0")
    End Function

    Public Function Mayor(xEjercicio As String, ByVal xCuenta As String, xFechaDesde As DateTime, xFechaHasta As DateTime) As List(Of Mayor)
        Dim sDR As DbDataReader
        sDR = Conn.ExecuteDataReader("SELECT * FROM AsientosDetalle INNER JOIN Asientos ON adeasi_Id = asi_Id AND adeeje_Cod = asieje_Cod INNER JOIN Cuentas ON adecue_Cod = cue_Cod AND adeeje_Cod = cueeje_Cod WHERE adeeje_Cod = " & SqlStr(xEjercicio) & IIf(Not String.IsNullOrEmpty(xCuenta), " AND adecue_Cod = " & SqlStr(xCuenta), "") & " AND asi_Fecha >= " & SqlDate(xFechaDesde) & " AND asi_Fecha <= " & SqlDate(xFechaHasta) & " ORDER BY adecue_Cod, asi_Fecha, asi_Id")

        Try
            Dim sLista As New List(Of Mayor)
            While sDR.Read()
                Dim sDetalle As New Mayor()
                sDetalle.Ejercicio = xEjercicio
                sDetalle.IdAsiento = sDR("adeasi_Id")
                sDetalle.Id = sDR("ade_Id")
                sDetalle.Fecha = sDR("asi_Fecha")
                sDetalle.Numero = sDR("asi_Nro")
                sDetalle.Descripcion = sDR("asi_Desc")
                sDetalle.Cuenta = sDR("adecue_Cod")
                sDetalle.DescripcionCuenta = sDR("cue_Desc")
                sDetalle.Columna = IIf(sDR("ade_Columna") = False, EnumColumna.eDebe, EnumColumna.eHaber)
                sDetalle.Importe = sDR("ade_Importe")
                sLista.Add(sDetalle)
            End While
            Return sLista
        Catch ex As Exception
            Throw New ApplicationException("Error al Cargar el Mayor: " & ex.Message & ".")
        Finally
            'cierro conexiones
            sDR.Close()
            sDR = Nothing

        End Try
    End Function

    Public Function SumasYSaldos(xEjercicio As String, xFechaDesde As DateTime, xFechaHasta As DateTime) As List(Of SumasYSaldos)
        Dim sDR As DbDataReader
        sDR = Conn.ExecuteDataReader("SELECT cue_Cod, cue_Desc, SUM(CASE WHEN asi_Fecha < " & SqlDate(xFechaDesde) & " THEN CASE WHEN ade_Columna = 0 THEN ade_Importe ELSE ade_Importe * -1 END ELSE 0 END) AS SaldoAnterior, SUM(CASE WHEN asi_Fecha >= " & SqlDate(xFechaDesde) & " THEN CASE WHEN ade_Columna = 0 THEN ade_Importe ELSE 0 END ELSE 0 END) AS Debe, SUM(CASE WHEN asi_Fecha >= " & SqlDate(xFechaDesde) & " THEN CASE WHEN ade_Columna = 0 THEN 0 ELSE ade_Importe * -1 END ELSE 0 END) AS Haber, SUM(CASE WHEN ade_Columna = 0 THEN ade_Importe ELSE ade_Importe * -1 END) AS Saldo FROM AsientosDetalle INNER JOIN Asientos ON adeasi_Id = asi_Id AND adeeje_Cod = asieje_Cod INNER JOIN Cuentas ON cueeje_Cod = adeeje_Cod AND cue_Cod = adecue_Cod WHERE adeeje_Cod = " & SqlStr(xEjercicio) & " AND asi_Fecha <= " & SqlDate(xFechaHasta) & " GROUP BY cue_Cod, cue_Desc ORDER BY cue_Cod")

        Try
            Dim sLista As New List(Of SumasYSaldos)
            While sDR.Read()
                Dim sDetalle As New SumasYSaldos()
                sDetalle.Capitulo = Cuenta.Capitulo(sDR("cue_Cod"))
                sDetalle.Codigo = sDR("cue_Cod")
                sDetalle.Descripcion = sDR("cue_Desc")
                sDetalle.SaldoAnterior = sDR("SaldoAnterior")
                sDetalle.Debe = sDR("Debe")
                sDetalle.Haber = sDR("Haber")
                sDetalle.Saldo = sDR("Saldo")
                sLista.Add(sDetalle)
            End While
            Return sLista
        Catch ex As Exception
            Throw New ApplicationException("Error al Cargar el Sumas y Saldos: " & ex.Message & ".")
        Finally
            'cierro conexiones
            sDR.Close()
            sDR = Nothing

        End Try
    End Function

    Public Function SaldoCuenta(xEjercicio As String, ByVal xCuenta As String, xFechaHasta As DateTime) As Double
        Try
            Return Conn.ExecuteScalar("SELECT SUM(CASE WHEN ade_Columna = 0 THEN ade_Importe ELSE ade_Importe * -1 END) AS Saldo FROM AsientosDetalle INNER JOIN Asientos ON adeasi_Id = asi_Id AND adeeje_Cod = asieje_Cod WHERE adecue_Cod = " & SqlStr(xCuenta) & " AND adeeje_Cod = " & SqlStr(xEjercicio) & " AND asi_Fecha <= " & SqlDate(xFechaHasta))
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Function ProximoNumero(xEjercicio As String) As Integer
        Try
            Dim sNro As Integer = Conn.ExecuteScalar("SELECT MAX(asi_Nro) FROM Asientos WHERE asieje_Cod = " & SqlStr(xEjercicio))
            Return sNro + 1
        Catch ex As Exception
            Return 1
        End Try
    End Function

    Public Sub Renumerar(xEjercicio As String)
        Dim sDR As DbDataReader
        sDR = Conn.ExecuteDataReader("SELECT * FROM Asientos WHERE asieje_Cod = " & SqlStr(xEjercicio) & " ORDER BY asi_Fecha, asi_Id")

        Try
            Dim sLista As New Dictionary(Of Integer, Integer)
            Dim sNro As Integer = 1
            While sDR.Read()
                sLista.Add(sDR("asi_Id"), sNro)
                sNro += 1
            End While
            sDR.Close()
            For Each sAsiento As KeyValuePair(Of Integer, Integer) In sLista
                Conn.ExecuteNonQuery("UPDATE Asientos SET asi_Nro = " & sAsiento.Value & "WHERE asieje_Cod = " & SqlStr(xEjercicio) & " AND asi_Id = " & SqlNum(sAsiento.Key))
            Next
        Catch ex As Exception
            Throw New ApplicationException("Error al Renumerar Asientos: " & ex.Message & ".")
        Finally
            'cierro conexiones
            If sDR IsNot Nothing AndAlso Not sDR.IsClosed Then sDR.Close()
            sDR = Nothing

        End Try
    End Sub

    Public Function GenerarAsientosVentas(xEjercicio As String, xFechaInicio As DateTime, xFechaFin As DateTime) As List(Of String)
        Dim sDR As DbDataReader = Nothing
        Dim sErrores As New List(Of String)
        Try
            Dim sCuentas As New SortedDictionary(Of String, AsientoDetalle)
            Dim sAsiento As New Asiento(StrConn)
            sAsiento.Ejercicio = xEjercicio
            sAsiento.Fecha = xFechaFin
            sAsiento.Descripcion = "Asiento Resumen de Ventas " & xFechaInicio.ToString("dd/MM/yyyy") & " - " & xFechaFin.ToString("dd/MM/yyyy")
            Dim sQuery As String = ""

            'Deudores por Ventas
            sQuery = "SELECT ROUND(ISNULL(SUM(cve_ImpMonLoc),0),2) AS [Deudores] "
            sQuery &= "FROM CabVenta "
            sQuery &= "INNER JOIN TipComp ON cvetco_Cod = tco_Cod AND cvecir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN CondVta ON cvecvt_Cod = cvt_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'V' AND cve_PasadoCG = 'N' AND cve_FContab >= " & SqlDate(xFechaInicio) & " AND cve_FContab <= " & SqlDate(xFechaFin) & " AND cve_Anulado = 0 AND cvt_TipCond = 2 "
            sDR = Conn.ExecuteDataReader(sQuery)
            If sDR.Read AndAlso sDR("Deudores") Then
                Dim sCuenta = ParametrosCG.CuentaDeudoresPorVenta(StrConn, xEjercicio)
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Deudores")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para Deudores por Venta.")
                End If
            End If
            sDR.Close()

            'Ventas en Efectivo - Caja
            sQuery = "SELECT mfocaj_Cod, caj_Desc, cajccb_Cod, SUM(mfo_ImpMonLoc) AS Caja "
            sQuery &= "FROM CabVenta "
            sQuery &= "INNER JOIN TipComp ON cvetco_Cod = tco_Cod AND cvecir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN CondVta ON cvecvt_Cod = cvt_Cod "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = cveemp_Cod AND mfosuc_Cod = cvesuc_Cod AND mfocmf_Id = cvecmf_ID "
            sQuery &= "INNER JOIN Cajas ON mfocaj_Cod = caj_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'V' AND cve_PasadoCG = 'N' AND cve_FContab >= " & SqlDate(xFechaInicio) & " AND cve_FContab <= " & SqlDate(xFechaFin) & " AND cve_Anulado = 0 AND cvt_TipCond = 1 "
            sQuery &= "AND mfo_CodConcepto = 1 GROUP BY mfocaj_Cod, cajccb_Cod, caj_Desc"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("cajccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Caja")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para la Caja: " & sDR("caj_Desc") & ".")
                End If
            End While
            sDR.Close()

            'Ventas en Efectivo - Cuenta Bancaria
            sQuery = "SELECT mfobco_Cod, mfobco_Suc, mfoctb_Cod, ctbccb_Cod, ctb_Desc, SUM(mfo_ImpMonLoc) AS Banco "
            sQuery &= "FROM CabVenta "
            sQuery &= "INNER JOIN TipComp ON cvetco_Cod = tco_Cod AND cvecir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN CondVta ON cvecvt_Cod = cvt_Cod "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = cveemp_Cod AND mfosuc_Cod = cvesuc_Cod AND mfocmf_Id = cvecmf_ID "
            sQuery &= "INNER JOIN CtaBan ON mfobco_Cod = ctbbco_Cod AND mfobco_Suc = ctbbco_Suc AND mfoctb_Cod = ctb_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'V' AND cve_PasadoCG = 'N' AND cve_FContab >= " & SqlDate(xFechaInicio) & " AND cve_FContab <= " & SqlDate(xFechaFin) & " AND cve_Anulado = 0 AND cvt_TipCond = 1 "
            sQuery &= "AND mfo_CodConcepto = 4 GROUP BY mfobco_Cod, mfobco_Suc, mfoctb_Cod, ctbccb_Cod, ctb_Desc"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("ctbccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Banco")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para la Cuenta Bancaria: " & sDR("ctb_Desc") & ".")
                End If
            End While
            sDR.Close()

            'Ventas en Efectivo - Cheques de Terceros
            sQuery = "SELECT tch_Cod, tchccb_Cod, tch_Desc, SUM(ch3_Importe) AS [Cheques] "
            sQuery &= "FROM CabVenta "
            sQuery &= "INNER JOIN TipComp ON cvetco_Cod = tco_Cod AND cvecir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN CondVta ON cvecvt_Cod = cvt_Cod "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = cveemp_Cod AND mfosuc_Cod = cvesuc_Cod AND mfocmf_Id = cvecmf_ID "
            sQuery &= "INNER JOIN RelaChq3 ON mfoemp_Cod = rchemp_CodMF AND mfosuc_Cod = rchsuc_CodMF AND mfocmf_ID = rchcmf_ID "
            sQuery &= "INNER JOIN Cheques3 On ch3emp_Cod = rchemp_Cod And ch3suc_Cod = rchsuc_Cod And ch3_Id = rchch3_Id "
            sQuery &= "INNER JOIN TipCheq ON tch_Cod = ch3tch_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'V' AND cve_PasadoCG = 'N' AND cve_FContab >= " & SqlDate(xFechaInicio) & " AND cve_FContab <= " & SqlDate(xFechaFin) & " AND cve_Anulado = 0 AND cvt_TipCond = 1 "
            sQuery &= "AND mfo_CodConcepto = 2 GROUP BY tch_Cod, tchccb_Cod, tch_Desc"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("tchccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Cheques")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para el Tipo de Cheque: " & sDR("tch_Desc") & ".")
                End If
            End While
            sDR.Close()

            'Neto / IVA / Impuestos Internos
            sQuery = "SELECT "
            sQuery &= "ROUND(ISNULL(SUM(Case When div_TipoTot != '4' THEN div_ImpNeto ELSE 0 END),0)*-1,2) AS Neto, "
            sQuery &= "ROUND(ISNULL(SUM(Case When div_TipoTot != '4' THEN div_Imp1 + div_Imp2 ELSE 0 END),0)*-1,2) AS IVA, "
            sQuery &= "ROUND(ISNULL(SUM(Case When div_TipoTot = '4' THEN div_ImpNeto + div_Imp1 + div_Imp2 ELSE 0 END),0)*-1,2) AS ImpInt "
            sQuery &= "FROM CabVenta "
            sQuery &= "INNER JOIN DetIvaVta On cveemp_Cod = divemp_Cod And cvesuc_Cod = divsuc_Cod And cve_ID = divcve_ID "
            sQuery &= "INNER JOIN TipComp On cvetco_Cod = tco_Cod And cvecir_Cod = tcocir_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'V' AND cve_PasadoCG = 'N' AND cve_FContab >= " & SqlDate(xFechaInicio) & " AND cve_FContab <= " & SqlDate(xFechaFin) & " AND cve_Anulado = 0 AND div_TipoTot <> '6'"
            sDR = Conn.ExecuteDataReader(sQuery)
            If sDR.Read Then
                If sDR("Neto") <> 0 Then
                    Dim sCuenta As String = ParametrosCG.CuentaVentas(StrConn, xEjercicio)
                    If Not String.IsNullOrEmpty(sCuenta) Then
                        Dim sAsientoDetalle As New AsientoDetalle()
                        sAsientoDetalle.Cuenta = sCuenta
                        sAsientoDetalle.Importe = sDR("Neto")
                        If sCuentas.ContainsKey(sCuenta) Then
                            sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                        Else
                            sCuentas.Add(sCuenta, sAsientoDetalle)
                        End If
                    Else
                        sErrores.Add("No se encuentra configurada la Cuenta Contable para Ventas.")
                    End If
                End If

                If sDR("IVA") <> 0 Then
                    Dim sCuenta As String = ParametrosCG.CuentaIVADebitoFiscal(StrConn, xEjercicio)
                    If Not String.IsNullOrEmpty(sCuenta) Then
                        Dim sAsientoDetalle As New AsientoDetalle()
                        sAsientoDetalle.Cuenta = sCuenta
                        sAsientoDetalle.Importe = sDR("IVA")
                        If sCuentas.ContainsKey(sCuenta) Then
                            sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                        Else
                            sCuentas.Add(sCuenta, sAsientoDetalle)
                        End If
                    Else
                        sErrores.Add("No se encuentra configurada la Cuenta Contable para IVA Débito Fiscal.")
                    End If
                End If

                If sDR("ImpInt") <> 0 Then
                    Dim sCuenta As String = ParametrosCG.CuentaImpuestosInternos(StrConn, xEjercicio)
                    If Not String.IsNullOrEmpty(sCuenta) Then
                        Dim sAsientoDetalle As New AsientoDetalle()
                        sAsientoDetalle.Cuenta = sCuenta
                        sAsientoDetalle.Importe = sDR("ImpInt")
                        If sCuentas.ContainsKey(sCuenta) Then
                            sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                        Else
                            sCuentas.Add(sCuenta, sAsientoDetalle)
                        End If
                    Else
                        sErrores.Add("No se encuentra configurada la Cuenta Contable para Impuestos Internos.")
                    End If
                End If
            End If
            sDR.Close()

            'Retenciones
            sQuery = "SELECT divres_Cod, divres_Art, resccb_CodVta, res_Desc, ROUND(ISNULL(SUM(div_ImpNeto + div_Imp1 + div_Imp2),0)*-1,2) AS [Retenciones] "
            sQuery &= "FROM CabVenta "
            sQuery &= "INNER JOIN DetIvaVta ON cveemp_Cod = divemp_Cod And cvesuc_Cod = divsuc_Cod And cve_ID = divcve_ID "
            sQuery &= "INNER JOIN TipComp ON cvetco_Cod = tco_Cod And cvecir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN RegEsp ON divres_Cod = res_Cod AND divres_Art = res_Art "
            sQuery &= "WHERE tcocir_CodCG = 'V' AND cve_PasadoCG = 'N' AND cve_FContab >= " & SqlDate(xFechaInicio) & " AND cve_FContab <= " & SqlDate(xFechaFin) & " AND cve_Anulado = 0 AND div_TipoTot = '6' "
            sQuery &= "GROUP BY divres_Cod, divres_Art, resccb_CodVta, res_Desc"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                If sDR("Retenciones") <> 0 Then
                    Dim sCuenta As String = ReplaceDBNull(sDR("resccb_CodVta"), "")
                    If Not String.IsNullOrEmpty(sCuenta) Then
                        Dim sAsientoDetalle As New AsientoDetalle()
                        sAsientoDetalle.Cuenta = sCuenta
                        sAsientoDetalle.Importe = sDR("Retenciones")
                        If sCuentas.ContainsKey(sCuenta) Then
                            sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                        Else
                            sCuentas.Add(sCuenta, sAsientoDetalle)
                        End If
                    Else
                        sErrores.Add("No se encuentra configurada la Cuenta Contable para el Régimen Especial: " & sDR("res_Desc") & ".")
                    End If
                End If
            End While
            sDR.Close()

            Dim sVentas As New List(Of String)
            sQuery = "SELECT cveemp_Cod, cvesuc_Cod, cve_Id FROM CabVenta "
            sQuery &= "INNER JOIN TipComp ON cvetco_Cod = tco_Cod AND cvecir_Cod = tcocir_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'V' AND cve_PasadoCG = 'N' AND cve_FContab >= " & SqlDate(xFechaInicio) & " AND cve_FContab <= " & SqlDate(xFechaFin) & " AND cve_Anulado = 0"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                sVentas.Add(sDR("cveemp_Cod") & "|" & sDR("cvesuc_Cod") & "|" & sDR("cve_Id"))
            End While
            sDR.Close()

            If sCuentas.Count <> 0 Then
                Try
                    For Each sCuenta As AsientoDetalle In sCuentas.Values
                        sCuenta.Columna = IIf(sCuenta.Importe > 0, EnumColumna.eDebe, EnumColumna.eHaber)
                        sCuenta.Importe = Math.Abs(sCuenta.Importe)
                        sAsiento.AgregarCuenta(sCuenta)
                    Next
                    sAsiento.Numero = Asiento.ProximoNumero(StrConn, xEjercicio)
                    sAsiento.Grabar(sVentas)
                Catch ex As Exception
                    sErrores.Add(ex.Message)
                End Try
            Else
                sErrores.Add("No hay movimientos de Ventas para el período seleccionado.")
            End If

            Return sErrores

        Catch ex As Exception
            Throw New ApplicationException("Error al Generar los Asientos de Ventas: " & ex.Message & ".")
        Finally
            If sDR IsNot Nothing AndAlso Not sDR.IsClosed Then sDR.Close()
            sDR = Nothing
        End Try
    End Function

    Public Function GenerarAsientosCobros(xEjercicio As String, xFechaInicio As DateTime, xFechaFin As DateTime) As List(Of String)
        Dim sDR As DbDataReader = Nothing
        Dim sErrores As New List(Of String)
        Try
            Dim sCuentas As New SortedDictionary(Of String, AsientoDetalle)
            Dim sAsiento As New Asiento(StrConn)
            sAsiento.Ejercicio = xEjercicio
            sAsiento.Fecha = xFechaFin
            sAsiento.Descripcion = "Asiento Resumen de Cobros " & xFechaInicio.ToString("dd/MM/yyyy") & " - " & xFechaFin.ToString("dd/MM/yyyy")
            Dim sQuery As String = ""


            'Cobros - Retenciones
            sQuery = "SELECT divres_Cod, divres_Art, resccb_CodVta, res_Desc, ROUND(ISNULL(SUM(div_ImpNeto + div_Imp1 + div_Imp2),0)*-1,2) AS [Retenciones] "
            sQuery &= "FROM CabVenta "
            sQuery &= "INNER JOIN TipComp ON cvetco_Cod = tco_Cod AND cvecir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN CondVta On cvecvt_Cod = cvt_Cod "
            sQuery &= "INNER JOIN DetIvaVta ON cveemp_Cod = divemp_Cod AND cvesuc_Cod = divsuc_Cod AND cve_ID = divcve_ID "
            sQuery &= "INNER JOIN RegEsp ON divres_Cod = res_Cod AND divres_Art = res_Art "
            sQuery &= "WHERE tcocir_CodCG = 'B' AND cve_PasadoCG = 'N' AND cve_FContab >= " & SqlDate(xFechaInicio) & " AND cve_FContab <= " & SqlDate(xFechaFin) & " AND div_TipoTot = '6' "
            sQuery &= "GROUP BY divres_Cod, divres_Art, resccb_CodVta, res_Desc"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                If sDR("Retenciones") <> 0 Then
                    Dim sCuenta As String = ReplaceDBNull(sDR("resccb_CodVta"), "")
                    If Not String.IsNullOrEmpty(sCuenta) Then
                        Dim sAsientoDetalle As New AsientoDetalle()
                        sAsientoDetalle.Cuenta = sCuenta
                        sAsientoDetalle.Importe = sDR("Retenciones")
                        If sCuentas.ContainsKey(sCuenta) Then
                            sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                        Else
                            sCuentas.Add(sCuenta, sAsientoDetalle)
                        End If
                    Else
                        sErrores.Add("No se encuentra configurada la Cuenta Contable para el Régimen Especial: " & sDR("res_Desc") & ".")
                    End If
                End If
            End While
            sDR.Close()

            'Cobros - Caja
            sQuery = "SELECT mfocaj_Cod, caj_Desc, cajccb_Cod, ROUND(SUM(mfo_ImpMonLoc),2) AS Caja "
            sQuery &= "FROM CabVenta "
            sQuery &= "INNER JOIN TipComp ON cvetco_Cod = tco_Cod AND cvecir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN CondVta ON cvecvt_Cod = cvt_Cod "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = cveemp_Cod AND mfosuc_Cod = cvesuc_Cod AND mfocmf_Id = cvecmf_ID "
            sQuery &= "INNER JOIN Cajas ON mfocaj_Cod = caj_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'B' AND cve_PasadoCG = 'N' AND cve_FContab >= " & SqlDate(xFechaInicio) & " AND cve_FContab <= " & SqlDate(xFechaFin) & " "
            sQuery &= "AND mfo_CodConcepto = 1 GROUP BY mfocaj_Cod, cajccb_Cod, caj_Desc"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("cajccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Caja")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para la Caja: " & sDR("caj_Desc") & ".")
                End If
            End While
            sDR.Close()

            'Cobros - Cuenta Bancaria
            sQuery = "SELECT mfobco_Cod, mfobco_Suc, mfoctb_Cod, ctbccb_Cod, ctb_Desc, ROUND(SUM(mfo_ImpMonLoc),2) AS Banco "
            sQuery &= "FROM CabVenta "
            sQuery &= "INNER JOIN TipComp ON cvetco_Cod = tco_Cod AND cvecir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN CondVta ON cvecvt_Cod = cvt_Cod "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = cveemp_Cod AND mfosuc_Cod = cvesuc_Cod AND mfocmf_Id = cvecmf_ID "
            sQuery &= "INNER JOIN CtaBan ON mfobco_Cod = ctbbco_Cod AND mfobco_Suc = ctbbco_Suc AND mfoctb_Cod = ctb_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'B' AND cve_PasadoCG = 'N' AND cve_FContab >= " & SqlDate(xFechaInicio) & " AND cve_FContab <= " & SqlDate(xFechaFin) & " "
            sQuery &= "AND mfo_CodConcepto = 4 GROUP BY mfobco_Cod, mfobco_Suc, mfoctb_Cod, ctbccb_Cod, ctb_Desc"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("ctbccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Banco")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para la Cuenta Bancaria: " & sDR("ctb_Desc") & ".")
                End If
            End While
            sDR.Close()

            'Cobros - Cheques de Terceros
            sQuery = "SELECT tch_Cod, tchccb_Cod, tch_Desc, ROUND(SUM(ch3_Importe),2) AS [Cheques] "
            sQuery &= "FROM CabVenta "
            sQuery &= "INNER JOIN TipComp ON cvetco_Cod = tco_Cod AND cvecir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN CondVta ON cvecvt_Cod = cvt_Cod "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = cveemp_Cod AND mfosuc_Cod = cvesuc_Cod AND mfocmf_Id = cvecmf_ID "
            sQuery &= "INNER JOIN RelaChq3 ON mfoemp_Cod = rchemp_CodMF AND mfosuc_Cod = rchsuc_CodMF AND mfocmf_ID = rchcmf_ID "
            sQuery &= "INNER JOIN Cheques3 On ch3emp_Cod = rchemp_Cod And ch3suc_Cod = rchsuc_Cod And ch3_Id = rchch3_Id "
            sQuery &= "INNER JOIN TipCheq ON tch_Cod = ch3tch_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'B' AND cve_PasadoCG = 'N' AND cve_FContab >= " & SqlDate(xFechaInicio) & " AND cve_FContab <= " & SqlDate(xFechaFin) & " "
            sQuery &= "AND mfo_CodConcepto = 2 GROUP BY tch_Cod, tchccb_Cod, tch_Desc"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("tchccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Cheques")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para el Tipo de Cheque: " & sDR("tch_Desc") & ".")
                End If
            End While
            sDR.Close()

            'Cobros - Deudores por Ventas
            sQuery = "SELECT ROUND(ISNULL(SUM(cve_ImpMonLoc),0),2) AS [Deudores] "
            sQuery &= "FROM CabVenta "
            sQuery &= "INNER JOIN TipComp ON cvetco_Cod = tco_Cod AND cvecir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN CondVta On cvecvt_Cod = cvt_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'B' AND cve_PasadoCG = 'N' AND cve_FContab >= " & SqlDate(xFechaInicio) & " AND cve_FContab <= " & SqlDate(xFechaFin)
            sDR = Conn.ExecuteDataReader(sQuery)
            If sDR.Read AndAlso sDR("Deudores") Then
                Dim sCuenta = ParametrosCG.CuentaDeudoresPorVenta(StrConn, xEjercicio)
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Deudores")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para Deudores por Venta.")
                End If
            End If
            sDR.Close()

            Dim sVentas As New List(Of String)
            sQuery = "SELECT cveemp_Cod, cvesuc_Cod, cve_Id FROM CabVenta "
            sQuery &= "INNER JOIN TipComp ON cvetco_Cod = tco_Cod AND cvecir_Cod = tcocir_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'B' AND cve_PasadoCG = 'N' AND cve_FContab >= " & SqlDate(xFechaInicio) & " AND cve_FContab <= " & SqlDate(xFechaFin) & " AND cve_Anulado = 0"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                sVentas.Add(sDR("cveemp_Cod") & "|" & sDR("cvesuc_Cod") & "|" & sDR("cve_Id"))
            End While
            sDR.Close()

            If sCuentas.Count <> 0 Then
                Try
                    For Each sCuenta As AsientoDetalle In sCuentas.Values
                        sCuenta.Columna = IIf(sCuenta.Importe > 0, EnumColumna.eDebe, EnumColumna.eHaber)
                        sCuenta.Importe = Math.Abs(sCuenta.Importe)
                        sAsiento.AgregarCuenta(sCuenta)
                    Next
                    sAsiento.Numero = Asiento.ProximoNumero(StrConn, xEjercicio)
                    sAsiento.Grabar(sVentas)
                Catch ex As Exception
                    sErrores.Add(ex.Message)
                End Try
            Else
                sErrores.Add("No hay movimientos de Cobros para el período seleccionado.")
            End If

            Return sErrores

        Catch ex As Exception
            Throw New ApplicationException("Error al Generar los Asientos de Cobros: " & ex.Message & ".")
        Finally
            If sDR IsNot Nothing AndAlso Not sDR.IsClosed Then sDR.Close()
            sDR = Nothing
        End Try
    End Function

    Public Function GenerarAsientosCompras(xEjercicio As String, xFechaInicio As DateTime, xFechaFin As DateTime) As List(Of String)
        Dim sDR As DbDataReader = Nothing
        Dim sErrores As New List(Of String)
        Try
            Dim sCuentas As New SortedDictionary(Of String, AsientoDetalle)
            Dim sAsiento As New Asiento(StrConn)
            sAsiento.Ejercicio = xEjercicio
            sAsiento.Fecha = xFechaFin
            sAsiento.Descripcion = "Asiento Resumen de Compras " & xFechaInicio.ToString("dd/MM/yyyy") & " - " & xFechaFin.ToString("dd/MM/yyyy")
            Dim sQuery As String = ""

            'Proveedores
            sQuery = "SELECT protpr_Cod, tprccb_Cod, tpr_Desc, ROUND(ISNULL(SUM(cco_ImpMonLoc),0),2) AS [Proveedores] "
            sQuery &= "FROM CabCompra "
            sQuery &= "INNER JOIN TipComp ON ccotco_Cod = tco_Cod AND ccocir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN Proveedores On ccopro_Cod = pro_Cod "
            sQuery &= "INNER JOIN CondPago ON ccocpg_Cod = cpg_Cod "
            sQuery &= "INNER JOIN TipProveedor On tpr_Cod = protpr_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'C' AND cco_PasadoCG = 'N' AND cco_FContab >= " & SqlDate(xFechaInicio) & " AND cco_FContab <= " & SqlDate(xFechaFin) & " AND cpg_TipCond = 2 "
            sQuery &= "GROUP BY protpr_Cod, tprccb_Cod, tpr_Desc "
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta = ReplaceDBNull(sDR("tprccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Proveedores")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para el Tipo de Proveedor: " & sDR("tpr_Desc") & ".")
                End If
            End While
            sDR.Close()

            'Compras en Efectivo - Caja
            sQuery = "SELECT mfocaj_Cod, cajccb_Cod, caj_Desc, ROUND(ISNULL(SUM(mfo_ImpMonLoc),0),2) AS Caja "
            sQuery &= "FROM CabCompra "
            sQuery &= "INNER JOIN TipComp ON ccotco_Cod = tco_Cod AND ccocir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN CondPago On ccocpg_Cod = cpg_Cod "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = ccoemp_Cod AND mfosuc_Cod = ccosuc_Cod AND mfocmf_Id = ccocmf_ID "
            sQuery &= "INNER JOIN Cajas On mfocaj_Cod = caj_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'C' AND cco_PasadoCG = 'N' AND cco_FContab >= " & SqlDate(xFechaInicio) & " AND cco_FContab <= " & SqlDate(xFechaFin) & " AND cpg_TipCond = 1 "
            sQuery &= " And mfo_CodConcepto = 1 GROUP BY mfocaj_Cod, cajccb_Cod, caj_Desc"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("cajccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Caja")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para la Caja: " & sDR("caj_Desc") & ".")
                End If
            End While
            sDR.Close()

            'Compras en Efectivo - Cuenta Bancaria
            sQuery = "SELECT mfobco_Cod, mfobco_Suc, mfoctb_Cod, ctbccb_Cod, ctb_Desc, ROUND(ISNULL(SUM(mfo_ImpMonLoc),0),2) AS Banco "
            sQuery &= "FROM CabCompra "
            sQuery &= "INNER JOIN TipComp ON ccotco_Cod = tco_Cod AND ccocir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN CondPago On ccocpg_Cod = cpg_Cod "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = ccoemp_Cod AND mfosuc_Cod = ccosuc_Cod AND mfocmf_Id = ccocmf_ID "
            sQuery &= "INNER JOIN CtaBan On mfobco_Cod = ctbbco_Cod And mfobco_Suc = ctbbco_Suc And mfoctb_Cod = ctb_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'C' AND cco_PasadoCG = 'N' AND cco_FContab >= " & SqlDate(xFechaInicio) & " AND cco_FContab <= " & SqlDate(xFechaFin) & " AND cpg_TipCond = 1 "
            sQuery &= " And mfo_CodConcepto = 4 GROUP BY mfobco_Cod, mfobco_Suc, mfoctb_Cod, ctbccb_Cod, ctb_Desc"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("ctbccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Banco")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para la Cuenta Bancaria: " & sDR("ctb_Desc") & ".")
                End If
            End While
            sDR.Close()

            'Compras en Efectivo - Cheques de Terceros
            sQuery = "SELECT tchccb_Cod, tch_Desc, ROUND(ISNULL(SUM(ch3_Importe*-1),0),2) AS [Cheque] "
            sQuery &= "FROM CabCompra "
            sQuery &= "INNER JOIN TipComp ON ccotco_Cod = tco_Cod AND ccocir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN CondPago On ccocpg_Cod = cpg_Cod "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = ccoemp_Cod AND mfosuc_Cod = ccosuc_Cod AND mfocmf_Id = ccocmf_ID "
            sQuery &= "INNER JOIN RelaChq3 ON mfoemp_Cod = rchemp_CodMF AND mfosuc_Cod = rchsuc_CodMF AND mfocmf_ID = rchcmf_ID "
            sQuery &= "INNER JOIN Cheques3 On ch3emp_Cod = rchemp_Cod And ch3suc_Cod = rchsuc_Cod And ch3_Id = rchch3_Id "
            sQuery &= "INNER JOIN TipCheq ON tch_Cod = ch3tch_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'C' AND cco_PasadoCG = 'N' AND cco_FContab >= " & SqlDate(xFechaInicio) & " AND cco_FContab <= " & SqlDate(xFechaFin) & " AND cpg_TipCond = 1 "
            sQuery &= "AND mfo_CodConcepto = 2 GROUP BY tchccb_Cod, tch_Desc"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("tchccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Cheque")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para el Tipo de Cheque: " & sDR("tch_Desc") & ".")
                End If
            End While
            sDR.Close()

            'Neto / IVA / Impuestos Internos
            sQuery = "SELECT proccb_Cod, pro_RazSoc, "
            sQuery &= "ROUND(ISNULL(SUM(CASE When dic_TipoTot != '4' THEN dic_ImpNeto ELSE 0 END),0)*-1,2) AS Neto, "
            sQuery &= "ROUND(ISNULL(SUM(CASE WHEN dic_TipoTot != '4' THEN dic_Imp1 + dic_Imp2 ELSE 0 END),0)*-1,2) AS IVA, "
            sQuery &= "ROUND(ISNULL(SUM(CASE When dic_TipoTot = '4' THEN dic_ImpNeto + dic_Imp1 + dic_Imp2 ELSE 0 END),0)*-1,2) AS ImpInt "
            sQuery &= "FROM CabCompra "
            sQuery &= "INNER JOIN DetIvaCpr ON ccoemp_Cod = dicemp_Cod And ccosuc_Cod = dicsuc_Cod And cco_ID = diccco_ID  "
            sQuery &= "INNER JOIN TipComp ON ccotco_Cod = tco_Cod AND ccocir_Cod = tcocir_Cod  "
            sQuery &= "INNER JOIN Proveedores ON ccopro_Cod = pro_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'C' AND cco_PasadoCG = 'N' AND cco_FContab >= " & SqlDate(xFechaInicio) & " AND cco_FContab <= " & SqlDate(xFechaFin) & " AND dic_TipoTot != '6' "
            sQuery &= "GROUP BY proccb_Cod, pro_RazSoc"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                If sDR("Neto") <> 0 OrElse sDR("ImpInt") <> 0 Then
                    Dim sCuenta As String = ReplaceDBNull(sDR("proccb_Cod"), "")
                    If Not String.IsNullOrEmpty(sCuenta) Then
                        Dim sAsientoDetalle As New AsientoDetalle()
                        sAsientoDetalle.Cuenta = sCuenta
                        sAsientoDetalle.Importe = sDR("Neto") + sDR("ImpInt")
                        If sCuentas.ContainsKey(sCuenta) Then
                            sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                        Else
                            sCuentas.Add(sCuenta, sAsientoDetalle)
                        End If
                    Else
                        sErrores.Add("No se encuentra configurada la Cuenta Contable para el Proveedor: " & sDR("pro_RazSoc") & ".")
                    End If
                End If

                If sDR("IVA") <> 0 Then
                    Dim sCuenta As String = ParametrosCG.CuentaIVACreditoFiscal(StrConn, xEjercicio)
                    If Not String.IsNullOrEmpty(sCuenta) Then
                        Dim sAsientoDetalle As New AsientoDetalle()
                        sAsientoDetalle.Cuenta = sCuenta
                        sAsientoDetalle.Importe = sDR("IVA")
                        If sCuentas.ContainsKey(sCuenta) Then
                            sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                        Else
                            sCuentas.Add(sCuenta, sAsientoDetalle)
                        End If
                    Else
                        sErrores.Add("No se encuentra configurada la Cuenta Contable para IVA Crédito Fiscal.")
                    End If
                End If

                'If sDR("ImpInt") <> 0 Then
                '    Dim sCuenta As String = ParametrosCG.CuentaImpuestosInternos(StrConn, xEjercicio)
                '    If Not String.IsNullOrEmpty(sCuenta) Then
                '        Dim sAsientoDetalle As New AsientoDetalle()
                '        sAsientoDetalle.Cuenta = sCuenta
                '        sAsientoDetalle.Importe = sDR("ImpInt")
                '        If sCuentas.ContainsKey(sCuenta) Then
                '            sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                '        Else
                '            sCuentas.Add(sCuenta, sAsientoDetalle)
                '        End If
                '    Else
                '        sErrores.Add("No se encuentra configurada la Cuenta Contable para Impuestos Internos.")
                '    End If
                'End If
            End While
            sDR.Close()

            'Retenciones
            sQuery = "SELECT dicres_Cod, dicres_Art, res_Desc, resccb_CodCpra, ROUND(ISNULL(SUM(dic_ImpNeto + dic_Imp1 + dic_Imp2),0)*-1,2) AS [Retenciones] "
            sQuery &= "FROM CabCompra "
            sQuery &= "INNER JOIN DetIvaCpr ON ccoemp_Cod = dicemp_Cod AND ccosuc_Cod = dicsuc_Cod AND cco_ID = diccco_ID  "
            sQuery &= "INNER JOIN TipComp On ccotco_Cod = tco_Cod And ccocir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN RegEsp ON dicres_Cod = res_Cod AND dicres_Art = res_Art "
            sQuery &= "WHERE tcocir_CodCG = 'C' AND cco_PasadoCG = 'N' AND cco_FContab >= " & SqlDate(xFechaInicio) & " AND cco_FContab <= " & SqlDate(xFechaFin) & " AND dic_TipoTot = '6' "
            sQuery &= "GROUP BY dicres_Cod, dicres_Art, res_Desc, resccb_CodCpra"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                If sDR("Retenciones") <> 0 Then
                    Dim sCuenta As String = ReplaceDBNull(sDR("resccb_CodCpra"), "")
                    If Not String.IsNullOrEmpty(sCuenta) Then
                        Dim sAsientoDetalle As New AsientoDetalle()
                        sAsientoDetalle.Cuenta = sCuenta
                        sAsientoDetalle.Importe = sDR("Retenciones")
                        If sCuentas.ContainsKey(sCuenta) Then
                            sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                        Else
                            sCuentas.Add(sCuenta, sAsientoDetalle)
                        End If
                    Else
                        sErrores.Add("No se encuentra configurada la Cuenta Contable para el Régimen Especial: " & sDR("res_Desc") & ".")
                    End If
                End If
            End While
            sDR.Close()

            Dim sCompras As New List(Of String)
            sQuery = "SELECT ccoemp_Cod, ccosuc_Cod, cco_Id FROM CabCompra "
            sQuery &= "INNER JOIN TipComp On ccotco_Cod = tco_Cod And ccocir_Cod = tcocir_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'C' AND cco_PasadoCG = 'N' AND cco_FContab >= " & SqlDate(xFechaInicio) & " AND cco_FContab <= " & SqlDate(xFechaFin)
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                sCompras.Add(sDR("ccoemp_Cod") & "|" & sDR("ccosuc_Cod") & "|" & sDR("cco_Id"))
            End While
            sDR.Close()

            If sCuentas.Count <> 0 Then
                Try
                    For Each sCuenta As AsientoDetalle In sCuentas.Values
                        sCuenta.Columna = IIf(sCuenta.Importe > 0, EnumColumna.eDebe, EnumColumna.eHaber)
                        sCuenta.Importe = Math.Abs(Math.Round(sCuenta.Importe, 2))
                        sAsiento.AgregarCuenta(sCuenta)
                    Next
                    sAsiento.Numero = Asiento.ProximoNumero(StrConn, xEjercicio)
                    sAsiento.Grabar(Nothing, sCompras)
                Catch ex As Exception
                    sErrores.Add(ex.Message)
                End Try
            Else
                sErrores.Add("No hay movimientos de Compras para el período seleccionado.")
            End If

            Return sErrores

        Catch ex As Exception
            Throw New ApplicationException("Error al Generar los Asientos de Compras: " & ex.Message & ".")
        Finally
            If sDR IsNot Nothing AndAlso Not sDR.IsClosed Then sDR.Close()
            sDR = Nothing
        End Try
    End Function

    Public Function GenerarAsientosPagos(xEjercicio As String, xFechaInicio As DateTime, xFechaFin As DateTime) As List(Of String)
        Dim sDR As DbDataReader = Nothing
        Dim sErrores As New List(Of String)
        Try
            Dim sCuentas As New SortedDictionary(Of String, AsientoDetalle)
            Dim sAsiento As New Asiento(StrConn)
            sAsiento.Ejercicio = xEjercicio
            sAsiento.Fecha = xFechaFin
            sAsiento.Descripcion = "Asiento Resumen de Pagos " & xFechaInicio.ToString("dd/MM/yyyy") & " - " & xFechaFin.ToString("dd/MM/yyyy")
            Dim sQuery As String = ""

            'Proveedores
            sQuery = "SELECT protpr_Cod, tprccb_Cod, tpr_Desc, ROUND(ISNULL(SUM(cco_ImpMonLoc),0),2) AS [Proveedores] "
            sQuery &= "FROM CabCompra "
            sQuery &= "INNER JOIN TipComp ON ccotco_Cod = tco_Cod AND ccocir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN Proveedores On ccopro_Cod = pro_Cod "
            sQuery &= "INNER JOIN CondPago ON ccocpg_Cod = cpg_Cod "
            sQuery &= "INNER JOIN TipProveedor On tpr_Cod = protpr_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'P' AND cco_PasadoCG = 'N' AND cco_FContab >= " & SqlDate(xFechaInicio) & " AND cco_FContab <= " & SqlDate(xFechaFin) & " "
            sQuery &= "GROUP BY protpr_Cod, tprccb_Cod, tpr_Desc "
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta = ReplaceDBNull(sDR("tprccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Proveedores")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para el Tipo de Proveedor: " & sDR("tpr_Desc") & ".")
                End If
            End While
            sDR.Close()

            'Compras en Efectivo - Caja
            sQuery = "SELECT mfocaj_Cod, cajccb_Cod, caj_Desc, ROUND(ISNULL(SUM(mfo_ImpMonLoc),0),2) AS Caja "
            sQuery &= "FROM CabCompra "
            sQuery &= "INNER JOIN TipComp ON ccotco_Cod = tco_Cod AND ccocir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN CondPago On ccocpg_Cod = cpg_Cod "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = ccoemp_Cod AND mfosuc_Cod = ccosuc_Cod AND mfocmf_Id = ccocmf_ID "
            sQuery &= "INNER JOIN Cajas On mfocaj_Cod = caj_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'P' AND cco_PasadoCG = 'N' AND cco_FContab >= " & SqlDate(xFechaInicio) & " AND cco_FContab <= " & SqlDate(xFechaFin) & " "
            sQuery &= " And mfo_CodConcepto = 1 GROUP BY mfocaj_Cod, cajccb_Cod, caj_Desc"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("cajccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Caja")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para la Caja: " & sDR("caj_Desc") & ".")
                End If
            End While
            sDR.Close()

            'Compras en Efectivo - Cuenta Bancaria
            sQuery = "SELECT mfobco_Cod, mfobco_Suc, mfoctb_Cod, ctbccb_Cod, ctb_Desc, ROUND(ISNULL(SUM(mfo_ImpMonLoc),0),2) AS Banco "
            sQuery &= "FROM CabCompra "
            sQuery &= "INNER JOIN TipComp ON ccotco_Cod = tco_Cod AND ccocir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN CondPago On ccocpg_Cod = cpg_Cod "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = ccoemp_Cod AND mfosuc_Cod = ccosuc_Cod AND mfocmf_Id = ccocmf_ID "
            sQuery &= "INNER JOIN CtaBan On mfobco_Cod = ctbbco_Cod And mfobco_Suc = ctbbco_Suc And mfoctb_Cod = ctb_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'P' AND cco_PasadoCG = 'N' AND cco_FContab >= " & SqlDate(xFechaInicio) & " AND cco_FContab <= " & SqlDate(xFechaFin) & " "
            sQuery &= " And mfo_CodConcepto = 4 GROUP BY mfobco_Cod, mfobco_Suc, mfoctb_Cod, ctbccb_Cod, ctb_Desc"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("ctbccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Banco")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para la Cuenta Bancaria: " & sDR("ctb_Desc") & ".")
                End If
            End While
            sDR.Close()

            'Compras en Efectivo - Cheques de Terceros
            sQuery = "SELECT tchccb_Cod, tch_Desc, ROUND(ISNULL(SUM(ch3_Importe*-1),0),2) AS [Cheque] "
            sQuery &= "FROM CabCompra "
            sQuery &= "INNER JOIN TipComp ON ccotco_Cod = tco_Cod AND ccocir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN CondPago On ccocpg_Cod = cpg_Cod "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = ccoemp_Cod AND mfosuc_Cod = ccosuc_Cod AND mfocmf_Id = ccocmf_ID "
            sQuery &= "INNER JOIN RelaChq3 ON mfoemp_Cod = rchemp_CodMF AND mfosuc_Cod = rchsuc_CodMF AND mfocmf_ID = rchcmf_ID "
            sQuery &= "INNER JOIN Cheques3 On ch3emp_Cod = rchemp_Cod And ch3suc_Cod = rchsuc_Cod And ch3_Id = rchch3_Id "
            sQuery &= "INNER JOIN TipCheq ON tch_Cod = ch3tch_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'P' AND cco_PasadoCG = 'N' AND cco_FContab >= " & SqlDate(xFechaInicio) & " AND cco_FContab <= " & SqlDate(xFechaFin) & " "
            sQuery &= "AND mfo_CodConcepto = 2 GROUP BY tchccb_Cod, tch_Desc"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("tchccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Cheque")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para el Tipo de Cheque: " & sDR("tch_Desc") & ".")
                End If
            End While
            sDR.Close()

            'Retenciones
            sQuery = "SELECT dicres_Cod, dicres_Art, res_Desc, resccb_CodCpra, ROUND(ISNULL(SUM(dic_ImpNeto + dic_Imp1 + dic_Imp2),0)*-1,2) AS [Retenciones] "
            sQuery &= "FROM CabCompra "
            sQuery &= "INNER JOIN DetIvaCpr ON ccoemp_Cod = dicemp_Cod AND ccosuc_Cod = dicsuc_Cod AND cco_ID = diccco_ID  "
            sQuery &= "INNER JOIN TipComp On ccotco_Cod = tco_Cod And ccocir_Cod = tcocir_Cod "
            sQuery &= "INNER JOIN RegEsp ON dicres_Cod = res_Cod AND dicres_Art = res_Art "
            sQuery &= "WHERE tcocir_CodCG = 'P' AND cco_PasadoCG = 'N' AND cco_FContab >= " & SqlDate(xFechaInicio) & " AND cco_FContab <= " & SqlDate(xFechaFin) & " AND dic_TipoTot = '6' "
            sQuery &= "GROUP BY dicres_Cod, dicres_Art, res_Desc, resccb_CodCpra"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                If sDR("Retenciones") <> 0 Then
                    Dim sCuenta As String = ReplaceDBNull(sDR("resccb_CodCpra"), "")
                    If Not String.IsNullOrEmpty(sCuenta) Then
                        Dim sAsientoDetalle As New AsientoDetalle()
                        sAsientoDetalle.Cuenta = sCuenta
                        sAsientoDetalle.Importe = sDR("Retenciones")
                        If sCuentas.ContainsKey(sCuenta) Then
                            sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                        Else
                            sCuentas.Add(sCuenta, sAsientoDetalle)
                        End If
                    Else
                        sErrores.Add("No se encuentra configurada la Cuenta Contable para el Régimen Especial: " & sDR("res_Desc") & ".")
                    End If
                End If
            End While
            sDR.Close()

            Dim sCompras As New List(Of String)
            sQuery = "SELECT ccoemp_Cod, ccosuc_Cod, cco_Id FROM CabCompra "
            sQuery &= "INNER JOIN TipComp On ccotco_Cod = tco_Cod And ccocir_Cod = tcocir_Cod "
            sQuery &= "WHERE tcocir_CodCG = 'C' AND cco_PasadoCG = 'N' AND cco_FContab >= " & SqlDate(xFechaInicio) & " AND cco_FContab <= " & SqlDate(xFechaFin)
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                sCompras.Add(sDR("ccoemp_Cod") & "|" & sDR("ccosuc_Cod") & "|" & sDR("cco_Id"))
            End While
            sDR.Close()

            If sCuentas.Count <> 0 Then
                Try
                    For Each sCuenta As AsientoDetalle In sCuentas.Values
                        sCuenta.Columna = IIf(sCuenta.Importe > 0, EnumColumna.eDebe, EnumColumna.eHaber)
                        sCuenta.Importe = Math.Abs(sCuenta.Importe)
                        sAsiento.AgregarCuenta(sCuenta)
                    Next
                    sAsiento.Numero = Asiento.ProximoNumero(StrConn, xEjercicio)
                    sAsiento.Grabar(Nothing, sCompras)
                Catch ex As Exception
                    sErrores.Add(ex.Message)
                End Try
            Else
                sErrores.Add("No hay movimientos de Pagos para el período seleccionado.")
            End If

            Return sErrores

        Catch ex As Exception
            Throw New ApplicationException("Error al Generar los Asientos de Pagos: " & ex.Message & ".")
        Finally
            If sDR IsNot Nothing AndAlso Not sDR.IsClosed Then sDR.Close()
            sDR = Nothing
        End Try
    End Function

    Public Function GenerarAsientosFinanzas(xEjercicio As String, xFechaInicio As DateTime, xFechaFin As DateTime) As List(Of String)
        Dim sDR As DbDataReader = Nothing
        Dim sErrores As New List(Of String)
        Try
            Dim sCuentas As New SortedDictionary(Of String, AsientoDetalle)
            Dim sAsiento As New Asiento(StrConn)
            sAsiento.Ejercicio = xEjercicio
            sAsiento.Fecha = xFechaFin
            sAsiento.Descripcion = "Asiento Resumen de Finanzas " & xFechaInicio.ToString("dd/MM/yyyy") & " - " & xFechaFin.ToString("dd/MM/yyyy")
            Dim sQuery As String = ""

            'Cajas
            sQuery = "SELECT mfocaj_Cod, cajccb_Cod, caj_Desc, ROUND(SUM(mfo_ImpMonLoc),2) AS Total, mfo_MarcaOA "
            sQuery &= "FROM CabMovF "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = cmfemp_Cod AND mfosuc_Cod = cmfsuc_Cod AND mfocmf_ID = cmf_ID "
            sQuery &= "INNER JOIN Cajas On mfocaj_Cod = caj_Cod "
            sQuery &= "WHERE cmf_FMov >= " & SqlDate(xFechaInicio) & " AND cmf_FMov <= " & SqlDate(xFechaFin) & " and cmf_CompAsoc = 0 AND mfo_CodConcepto = 1 AND cmf_PasadoCG = 'N' "
            sQuery &= "GROUP BY mfocaj_Cod, cajccb_Cod, caj_Desc, mfo_MarcaOA"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("cajccb_Cod"), "") & "|" & sDR("mfo_MarcaOA")
                If (sDR("Total") <> 0) Then
                    If Not String.IsNullOrEmpty(sCuenta) Then
                        Dim sAsientoDetalle As New AsientoDetalle()
                        sAsientoDetalle.Cuenta = sCuenta.Split("|")(0)
                        sAsientoDetalle.Importe = sDR("Total")
                        If sCuentas.ContainsKey(sCuenta) Then
                            sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                        Else
                            sCuentas.Add(sCuenta, sAsientoDetalle)
                        End If
                    Else
                        sErrores.Add("No se encuentra configurada la Cuenta Contable para la Caja: " & sDR("caj_Desc") & ".")
                    End If
                End If
            End While
            sDR.Close()

            'Cuentas Bancarias
            sQuery = "SELECT mfobco_Cod, mfobco_Suc, mfoctb_Cod, ctbccb_Cod, ctb_Desc, ROUND(SUM(mfo_ImpMonLoc),2) AS Total, mfo_MarcaOA "
            sQuery &= "FROM CabMovF "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = cmfemp_Cod AND mfosuc_Cod = cmfsuc_Cod AND mfocmf_ID = cmf_ID "
            sQuery &= "INNER JOIN CtaBan ON mfobco_Cod = ctbbco_Cod AND mfobco_Suc = ctbbco_Suc AND mfoctb_Cod = ctb_Cod "
            sQuery &= "WHERE cmf_FMov >= " & SqlDate(xFechaInicio) & " AND cmf_FMov <= " & SqlDate(xFechaFin) & " and cmf_CompAsoc = 0 AND mfo_CodConcepto = 4 AND cmf_PasadoCG = 'N' "
            sQuery &= "GROUP BY mfobco_Cod, mfobco_Suc, mfoctb_Cod, ctbccb_Cod, ctb_Desc, mfo_MarcaOA"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("ctbccb_Cod"), "") & "|" & sDR("mfo_MarcaOA")
                If (sDR("Total") <> 0) Then
                    If Not String.IsNullOrEmpty(sCuenta) Then
                        Dim sAsientoDetalle As New AsientoDetalle()
                        sAsientoDetalle.Cuenta = sCuenta.Split("|")(0)
                        sAsientoDetalle.Importe = sDR("Total")
                        If sCuentas.ContainsKey(sCuenta) Then
                            sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                        Else
                            sCuentas.Add(sCuenta, sAsientoDetalle)
                        End If
                    Else
                        sErrores.Add("No se encuentra configurada la Cuenta Contable para la Cuenta Bancaria: " & sDR("ctb_Desc") & ".")
                    End If
                End If
            End While
            sDR.Close()

            'Cheques de Terceros
            sQuery = "SELECT tch_Cod, tch_Desc, tchccb_Cod, ROUND(SUM(CASE WHEN mfo_MarcaOA = 1 THEN ch3_Importe * -1 ELSE ch3_Importe END),2) AS Total, mfo_MarcaOA "
            sQuery &= "FROM CabMovF "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = cmfemp_Cod AND mfosuc_Cod = cmfsuc_Cod AND mfocmf_ID = cmf_ID "
            sQuery &= "INNER JOIN RelaChq3 ON mfoemp_Cod = rchemp_CodMF AND mfosuc_Cod = rchsuc_CodMF AND mfocmf_ID = rchcmf_ID "
            sQuery &= "INNER JOIN Cheques3 On ch3emp_Cod = rchemp_Cod And ch3suc_Cod = rchsuc_Cod And ch3_Id = rchch3_Id "
            sQuery &= "INNER JOIN TipCheq ON tch_Cod = ch3tch_Cod "
            sQuery &= "WHERE cmf_FMov >= " & SqlDate(xFechaInicio) & " AND cmf_FMov <= " & SqlDate(xFechaFin) & " and cmf_CompAsoc = 0 AND mfo_CodConcepto = 2 AND cmf_PasadoCG = 'N' "
            sQuery &= "GROUP BY tch_Cod, tch_Desc, tchccb_Cod, mfo_MarcaOA"
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("tchccb_Cod"), "") & "|" & sDR("mfo_MarcaOA")
                If (sDR("Total") <> 0) Then
                    If Not String.IsNullOrEmpty(sCuenta) Then
                        Dim sAsientoDetalle As New AsientoDetalle()
                        sAsientoDetalle.Cuenta = sCuenta.Split("|")(0)
                        sAsientoDetalle.Importe = sDR("Total")
                        If sCuentas.ContainsKey(sCuenta) Then
                            sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                        Else
                            sCuentas.Add(sCuenta, sAsientoDetalle)
                        End If
                    Else
                        sErrores.Add("No se encuentra configurada la Cuenta Contable para el Tipo de Cheque: " & sDR("tch_Desc") & ".")
                    End If
                End If
            End While
            sDR.Close()

            'Origenes
            sQuery = "SELECT ori_Cod, ori_Desc, oriccb_Cod, ROUND(SUM(mfo_ImpMonLoc),2) AS Total "
            sQuery &= "FROM CabMovF "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = cmfemp_Cod AND mfosuc_Cod = cmfsuc_Cod AND mfocmf_ID = cmf_ID "
            sQuery &= "INNER JOIN Origenes On mfoori_Cod = ori_Cod "
            sQuery &= "WHERE cmf_FMov >= " & SqlDate(xFechaInicio) & " AND cmf_FMov <= " & SqlDate(xFechaFin) & " and cmf_CompAsoc = 0  AND mfo_CodConcepto = 6 AND cmf_PasadoCG = 'N' "
            sQuery &= "GROUP BY ori_Cod, ori_Desc, oriccb_Cod, mfo_MarcaOA "
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("oriccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Total")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para el Orígen: " & sDR("ori_Desc") & ".")
                End If
            End While
            sDR.Close()

            'Aplicaciones
            sQuery = "SELECT apl_Cod, apl_Desc, aplccb_Cod, ROUND(SUM(mfo_ImpMonLoc),2) AS Total "
            sQuery &= "FROM CabMovF "
            sQuery &= "INNER JOIN MovF ON mfoemp_Cod = cmfemp_Cod AND mfosuc_Cod = cmfsuc_Cod AND mfocmf_ID = cmf_ID "
            sQuery &= "INNER JOIN Aplicacion ON mfoapl_Cod = apl_Cod "
            sQuery &= "WHERE cmf_FMov >= " & SqlDate(xFechaInicio) & " AND cmf_FMov <= " & SqlDate(xFechaFin) & " and cmf_CompAsoc = 0  AND mfo_CodConcepto = 8 AND cmf_PasadoCG = 'N' "
            sQuery &= "GROUP BY apl_Cod, apl_Desc, aplccb_Cod "
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                Dim sCuenta As String = ReplaceDBNull(sDR("aplccb_Cod"), "")
                If Not String.IsNullOrEmpty(sCuenta) Then
                    Dim sAsientoDetalle As New AsientoDetalle()
                    sAsientoDetalle.Cuenta = sCuenta
                    sAsientoDetalle.Importe = sDR("Total")
                    If sCuentas.ContainsKey(sCuenta) Then
                        sCuentas(sCuenta).Importe = sCuentas(sCuenta).Importe + sAsientoDetalle.Importe
                    Else
                        sCuentas.Add(sCuenta, sAsientoDetalle)
                    End If
                Else
                    sErrores.Add("No se encuentra configurada la Cuenta Contable para la Aplicación: " & sDR("apl_Desc") & ".")
                End If
            End While
            sDR.Close()

            Dim sFinanzas As New List(Of String)
            sQuery = "SELECT cmfemp_Cod, cmfsuc_Cod, cmf_Id FROM CabMovF "
            sQuery &= "WHERE cmf_PasadoCG = 'N' AND cmf_FMov >= " & SqlDate(xFechaInicio) & " AND cmf_FMov <= " & SqlDate(xFechaFin)
            sDR = Conn.ExecuteDataReader(sQuery)
            While sDR.Read
                sFinanzas.Add(sDR("cmfemp_Cod") & "|" & sDR("cmfsuc_Cod") & "|" & sDR("cmf_Id"))
            End While
            sDR.Close()

            If sCuentas.Count <> 0 Then
                Try
                    For Each sCuenta As AsientoDetalle In sCuentas.Values
                        sCuenta.Columna = IIf(sCuenta.Importe > 0, EnumColumna.eDebe, EnumColumna.eHaber)
                        sCuenta.Importe = Math.Abs(sCuenta.Importe)
                        sAsiento.AgregarCuenta(sCuenta)
                    Next
                    sAsiento.Numero = Asiento.ProximoNumero(StrConn, xEjercicio)
                    sAsiento.Grabar(Nothing, Nothing, sFinanzas)
                Catch ex As Exception
                    sErrores.Add(ex.Message)
                End Try
            Else
                sErrores.Add("No hay movimientos de Finanzas para el período seleccionado.")
            End If

            Return sErrores

        Catch ex As Exception
            Throw New ApplicationException("Error al Generar los Asientos de Finanzas: " & ex.Message & ".")
        Finally
            If sDR IsNot Nothing AndAlso Not sDR.IsClosed Then sDR.Close()
            sDR = Nothing
        End Try
    End Function


End Class
