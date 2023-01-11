Imports SkY.RutinasDeBase

Public Class Asiento
    Inherits cEntityGenerico(Of AsientoMapper)

#Region "Propiedades"

    Private _Ejercicio As String
    ''' <summary>
    ''' Código del Ejercicio
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Ejercicio() As String
        Get
            Return _Ejercicio
        End Get
        Set(ByVal value As String)
            _Ejercicio = value
        End Set
    End Property

    Private _Id As Integer
    ''' <summary>
    ''' ID
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Id() As Integer
        Get
            Return _Id
        End Get
        Set(ByVal value As Integer)
            _Id = value
        End Set
    End Property

    Private _Numero As Integer
    ''' <summary>
    ''' Número de Asiento
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Numero() As Integer
        Get
            Return _Numero
        End Get
        Set(ByVal value As Integer)
            _Numero = value
        End Set
    End Property

    Private _Fecha As DateTime
    ''' <summary>
    ''' Fecha del Asiento
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Fecha() As DateTime
        Get
            Return _Fecha
        End Get
        Set(ByVal value As DateTime)
            _Fecha = value
        End Set
    End Property

    Private _Descripcion As String
    ''' <summary>
    ''' Descripción del Asiento
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Descripcion() As String
        Get
            Return _Descripcion
        End Get
        Set(ByVal value As String)
            _Descripcion = value
        End Set
    End Property

    Private _Cuentas As List(Of AsientoDetalle)
    ''' <summary>
    ''' Detalle del asiento
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Cuentas() As List(Of AsientoDetalle)
        Get
            Return _Cuentas
        End Get
        Set(ByVal value As List(Of AsientoDetalle))
            _Cuentas = value
        End Set
    End Property

#End Region

#Region "Metodos"

    Public Shared Function Grilla(ByVal xStrConn As cStrConnection, xEjercicio As String) As DataSet
        Dim sMapper As New AsientoMapper
        sMapper.StrConn = xStrConn
        Dim sResult As DataSet = sMapper.Grilla(xEjercicio)
        sMapper = Nothing
        Return sResult
    End Function

    Public Shared Function GrillaDetalleDummy(ByVal xStrConn As cStrConnection) As DataSet
        Dim sMapper As New AsientoMapper
        sMapper.StrConn = xStrConn
        Dim sResult As DataSet = sMapper.GrillaDetalleDummy()
        sMapper = Nothing
        Return sResult
    End Function

    Public Sub Cargar()
        _mapper.Cargar(Me)
    End Sub

    Public Sub Grabar(Optional xVentas As List(Of String) = Nothing, Optional xCompras As List(Of String) = Nothing, Optional xFinanzas As List(Of String) = Nothing)
        If Me.Validar(True) Then
            RenumerarDetalle()
            _mapper.Grabar(Me, xVentas, xCompras, xFinanzas)
        End If
    End Sub

    Public Sub Modificar()
        If Me.Validar(False) Then
            RenumerarDetalle()
            _mapper.Modificar(Me)
        End If
    End Sub

    Public Sub Eliminar()
        _mapper.Eliminar(Me)
    End Sub

    Public Function Validar(ByVal xAlta As Boolean) As Boolean
        If String.IsNullOrEmpty(Ejercicio) Then
            Dim sEx As New NullReferenceException("Debe ingresar el Ejercicio.")
            sEx.Source = "Ejercicio"
            Throw sEx
        ElseIf String.IsNullOrEmpty(Descripcion) Then
            Dim sEx As New NullReferenceException("Debe ingresar la Descripción del Asiento.")
            sEx.Source = "Descripcion"
            Throw sEx
        ElseIf Fecha = New DateTime Then
            Dim sEx As New NullReferenceException("Debe ingresar la Fecha del Asiento.")
            sEx.Source = "Fecha"
            Throw sEx
        ElseIf Not sky.Contable.Ejercicio.FechaEnEjercicio(Me.StrConn, Me.Ejercicio, Me.Fecha) Then
            Dim sEx As New NullReferenceException("La Fecha del Asiento debe ser válida para el Ejercicio.")
            sEx.Source = "Fecha"
            Throw sEx
        ElseIf Me.Cuentas.Count() = 0 Then
            Dim sEx As New NullReferenceException("Debe ingresar el detalle del Asiento.")
            sEx.Source = "Detalle"
            Throw sEx
        ElseIf Not Me.Balanceado Then
            Dim sEx As New NullReferenceException("El Asiento no se encuentra balanceado.")
            sEx.Source = "Detalle"
            Throw sEx
        End If

        For Each sDetalle As AsientoDetalle In Me.Cuentas
            Dim sCuenta As New Cuenta(StrConn)
            sCuenta.Ejercicio = Me.Ejercicio
            sCuenta.Codigo = sDetalle.Cuenta
            sCuenta.Cargar()

            If String.IsNullOrEmpty(sCuenta.Descripcion) Then
                Dim sEx As New NullReferenceException("La cuenta " & sDetalle.Cuenta & " no existe en el ejercicio " & Me.Ejercicio & ".")
                sEx.Source = "Detalle"
                Throw sEx
            End If
        Next

        Return True

    End Function

    Public Function Balancear() As Decimal
        If Me.Cuentas Is Nothing OrElse Me.Cuentas.Count = 0 Then Return 0
        Dim sDebe As Decimal = 0
        Dim sHaber As Decimal = 0
        For Each sCuenta As AsientoDetalle In Me.Cuentas
            If sCuenta.Columna = EnumColumna.eDebe Then
                sDebe = sDebe + sCuenta.Importe
            Else
                sHaber = sHaber + sCuenta.Importe
            End If
        Next
        If sDebe < sHaber Then Return 0
        Return sDebe - sHaber
    End Function

    Public Function TotalDebe() As Decimal
        If Me.Cuentas Is Nothing OrElse Me.Cuentas.Count = 0 Then Return 0
        Dim sDebe As Decimal = 0
        For Each sCuenta As AsientoDetalle In Me.Cuentas
            If sCuenta.Columna = EnumColumna.eDebe Then
                sDebe = sDebe + sCuenta.Importe
            End If
        Next
        Return sDebe
    End Function

    Public Function TotalHaber() As Decimal
        If Me.Cuentas Is Nothing OrElse Me.Cuentas.Count = 0 Then Return 0
        Dim sHaber As Decimal = 0
        For Each sCuenta As AsientoDetalle In Me.Cuentas
            If sCuenta.Columna = EnumColumna.eHaber Then
                sHaber = sHaber + sCuenta.Importe
            End If
        Next
        Return sHaber
    End Function

    Public Function Balanceado() As Boolean
        If Me.Cuentas Is Nothing OrElse Me.Cuentas.Count = 0 Then Return True
        Dim sDebe As Decimal = 0
        Dim sHaber As Decimal = 0
        For Each sCuenta As AsientoDetalle In Me.Cuentas
            If sCuenta.Columna = EnumColumna.eDebe Then
                sDebe = sDebe + sCuenta.Importe
            Else
                sHaber = sHaber + sCuenta.Importe
            End If
        Next
        Return (sDebe = sHaber)
    End Function

    Public Sub AgregarCuenta(xCuenta As AsientoDetalle)
        xCuenta.Id = Me.Cuentas.Count()
        Me.Cuentas.Add(xCuenta)
    End Sub

    Public Sub ModificarCuenta(xCuenta As AsientoDetalle)
        Me.Cuentas.Item(xCuenta.Id) = xCuenta
    End Sub

    Public Sub EliminarCuenta(xId As Integer)
        If Me.Cuentas.Count > xId Then Me.Cuentas.RemoveAt(xId)
        RenumerarDetalle()
    End Sub

    Public Sub RenumerarDetalle()
        Dim i As Integer = 0
        For Each sCuenta As AsientoDetalle In Me.Cuentas
            sCuenta.Id = i
            i += 1
        Next
    End Sub

    Public Shared Function Mayor(xStrConn As cStrConnection, xEjercicio As String, xCuenta As String, xFechaDesde As DateTime, xFechaHasta As DateTime) As List(Of Mayor)
        Dim sMapper As New AsientoMapper()
        sMapper.StrConn = xStrConn
        Return sMapper.Mayor(xEjercicio, xCuenta, xFechaDesde, xFechaHasta)
    End Function

    Public Shared Function SumasYSaldos(xStrConn As cStrConnection, xEjercicio As String, xFechaDesde As DateTime, xFechaHasta As DateTime) As List(Of SumasYSaldos)
        Dim sMapper As New AsientoMapper()
        sMapper.StrConn = xStrConn
        Return sMapper.SumasYSaldos(xEjercicio, xFechaDesde, xFechaHasta)
    End Function

    Public Shared Function SaldoCuenta(xStrConn As cStrConnection, xEjercicio As String, xCuenta As String, xFechaHasta As DateTime) As Double
        Dim sMapper As New AsientoMapper()
        sMapper.StrConn = xStrConn
        Return sMapper.SaldoCuenta(xEjercicio, xCuenta, xFechaHasta)
    End Function

    Public Shared Function ProximoNumero(xStrConn As cStrConnection, xEjercicio As String) As Integer
        Dim sMapper As New AsientoMapper()
        sMapper.StrConn = xStrConn
        Return sMapper.ProximoNumero(xEjercicio)
    End Function

    Public Shared Sub Renumerar(xStrConn As cStrConnection, xEjercicio As String)
        Dim sMapper As New AsientoMapper()
        sMapper.StrConn = xStrConn
        sMapper.Renumerar(xEjercicio)
    End Sub

    Public Shared Function GenerarAsientosVentas(xStrConn As cStrConnection, xEjercicio As String, xFechaInicio As DateTime, xFechaFin As DateTime) As List(Of String)
        Dim sMapper As New AsientoMapper()
        sMapper.StrConn = xStrConn
        Return sMapper.GenerarAsientosVentas(xEjercicio, xFechaInicio, xFechaFin)
    End Function

    Public Shared Function GenerarAsientosCobros(xStrConn As cStrConnection, xEjercicio As String, xFechaInicio As DateTime, xFechaFin As DateTime) As List(Of String)
        Dim sMapper As New AsientoMapper()
        sMapper.StrConn = xStrConn
        Return sMapper.GenerarAsientosCobros(xEjercicio, xFechaInicio, xFechaFin)
    End Function

    Public Shared Function GenerarAsientosCompras(xStrConn As cStrConnection, xEjercicio As String, xFechaInicio As DateTime, xFechaFin As DateTime) As List(Of String)
        Dim sMapper As New AsientoMapper()
        sMapper.StrConn = xStrConn
        Return sMapper.GenerarAsientosCompras(xEjercicio, xFechaInicio, xFechaFin)
    End Function

    Public Shared Function GenerarAsientosPagos(xStrConn As cStrConnection, xEjercicio As String, xFechaInicio As DateTime, xFechaFin As DateTime) As List(Of String)
        Dim sMapper As New AsientoMapper()
        sMapper.StrConn = xStrConn
        Return sMapper.GenerarAsientosPagos(xEjercicio, xFechaInicio, xFechaFin)
    End Function

    Public Shared Function GenerarAsientosFinanzas(xStrConn As cStrConnection, xEjercicio As String, xFechaInicio As DateTime, xFechaFin As DateTime) As List(Of String)
        Dim sMapper As New AsientoMapper()
        sMapper.StrConn = xStrConn
        Return sMapper.GenerarAsientosFinanzas(xEjercicio, xFechaInicio, xFechaFin)
    End Function

#End Region

#Region "Constructores"
    Public Sub New(ByVal xStrConn As cStrConnection)
        If xStrConn IsNot Nothing Then StrConn = xStrConn
        _Cuentas = New List(Of AsientoDetalle)
    End Sub

#End Region

End Class
