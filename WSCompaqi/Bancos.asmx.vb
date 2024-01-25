Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Xml.Serialization
Imports System.Web.Script.Services
Imports System.Xml
Imports System.Data.SqlTypes
Imports Newtonsoft.Json.JsonConvert
Imports SDKCONTPAQNGLib

' Para permitir que se llame a este servicio Web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class WSBancos
    Inherits System.Web.Services.WebService

    Public lSdkSesion As TSdkSesion
    Public ListaEmpresa As TSdkListaEmpresas
    Public proveedor As TSdkProveedor
    Public cliente As TSdkCliente
    Public cheque As TSdkCheque
    Public ingreso As TSdkIngreso
    Public ingresonodepositado As TSdkIngresoNoDepositado
    Public egreso As TSdkEgreso
    Public deposito As TSdkDeposito
    Public cuenta As TSdkCuenta
    Public cuentacheque As TSdkCuentaCheque
    Public asociacioncategoria As TSdkAsociacionCategoria
    Public BD_Empresa As String = ""
    Public sNombreEmpresa As String = ""
    Public vcCnx As String = ""

    <WebMethod()>
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=False, XmlSerializeString:=False)>
    Public Function Gestionmovtos(ByVal vpservidor As String, ByVal vpbase As String, ByVal vpusuario As String, ByVal vpclave As String,
                                   ByVal vpusuariocontpaqi As String, ByVal vpclavecontpaqi As String, ByVal vpmovimientos()() As String) As String
        Dim vpMsg As String = ""
        Dim vlSalida(3) As String
        vlSalida(0) = "ERROR"         'result_trans
        vlSalida(1) = "SIN PROCESAR"  'msg_trans
        vlSalida(2) = "ERROR"         'result_benpag
        vlSalida(3) = "SIN PROCESAR"  'msg_benpag
        vcCnx = "Server=" & vpservidor & ";Database=" & vpbase & ";Uid=" & vpusuario & ";Password=" & vpclave & ";Trusted_Connection=False;"

        If IniciaSDK(vpusuariocontpaqi, vpclavecontpaqi, vpbase, vpMsg) Then ' solo devuelve el json con campo MSG
            'aqui cuando inicializa eel sdk, es otro error, ahi se devuelve solo MSG, 
            Try
                For Each row In vpmovimientos

                    If VerificaProveedor(row, vlSalida(3)) Then
                        vlSalida(2) = "OK"
                        Select Case row(0) 'iddocumentode
                            Case "33" 'ingre9so No Depositado
                                Select Case row(16) 'tipomovimiento
                                    Case "i" : vlSalida(0) = SDK_Insert_Ingreso_No_Depositado(row, vlSalida(1))  'insert
                                    Case "u" : vlSalida(0) = SDK_Update_Ingreso_No_Depositado(row, vlSalida(1))   'update
                                    Case "d" : vlSalida(0) = SDK_Delete_Ingreso_No_Depositado(row, vlSalida(1))   'delete
                                    Case Else
                                        vlSalida(0) = "ERROR"
                                        vlSalida(1) = "Los Movimientos Permitidos son Insert, Update y Delete"
                                End Select
                            Case "34" 'ingreso
                                Select Case row(16) 'tipomovimiento
                                    Case "i" : vlSalida(0) = SDK_Insert_Ingreso(row, vlSalida(1))  'insert
                                    Case "u" : vlSalida(0) = SDK_Update_Ingreso(row, vlSalida(1))   'update
                                    Case "d" : vlSalida(0) = SDK_Delete_Ingreso(row, vlSalida(1))   'delete
                                    Case Else
                                        vlSalida(0) = "ERROR"
                                        vlSalida(1) = "Los Movimientos Permitidos son Insert, Update y Delete"
                                End Select
                            Case "35" 'depósito
                                Select Case row(16) 'tipomovimiento
                                    Case "i" : vlSalida(0) = SDK_Insert_Deposito(row, vlSalida(1))  'insert
                                    Case "u" : vlSalida(0) = SDK_Update_Deposito(row, vlSalida(1))   'update
                                    Case "d" : vlSalida(0) = SDK_Delete_Deposito(row, vlSalida(1))   'delete
                                    Case Else
                                        vlSalida(0) = "ERROR"
                                        vlSalida(1) = "Los Movimientos Permitidos son Insert, Update y Delete"
                                End Select

                            Case "36" ' egreso
                                Select Case row(16) 'tipomovimiento
                                    Case "i" : vlSalida(0) = SDK_Insert_Egreso(row, vlSalida(1))    'insert
                                    Case "u" : vlSalida(0) = SDK_Update_Egreso(row, vlSalida(1))   'update
                                    Case "d" : vlSalida(0) = SDK_Delete_Egreso(row, vlSalida(1))   'delete
                                    Case Else
                                        vlSalida(0) = "ERROR"
                                        vlSalida(1) = "Los Movimientos Permitidos son Insert, Update y Delete"
                                End Select
                            Case "37" 'cheque
                                Select Case row(16) 'tipomovimiento
                                    Case "i" : vlSalida(0) = SDK_Insert_Cheque(row, vlSalida(1))    'insert
                                    Case "u" : vlSalida(0) = SDK_Update_Cheque(row, vlSalida(1))    'update
                                    Case "c" : vlSalida(0) = SDK_Cancel_Cheque(row, vlSalida(1))    'cancel
                                    Case "d" : vlSalida(0) = SDK_Delete_Cheque(row, vlSalida(1))    'delete
                                    Case Else
                                        vlSalida(0) = "ERROR"
                                        vlSalida(1) = "Los Movimientos Permitidos son Insert, Update y Cancel"
                                End Select
                        End Select
                    End If
                Next
                Kill_SDKCONTPAQNG()
                Return Newtonsoft.Json.JsonConvert.SerializeObject(MsgArreglo2(vpmovimientos, vlSalida))
            Catch ex As Exception
                Kill_SDKCONTPAQNG()
                Return Newtonsoft.Json.JsonConvert.SerializeObject(MsgError(ex.Source & "-" & ex.Message))
            End Try
        Else
            Kill_SDKCONTPAQNG()
            Return Newtonsoft.Json.JsonConvert.SerializeObject(MsgError(vpMsg))
        End If
    End Function

    Function SDK_Insert_Cheque(vprow() As String, ByRef vpMsg As String) As String
        Dim vpError As Integer
        Try
            With cheque
                .iniciarInfo()
                .TipoDocumento = vprow(1)
                .Folio = vprow(2)
                .Fecha = CDate(vprow(3))
                .FechaAplicacion = CDate(vprow(4))
                .CodigoPersona = vprow(5)
                .BeneficiarioPagador = vprow(6)
                .IdCuentaCheques = CInt(vprow(12))
                .Total = CDbl(vprow(13))
                .Referencia = vprow(14)
                .Concepto = vprow(15)
                .Origen = 202
                .CodigoMonedaTipoCambio = 2
                .TipoCambio = 1
                .EsImpreso = 1
                .EsAsociado = 0
                vpError = .crea()
            End With
            If vpError = 0 Then
                vpMsg = "SDK: " & cheque.getMensajeError
                Return "ERROR"
            Else
                vpMsg = "ID: " & cheque.Id
                Return "OK"
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function

    Function SDK_Update_Cheque(vprow() As String, ByRef vpMsg As String) As String

        Dim vpError As Integer
        Dim IdCheque As Integer
        Dim vlMsg As String = ""
        Try
            cheque.iniciarInfo()
            IdCheque = Busca_Cheque(vprow(1), vprow(2), vprow(12), vlMsg)
            If IdCheque = 0 Then
                vpMsg = vlMsg ' jlara 12-jun-2023 se debe devolver el mensaje de busqueda
                Return "ERROR"
            Else
                With cheque
                    .CodigoPersona = vprow(5)
                    .Referencia = vprow(14)
                    .Concepto = vprow(15)
                    vpError = .modifica()
                End With
                If vpError = 0 Then
                    vpMsg = "SDK: " & cheque.getMensajeError
                    Return "ERROR"
                Else
                    vpMsg = "ID: " & cheque.Id
                    Return "OK"
                End If
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function

    Function SDK_Cancel_Cheque(vprow() As String, ByRef vpMsg As String) As String
        Dim vpError As Integer
        Dim IdCheque As Integer
        Dim vlMsg As String = ""
        Try
            cheque.iniciarInfo()
            IdCheque = Busca_Cheque(vprow(1), vprow(2), vprow(12), vlMsg) ' TipoDocumento, Folio, IdCuentaCheques , mensaje de retorno
            If IdCheque = 0 Then
                vpMsg = vlMsg ' jlara 12-jun-2023 se debe devolver el mensaje de busqueda
                Return "ERROR"
            Else
                vpError = cheque.cancela(IdCheque, 0)

                If vpError = 0 Then
                    vpMsg = "SDK: " & cheque.getMensajeError
                    Return "ERROR"
                Else
                    vpMsg = "ID: " & cheque.Id
                    Return "OK"
                End If
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function
    Function SDK_Delete_Cheque(vprow() As String, ByRef vpMsg As String) As String
        Dim vpError As Integer
        Dim IdCheque As Integer
        Dim vlMsg As String = ""
        Try
            cheque.iniciarInfo()
            IdCheque = Busca_Cheque(vprow(1), vprow(2), vprow(12), vlMsg)
            If IdCheque = 0 Then
                vpMsg = vlMsg
                Return "ERROR"
            Else
                vpError = cheque.borra()
                If vpError = 0 Then
                    vpMsg = "SDK: " & cheque.getMensajeError
                    Return "ERROR"
                Else
                    vpMsg = "ID: " & cheque.Id
                    Return "OK"
                End If
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function
    Function SDK_Insert_Egreso(vprow() As String, ByRef vpMsg As String) As String
        Dim vpError As Integer
        Try
            With egreso
                .iniciarInfo()
                .TipoDocumento = vprow(1)
                .Folio = vprow(2)
                .Fecha = CDate(vprow(3))
                .FechaAplicacion = CDate(vprow(4))
                .CodigoPersona = vprow(5)
                .BeneficiarioPagador = vprow(6)
                .IdCuentaCheques = CInt(vprow(12))
                .Total = CDbl(vprow(13))
                .Referencia = vprow(14)
                .Concepto = vprow(15)
                .Origen = 202
                .CodigoMonedaTipoCambio = 2
                .TipoCambio = 1
                vpError = .crea()
            End With
            If vpError = 0 Then
                vpMsg = "SDK: " & egreso.getMensajeError
                Return "ERROR"
            Else
                vpMsg = "ID: " & egreso.Id
                Return "OK"
            End If

        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function
    Function SDK_Update_Egreso(vprow() As String, ByRef vpMsg As String) As String
        Dim vpError As Integer
        Dim IdEgreso As Integer
        Dim vlMsg As String = ""
        Try
            egreso.iniciarInfo()
            IdEgreso = Busca_Egreso(vprow(0), vprow(1), vprow(2), vprow(12), vlMsg)
            If IdEgreso = 0 Then
                vpMsg = vlMsg ' jlara 12-jun-2023 se debe devolver el mensaje de busqueda
                Return "ERROR"
            Else
                With egreso
                    .CodigoPersona = vprow(5)
                    .Referencia = vprow(14)
                    .Concepto = vprow(15)
                    vpError = .modifica()
                End With

                If vpError = 0 Then
                    vpMsg = "SDK: " & egreso.getMensajeError
                    Return "ERROR"
                Else
                    vpMsg = "ID: " & egreso.Id
                    Return "OK"
                End If
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function
    Function SDK_Delete_Egreso(vprow() As String, ByRef vpMsg As String) As String
        Dim vpError As Integer
        Dim IdEgreso As Integer
        Dim vlMsg As String = ""
        Try
            egreso.iniciarInfo()
            IdEgreso = Busca_Egreso(vprow(0), vprow(1), vprow(2), vprow(12), vlMsg)
            If IdEgreso = 0 Then
                vpMsg = vlMsg ' jlara 12-jun-2023 se debe devolver el mensaje de busqueda
                Return "ERROR"
            Else
                vpError = egreso.borra()
                If vpError = 0 Then
                    vpMsg = "SDK: " & egreso.getMensajeError
                    Return "ERROR"
                Else
                    vpMsg = "ID: " & egreso.Id
                    Return "OK"
                End If
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function



    Function SDK_Insert_Ingreso_No_depositado(vprow() As String, ByRef vpMsg As String) As String
        Dim vpError As Integer
        Try
            With ingresonodepositado
                .iniciarInfo()
                .TipoDocumento = vprow(1)
                .Folio = vprow(2)
                .Fecha = CDate(vprow(3))
                .FechaAplicacion = CDate(vprow(4))
                .CodigoPersona = vprow(5)
                .BeneficiarioPagador = vprow(6)
                .IdCuentaCheques = CInt(vprow(12))
                .Total = CDbl(vprow(13))
                .Referencia = vprow(14)
                .Concepto = vprow(15)
                .Origen = 202
                .CodigoMonedaTipoCambio = 2
                .TipoCambio = 1
                '.BancoOrigen
                '.CuentaOrigen
                vpError = .crea()
            End With
            If vpError = 0 Then
                vpMsg = "SDK: " & ingresonodepositado.getMensajeError
                Return "ERROR"
            Else
                vpMsg = "ID: " & ingresonodepositado.Id
                Return "OK"
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function
    Function SDK_Update_Ingreso_No_Depositado(vprow() As String, ByRef vpMsg As String) As String
        Dim vpError As Integer
        Dim IdIngreso As Integer
        Dim vlMsg As String = ""
        Try
            ingresonodepositado.iniciarInfo()
            IdIngreso = Busca_Ingreso_no_depositado(vprow(1), vprow(2), vlMsg) ' vpTipoDocto, vpFolio, vpMsg 
            If IdIngreso = 0 Then
                vpMsg = vlMsg '"No se encontró el ingreso "
                Return "ERROR"
            Else
                With ingresonodepositado
                    .CodigoPersona = vprow(5)
                    .Referencia = vprow(14)
                    .Concepto = vprow(15)
                    vpError = .modifica()
                End With
                If vpError = 0 Then
                    vpMsg = "SDK: " & ingresonodepositado.getMensajeError
                    Return "ERROR"
                Else
                    vpMsg = "ID: " & ingresonodepositado.Id
                    Return "OK"
                End If
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function

    Function SDK_Delete_Ingreso_No_Depositado(vprow() As String, ByRef vpMsg As String) As String
        Dim vpError As Integer
        Dim IdIngreso As Integer
        Dim vlMsg As String = ""
        Try
            ingresonodepositado.iniciarInfo()
            IdIngreso = Busca_Ingreso_no_depositado(vprow(1), vprow(2), vlMsg) ' vpTipoDocto, vpFolio, vpMsg 
            If IdIngreso = 0 Then
                vpMsg = vlMsg ' jlara 12-jun-2023 se debe devolver el mensaje de busqueda
                Return "ERROR"
            Else
                vpError = ingresonodepositado.borra()
                If vpError = 0 Then
                    vpMsg = "SDK: " & ingresonodepositado.getMensajeError
                    Return "ERROR"
                Else
                    vpMsg = "ID: " & ingresonodepositado.Id
                    Return "OK"
                End If
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function

    Function SDK_Insert_Ingreso(vprow() As String, ByRef vpMsg As String) As String
        Dim vpError As Integer
        Try
            With ingreso
                .iniciarInfo()
                .TipoDocumento = vprow(1)
                .Folio = vprow(2)
                .Fecha = CDate(vprow(3))
                .FechaAplicacion = CDate(vprow(4))
                .CodigoPersona = vprow(5)
                .BeneficiarioPagador = vprow(6)
                .IdCuentaCheques = CInt(vprow(12))
                .Total = CDbl(vprow(13))
                .Referencia = vprow(14)
                .Concepto = vprow(15)
                .Origen = 202
                .CodigoMonedaTipoCambio = 2
                .TipoCambio = 1
                '.BancoOrigen
                '.CuentaOrigen
                vpError = .crea()
            End With
            If vpError = 0 Then
                vpMsg = "SDK: " & ingreso.getMensajeError
                Return "ERROR"
            Else
                vpMsg = "ID: " & ingreso.Id
                Return "OK"
            End If

        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function
    Function SDK_Update_Ingreso(vprow() As String, ByRef vpMsg As String) As String
        Dim vpError As Integer
        Dim IdIngreso As Integer
        Dim vlMsg As String = ""
        Try
            ingreso.iniciarInfo()
            IdIngreso = Busca_Ingreso(vprow(0), vprow(1), vprow(2), vprow(12), vlMsg) ' vpIdDocumentoDe, vpTipoDocto, vpFolio,vpIdCuentaCheques, vpMsg 
            If IdIngreso = 0 Then
                vpMsg = vlMsg ' jlara 12-jun-2023 se debe devolver el mensaje de busqueda
                Return "ERROR"
            Else
                With ingreso
                    .CodigoPersona = vprow(5)
                    .Referencia = vprow(14)
                    .Concepto = vprow(15)
                    vpError = .modifica()
                End With
                If vpError = 0 Then
                    vpMsg = "SDK: " & ingreso.getMensajeError
                    Return "ERROR"
                Else
                    vpMsg = "ID: " & ingreso.Id
                    Return "OK"
                End If
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function
    Function SDK_Delete_Ingreso(vprow() As String, ByRef vpMsg As String) As String
        Dim vpError As Integer
        Dim IdIngreso As Integer
        Dim vlMsg As String = ""
        Try
            ingreso.iniciarInfo()
            IdIngreso = Busca_Ingreso(vprow(0), vprow(1), vprow(2), vprow(12), vlMsg)
            If IdIngreso = 0 Then
                vpMsg = vlMsg ' jlara 12-jun-2023 se debe devolver el mensaje de busqueda
                Return "ERROR"
            Else
                vpError = ingreso.borra()
                If vpError = 0 Then
                    vpMsg = "SDK: " & ingreso.getMensajeError
                    Return "ERROR"
                Else
                    vpMsg = "ID: " & ingreso.Id
                    Return "OK"
                End If
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function

    Function SDK_Insert_Deposito(vprow() As String, ByRef vpMsg As String) As String
        Dim vpError As Integer
        Try
            With deposito
                .iniciarInfo()
                .TipoDocumento = vprow(1)
                .Folio = vprow(2)
                .Fecha = CDate(vprow(3))
                .FechaAplicacion = CDate(vprow(4))
                .IdCuentaCheques = CInt(vprow(12))
                .Total = CDbl(vprow(13))
                .Referencia = vprow(14)
                .Concepto = vprow(15)
                .Origen = 202
                '.BancoOrigen
                '.CuentaOrigen
                vpError = .crea()
            End With
            If vpError = 0 Then
                vpMsg = "SDK: " & deposito.getMensajeError
                Return "ERROR"
            Else
                vpMsg = "ID: " & deposito.Id
                Return "OK"
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function

    Function SDK_Update_Deposito(vprow() As String, ByRef vpMsg As String) As String
        Dim vpError As Integer
        Dim IdDeposito As Integer
        Dim vlMsg As String = ""
        Try
            deposito.iniciarInfo()
            IdDeposito = Busca_Deposito(vprow(1), vprow(2), vlMsg) '  vpTipoDocto, vpFolio, vpMsg 
            If IdDeposito = 0 Then
                vpMsg = vlMsg ' jlara 12-jun-2023 se debe devolver el mensaje de busqueda
                Return "ERROR"
            Else
                With deposito
                    .Referencia = vprow(14)
                    .Concepto = vprow(15)
                    vpError = .modifica()
                End With
                If vpError = 0 Then
                    vpMsg = "SDK: " & deposito.getMensajeError
                    Return "ERROR"
                Else
                    vpMsg = "ID: " & deposito.Id
                    Return "OK"
                End If
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function
    Function SDK_Delete_Deposito(vprow() As String, ByRef vpMsg As String) As String
        Dim vpError As Integer
        Dim IdDeposito As Integer
        Dim vlMsg As String = ""
        Try
            deposito.iniciarInfo()
            IdDeposito = Busca_Deposito(vprow(1), vprow(2), vlMsg) '  vpTipoDocto, vpFolio, vpMsg 
            If IdDeposito = 0 Then
                vpMsg = vlMsg ' jlara 12-jun-2023 se debe devolver el mensaje de busqueda
                Return "ERROR"
            Else
                vpError = deposito.borra()
                If vpError = 0 Then
                    vpMsg = "SDK: " & deposito.getMensajeError
                    Return "ERROR"
                Else
                    vpMsg = "ID: " & deposito.Id
                    Return "OK"
                End If
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function

    Private Function Busca_Ingreso_no_depositado(vpTipoDocto As String, vpFolio As String, ByRef vpMsg As String) As Integer
        Dim vlResult As Integer
        Try
            vlResult = ingresonodepositado.buscaPorTipoDocumentoFolio(vpTipoDocto, vpFolio)
            If vlResult = 0 Then
                vpMsg = "SDK: " & ingresonodepositado.getMensajeError 'no se encontró ningún registro 
                Return 0
            Else
                vpMsg = ""
                Return ingresonodepositado.Id
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return 0
        End Try
    End Function

    Private Function Busca_Ingreso(vpIdDocumentoDe As Integer, vpTipoDocto As String, vpFolio As String,
                                   vpIdCuentaCheques As Integer, ByRef vpMsg As String) As Integer
        Dim vlResult As Integer
        Try
            vlResult = ingreso.buscaPorCuentaTipoDoctoFolio(vpIdCuentaCheques, vpIdDocumentoDe, vpTipoDocto, vpFolio)
            If vlResult = 0 Then
                vpMsg = "SDK: " & ingreso.getMensajeError 'no se encontró ningún registro 
                Return 0
            Else
                vpMsg = ""
                Return ingreso.Id
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return 0
        End Try
    End Function
    Private Function Busca_Deposito(vpTipoDocto As String, vpFolio As String, ByRef vpMsg As String) As Integer
        Dim vlResult As Integer
        Try
            vlResult = deposito.buscaPorTipoDocumentoFolio(vpTipoDocto, vpFolio) ' 
            If vlResult = 0 Then
                vpMsg = "SDK: " & deposito.getMensajeError 'no se encontró ningún registro 
                Return 0
            Else
                vpMsg = ""
                Return deposito.Id
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return 0
        End Try
    End Function
    Private Function Busca_Egreso(vpIdDocumentoDe As Integer, vpTipoDocto As String, vpFolio As String,
                                  vpIdCuentaCheques As Integer, ByRef vpMsg As String) As Integer
        Dim vlResult As Integer
        Try
            vlResult = egreso.buscaPorCuentaTipoDoctoFolio(vpIdCuentaCheques, vpIdDocumentoDe, vpTipoDocto, vpFolio)
            If vlResult = 0 Then
                vpMsg = "SDK: " & egreso.getMensajeError 'no se encontró ningún registro 
                Return 0
            Else
                vpMsg = ""
                Return egreso.Id
            End If
        Catch ex As Exception
            vpMsg = ex.Source & " - " & ex.Message
            Return 0
        End Try
    End Function
    Private Function Busca_Cheque(vpTipoDocto As String, vpFolio As String, vpIdCuentaCheques As Integer, ByRef vpMsg As String) As Integer
        Dim vlResult As Integer
        Try

            vlResult = cheque.consultaPorTipoDocumentoFolio_buscaPorLlave(vpTipoDocto, vpFolio)
            If vlResult = 0 Then
                vpMsg = "SDK: " & cheque.getMensajeError 'no se encontró ningún registro 
                Return 0
            End If
            Do While cheque.Id > 0
                If vpIdCuentaCheques = cheque.IdCuentaCheques Then
                    vpMsg = ""
                    Return cheque.Id
                End If
                cheque.consultaPorTipoDocumentoFolio_buscaSiguiente()
            Loop
            vpMsg = "Cheque no encontrado, Folio: " & vpFolio 'no se encontró ningún registro 
            Return 0
        Catch ex As Exception
            vpMsg = ex.Source & "-" & ex.Message
            Return 0
        End Try
    End Function

    Function VerificaProveedor(ByVal vpProveedor() As String, ByRef vpmsg As String) As Boolean
        Dim vlResult As Integer
        VerificaProveedor = False
        vpmsg = ""
        Try
            If vpProveedor(0) <> "35" Then ' el movimiento 35 depósito, no requiere beneficiario pagador jlara 22/06/23
                If vpProveedor(16) = "i" Or vpProveedor(16) = "u" Then
                    proveedor.iniciarInfo()
                    vlResult = proveedor.buscaPorCodigo(vpProveedor(5))
                    If vlResult = 0 Then
                        Return Insert_BeneficiarioPagador(vpProveedor, vpmsg)
                    Else
                        'actualizar cuenta
                        If proveedor.Nombre <> vpProveedor(6) Or proveedor.RFC <> vpProveedor(7) Or
                        proveedor.CURP <> vpProveedor(8) Or
                        proveedor.CodigoCuenta <> vpProveedor(9) Or
                        proveedor.CodigoSegNeg <> vpProveedor(10) Or
                        proveedor.CodigoPrepoliza <> vpProveedor(11) Then
                            Return Update_BeneficiarioPagador(vpProveedor, vpmsg)
                        Else
                            vpmsg = "ID: " & proveedor.Id
                            Return True
                        End If

                    End If

                Else
                    vpmsg = "ID: "
                    Return True
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            vpmsg = ex.Source & " - " & ex.Message
        End Try
    End Function
    Function Insert_BeneficiarioPagador(vprow() As String, ByRef vpmsg As String) As Boolean
        Dim vpError As Integer
        Insert_BeneficiarioPagador = False
        Try
            proveedor.iniciarInfo()
            cliente.iniciarInfo()
            proveedor.EsProveedor = 1
            proveedor.EsCliente = 1
            proveedor.Codigo = vprow(5)
            proveedor.Nombre = vprow(6)
            proveedor.RFC = vprow(7) '"" ' marca error la libreria al insertar el registro "ER: El índice estaba fuera del intervalo. Debe ser un valor no negativo e inferior al tamaño de la colección. Nombre del parámetro: Index " al usar el rfc generico "XAXX010101000"
            proveedor.CURP = vprow(8) '""
            proveedor.CodigoCuenta = vprow(9) '""

            proveedor.CodigoSegNeg = vprow(10)
            proveedor.CodigoPrepoliza = vprow(11)

            proveedor.TipoTercero = SDKCONTPAQNGLib.ETIPOTERCERO.TIPOTER_PROVNACIONAL '4
            proveedor.TipoOperacion = SDKCONTPAQNGLib.ETIPOOPERACION.TIPOOPE_PRESERVPRO ' 3
            proveedor.FechaRegistro = DateTime.Now
            proveedor.Nacionalidad = "Mexicana"
            proveedor.TasaAsumida = SDKCONTPAQNGLib.ETASAASUMIDAPROV.TAASUPRO_TASAOTRA1  '5
            proveedor.UsaTasaIVAOtra1 = ETASAASUMIDAPROV.TAASUPRO_TASA0
            proveedor.SistOrig = 201
            proveedor.IdFiscal = ""
            proveedor.EsParaAbonoCta = 1    'jlara 27-abr-2023 checar si es funcional la activacion
            proveedor.GenerarPolizaAuto = 1 'jlara 27-abr-2023 checar si es funcional la activacion
            cliente.Codigo = proveedor.Codigo
            cliente.IdSegNeg = 0
            cliente.agregaCliente(cliente)
            proveedor.agregaProveedor(proveedor)
            vpError = proveedor.crea()
            If vpError = 0 Then
                vpmsg = proveedor.getMensajeError
                Return False
            Else
                If Update_BeneficiarioPagador_CodigoCuenta(proveedor.Id, vprow(9), vpmsg) Then
                    vpmsg = "ID: " & proveedor.Id
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            vpmsg = ex.Source & " - " & ex.Message
            Return "ERROR"
        End Try
    End Function
    Function Update_BeneficiarioPagador(vprow() As String, ByRef vpmsg As String) As Boolean
        Dim vpError As Integer
        vpmsg = ""
        Try
            proveedor.Nombre = vprow(6)
            proveedor.RFC = vprow(7)
            proveedor.CURP = vprow(8)
            proveedor.CodigoCuenta = vprow(9)
            proveedor.CodigoSegNeg = vprow(10)
            proveedor.CodigoPrepoliza = vprow(11)
            vpError = proveedor.modifica
            If vpError = 0 Then
                vpmsg = proveedor.getMensajeError
                Return False
            Else
                If Update_BeneficiarioPagador_CodigoCuenta(proveedor.Id, vprow(9), vpmsg) Then
                    vpmsg = "ID: " & proveedor.Id
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            vpmsg = ex.Source & " - " & ex.Message
            Return False
        End Try
    End Function
    Public Function Update_BeneficiarioPagador_CodigoCuenta(vpProveedorId As Integer, vpCodigoCuenta As String, ByRef vpmsg As String) As Boolean
        Dim vmCNX As New SqlConnection
        Dim vmSQLDA As New SqlDataAdapter
        Dim vmDS As New DataSet
        Dim vmSQLCMD As New SqlCommand

        Dim vlComando As String
        Dim vlResult As Integer
        Try
            cuenta.iniciarInfo()
            vlResult = cuenta.buscaPorCodigo(vpCodigoCuenta)
            If vlResult = 0 Then
                vpmsg = "SDK: " & cuenta.getMensajeError 'no se encontró ningún registro 
                Return 0
            Else
                vmCNX.ConnectionString = vcCnx
                'vlComando = "Update Proveedores set cuenta = '" & vpCodigoCuenta & "', IdCuenta = " & CStr(cuenta.Id) & " Where id = " & proveedor.Id
                vlComando = "BEGIN TRANSACTION" &
                         " BEGIN TRY" &
                         " Update Proveedores set CodigoCuenta = '" & vpCodigoCuenta & "', IdCuenta = " & CStr(cuenta.Id) & " Where id = " & CStr(vpProveedorId) &
                         " COMMIT TRANSACTION" &
                         " select 1 as RowsAfected" &
                         " END TRY" &
                         " BEGIN CATCH" &
                         " ROLLBACK TRANSACTION" &
                         " Select 0 as RowsAfected" &
                         " END CATCH"
                vmSQLCMD.CommandText = vlComando
                vmSQLCMD.CommandType = CommandType.Text
                vmSQLCMD.Connection = vmCNX
                vmSQLDA.SelectCommand = vmSQLCMD
                vmSQLDA.Fill(vmDS)
                vpmsg = ""
                Return CBool(vmDS.Tables(0).Rows(0)("RowsAfected"))
            End If
        Catch ex As Exception
            vpmsg = ex.Source & " - " & ex.Message
            Return False
        End Try
    End Function

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    <WebMethod()>
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=False, XmlSerializeString:=False)>
    Public Function GestionBeneficiarioPagador(ByVal vpservidor As String, ByVal vpbase As String, ByVal vpusuario As String, ByVal vpclave As String,
                                   ByVal vpusuariocontpaqi As String, ByVal vpclavecontpaqi As String, ByVal vpmovimientos()() As String) As String
        Dim vpMsg As String = ""
        Dim vlSalida(3) As String
        vlSalida(0) = "ERROR"         'result_trans
        vlSalida(1) = "SIN PROCESAR"  'msg_trans
        vlSalida(2) = "ERROR"         'result_benpag
        vlSalida(3) = "SIN PROCESAR"  'msg_benpag
        vcCnx = "Server=" & vpservidor & ";Database=" & vpbase & ";Uid=" & vpusuario & ";Password=" & vpclave & ";Trusted_Connection=False;"

        If IniciaSDK(vpusuariocontpaqi, vpclavecontpaqi, vpbase, vpMsg) Then ' solo devuelve el json con campo MSG
            'aqui cuando inicializa eel sdk, es otro error, ahi se devuelve solo MSG, 
            Try
                For Each row In vpmovimientos
                    Select Case row(5) 'tipomovimiento
                        ' Case "i" : vlSalida(0) = SDK_Insert_BeneficiarioPagador(row, vlSalida(1))  'insert
                        ' Case "u" : vlSalida(0) = SDK_Update_BeneficiarioPagador(row, vlSalida(1))   'update
                        ' Case "d" : vlSalida(0) = SDK_Delete_BeneficiarioPagador(row, vlSalida(1))   'delete
                        ' Case Else
                        ' vlSalida(0) = "ERROR"
                        'vlSalida(1) = "Los Movimientos Permitidos son Insert, Update y Delete"
                    End Select
                Next
                Kill_SDKCONTPAQNG()
                Return Newtonsoft.Json.JsonConvert.SerializeObject(MsgArreglo2(vpmovimientos, vlSalida))
            Catch ex As Exception
                Kill_SDKCONTPAQNG()
                Return Newtonsoft.Json.JsonConvert.SerializeObject(MsgError(ex.Source & "-" & ex.Message))
            End Try
        Else
            Kill_SDKCONTPAQNG()
            Return Newtonsoft.Json.JsonConvert.SerializeObject(MsgError(vpMsg))
        End If
    End Function

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    <WebMethod()>
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=False, XmlSerializeString:=False)>
    Public Function Catcuentascheques(ByVal vpservidor As String, ByVal vpbase As String, ByVal vpusuario As String, ByVal vpclave As String) As String
        Dim vmCNX As New SqlConnection
        Dim vmSQLDA As New SqlDataAdapter
        Dim vmDS As New DataSet
        Dim vmSQLCMD As New SqlCommand

        Dim vlComando As String
        Try
            Dim vlCnx As String
            vlCnx = "Server=" & vpservidor & ";Database=" & vpbase & ";Uid=" & vpusuario & ";Password=" & vpclave & ";Trusted_Connection=False;"
            vmCNX.ConnectionString = vlCnx
            vlComando = "select Id, Codigo, Nombre from CuentasCheques"
            vmSQLCMD.CommandText = vlComando
            vmSQLCMD.CommandType = CommandType.Text
            vmSQLCMD.Connection = vmCNX
            vmSQLDA.SelectCommand = vmSQLCMD
            vmSQLDA.Fill(vmDS)
            Return Newtonsoft.Json.JsonConvert.SerializeObject(vmDS.Tables(0))
        Catch ex As Exception
            Return Newtonsoft.Json.JsonConvert.SerializeObject(MsgError(ex.Source & "-" & ex.Message))
        End Try
    End Function

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    <WebMethod()>
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=False, XmlSerializeString:=False)>
    Public Function Catprepolizas(ByVal vpservidor As String, ByVal vpbase As String, ByVal vpusuario As String, ByVal vpclave As String) As String
        Dim vmCNX As New SqlConnection
        Dim vmSQLDA As New SqlDataAdapter
        Dim vmDS As New DataSet
        Dim vmSQLCMD As New SqlCommand

        Dim vlComando As String
        Try
            Dim vlCnx As String
            vlCnx = "Server=" & vpservidor & ";Database=" & vpbase & ";Uid=" & vpusuario & ";Password=" & vpclave & ";Trusted_Connection=False;"
            vmCNX.ConnectionString = vlCnx
            vlComando = "select Trim(Codigo) as Codigo, Trim(Nombre) as Nombre from prepolizas"
            vmSQLCMD.CommandText = vlComando
            vmSQLCMD.CommandType = CommandType.Text
            vmSQLCMD.Connection = vmCNX
            vmSQLDA.SelectCommand = vmSQLCMD
            vmSQLDA.Fill(vmDS)
            Return Newtonsoft.Json.JsonConvert.SerializeObject(vmDS.Tables(0))
        Catch ex As Exception
            Return Newtonsoft.Json.JsonConvert.SerializeObject(MsgError(ex.Source & "-" & ex.Message))
        End Try
    End Function

    '---------------------------------------------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------------------------------

    <WebMethod()>
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=False, XmlSerializeString:=False)>
    Public Function GestionmovtosEdoCuenta(ByVal vpservidor As String, ByVal vpbase As String, ByVal vpusuario As String, ByVal vpclave As String,
                                   ByVal vpusuariocontpaqi As String, ByVal vpclavecontpaqi As String, ByVal vpmovimientos()() As String) As String

        Dim associatedChar As Char = Chr(34)

        Dim vpMsg As String = ""
        Dim vlSalida(3) As String
        vlSalida(0) = "ERROR"         'result_trans
        vlSalida(1) = "SIN PROCESAR"  'msg_trans
        vlSalida(2) = "ERROR"         'result_benpag
        vlSalida(3) = "SIN PROCESAR"  'msg_benpag
        vcCnx = "Server=" & vpservidor & ";Database=" & vpbase & ";Uid=" & vpusuario & ";Password=" & vpclave & ";Trusted_Connection=False;"

        Try
            For Each row In vpmovimientos
                Select Case row(2).ToUpper 'tipomovimiento
                    Case "I", "C", "E" : vlSalida(0) = Insert_EdoCtaBancos(row, vlSalida(1))
                    Case Else
                        vlSalida(0) = "ERROR"
                        vlSalida(1) = "Los Movimientos Permitidos son Ingreso, Egreso y Cheque"
                End Select
            Next
            '  Return Newtonsoft.Json.JsonConvert.SerializeObject(MsgArreglo2(vpmovimientos, vlSalida)) 'MsgArregloEdoCuenta
            'Return Newtonsoft.Json.JsonConvert.SerializeObject(vlSalida(1))
            Return Newtonsoft.Json.JsonConvert.SerializeObject(MsgArregloEdoCuenta(vpmovimientos, vlSalida)) '

        Catch ex As Exception
            Return Newtonsoft.Json.JsonConvert.SerializeObject(MsgError(ex.Source & "-" & ex.Message))
        End Try

    End Function

    Public Function Insert_EdoCtaBancos(vprow() As String, ByRef vpmsg As String) As Boolean

        Dim TR_sTipoMovto As String, TR_sImporte As String, TR_sNumero As String, sIdEdoCtaBanco As String, TR_sFecha As String
        Dim TR_sReferencia As String, TR_sConcepto As String, TR_nidCuentaCheques As String
        Dim associatedChar As Char = Chr(34)

        Dim vmCNX As New SqlConnection
        Dim vmSQLDA As New SqlDataAdapter
        Dim vmDS As New DataSet
        Dim vmSQLCMD As New SqlCommand

        Dim vlComando As String
        Dim vlResult As Integer
        Try
            vmCNX.ConnectionString = vcCnx

            TR_nidCuentaCheques = vprow(0)
            TR_sFecha = "Convert(VARCHAR(10), CAST('" & vprow(1) & "' AS DATE), 103)"
            'TR_sFecha = "'" & vprow(1) & "'"
            TR_sTipoMovto = "CAST('" & vprow(2) & "' as NVARCHAR(1))"
            TR_sNumero = "CAST('" & vprow(3) & "' as NVARCHAR(20))"
            TR_sReferencia = "CAST('" & Trim(vprow(4)) & "' as NVARCHAR(20))"
            TR_sConcepto = "CAST('" & Trim(vprow(5)) & "' as NVARCHAR(100))"

            Select Case vprow(2).ToUpper
                Case "I" : TR_sImporte = "CAST('" & vprow(6) & "' as DECIMAL(15,2))" 'Ingreso 
                Case "C", "E" : TR_sImporte = "CAST('" & vprow(7) & "' as DECIMAL(15,2))"  'Egreso y Cheque
            End Select

            vlComando = "BEGIN TRANSACTION" &
            " BEGIN TRY" &
            " declare @next int,@nextEdoCtaBanco INT, @estadoconciliacion int, @Numero int, @FechaInicial  date, @FechaFinal date" &
            " If Not exists (SELECT * FROM  dbo.EdoCtaBancos Where IdCuentaCheques = " & TR_nidCuentaCheques & ")" &
                " Begin" &
                " Select @FechaInicial = CAST(FechaSaldoInicial AS date) FROM CUENTASCHEQUES Where Id = " & TR_nidCuentaCheques &
                " Select @FechaFinal = EOMONTH(@FechaInicial)  " &
                " end" &
            " else" &
                " Begin" &
                " Select @FechaInicial = CAST(getdate() AS date) " &
                " end" &
            " SELECT Top 1 @next=MAX(ID), @estadoconciliacion=EstadoConciliacion, @FechaFinal = FechaFinal  FROM  dbo.EdoCtaBancos Where IdCuentaCheques =" & TR_nidCuentaCheques & " group by EstadoConciliacion,FechaFinal order by max(ID) desc " &
            " SET DATEFORMAT dmy " &
            " IF @estadoconciliacion = 0 " &
                " begin" &
                " Set @nextEdoCtaBanco = @next " &
                " end " &
            " else" &
                " begin" &
                " Select @FechaInicial = DATEADD(DAY, 1, @FechaFinal) " &
                " Select @FechaFinal = EOMONTH(@FechaInicial )  " &
                " select @nextEdoCtaBanco = next from counters where name =  'Id_EdoCtaBanco'" &
                " SELECT @numero=ISNULL(MAX(Numero),1)+1 FROM EdoCtaBancos Where IdCuentaCheques = " & TR_nidCuentaCheques &
                " INSERT INTO dbo.EdoCtaBancos (Id,RowVersion,Numero,IdCuentaCheques,Fecha,FechaInicial,FechaFinal,EstadoConciliacion,SaldoInicial,SaldoFinal,TimeStamp) " &
                " VALUES(@nextEdoCtaBanco,RAND(@nextEdoCtaBanco) * 12583646,@numero," & TR_nidCuentaCheques & ",@FechaInicial,@FechaInicial,Cast(getdate() as date),0,0,0,'')" &
                " Update counters set next = @nextEdoCtaBanco + 1  where name =  'Id_EdoCtaBanco'" &
                " end" &
            " select @next = next from counters where name =  'Id_MovtoEdoCtaBanco'" &
            " insert into movtosedoctabancos (Id,RowVersion,IdMovto,Numero,IdEdoCtaBanco,Fecha,TipoMovto,Referencia,Concepto,Total,EsConciliado) " &
            " Values (@next,RAND(@next) * 12583646, @next," & TR_sNumero & ",@nextEdoCtaBanco," & TR_sFecha & "," & TR_sTipoMovto & "," & TR_sReferencia & "," & TR_sConcepto & "," & TR_sImporte & "," & "0" & ")" &
            " Update counters set next = next + 1  where name =  'Id_MovtoEdoCtaBanco'" &
            " COMMIT TRANSACTION" &
            " SELECT 0 as CODIGO, '' AS MESSAGE " &
            " END TRY" &
            " BEGIN CATCH" &
            " ROLLBACK TRANSACTION" &
            " SELECT 1 as CODIGO, ERROR_MESSAGE() as MESSAGE" &
            " END CATCH"

            vmSQLCMD.CommandText = vlComando
            vmSQLCMD.CommandType = CommandType.Text
            vmSQLCMD.Connection = vmCNX
            vmSQLDA.SelectCommand = vmSQLCMD
            vmSQLDA.Fill(vmDS)
            'vpmsg = ""
            vpmsg = vlComando
            Return CBool(vmDS.Tables(0).Rows(0)("MESSAGE"))
            'Return 1
        Catch ex As Exception
            vpmsg = ex.Source & " - " & ex.Message
            Return False
        End Try
    End Function


    Private Function IniciaSDK(ByVal vpusuariocontpaqi As String, ByVal vpclavecontpaqi As String, ByVal vpempresa As String, ByRef vpmsg As String) As Boolean
        IniciaSDK = False
        Try
            lSdkSesion = New TSdkSesion()
            ListaEmpresa = New TSdkListaEmpresas()
            proveedor = New TSdkProveedor()
            cliente = New TSdkCliente()
            cuenta = New TSdkCuenta()
            cuentacheque = New TSdkCuentaCheque()
            ingresonodepositado = New TSdkIngresoNoDepositado()
            ingreso = New TSdkIngreso()
            deposito = New TSdkDeposito()
            egreso = New TSdkEgreso()
            cheque = New TSdkCheque()

            lSdkSesion.iniciaConexion()
            lSdkSesion.firmaUsuarioParams(vpusuariocontpaqi, vpclavecontpaqi) ' lSdkSesion.firmaUsuario()
            If lSdkSesion.ingresoUsuario = 1 Then
                lSdkSesion.abreEmpresa(vpempresa) ' viene de listaEmpresas
                proveedor.setSesion(lSdkSesion)
                cliente.setSesion(lSdkSesion)
                cuenta.setSesion(lSdkSesion)
                cuentacheque.setSesion(lSdkSesion)
                ingresonodepositado.setSesion(lSdkSesion)
                ingreso.setSesion(lSdkSesion)
                deposito.setSesion(lSdkSesion)
                egreso.setSesion(lSdkSesion)
                cheque.setSesion(lSdkSesion)
                asociacioncategoria.setSesion(lSdkSesion)

                vpmsg = ""
                IniciaSDK = True
            Else
                Kill_SDKCONTPAQNG()
                vpmsg = "Error en conexión a la empresa de contpaqi"
            End If
        Catch ex As Exception
            Kill_SDKCONTPAQNG()
            vpmsg = ex.Source & "-" & ex.Message
        End Try

    End Function

    Private Sub Kill_SDKCONTPAQNG()
        Try
            lSdkSesion.cierraEmpresa() 'Cerrar Empresa
            lSdkSesion.finalizaConexion()
            For Each P As Process In System.Diagnostics.Process.GetProcessesByName("SDKCONTPAQNG")
                P.CloseMainWindow()
                P.Kill()
            Next
        Catch
            For Each P As Process In System.Diagnostics.Process.GetProcessesByName("SDKCONTPAQNG")
                P.CloseMainWindow()
                P.Kill()
            Next
        End Try
    End Sub

    Public Function MsgError(ByVal vpMensaje As String) As DataTable
        Dim vlDt As New DataTable()
        vlDt.Columns.Add("MSG", GetType(System.String))
        vlDt.Rows.Add(vpMensaje)
        MsgError = vlDt
    End Function

    Public Function MsgArreglo(ByVal vpArray()() As String) As DataTable
        Dim vlDt As New DataTable()
        vlDt.Columns.Add("iddocumentode", GetType(System.String))       '(0)
        vlDt.Columns.Add("tipodocumento", GetType(System.String))       '(1)
        vlDt.Columns.Add("folio", GetType(System.String))               '(2)
        vlDt.Columns.Add("tipomovimiento", GetType(System.String))      '(16)
        vlDt.Columns.Add("result_trans", GetType(System.String))        '(17)
        For Each row In vpArray
            Dim vldr As DataRow = vlDt.NewRow
            vldr("iddocumentode") = row(0)
            vldr("tipodocumento") = row(1)
            vldr("folio") = row(2)
            vldr("tipomovimiento") = row(16)
            vldr("result_trans") = ""
            vlDt.Rows.Add(vldr)
        Next
        MsgArreglo = vlDt
    End Function

    Public Function MsgArreglo2(ByVal vpArray()() As String, ByRef vpMsg() As String) As DataTable
        Dim vlDt As New DataTable()

        vlDt.Columns.Add("idtransaccion", GetType(System.String))      '(14)
        vlDt.Columns.Add("result_trans", GetType(System.String))        '(15)
        vlDt.Columns.Add("msg_trans", GetType(System.String))           '(15)
        vlDt.Columns.Add("result_benpag", GetType(System.String))       '(15)
        vlDt.Columns.Add("msg_benpag", GetType(System.String))          '(15)
        For Each row In vpArray
            Dim vldr As DataRow = vlDt.NewRow
            vldr("idtransaccion") = row(17)
            vldr("result_trans") = vpMsg(0)
            vldr("msg_trans") = vpMsg(1)
            vldr("result_benpag") = vpMsg(2)
            vldr("msg_benpag") = vpMsg(3)
            vlDt.Rows.Add(vldr)
        Next
        MsgArreglo2 = vlDt
    End Function


    Public Function MsgArregloEdoCuenta(ByVal vpArray()() As String, ByRef vpMsg() As String) As DataTable
        Dim vlDt As New DataTable()
        vlDt.Columns.Add("iddocumentode", GetType(System.String))       '(0)
        vlDt.Columns.Add("tipodocumento", GetType(System.String))       '(1)
        vlDt.Columns.Add("folio", GetType(System.String))               '(2)
        vlDt.Columns.Add("tipomovimiento", GetType(System.String))      '(16)
        vlDt.Columns.Add("result_trans", GetType(System.String))        '(17)
        For Each row In vpArray
            Dim vldr As DataRow = vlDt.NewRow
            vldr("iddocumentode") = row(0)
            vldr("tipodocumento") = row(1)
            vldr("folio") = row(2)
            vldr("tipomovimiento") = row(3)
            vldr("result_trans") = ""
            vlDt.Rows.Add(vldr)
        Next
        MsgArregloEdoCuenta = vlDt
    End Function


End Class