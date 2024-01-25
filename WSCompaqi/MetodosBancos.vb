'Module MetodosBancos

'    Private Sub SDK_Inserta_Cheque(ByVal Prow As DataGridViewRow)
'        Try
'            With cheque
'                .iniciarInfo()
'                .TipoDocumento = Prow.Cells("idTipoDocumento").Value.ToString
'                .Folio = CLng(Prow.Cells("Folio").Value.ToString.Trim)
'                .Fecha = Prow.Cells("Fecha").Value.ToString
'                .FechaAplicacion = Prow.Cells("Fecha").Value.ToString
'                .CodigoPersona = Prow.Cells("CodigoPersona").Value.ToString.Trim
'                .BeneficiarioPagador = Prow.Cells("BeneficiarioPagador").Value.ToString.Trim
'                .IdCuentaCheques = Prow.Cells("IdCuentaCheques").Value.ToString
'                .Total = Prow.Cells("Total").Value
'                .Referencia = CLng(Prow.Cells("Folio").Value.ToString.Trim)
'                .Concepto = Prow.Cells("Concepto").Value.ToString.Trim
'                .Origen = 202
'                .CodigoMonedaTipoCambio = 2
'                .TipoCambio = 1
'                .EsImpreso = 1
'                .EsAsociado = 0
'                vpError = .crea()
'            End With
'            If vpError = 0 Then
'                sRowLog = "Error: " & cheque.getMensajeError & "; Registro: " & Trim(Prow.Cells(2).Value.ToString) & "-" & Trim(Prow.Cells(1).Value.ToString)
'                Me.txtLog.AppendText("...ERROR" & vbCrLf)
'                Me.txtLogError.AppendText(sRowLog & vbCrLf)
'            Else
'                sRowLog = "OK" & " - "
'                Me.txtLog.AppendText("...OK" & vbCrLf)
'            End If

'        Catch ex As Exception
'            txtLog.AppendText(" *** ERROR SDK_Inserta_Cheque ***" & vbCrLf)
'            txtLog.AppendText("Error: " & ex.Message.ToString & vbCrLf)
'        End Try
'    End Sub
'    Private Sub SDK_Cancela_Cheque(ByVal Prow As DataGridViewRow)
'        Dim dt As New DataSet
'        Dim nRows As Integer
'        Dim nConciliado As Integer
'        Dim nFolio As Integer
'        Dim nIdCuentaCheques As Integer
'        Dim nidTipoDocumento As Integer
'        Dim nEjercicio As Integer
'        Dim nPeriodo As Integer
'        Dim nTotal As Decimal
'        Dim nEscancelado As Integer
'        Try
'            nIdCuentaCheques = CInt(Prow.Cells("IdCuentaCheques").Value)
'            nidTipoDocumento = CInt(Prow.Cells("idTipoDocumento").Value)
'            nEjercicio = Year(Prow.Cells("Fecha").Value.ToString)
'            nPeriodo = Month(Prow.Cells("Fecha").Value.ToString)
'            nTotal = Prow.Cells("Total").Value
'            nFolio = CInt(Prow.Cells("Folio").Value.ToString)
'            Select Case nIdCuentaCheques
'                Case 8, 9, 19, 25, 29, 30, 31 'BANORTE 8-192994129, 9-645823239,19-493116981,25-617297882,29-493116954,30-493116972,31-495743804
'                    nConciliado = 1 ' BANORTE CONCILIA, NO CANCELA
'                    strCommand = "declare @RowsAfected as integer BEGIN TRANSACTION" &
'                         " BEGIN TRY" &
'                         " Update  " & txtBase.Text & "..Cheques set EsConciliado = " & nConciliado.ToString & " , Referencia= '" & "*CAN*" & "' ,Concepto = '" & "-CANCELAR-" & "'" & " + Concepto  " &
'                         " where EsConciliado = 0 and IdCuentaCheques= " & nIdCuentaCheques.ToString & " and " &
'                         "       IdDocumentoDe = 37 and Tipodocumento = " & nidTipoDocumento.ToString & " and " &
'                         "       folio = " & nFolio.ToString &
'                         " set @RowsAfected = @@ROWCOUNT " &
'                         " COMMIT TRANSACTION" &
'                         " Select 0 as Ok_Exec, @RowsAfected as RowsAfected" &
'                         " END TRY" &
'                         " BEGIN CATCH" &
'                         " ROLLBACK TRANSACTION" &
'                         " Select 1 as Ok_Exec, 0 as RowsAfected" &
'                         " END CATCH"

'                    If xExec(strCommand, sRowLog) Then
'                        Me.txtLog.AppendText("...OK Cancelación" & vbCrLf)
'                    Else
'                        Me.txtLog.AppendText("Error: " & sRowLog & vbCrLf)
'                        Me.txtLogError.AppendText(sgError & vbCrLf)
'                    End If
'                Case 1, 2, 5, 6, 7, 14, 18, 28 'BANCOMER 1-184130758,2-184120477,5-184540751, 6-184122631,7-184541189,14-184540999,18-184129466,28-106861482
'                    nEscancelado = 1 ' BANCOMER CANCELA, NO CONCILIA
'                    strCommand = "declare @RowsAfected as integer, @Identity as Integer, @Ejercicio as integer, @Periodo as integer  " &
'                        " BEGIN TRANSACTION" &
'                        " BEGIN TRY" &
'                        " select top 1 @Identity= Id, @Ejercicio = Ejercicio,@Periodo = Periodo  from " & txtBase.Text & "..Cheques " &
'                        " where EsConciliado = 0 And IdCuentaCheques= " & nIdCuentaCheques.ToString & " And " &
'                        "      IdDocumentoDe = 37 And Tipodocumento = " & nidTipoDocumento.ToString & " And " &
'                        "      folio = " & nFolio.ToString &
'                        " IF @Identity Is Not null " &
'                            " begin " &
'                                " Update  " & txtBase.Text & "..Cheques Set EsCancelado = " & nEscancelado.ToString & " , " &
'                                        " Referencia = '" & "*CAN SDK*" & "' ,Concepto = '" & "-CANCELAR-" & "'" & " + Concepto  " &
'                                " where Id = @Identity " &
'                            " end " &
'                         " set @RowsAfected = @@ROWCOUNT " &
'                         " COMMIT TRANSACTION" &
'                         " Select 0 as Ok_Exec, @RowsAfected as RowsAfected, @Ejercicio as Ejercicio, @Periodo as Periodo " &
'                         " END TRY" &
'                         " BEGIN CATCH" &
'                         " ROLLBACK TRANSACTION" &
'                          " Select 1 as Ok_Exec, 0 as RowsAfected, 0 as Ejercicio, 0 as Periodo" &
'                         " END CATCH"

'                    If xExec(strCommand, sRowLog, dt) Then
'                        nRows = dt.Tables(0).Rows(0)("RowsAfected").ToString()
'                        nEjercicio = dt.Tables(0).Rows(0)("Ejercicio").ToString()
'                        nPeriodo = dt.Tables(0).Rows(0)("Periodo").ToString()
'                        If nRows = 1 Then
'                            Update_SaldosCtasCheques(nIdCuentaCheques, nEjercicio, nPeriodo, nTotal) ' actualiza saldos (simulando cancelación)
'                            Me.txtLog.AppendText("...OK Cancelación" & vbCrLf)
'                        Else
'                            'se debe mostrar un aviso de error
'                            Me.txtLog.AppendText("...Error, no se encontró el cheque: " & nFolio.ToString & " Idcuenta: " & nIdCuentaCheques.ToString & " Total: " & nTotal.ToString & vbCrLf)
'                            Me.txtLogError.AppendText("Error al cancelar el cheque: " & nFolio.ToString & " Idcuenta: " & nIdCuentaCheques.ToString & " Total: " & nTotal.ToString & " Folio no localizado" & vbCrLf)
'                            sRowLog = "Aviso de Mov. Cancelación: No se encontró el Número de cheque en las tabla Cheques"
'                        End If

'                    Else
'                        Me.txtLog.AppendText("Error: " & sRowLog & vbCrLf)
'                        Me.txtLogError.AppendText(sgError & vbCrLf)
'                    End If

'                Case Else
'                    'no ejecuta acción alguna
'            End Select

'            'If xExec(strCommand, sRowLog, dt) Then
'            '    nIdFolio = dt.Tables(0).Rows(0)("Id").ToString
'            'Else
'            '    nIdFolio = -1
'            'End If

'            'Dim lResult As Integer
'            'nIdFolio = 123641
'            'lResult = cheque.buscaPorId(nIdFolio)

'            'If lResult = 0 Then
'            '    MessageBox.Show("No cancelado")
'            '    sRowLog = "Error: " & cheque.getMensajeError
'            'Else
'            '    cheque.Concepto = " *Cancelado * " & cheque.Concepto
'            '    vpError = cheque.modifica()
'            '    sRowLog = "Error: " & cheque.getMensajeError

'            '    lResult = cheque.cancela(nIdCuentaCheques, 1)
'            '    sRowLog = "Error: " & cheque.getMensajeError
'            '    ' lResult = cheque.devuelve(nIdCuentaCheques, 0)
'            '    ' sRowLog = "Error: " & cheque.getMensajeError
'            'End If


'            'row.Cells("Concepto").Value = "Cancelación: " & row.Cells("Concepto").Value
'            ' row.Cells("idTipoDocumento").Value = 42 ' se cambia el tipodocumento cheque, por Ingresos
'            ' SDK_Inserta_Ingreso(row)
'            'strCommand = "BEGIN TRANSACTION" &
'            '    " BEGIN TRY" &
'            '    " declare @next int" &
'            '    " Update Cheques set EsCancelado= 1,EsImpreso = 1  where EsConciliado = 0 and IdCuentaCheques= " & row.Cells("IdCuentaCheques").Value.ToString & " and IdDocumentoDe = 37 and Tipodocumento = " & row.Cells("idTipoDocumento").Value.ToString & " and folio = " & CInt(row.Cells("Folio").Value.ToString) &
'            '    " Delete from asocDoctoCategorias  where IdCuentaCheque = " & row.Cells("IdCuentaCheques").Value.ToString & " and IdDocumentoDe = 37 and Tipodocumento = " & row.Cells("idTipoDocumento").Value.ToString & " and folio = " & CInt(row.Cells("Folio").Value.ToString) &
'            '    " update SALDOSCATEGORIAS set " & colSaldoCategoria & " = " & colSaldoCategoria & " - " & row.Cells(9).Value & "  where IdCategoria = " & row.Cells(14).Value.ToString & " AND ejercicio = " & nEjercicio &
'            '    " COMMIT TRANSACTION" &
'            '    " Select 0" &
'            '    " END TRY" &
'            '    " BEGIN CATCH" &
'            '    " ROLLBACK TRANSACTION" &
'            '    " Select 1" &
'            '    " END CATCH"






'            'Dim lResult As Integer
'            'Dim lfolio As Integer = CInt("58596")

'            'lResult = cheque.buscaPorId(lfolio)

'            'If lResult = 0 Then
'            '    MessageBox.Show("No cancelado")
'            '    sRowLog = "Error: " & cheque.getMensajeError
'            'Else
'            '    lResult = cheque.cancela(184120477, 0)
'            '    sRowLog = "Error: " & cheque.getMensajeError

'            '    MessageBox.Show("Cheque cancelado")
'            'End If


'            'vpError = Nothing
'            'vpError = cheque.buscaPorId(79975)
'            'If vpError = True Then
'            '    With cheque
'            '        .Concepto = " *Cancelado * " & .Concepto
'            '        vpError = .modifica()
'            '        sRowLog = "Error: " & cheque.getMensajeError
'            '    End With
'            '    vpError = cheque.cancela(8, 0)
'            '    sRowLog = "Error: " & cheque.getMensajeError
'            'End If

'        Catch ex As Exception
'            Me.txtLog.AppendText(" *** ERROR SDK_Cancela_Cheque ***" & vbCrLf)
'            Me.txtLog.AppendText("Error: " & ex.Message.ToString & vbCrLf)
'        End Try
'    End Sub
'    Private Sub SDK_Inserta_Ingreso(ByVal Prow As DataGridViewRow)
'        Try
'            With ingreso
'                .iniciarInfo()
'                .TipoDocumento = Prow.Cells("idTipoDocumento").Value.ToString
'                '.Folio = Prow.Cells("Folio").Value.ToString
'                .Fecha = Prow.Cells("Fecha").Value.ToString
'                .FechaAplicacion = Prow.Cells("Fecha").Value.ToString
'                .CodigoPersona = Prow.Cells("CodigoPersona").Value.ToString
'                .BeneficiarioPagador = Prow.Cells("BeneficiarioPagador").Value.ToString
'                .IdCuentaCheques = Prow.Cells("IdCuentaCheques").Value.ToString
'                .Total = Prow.Cells("Total").Value
'                .Referencia = Trim(Prow.Cells(10).Value.ToString)
'                .Concepto = "IdTran: " & Prow.Cells("Folio").Value.ToString & " - " & Prow.Cells("Concepto").Value.ToString
'                .Origen = 202
'                .CodigoMonedaTipoCambio = 2
'                .TipoCambio = 1

'                vpError = .crea()
'            End With

'            If vpError = 0 Then
'                sRowLog = "Error: " & ingreso.getMensajeError & "; Registro: " & Trim(Prow.Cells(2).Value.ToString) & "-" & Trim(Prow.Cells(1).Value.ToString)
'                Update_BD_Movimientos(True, sRowLog, Prow) 'Actualiza Columnas de la tabla movimientos
'                Me.txtLog.AppendText("...ERROR" & vbCrLf)
'                Me.txtLogError.AppendText(sRowLog & vbCrLf)
'            Else
'                'nDocumentoDe = 34
'                'strCommand = "BEGIN TRANSACTION" &
'                '" BEGIN TRY" &
'                '" declare @next int" &
'                '" select @next = next from counters where name =  'Id_AsocDoctoCategoria'" &
'                '" insert into asocDoctoCategorias (Id,RowVersion,IdDocumento,IdDocumentoDe,TipoDocumento,IdCuentaCheque,Folio,IdCategoria,IdSubCategoria,Porcentaje,	Total) " &
'                '                        "Values (@next," & "RAND(" & ingreso.Id & ") * 12583646," & ingreso.Id & "," & nDocumentoDe & ",CAST(" & ingreso.TipoDocumento & " AS NVARCHAR(30))," & ingreso.IdCuentaCheques & "," &
'                '                        ingreso.Folio & "," & row.Cells(12).Value.ToString & "," & row.Cells(14).Value.ToString & "," & "100" & "," & ingreso.Total & ")" &
'                '" Update counters set next = next + 1  where name =  'Id_AsocDoctoCategoria'" &
'                '" If NOT exists (select * from dbo.SaldosCategorias where IdCategoria = " & row.Cells(14).Value.ToString & " AND ejercicio = " & nEjercicio & ")" &
'                '" BEGIN" &
'                '" select @next = next from counters where name =  'Id_SaldoCategoria'" &
'                '" insert into dbo.SaldosCategorias (Id,RowVersion,IdCategoria,Ejercicio,SaldoInicial,Saldo1,Saldo2,Saldo3,Saldo4,Saldo5,Saldo6,Saldo7,Saldo8,Saldo9,Saldo10,Saldo11,Saldo12,Saldo13,Saldo14)" &
'                '" Values (@next,RAND(" & ingreso.Id & ") * 12583646," & row.Cells(14).Value.ToString & "," & nEjercicio & "," & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)" &
'                '" Update counters set next = next + 1  where name =  'Id_SaldoCategoria'" &
'                '" END" &
'                '" update SALDOSCATEGORIAS set " & colSaldoCategoria & " = " & colSaldoCategoria & " + " & ingreso.Total & "  where IdCategoria = " & row.Cells(14).Value.ToString & " AND ejercicio = " & nEjercicio &
'                '" COMMIT TRANSACTION" &
'                '" Select 0" &
'                '" END TRY" &
'                '" BEGIN CATCH" &
'                '" ROLLBACK TRANSACTION" &
'                '" Select 1" &
'                '" END CATCH"
'                'xExec(strCommand)
'                sRowLog = "OK"
'                Update_BD_Movimientos(True, sRowLog, Prow) 'Actualiza Columnas de la tabla movimientos
'                Me.txtLog.AppendText("...OK" & vbCrLf)
'            End If


'        Catch ex As Exception
'            Me.txtLog.AppendText(" *** ERROR SDK_Inserta_Ingreso ***" & vbCrLf)
'            Me.txtLog.AppendText("Error: " & ex.Message.ToString & vbCrLf)
'        End Try
'    End Sub
'    Private Sub SDK_Inserta_Egreso(ByVal Prow As DataGridViewRow)
'        Try
'            With egreso
'                .iniciarInfo()
'                .TipoDocumento = Prow.Cells("idTipoDocumento").Value.ToString
'                '.Folio = Prow.Cells("Folio").Value.ToString
'                .Fecha = Prow.Cells("Fecha").Value.ToString
'                .FechaAplicacion = Prow.Cells("Fecha").Value.ToString
'                .CodigoPersona = Prow.Cells("CodigoPersona").Value.ToString
'                .BeneficiarioPagador = Prow.Cells("BeneficiarioPagador").Value.ToString
'                .IdCuentaCheques = Prow.Cells("IdCuentaCheques").Value.ToString
'                .Total = Prow.Cells("Total").Value
'                .Referencia = Prow.Cells(10).Value.ToString.Trim
'                .Concepto = "IdTran: " & Prow.Cells("Folio").Value.ToString & " - " & Prow.Cells("Concepto").Value.ToString
'                .Origen = 202
'                .CodigoMonedaTipoCambio = 2
'                .TipoCambio = 1
'                vpError = .crea()
'            End With

'            If vpError = 0 Then
'                sRowLog = "Error: " & egreso.getMensajeError & "; Registro: " & Trim(Prow.Cells(2).Value.ToString) & "-" & Trim(Prow.Cells(1).Value.ToString)
'                txtLog.AppendText("...ERROR" & vbCrLf)
'                txtLogError.AppendText(sRowLog & vbCrLf)
'            Else

'                sRowLog = "OK"
'                txtLog.AppendText("...OK" & vbCrLf)
'            End If

'        Catch ex As Exception
'            txtLog.AppendText(" *** ERROR SDK_Inserta_Egreso ***" & vbCrLf)
'            txtLog.AppendText("Error: " & ex.Message.ToString & vbCrLf)
'        End Try
'    End Sub

'End Module



'<WebMethod()>
'<ScriptMethod(ResponseFormat:=ResponseFormat.Json, UseHttpGet:=False, XmlSerializeString:=False)>
'Public Function Empleados(ByVal vpServidor As String, ByVal vpBase As String, ByVal vpUsuario As String, ByVal vpClave As String,
'                                   ByVal vpIdTipoPeriodo As String, ByVal vpIdEmpleado As String, ByVal vpCodigoEmpleado As String, ByVal vpNombre As String) As String
'    Dim vmCNX As New SqlConnection
'    Dim vmSQLDA As New SqlDataAdapter
'    Dim vmDS As New DataSet
'    Dim vmSQLCMD As New SqlCommand
'    Dim vlComando As String
'    Try
'        Dim vlCnx As String
'        vlCnx = "Server=" & vpServidor & ";Database=" & vpBase & ";Uid=" & vpUsuario & ";Password=" & vpClave & ";Trusted_Connection=False;"
'        vmCNX.ConnectionString = vlCnx
'        vlComando =
'                "select Emps.idempleado, emps.idtipoperiodo, Periodo.nombretipoperiodo, Periodo.diasdelperiodo, Periodo.diasdepago, Emps.codigoempleado, Emps.apellidopaterno, " &
'                "Emps.apellidomaterno, Emps.nombre,Emps.nombrelargo, Emps.estadoempleado, Emps.telefono, Emps.CorreoElectronico, " &
'                "Emps.numerosegurosocial, CONVERT(VARCHAR, Emps.fechanacimiento,126) as fechanacimiento, " &
'                "COALESCE(RTRIM(Emps.curpi), '')+COALESCE(CONVERT(VARCHAR, Emps.fechanacimiento,12), '')+COALESCE(RTRIM(Emps.curpf), '') as CURP, " &
'                "COALESCE(RTRIM(Emps.rfc), '')+COALESCE(CONVERT(VARCHAR, Emps.fechanacimiento,12), '')+COALESCE(RTRIM(Emps.homoclave), '') as RFC, " &
'                "CONVERT(VARCHAR, Emps.fechaalta,126) as fechaalta, CONVERT(VARCHAR, Emps.fechareingreso,126) as fechareingreso, " &
'                "CONVERT(VARCHAR, Emps.fechabaja,126) as fechabaja, Emps.causabaja, EmpXPer.cidperiodo as cidperiodo_aPagar, EmpXPer.cdiastrabajados,  " &
'                "EmpXPer.cdiaspagados, Bancos.Descripcion, Emps.cuentapagoelectronico, Emps.sucursalpagoelectronico, Emps.bancopagoelectronico, " &
'                "EmpXPer.sueldodiario, EmpXPer.sueldointegrado,  MovsXPer.importetotal " &
'                "from " &
'                "nom10001 Emps  " &
'                "left join   " &
'                "( " &
'                 "select pp.idperiodo, pp.idtipoperiodo, pp.numeroperiodo, t.nombretipoperiodo, t.diasdelperiodo, t.diasdepago, pp.ejercicio, pp.mes, pp.fechainicio, pp.fechafin  " &
'                 "from nom10023 as T inner join  " &
'                    "(select p.idperiodo, p.idtipoperiodo,  " &
'                   "p.numeroperiodo, p.ejercicio,  " &
'                   "p.mes, p.fechainicio, p.fechafin from nom10002 as P  " &
'                   "inner join  " &
'                            "(select idtipoperiodo, ejercicio, min(idperiodo) as idperiodo  " &
'                     "from Nom10002 where afectado=0 group by idtipoperiodo, ejercicio) as M  " &
'                   "on m.idtipoperiodo=p.idtipoperiodo and m.idperiodo=P.idperiodo " &
'                 ")as pp  " &
'                 "on pp.idtipoperiodo = t.idtipoperiodo and pp.ejercicio=t.ejercicio  " &
'                ") DelPeriodo  on emps.idtipoperiodo=DelPeriodo.idtipoperiodo and Emps.estadoempleado='A'  " &
'                "left join  " &
'                "nom10034 EmpXPer on EmpXPer.cidperiodo = DelPeriodo.idperiodo and EmpXPer.idempleado = Emps.idempleado and Emps.estadoempleado='A'  " &
'                "left join  " &
'                "nom10008 MovsXPer on EmpXPer.cidperiodo = MovsXPer.idperiodo and EmpXPer.idempleado = MovsXPer.idempleado and MovsXPer.idconcepto=1 " &
'                "left join [nomGenerales].[dbo].[SATCatBancos] as Bancos on emps.bancopagoelectronico = bancos.ClaveBanco " &
'                "left join nom10023 as Periodo on Periodo.idtipoperiodo = emps.idtipoperiodo "

'        Select Case True
'            Case vpIdTipoPeriodo <> ""
'                vlComando = vlComando & " where Emps.idtipoperiodo = " & vpIdTipoPeriodo
'            Case vpIdEmpleado <> ""
'                vlComando = vlComando & " where Emps.idempleado = " & vpIdEmpleado
'            Case vpCodigoEmpleado <> ""
'                vlComando = vlComando & " where Emps.codigoempleado = '" & vpCodigoEmpleado & "'"
'            Case vpNombre <> ""
'                vlComando = vlComando & " where Emps.nombrelargo like '%" & vpNombre & "%'"
'        End Select
'        vmSQLCMD.CommandText = vlComando
'        vmSQLCMD.CommandType = CommandType.Text
'        vmSQLCMD.Connection = vmCNX
'        vmSQLDA.SelectCommand = vmSQLCMD
'        vmSQLDA.Fill(vmDS)
'        Return Newtonsoft.Json.JsonConvert.SerializeObject(vmDS.Tables(0))
'    Catch ex As Exception
'        Return Newtonsoft.Json.JsonConvert.SerializeObject(MsgError(ex.Source & "-" & ex.Message))
'    End Try
'End Function