## `DetalleEconomicoProyectoObraFrm.frm`

### `CargarDetalleCostes`
Reemplazar por
```vb
Private Sub CargarDetalleCostes()

Dim MiRc As Recordset
Dim MiRc1 As Recordset
Dim MiRc2 As Recordset
Dim MiRc3 As Recordset
Dim Fecha1 As Date
Dim Fecha2 As Date
Dim FiltroFechasTxt As String
Dim FechaFactura As Date
Dim ImporteAsignado As Double
Dim Posicion As Double
Dim Auxiliar As Double
Dim EstaCerrada As Boolean

'Borrando tabla auxiliar de carga
Base.Execute ("Delete * from Auxiliar_Carga Where Usuario_Id=" & IdUsuarioGeneral)
Me.ProveedoresCombo.Clear
ReDim IdProyectoCosteProveedor(0)

'Etableciendo filtro de fechas
If Me.TipoListadoCombo.ListIndex = 0 Then
    Fecha1 = 0
    If Me.FechaLimiteListadoTxt.Text <> "__/__/____" Then
        Fecha1 = Me.FechaLimiteListadoTxt.Text
        Fecha1 = "01/" & Format(Fecha1, "mm/yyyy")
        Fecha1 = DateAdd("m", 1, Fecha1)
        Fecha1 = DateAdd("d", -1, Fecha1)
        FiltroFechasTxt = " and Fecha<=#" & Format(Fecha1, "m/d/yy") & "#"
    Else
        FiltroFechasTxt = ""
    End If
Else
    FiltroFechasTxt = ""
    
    If Me.FechaLimiteListadoTxt.Text <> "__/__/____" Then
        Fecha1 = Me.FechaLimiteListadoTxt.Text
        Fecha1 = "01/" & Format(Fecha1, "mm/yyyy")
        FiltroFechasTxt = " AND Fecha >= #" & Format(Fecha1, "m/d/yy") & "#"
    End If
    
    If Me.FechaLimite1Txt.Text <> "__/__/____" Then
        Fecha2 = Me.FechaLimite1Txt.Text
        Fecha2 = "01/" & Format(Fecha2, "mm/yyyy")
        Fecha2 = DateAdd("m", 1, Fecha2)
        Fecha2 = DateAdd("d", -1, Fecha2)
        FiltroFechasTxt = FiltroFechasTxt & " AND Fecha <= #" & Format(Fecha2, "m/d/yy") & "#"
    End If
End If

'Preparando grid
Me.GridCostes.Rows = 1
Me.GridCostes.Rows = 2

Me.GridCostes.Redraw = False

Me.TipoCosteCombo.ListIndex = 0

Set MiRc1 = Base.OpenRecordset("Select *" _
    & " from Auxiliar_Carga" _
    & " where Id=-1")

Set MiRc = Base.OpenRecordset("Select Proyecto_Coste_Proveedor.Id as IdCosteProyecto, Cabecera_Documentos_Compra.Id,Cabecera_Documentos_Compra.Fecha,Cabecera_Documentos_Compra.Serie_Id," _
    & "Cabecera_Documentos_Compra.Numero,Cabecera_Documentos_Compra.Entidad_Id,Cabecera_Documentos_Compra.Factura_Proveedor_Rapido,Cabecera_Documentos_Compra.Ruta_Archivo," _
    & "Cabecera_Documentos_Compra.Referencia,Proyecto_Coste_Proveedor.Coste_Asignado,Proyecto_Coste_Proveedor.Autorizado_Tecnico_Id,Proyecto_Coste_Proveedor.Autorizado_Responsable_Tecnico_Id," _
    & "Cabecera_Documentos_Compra.Autorizado_Administracion_Id,Cabecera_Documentos_Compra.Tipo_Gasto,Cabecera_Documentos_Compra.Base_Documento" _
    & " from Cabecera_Documentos_Compra inner join Proyecto_Coste_Proveedor" _
    & " on Cabecera_Documentos_Compra.Id=Proyecto_Coste_Proveedor.Cabecera_Id" _
    & " where Proyecto_Coste_Proveedor.Proyecto_Id=" & IdProyecto _
    & " order by Cabecera_Documentos_Compra.Fecha asc")

If MiRc.EOF = False Then

    MiRc.MoveFirst
    While Not MiRc.EOF
        
        'Comprobando si hay desglose temporal,
            'Si no lo hay se cogen los datos de la factura
            'Si lo hay y no está en fecha se sale del proces
            'Si lo hay y está en fecha se coge sólo el importe correspondiente
        FechaFactura = 0
        ImporteAsignado = 0
        Set MiRc2 = Base.OpenRecordset("Select sum(Coste_Asignado) as coste" _
            & " from Desglose_Temporal_Asignacion_Coste" _
            & " where Cabecera_Factura_Id=" & MiRc!Id _
            & " and Proyecto_Id=" & IdProyecto _
            & " and Tipo_Item=0" _
            & FiltroFechasTxt)
        MiRc2.MoveFirst
        If IsNull(MiRc2!Coste) = True Then  'No existe desglose
            FechaFactura = MiRc!Fecha
            ImporteAsignado = MiRc!Coste_Asignado
            If Me.TipoListadoCombo.ListIndex = 0 Then   'A origen
                If FechaFactura > Fecha1 Then
                    MiRc2.Close
                    GoTo SiguienteFactura
                End If
            Else    'Entre fechas
                If FechaFactura < Fecha1 Or FechaFactura > Fecha2 Then
                    MiRc2.Close
                    GoTo SiguienteFactura
                End If
            End If
        Else    'Existe desglose
            'Hay que crear una línea por cada mes de asignación
            MiRc2.Close
            Set MiRc2 = Base.OpenRecordset("Select Fecha,sum(Coste_Asignado) as Coste" _
                & " from Desglose_Temporal_Asignacion_Coste" _
                & " where Cabecera_Factura_Id=" & MiRc!Id _
                & " and Coste_Asignado<>0" _
                & " and Proyecto_Id=" & IdProyecto _
                & " and Tipo_Item=0" _
                & FiltroFechasTxt _
                & " group by Fecha")
            If MiRc2.EOF = False Then
                MiRc2.MoveFirst
                While Not MiRc2.EOF
                    MiRc1.AddNew
                    MiRc1!Usuario_Id = IdUsuarioGeneral
                    MiRc1!Col_0 = MiRc!Id
                    MiRc1!Fecha = MiRc2!Fecha
                    MiRc1!Col_1 = MiRc2!Fecha
                    Select Case MiRc!Tipo_Gasto
                        Case 1  'Material
                            MiRc1!Col_2 = "1"
                        Case 2  'Mano de obra
                            MiRc1!Col_2 = "2"
                        Case 3, 4, 5, 6 'Otros
                            MiRc1!Col_2 = "3"
                        Case Else
                            MiRc1!Col_2 = "4"
                    End Select
                    MiRc1!Col_3 = CapturaNumeroDocumento(MiRc!Serie_Id, MiRc!Numero)
                    If IsNull(MiRc!Factura_Proveedor_Rapido) = False And MiRc!Factura_Proveedor_Rapido <> "" Then
                        MiRc1!Col_4 = MiRc!Factura_Proveedor_Rapido
                    Else
                        If IsNull(MiRc!Referencia) = False And MiRc!Referencia <> "" Then
                            MiRc1!Col_4 = MiRc!Referencia
                        End If
                    End If
                    MiRc1!Col_5 = NombreFicha(MiRc!Entidad_Id)
                    If Trim(MiRc1!Col_5) = "" Then MiRc1!Col_5 = " "
                    MiRc1!Col_6 = Format(MiRc2!Coste, FormatoDecimalStandar)
                    If IsNull(MiRc!Ruta_Archivo) = True Or Trim(MiRc!Ruta_Archivo) = "" Then
                        MiRc1!Col_7 = "N"
                    Else
                        MiRc1!Col_7 = "S"
                    End If
                    If IsNull(MiRc!Autorizado_Tecnico_Id) = True Or MiRc!Autorizado_Tecnico_Id = 0 Then
                        MiRc1!Col_8 = "N"
                    Else
                        MiRc1!Col_8 = "S"
                    End If
                    If IsNull(MiRc!Autorizado_Responsable_Tecnico_Id) = True Or MiRc!Autorizado_Responsable_Tecnico_Id = 0 Then
                        MiRc1!Col_9 = "N"
                    Else
                        MiRc1!Col_9 = "S"
                    End If
                    If IsNull(MiRc!Autorizado_Administracion_Id) = True Or MiRc!Autorizado_Administracion_Id = 0 Then
                        MiRc1!Col_10 = "N"
                    Else
                        MiRc1!Col_10 = "S"
                    End If
                    MiRc1!Col_11 = MiRc!Fecha
                    MiRc1!Col_12 = Format(MiRc!Base_Documento, FormatoDecimalStandar)
                    MiRc1!Col_13 = MiRc!IdCosteProyecto
                    Auxiliar = 0
                    If IsNull(MiRc!Coste_Asignado) = False Then Auxiliar = MiRc!Coste_Asignado
                    MiRc1!Col_14 = Format(Auxiliar - CargarAsignacionPorPartidas(MiRc!Id, IdProyecto), FormatoDecimalStandar)
                    
                    MiRc1.Update
                    'Siguiente asignación temporal
                    MiRc2.MoveNext
                Wend
            End If
            MiRc2.Close
            GoTo SiguienteFactura
        End If
        MiRc2.Close
        
        'Grabando
        MiRc1.AddNew
        
        MiRc1!Usuario_Id = IdUsuarioGeneral
        MiRc1!Col_0 = MiRc!Id
        MiRc1!Fecha = FechaFactura
        MiRc1!Col_1 = FechaFactura
        Select Case MiRc!Tipo_Gasto
            Case 1  'Material
                MiRc1!Col_2 = "1"
            Case 2  'Mano de obra
                MiRc1!Col_2 = "2"
            Case 3, 4, 5, 6 'Otros
                MiRc1!Col_2 = "3"
            Case Else
                MiRc1!Col_2 = "4"
        End Select
        MiRc1!Col_3 = CapturaNumeroDocumento(MiRc!Serie_Id, MiRc!Numero)
        If IsNull(MiRc!Factura_Proveedor_Rapido) = False And MiRc!Factura_Proveedor_Rapido <> "" Then
            MiRc1!Col_4 = MiRc!Factura_Proveedor_Rapido
        Else
            If IsNull(MiRc!Referencia) = False And MiRc!Referencia <> "" Then
                MiRc1!Col_4 = MiRc!Referencia
            End If
        End If
        MiRc1!Col_5 = NombreFicha(MiRc!Entidad_Id)
        If MiRc1!Col_5 = "" Then MiRc1!Col_5 = " "
        MiRc1!Col_6 = Format(ImporteAsignado, FormatoDecimalStandar)
        If IsNull(MiRc!Ruta_Archivo) = True Or Trim(MiRc!Ruta_Archivo) = "" Then
            MiRc1!Col_7 = "N"
        Else
            MiRc1!Col_7 = "S"
        End If
        If IsNull(MiRc!Autorizado_Tecnico_Id) = True Or MiRc!Autorizado_Tecnico_Id = 0 Then
            MiRc1!Col_8 = "N"
        Else
            MiRc1!Col_8 = "S"
        End If
        If IsNull(MiRc!Autorizado_Responsable_Tecnico_Id) = True Or MiRc!Autorizado_Responsable_Tecnico_Id = 0 Then
            MiRc1!Col_9 = "N"
        Else
            MiRc1!Col_9 = "S"
        End If
        If IsNull(MiRc!Autorizado_Administracion_Id) = True Or MiRc!Autorizado_Administracion_Id = 0 Then
            MiRc1!Col_10 = "N"
        Else
            MiRc1!Col_10 = "S"
        End If
        MiRc1!Col_11 = MiRc!Fecha
        MiRc1!Col_12 = Format(MiRc!Base_Documento, FormatoDecimalStandar)
        MiRc1!Col_13 = MiRc!IdCosteProyecto
        Auxiliar = 0
        If IsNull(MiRc!Coste_Asignado) = False Then Auxiliar = MiRc!Coste_Asignado
        MiRc1!Col_14 = Format(Auxiliar - CargarAsignacionPorPartidas(MiRc!Id, IdProyecto), FormatoDecimalStandar)
        
        MiRc1.Update

SiguienteFactura:
        'Siguiente Factura
        MiRc.MoveNext

    Wend

End If

MiRc.Close

'Cargando datos previsiones (positivas - Fecha de Alta)
FiltroFechasTxt = ""
If Me.TipoListadoCombo.ListIndex = 0 Then
    FiltroFechasTxt = " and Fecha_Alta<=#" & Format(Fecha1, "m/d/yy") & "#"
Else
    FiltroFechasTxt = " and Fecha_Alta>=#" & Format(Fecha1, "m/d/yy") & "# and Fecha_Alta<=#" & Format(Fecha2, "m/d/yy") & "#"
End If
Set MiRc = Base.OpenRecordset("Select *" _
    & " from Ajustes_Cuenta_Resultado" _
    & " where Proyecto_Id=" & IdProyecto _
    & " and Fecha_Alta<>0" _
    & FiltroFechasTxt)
If MiRc.EOF = False Then
    MiRc.MoveFirst
    
    While Not MiRc.EOF
    
        MiRc1.AddNew
        MiRc1!Usuario_Id = IdUsuarioGeneral
        MiRc1!Col_0 = MiRc!Id
        MiRc1!Fecha = MiRc!Fecha_Alta
        MiRc1!Col_1 = MiRc!Fecha_Alta
        MiRc1!Col_2 = "P"
        MiRc1!Col_3 = " "
        MiRc1!Col_4 = MiRc!concepto
        If IsNull(MiRc!Ficha_ID) = False Then MiRc1!Col_5 = NombreFicha(MiRc!Ficha_ID)
        If Trim(MiRc1!Col_5) = "" Then MiRc1!Col_5 = " "
        MiRc1!Col_6 = Format(MiRc!importe, FormatoDecimalStandar)
        MiRc1!Col_7 = "N"
        MiRc1!Col_8 = "N"
        MiRc1!Col_9 = "N"
        MiRc1!Col_10 = "N"
        MiRc1!Col_11 = " "
        MiRc1!Col_12 = " "
        MiRc1!Col_13 = " "
        MiRc1!Col_14 = " "
        
        MiRc1.Update
        
        MiRc.MoveNext
    
    Wend
    
End If
MiRc.Close

'Cargando datos previsiones (negativas)
FiltroFechasTxt = ""
If Me.TipoListadoCombo.ListIndex = 0 Then
    FiltroFechasTxt = " and Fecha_Baja<=#" & Format(Fecha1, "m/d/yy") & "#"
Else
    FiltroFechasTxt = " and Fecha_Baja>=#" & Format(Fecha1, "m/d/yy") & "# and Fecha_Baja<=#" & Format(Fecha2, "m/d/yy") & "#"
End If
Set MiRc = Base.OpenRecordset("Select *" _
    & " from Ajustes_Cuenta_Resultado" _
    & " where Proyecto_Id=" & IdProyecto _
    & " and Fecha_Baja<>0" _
    & FiltroFechasTxt)
If MiRc.EOF = False Then
    MiRc.MoveFirst
    
    While Not MiRc.EOF
    
        MiRc1.AddNew
        MiRc1!Usuario_Id = IdUsuarioGeneral
        MiRc1!Col_0 = MiRc!Id
        MiRc1!Fecha = MiRc!Fecha_Baja
        MiRc1!Col_1 = MiRc!Fecha_Baja
        MiRc1!Col_2 = "P"
        MiRc1!Col_3 = " "
        MiRc1!Col_4 = MiRc!concepto
        MiRc1!Col_5 = NombreFicha(MiRc!Ficha_ID)
        If Trim(MiRc1!Col_5) = "" Then MiRc1!Col_5 = " "
        MiRc1!Col_6 = Format(MiRc!importe * -1, FormatoDecimalStandar)
        MiRc1!Col_7 = "N"
        MiRc1!Col_8 = "N"
        MiRc1!Col_9 = "N"
        MiRc1!Col_10 = "N"
        MiRc1!Col_11 = " "
        MiRc1!Col_12 = " "
        MiRc1!Col_13 = " "
        MiRc1!Col_14 = " "
        MiRc1.Update
        
        MiRc.MoveNext
    
    Wend
    
End If
MiRc.Close

'Cerrando captura de datos
MiRc1.Close

'Mostrando datos en el grid
Set MiRc = Base.OpenRecordset("Select *" _
    & " from Auxiliar_Carga" _
    & " where Usuario_Id=" & IdUsuarioGeneral _
    & " order by Fecha asc")

If MiRc.EOF = False Then

    MiRc.MoveFirst
    While Not MiRc.EOF
    
        'Creando línea
        Me.GridCostes.Rows = Me.GridCostes.Rows + 1
        ReDim Preserve IdProyectoCosteProveedor(Me.GridCostes.Rows)
        
        'Pasando datos
        Me.GridCostes.TextMatrix(Me.GridCostes.Rows - 2, 0) = MiRc!Col_0
        Me.GridCostes.TextMatrix(Me.GridCostes.Rows - 2, 1) = MiRc!Col_1
        Me.GridCostes.Row = Me.GridCostes.Rows - 2
        Select Case UCase(Trim(MiRc!Col_2))
            Case "1"       'Material
                Me.GridCostes.TextMatrix(Me.GridCostes.Rows - 2, 3) = "Material"
                Me.GridCostes.Col = 2
                Me.GridCostes.CellPictureAlignment = 4
                Set Me.GridCostes.CellPicture = Me.Icono(1).Picture
            Case "2"        'Mano Obra
                Me.GridCostes.TextMatrix(Me.GridCostes.Rows - 2, 3) = "Mano Obra"
                Me.GridCostes.Col = 2
                Me.GridCostes.CellPictureAlignment = 4
                Set Me.GridCostes.CellPicture = Me.Icono(2).Picture
            Case "3"        'Otros Gastos
                Me.GridCostes.TextMatrix(Me.GridCostes.Rows - 2, 3) = "Otros"
                Me.GridCostes.Col = 2
                Me.GridCostes.CellPictureAlignment = 4
                Set Me.GridCostes.CellPicture = Me.Icono(3).Picture
            Case "4"        'No Clasificados
                Me.GridCostes.TextMatrix(Me.GridCostes.Rows - 2, 3) = "No Clasif."
                Me.GridCostes.Col = 2
                Me.GridCostes.CellPictureAlignment = 4
                Set Me.GridCostes.CellPicture = Me.Icono(4).Picture
            Case "P"        'Previsiones
                EstaCerrada = False
                Set MiRc3 = Base.OpenRecordset("Select Fecha_Baja" _
                    & " from Ajustes_Cuenta_Resultado" _
                    & " where Id=" & MiRc!Col_0)
                If MiRc3.EOF = False Then
                    MiRc3.MoveFirst
                    If IsNull(MiRc3!Fecha_Baja) = False And MiRc3!Fecha_Baja <> 0 Then EstaCerrada = True
                End If
                MiRc3.Close
                Me.GridCostes.TextMatrix(Me.GridCostes.Rows - 2, 3) = "Previsión"
                Me.GridCostes.Col = 2
                Me.GridCostes.CellPictureAlignment = 4
                Set Me.GridCostes.CellPicture = Me.Icono(5).Picture
                Me.GridCostes.Col = 3
                Me.GridCostes.CellBackColor = GrisClaro1
                If EstaCerrada = False Then
                    Me.GridCostes.CellForeColor = &H80&
                Else
                    Me.GridCostes.CellForeColor = VerdeOscuro
                End If
                Me.GridCostes.Col = 9
                Me.GridCostes.CellBackColor = GrisClaro1
                If EstaCerrada = False Then
                    Me.GridCostes.CellForeColor = &H80&
                Else
                    Me.GridCostes.CellForeColor = VerdeOscuro
                End If
        End Select
        Me.GridCostes.TextMatrix(Me.GridCostes.Rows - 2, 4) = MiRc!Col_11
        Me.GridCostes.TextMatrix(Me.GridCostes.Rows - 2, 5) = MiRc!Col_3
        Me.GridCostes.TextMatrix(Me.GridCostes.Rows - 2, 6) = MiRc!Col_4
        If IsNull(MiRc!Col_5) = False Then Me.GridCostes.TextMatrix(Me.GridCostes.Rows - 2, 7) = MiRc!Col_5
        Me.GridCostes.TextMatrix(Me.GridCostes.Rows - 2, 8) = MiRc!Col_12
        Me.GridCostes.TextMatrix(Me.GridCostes.Rows - 2, 9) = MiRc!Col_6
        Select Case VersionAplicativo
            Case 1      'Obras, se pone alerta si el coste asignado no está repartido por meses y partidas
                Select Case NumSM(MiRc!Col_14)
                    Case 0
                        'No se hace nada
                    Case Is <> 0
                        If NumSM(MiRc!Col_14) = NumSM(MiRc!Col_6) Then
                            Me.GridCostes.Row = Me.GridCostes.Rows - 2
                            Me.GridCostes.Col = 10
                            Me.GridCostes.Text = "."
                            Me.GridCostes.CellAlignment = 4
                            Me.GridCostes.CellPictureAlignment = 4
                            Set Me.GridCostes.CellPicture = Me.Icono(7).Picture
                        End If
                End Select
            
            Case 5      'Proyecto técnico, se pone aviso si el coste asignado y el coste total de la factura no coinciden
                If NumSM(MiRc!Col_6) - NumSM(MiRc!Col_12) <> 0 Then
                    Me.GridCostes.Row = Me.GridCostes.Rows - 2
                    Me.GridCostes.Col = 10
                    Me.GridCostes.Text = "."
                    Me.GridCostes.CellAlignment = 4
                    Me.GridCostes.CellPictureAlignment = 4
                    Set Me.GridCostes.CellPicture = Me.Icono(7).Picture
                End If
        End Select
        Select Case Trim(UCase(MiRc!Col_7))
            Case "N"
                'No se hace nada
            Case "S"
                Me.GridCostes.Row = Me.GridCostes.Rows - 2
                Me.GridCostes.Col = 11
                Me.GridCostes.Text = "."
                Me.GridCostes.CellAlignment = 4
                Me.GridCostes.CellPictureAlignment = 4
                Set Me.GridCostes.CellPicture = Me.Icono(11).Picture
        End Select
        Select Case Trim(UCase(MiRc!Col_8))
            Case "N"
                'No se hace nada
            Case "S"
                Me.GridCostes.Row = Me.GridCostes.Rows - 2
                Me.GridCostes.Col = 12
                Me.GridCostes.Text = "."
                Me.GridCostes.CellAlignment = 4
                Me.GridCostes.CellPictureAlignment = 4
                Set Me.GridCostes.CellPicture = Me.Icono(10).Picture
        End Select
        Select Case Trim(UCase(MiRc!Col_9))
            Case "N"
                'No se hace nada
            Case "S"
                Me.GridCostes.Row = Me.GridCostes.Rows - 2
                Me.GridCostes.Col = 13
                Me.GridCostes.Text = "."
                Me.GridCostes.CellAlignment = 4
                Me.GridCostes.CellPictureAlignment = 4
                Set Me.GridCostes.CellPicture = Me.Icono(10).Picture
        End Select
        Select Case Trim(UCase(MiRc!Col_10))
            Case "N"
                'No se hace nada
            Case "S"
                Me.GridCostes.Row = Me.GridCostes.Rows - 2
                Me.GridCostes.Col = 14
                Me.GridCostes.Text = "."
                Me.GridCostes.CellAlignment = 4
                Me.GridCostes.CellPictureAlignment = 4
                Set Me.GridCostes.CellPicture = Me.Icono(10).Picture
        End Select
        IdProyectoCosteProveedor(Me.GridCostes.Rows - 2) = Val(MiRc!Col_13)
        
        'Siguiente fila del grid
        MiRc.MoveNext
        
    Wend
    
End If

MiRc.Close

'Calculando total del listado
Call CalculaTotalListadoCostes

'Cargar Filtro de Proveedores
Base.Execute ("Delete * from Auxiliar_Carga Where Usuario_Id=" & IdUsuarioGeneral)
Me.ProveedoresCombo.Clear
Me.ProveedoresCombo.AddItem "<Todos>", 0

Set MiRc = Base.OpenRecordset("Select *" _
    & " from Auxiliar_Carga" _
    & " where Id=-1")

For Posicion = 1 To Me.GridCostes.Rows - 2
    If Trim(Trim(Me.GridCostes.TextMatrix(Posicion, 7))) <> "" Then
        MiRc.AddNew
        MiRc!Usuario_Id = IdUsuarioGeneral
        MiRc!Col_0 = Trim(Me.GridCostes.TextMatrix(Posicion, 7))
        MiRc.Update
    End If
Next Posicion

MiRc.Close

Posicion = 1
Set MiRc = Base.OpenRecordset("Select Col_0" _
    & " from Auxiliar_Carga" _
    & " where Usuario_Id=" & IdUsuarioGeneral _
    & " group by Col_0" _
    & " order by Col_0 Asc")

If MiRc.EOF = False Then
    MiRc.MoveFirst
    While Not MiRc.EOF
        Me.ProveedoresCombo.AddItem Trim(MiRc!Col_0), Posicion
        MiRc.MoveNext
        Posicion = Posicion + 1
    Wend
End If

MiRc.Close

Me.ProveedoresCombo.ListIndex = 0

'Saliendo
Me.GridCostes.Redraw = True

Me.GridCostes.Row = 1
Me.GridCostes.Col = 0

End Sub
```

### `CargaGastosTesoreria`
Reemplazar por

```vb
Private Sub CargaGastosTesoreria()

Dim MiRc As Recordset
Dim MiRc1 As Recordset
Dim MiRc2 As Recordset
Dim Fecha1 As Date
Dim Fecha2 As Date
Dim FiltroFechasTxt As String
Dim FechaFactura As Date
Dim ImporteAsignado As Double
Dim Posicion As Double
Dim Auxiliar As Double

'Borrando tabla auxiliar de carga
Base.Execute ("Delete * from Auxiliar_Carga Where Usuario_Id=" & IdUsuarioGeneral)
Me.FiltroProveedorCombo.Clear
ReDim IdProyectoCosteOperacion(0)
ReDim ImportePendienteRepartir(0)

'Etableciendo filtro de fechas
If Me.TipoListadoCombo.ListIndex = 0 Then
    Fecha1 = 0
    If Me.FechaLimiteListadoTxt.Text <> "__/__/____" Then
        Fecha1 = Me.FechaLimiteListadoTxt.Text
        Fecha1 = "01/" & Format(Fecha1, "mm/yyyy")
        Fecha1 = DateAdd("m", 1, Fecha1)
        Fecha1 = DateAdd("d", -1, Fecha1)
        FiltroFechasTxt = " and Fecha<=#" & Format(Fecha1, "m/d/yy") & "#"
    Else
        FiltroFechasTxt = ""
    End If
Else
    FiltroFechasTxt = ""

    ' Desde
    If Me.FechaLimiteListadoTxt.Text <> "__/__/____" Then
        Fecha1 = Me.FechaLimiteListadoTxt.Text
        Fecha1 = "01/" & Format(Fecha1, "mm/yyyy")
        FiltroFechasTxt = " AND Fecha >= #" & Format(Fecha1, "m/d/yy") & "#"
    End If
    
    ' Hasta
    If Me.FechaLimite1Txt.Text <> "__/__/____" Then
        Fecha2 = "01/" & Format(Fecha2, "mm/yyyy")
        Fecha2 = DateAdd("m", 1, Fecha2)
        Fecha2 = DateAdd("d", -1, Fecha2)
        FiltroFechasTxt = FiltroFechasTxt + " AND FECHA <= #" & Format(Fecha2, "m/d/yy") & "#"
    End If
End If

'Preparando grid
Me.GridOtrosGastos.Rows = 1
Me.GridOtrosGastos.Rows = 2

Me.GridOtrosGastos.Redraw = False

Me.FiltroTipoCosteCombo.ListIndex = 0

Set MiRc1 = Base.OpenRecordset("Select *" _
    & " from Auxiliar_Carga" _
    & " where Id=-1")

Set MiRc = Base.OpenRecordset("Select Operaciones_Bancarias.Expediente_Id as IdCosteProyecto, Proyecto_Coste_Proveedor.Coste_Asignado," _
    & "Operaciones_Bancarias.Id,Operaciones_Bancarias.Fecha_Operacion,Operaciones_Bancarias.Registro_Entrada,Operaciones_Bancarias.Descripcion," _
    & "Operaciones_Bancarias.Entidad_Id,Operaciones_Bancarias.Trabajador_Id,Operaciones_Bancarias.Importe_Operacion," _
    & "Operaciones_Bancarias.Pago,Operaciones_Bancarias.Clasificacion_ACR" _
    & " from Operaciones_Bancarias inner join Proyecto_Coste_Proveedor" _
    & " on Operaciones_Bancarias.Id=Proyecto_Coste_Proveedor.Operacion_Tesoreria_Directa_Id" _
    & " where Proyecto_Coste_Proveedor.Proyecto_Id=" & IdProyecto _
    & " order by Operaciones_Bancarias.Fecha_Operacion asc")

If MiRc.EOF = False Then

    MiRc.MoveFirst
    While Not MiRc.EOF

        'Comprobando si hay desglose temporal,
            'Si no lo hay se cogen los datos de la operación
            'Si lo hay y no está en fecha se sale del proces
            'Si lo hay y está en fecha se coge sólo el importe correspondiente
        FechaFactura = 0
        ImporteAsignado = 0
        Set MiRc2 = Base.OpenRecordset("Select sum(Coste_Asignado) as coste " _
            & " from Desglose_Temporal_Asignacion_Coste" _
            & " where Operacion_Tesoreria_Directa_Id=" & MiRc!Id _
            & " and Proyecto_Id=" & IdProyecto _
            & " and Tipo_Item=0" _
            & FiltroFechasTxt)
        MiRc2.MoveFirst
        If IsNull(MiRc2!Coste) = True Then  'No existe desglose
            FechaFactura = MiRc!fecha_operacion
            ImporteAsignado = MiRc!Coste_Asignado
            If Me.TipoListadoCombo.ListIndex = 0 Then   'A origen
                If FechaFactura > Fecha1 Then
                    MiRc2.Close
                    GoTo SiguienteFactura
                End If
            Else    'Entre fechas
                If FechaFactura < Fecha1 Or FechaFactura > Fecha2 Then
                    MiRc2.Close
                    GoTo SiguienteFactura
                End If
            End If
        Else    'Existe desglose
            'Hay que crear una línea por cada mes de asignación
            MiRc2.Close
            Set MiRc2 = Base.OpenRecordset("Select Fecha,sum(Coste_Asignado) as Coste" _
                & " from Desglose_Temporal_Asignacion_Coste" _
                & " where Operacion_Tesoreria_Directa_Id=" & MiRc!Id _
                & " and Coste_Asignado<>0" _
                & " and Proyecto_Id=" & IdProyecto _
                & " and Tipo_Item=0" _
                & FiltroFechasTxt _
                & " group by Fecha")
            If MiRc2.EOF = False Then
                MiRc2.MoveFirst
                While Not MiRc2.EOF
                    MiRc1.AddNew
                    MiRc1!Usuario_Id = IdUsuarioGeneral
                    MiRc1!Col_0 = MiRc!Id
                    MiRc1!Fecha = MiRc2!Fecha
                    MiRc1!Col_1 = MiRc2!Fecha
                    Select Case MiRc!Clasificacion_ACR
                        Case 1  'Material
                            MiRc1!Col_2 = "1"
                        Case 2  'Mano de obra
                            MiRc1!Col_2 = "2"
                        Case 3, 4, 5, 6 'Otros
                            MiRc1!Col_2 = "3"
                        Case Else
                            MiRc1!Col_2 = "4"
                    End Select
                    If IsNull(MiRc!Registro_Entrada) = False And MiRc!Registro_Entrada <> 0 Then
                        MiRc1!Col_3 = "OP" & Format(MiRc!fecha_operacion, "yy") & "-" & Trim(str(MiRc!Registro_Entrada))
                    Else
                        MiRc1!Col_3 = " "
                    End If
                    MiRc1!Col_4 = MiRc!Descripcion
                    If MiRc!Entidad_Id <> 0 Then
                        MiRc1!Col_5 = NombreFicha(MiRc!Entidad_Id)
                    Else
                        If MiRc!Trabajador_id <> 0 Then
                            MiRc1!Col_5 = CapturaOperario(MiRc!Trabajador_id)
                        End If
                    End If
                    If MiRc1!Col_5 = "" Then MiRc1!Col_5 = " "
                    MiRc1!Col_6 = Format(MiRc2!Coste, FormatoDecimalStandar)
                    MiRc1!Col_11 = MiRc!fecha_operacion
                    If MiRc!Pago = True Then
                        MiRc1!Col_12 = Format(MiRc!Importe_Operacion, FormatoDecimalStandar)
                    Else
                        MiRc1!Col_12 = Format(MiRc!Importe_Operacion * -1, FormatoDecimalStandar)
                    End If
                    MiRc1!Col_13 = MiRc!IdCosteProyecto
                    Auxiliar = 0
                    If IsNull(MiRc!Coste_Asignado) = False Then Auxiliar = MiRc!Coste_Asignado
                    MiRc1!Col_14 = Format(Auxiliar - CargarAsignacionPorPartidasTesoreria(MiRc!Id, IdProyecto), FormatoDecimalStandar)

                    MiRc1.Update
                    'Siguiente asignación temporal
                    MiRc2.MoveNext
                Wend
            End If
            MiRc2.Close
            GoTo SiguienteFactura
        End If
        MiRc2.Close

        'Grabando
        MiRc1.AddNew

        MiRc1!Usuario_Id = IdUsuarioGeneral
        MiRc1!Col_0 = MiRc!Id
        MiRc1!Fecha = FechaFactura
        MiRc1!Col_1 = FechaFactura
        Select Case MiRc!Clasificacion_ACR
            Case 1  'Material
                MiRc1!Col_2 = "1"
            Case 2  'Mano de obra
                MiRc1!Col_2 = "2"
            Case 3, 4, 5, 6 'Otros
                MiRc1!Col_2 = "3"
            Case Else
                MiRc1!Col_2 = "4"
        End Select
        If IsNull(MiRc!Registro_Entrada) = False And MiRc!Registro_Entrada <> 0 Then
            MiRc1!Col_3 = "OP" & Format(MiRc!fecha_operacion, "yy") & "-" & Trim(str(MiRc!Registro_Entrada))
        Else
            MiRc1!Col_3 = " "
        End If
        MiRc1!Col_4 = MiRc!Descripcion
        If MiRc!Entidad_Id <> 0 Then
            MiRc1!Col_5 = NombreFicha(MiRc!Entidad_Id)
        Else
            If MiRc!Trabajador_id <> 0 Then
                MiRc1!Col_5 = CapturaOperario(MiRc!Trabajador_id)
            End If
        End If
        If MiRc1!Col_5 = "" Then MiRc1!Col_5 = " "
        MiRc1!Col_6 = Format(ImporteAsignado, FormatoDecimalStandar)
        MiRc1!Col_11 = MiRc!fecha_operacion
        If MiRc!Pago = True Then
            MiRc1!Col_12 = Format(MiRc!Importe_Operacion, FormatoDecimalStandar)
        Else
            MiRc1!Col_12 = Format((MiRc!Importe_Operacion * -1), FormatoDecimalStandar)
        End If
        MiRc1!Col_13 = MiRc!IdCosteProyecto
        Auxiliar = 0
        If IsNull(MiRc!Coste_Asignado) = False Then Auxiliar = MiRc!Coste_Asignado
        MiRc1!Col_14 = Format(Auxiliar - CargarAsignacionPorPartidasTesoreria(MiRc!Id, IdProyecto), FormatoDecimalStandar)

        MiRc1.Update

SiguienteFactura:
        'Siguiente Factura
        MiRc.MoveNext

    Wend

End If

MiRc.Close

'Capturando Operaciones de partes de trabajo incluidas en concepto de otros de forma automática
Set MiRc = Base.OpenRecordset("Select *" _
    & " from Desglose_Parte_Trabajo" _
    & " where Proyecto_Id=" & IdProyecto _
    & FiltroFechasTxt _
    & " order by Fecha asc")
    
If MiRc.EOF = False Then

    While Not MiRc.EOF
  
        'Grabando
        MiRc1.AddNew

        MiRc1!Usuario_Id = IdUsuarioGeneral
        MiRc1!Col_0 = MiRc!Id * -1
        MiRc1!Fecha = MiRc!Fecha
        MiRc1!Col_1 = MiRc!Fecha
        MiRc1!Col_2 = "2"
        MiRc1!Col_3 = " "
        MiRc1!Col_4 = "Mano de Obra Directa"
        MiRc1!Col_5 = CapturaOperario(MiRc!Trabajador_id)
        If MiRc1!Col_5 = "" Then MiRc1!Col_5 = " "
        MiRc1!Col_6 = Format(MiRc!importe_total, FormatoDecimalStandar)
        MiRc1!Col_11 = MiRc!Fecha
        MiRc1!Col_12 = Format(MiRc!importe_total, FormatoDecimalStandar)
        MiRc1!Col_13 = IdProyecto
        Auxiliar = 0
        If IsNull(MiRc!importe_total) = False Then Auxiliar = MiRc!importe_total
        MiRc1!Col_14 = Format(Auxiliar - CargarAsignacionPorPartidasNomina(MiRc!Id, IdProyecto), FormatoDecimalStandar)

        MiRc1.Update
        
        'Siguiente registro
        MiRc.MoveNext
        
    Wend
        
End If

MiRc.Close

'Cerrando captura de datos
MiRc1.Close

'Mostrando datos en el grid
Set MiRc = Base.OpenRecordset("Select *" _
    & " from Auxiliar_Carga" _
    & " where Usuario_Id=" & IdUsuarioGeneral _
    & " order by Fecha asc")

If MiRc.EOF = False Then

    MiRc.MoveFirst
    While Not MiRc.EOF

        'Creando línea
        Me.GridOtrosGastos.Rows = Me.GridOtrosGastos.Rows + 1
        ReDim Preserve IdProyectoCosteOperacion(Me.GridOtrosGastos.Rows)
        ReDim Preserve ImportePendienteRepartir(Me.GridOtrosGastos.Rows)

        'Pasando datos
        Me.GridOtrosGastos.TextMatrix(Me.GridOtrosGastos.Rows - 2, 0) = MiRc!Col_0
        Me.GridOtrosGastos.TextMatrix(Me.GridOtrosGastos.Rows - 2, 1) = MiRc!Col_1
        Me.GridOtrosGastos.Row = Me.GridOtrosGastos.Rows - 2
        Select Case UCase(Trim(MiRc!Col_2))
            Case "1"       'Material
                Me.GridOtrosGastos.TextMatrix(Me.GridOtrosGastos.Rows - 2, 3) = "Material"
                Me.GridOtrosGastos.Col = 2
                Me.GridOtrosGastos.CellPictureAlignment = 4
                Set Me.GridOtrosGastos.CellPicture = Me.Icono(1).Picture
            Case "2"        'Mano Obra
                Me.GridOtrosGastos.TextMatrix(Me.GridOtrosGastos.Rows - 2, 3) = "Mano Obra"
                Me.GridOtrosGastos.Col = 2
                Me.GridOtrosGastos.CellPictureAlignment = 4
                Set Me.GridOtrosGastos.CellPicture = Me.Icono(2).Picture
            Case "3"        'Otros Gastos
                Me.GridOtrosGastos.TextMatrix(Me.GridOtrosGastos.Rows - 2, 3) = "Otros"
                Me.GridOtrosGastos.Col = 2
                Me.GridOtrosGastos.CellPictureAlignment = 4
                Set Me.GridOtrosGastos.CellPicture = Me.Icono(3).Picture
            Case "4"        'No Clasificados
                Me.GridOtrosGastos.TextMatrix(Me.GridOtrosGastos.Rows - 2, 3) = "No Clasif."
                Me.GridOtrosGastos.Col = 2
                Me.GridOtrosGastos.CellPictureAlignment = 4
                Set Me.GridOtrosGastos.CellPicture = Me.Icono(4).Picture
            Case "P"        'Previsiones
                Me.GridOtrosGastos.TextMatrix(Me.GridOtrosGastos.Rows - 2, 3) = "Previsión"
                Me.GridOtrosGastos.Col = 2
                Me.GridOtrosGastos.CellPictureAlignment = 4
                Set Me.GridOtrosGastos.CellPicture = Me.Icono(5).Picture
                Me.GridOtrosGastos.Col = 3
                Me.GridOtrosGastos.CellBackColor = GrisClaro1
                Me.GridOtrosGastos.CellForeColor = &H80&
        End Select
        Me.GridOtrosGastos.TextMatrix(Me.GridOtrosGastos.Rows - 2, 4) = MiRc!Col_11
        Me.GridOtrosGastos.TextMatrix(Me.GridOtrosGastos.Rows - 2, 5) = MiRc!Col_3
        Me.GridOtrosGastos.TextMatrix(Me.GridOtrosGastos.Rows - 2, 6) = MiRc!Col_4
        If IsNull(MiRc!Col_5) = False Then Me.GridOtrosGastos.TextMatrix(Me.GridOtrosGastos.Rows - 2, 7) = MiRc!Col_5
        If VersionAplicativo <> 5 Then Me.GridOtrosGastos.TextMatrix(Me.GridOtrosGastos.Rows - 2, 8) = MiRc!Col_12
        Me.GridOtrosGastos.TextMatrix(Me.GridOtrosGastos.Rows - 2, 9) = MiRc!Col_6
        If VersionAplicativo <> 5 Then
            ImportePendienteRepartir(Me.GridOtrosGastos.Rows - 2) = NumSM(MiRc!Col_12) - NumSM(MiRc!Col_6)
            Select Case ImportePendienteRepartir(Me.GridOtrosGastos.Rows - 2)
                Case 0
                    Me.GridOtrosGastos.Row = Me.GridOtrosGastos.Rows - 2
                    Me.GridOtrosGastos.Col = 10
                    Me.GridOtrosGastos.Text = ""
                    Me.GridOtrosGastos.CellAlignment = 4
                    Set Me.GridOtrosGastos.CellPicture = Nothing
                Case Is <> 0
                    Me.GridOtrosGastos.Row = Me.GridOtrosGastos.Rows - 2
                    Me.GridOtrosGastos.Col = 10
                    Me.GridOtrosGastos.Text = "."
                    Me.GridOtrosGastos.CellAlignment = 4
                    Me.GridOtrosGastos.CellPictureAlignment = 4
                    Set Me.GridOtrosGastos.CellPicture = Me.Icono(7).Picture
            End Select
        End If
        
        IdProyectoCosteOperacion(Me.GridOtrosGastos.Rows - 2) = Val(MiRc!Col_13)

        'Siguiente fila del grid
        MiRc.MoveNext

    Wend

End If

MiRc.Close

'Calculando total del listado
Call CalculaTotalListadoOtrosGastos

'Cargar Filtro de Proveedores
Base.Execute ("Delete * from Auxiliar_Carga Where Usuario_Id=" & IdUsuarioGeneral)
Me.FiltroProveedorCombo.Clear
Me.FiltroProveedorCombo.AddItem "<Todos>", 0

Set MiRc = Base.OpenRecordset("Select *" _
    & " from Auxiliar_Carga" _
    & " where Id=-1")

For Posicion = 1 To Me.GridOtrosGastos.Rows - 2
    If Trim(Trim(Me.GridOtrosGastos.TextMatrix(Posicion, 7))) <> "" Then
        MiRc.AddNew
        MiRc!Usuario_Id = IdUsuarioGeneral
        MiRc!Col_0 = Trim(Me.GridOtrosGastos.TextMatrix(Posicion, 7))
        MiRc.Update
    End If
Next Posicion

MiRc.Close

Posicion = 1
Set MiRc = Base.OpenRecordset("Select Col_0" _
    & " from Auxiliar_Carga" _
    & " where Usuario_Id=" & IdUsuarioGeneral _
    & " group by Col_0" _
    & " order by Col_0 Asc")

If MiRc.EOF = False Then
    MiRc.MoveFirst
    While Not MiRc.EOF
        Me.FiltroProveedorCombo.AddItem Trim(MiRc!Col_0), Posicion
        MiRc.MoveNext
        Posicion = Posicion + 1
    Wend
End If

MiRc.Close

Me.FiltroProveedorCombo.ListIndex = 0

'Saliendo
Me.GridOtrosGastos.Redraw = True

Me.GridOtrosGastos.Row = 1
Me.GridOtrosGastos.Col = 0

End Sub
```
