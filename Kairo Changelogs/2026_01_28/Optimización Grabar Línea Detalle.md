Miércoles 28/01/2026. ==Pendiente pruebas en Kedeke==, rama `optimizacion_ventas`.

## `DocumentoVentaFrm.frm`

### `GrabarLineaDetalle`
Reemplazar por
```vb
Private Sub GrabarLineaDetalle(ByVal Fila As Double, Optional IdLineaModificar As Double = 0)

Dim MiRc As Recordset
Dim MiRc1 As Recordset
Dim IdLinea As Double
Dim ImporteTotal As Double
Dim CobradoPrevio As Double
Dim ImporteTotalPrevio As Double
Dim ImportePendientePrevio As Double
Dim Auxiliar As String
Dim Unidades As Double
Dim precio As Double

Me.TextoEdicionTxt.Tag = ""

'Grabando
If IdLineaModificar = 0 Then
    Set MiRc = Base.OpenRecordset("Select *" _
        & " from Detalle_Documentos_Venta" _
        & " where Id=-1")
    MiRc.AddNew
Else
    Set MiRc = Base.OpenRecordset("Select *" _
        & " from Detalle_Documentos_Venta" _
        & " where Id=" & IdLineaModificar)
    MiRc.MoveFirst
    MiRc.Edit
End If

IdLinea = MiRc!Id
Me.GridDetalle.TextMatrix(Me.GridDetalle.Row, 0) = IdLinea

precio = PorcentajeIVAArticuloDefecto(IdArticuloTemporal)

MiRc!cabecera_Id = IdDocumento
MiRc!articulo_id = IdArticuloTemporal
MiRc!Descripcion_Articulo = Trim(Me.GridDetalle.TextMatrix(Fila, 3))
MiRc!Descripcion_Larga_Unificada = MiRc!Descripcion_Articulo
If Trim(Me.GridDetalle.TextMatrix(Fila, 4)) <> "" Then
    MiRc!Descripcion_Larga = Trim(Me.GridDetalle.TextMatrix(Fila, 4))
    MiRc!Descripcion_Larga_Unificada = MiRc!Descripcion_Larga_Unificada & chr(13) & chr(10) & MiRc!Descripcion_Larga
Else
    MiRc!Descripcion_Larga = " "
End If
MiRc!Cantidad = NumSM(Me.GridDetalle.TextMatrix(Fila, 8))
Unidades = MiRc!Cantidad
MiRc!precio = MascaraImporte(NumSM(Me.GridDetalle.TextMatrix(Fila, 9)), precio, 2)
MiRc!Descuento_1 = NumSM(Me.GridDetalle.TextMatrix(Fila, 10))
MiRc!Descuento_2 = NumSM(Me.GridDetalle.TextMatrix(Fila, 11))
MiRc!Importe_Neto = MascaraImporte(NumSM(Me.GridDetalle.TextMatrix(Fila, 12)), precio, 2)
ImporteTotal = MiRc!Importe_Neto
MiRc!Iva_id = CapturaIdIvaArticulo(IdArticuloTemporal)
MiRc!Iva_Porcentaje = CapturaPorcentajeIVA(MiRc!Iva_id)
MiRc!Cuota_Iva = MiRc!Importe_Neto * (MiRc!Iva_Porcentaje / 100)
MiRc!total_linea = MiRc!Importe_Neto + MiRc!Cuota_Iva
MiRc!Orden = Fila
MiRc!Lote_Id = IdLoteTemporal

MiRc.Update

MiRc.Close

If TipoDocumento = 4 Then
    'Importe Total Certificado
    Set MiRc = Base.OpenRecordset("Select Sum(Importe_Neto) as Importe" _
        & " from Detalle_Documentos_Venta" _
        & " where Cabecera_Id=" & IdDocumento _
        & " and Articulo_Id>0")
    If MiRc.EOF = False Then
        MiRc.MoveFirst
        If IsNull(MiRc!importe) = False Then
            Base.Execute ("Update Cabecera_Documentos_Venta set Importe_Certificacion_Origen=" & CadenaSQL(MiRc!importe) & " where Id=" & IdDocumento)
        Else
            Base.Execute ("Update Cabecera_Documentos_Venta set Importe_Certificacion_Origen=0 where Id=" & IdDocumento)
        End If
    End If
    MiRc.Close
    'Importe Certificaciones Anteriores
    Set MiRc = Base.OpenRecordset("Select Sum(Importe_Neto) as Importe" _
        & " from Detalle_Documentos_Venta" _
        & " where Cabecera_Id=" & IdDocumento _
        & " and Articulo_Id=-999")
    If MiRc.EOF = False Then
        MiRc.MoveFirst
        If IsNull(MiRc!importe) = False Then
            Base.Execute ("Update Cabecera_Documentos_Venta set Importe_Certificaciones_Anteriores=" & CadenaSQL(MiRc!importe * -1) & " where Id=" & IdDocumento)
        Else
            Base.Execute ("Update Cabecera_Documentos_Venta set Importe_Certificaciones_Anteriores=0 where Id=" & IdDocumento)
        End If
    End If
    MiRc.Close
End If

'Actualizando Pie de Factura (y grabando en base de datos)
Fila = 1
Me.IVA1Txt.Text = ""
Me.IVA2Txt.Text = ""
Me.IVA3Txt.Text = ""
Me.IVA4Txt.Text = ""
Me.IVA5Txt.Text = ""
Me.Base1Txt.Text = ""
Me.Base2Txt.Text = ""
Me.Base3Txt.Text = ""
Me.Base4Txt.Text = ""
Me.Base5Txt.Text = ""
Me.Cuota1Txt.Text = ""
Me.Cuota2Txt.Text = ""
Me.Cuota3Txt.Text = ""
Me.Cuota4Txt.Text = ""
Me.Cuota5Txt.Text = ""
Me.Label11(10).Caption = ""
Me.BaseImponibleTxt.Text = ""
Me.CuotaIvaTxt.Text = ""
Me.TotalFacturaTxt.Text = ""
Me.RetencionesTxt.Text = ""
Me.ParaCobrarTxt.Text = ""
Me.ProvisionFondosTxt.Text = ""
Me.CertificacionActualTxt.Text = ""
Me.CertificacionAOrigenTxt.Text = ""
Me.CertificacionesAnterioresTxt.Text = ""

Set MiRc = Base.OpenRecordset("Select Iva_Porcentaje" _
    & " from Detalle_Documentos_Venta" _
    & " where Cabecera_Id=" & IdDocumento _
    & " group by Iva_Porcentaje" _
    & " order by Iva_Porcentaje desc")

If MiRc.EOF = False Then

    MiRc.MoveFirst
    While Not MiRc.EOF

        'Capturando tipo de iva
        Select Case Fila
            Case 1
                Me.IVA1Txt.Text = Format(MiRc!Iva_Porcentaje, FormatoDecimalStandar)
            Case 2
                Me.IVA2Txt.Text = Format(MiRc!Iva_Porcentaje, FormatoDecimalStandar)
            Case 3
                Me.IVA3Txt.Text = Format(MiRc!Iva_Porcentaje, FormatoDecimalStandar)
            Case 4
                Me.IVA4Txt.Text = Format(MiRc!Iva_Porcentaje, FormatoDecimalStandar)
            Case 5
                Me.IVA5Txt.Text = Format(MiRc!Iva_Porcentaje, FormatoDecimalStandar)
        End Select

        'Capturando datos agregados
        Set MiRc1 = Base.OpenRecordset("Select sum(Importe_Neto) as Neto, sum(Cuota_IVA) as IVA" _
            & " from Detalle_Documentos_Venta" _
            & " where Cabecera_Id=" & IdDocumento _
            & " and IVA_Porcentaje=" & CadenaSQL(MiRc!Iva_Porcentaje))

        If MiRc1.EOF = False Then

            MiRc1.MoveFirst

            Select Case Fila
                Case 1
                    If IsNull(MiRc1!Neto) = False Then Me.Base1Txt.Text = Format(MiRc1!Neto, FormatoDecimalPrecio)
                    If IsNull(MiRc1!iva) = False Then Me.Cuota1Txt.Text = Format(MiRc1!iva, FormatoDecimalPrecio)
                Case 2
                    If IsNull(MiRc1!Neto) = False Then Me.Base2Txt.Text = Format(MiRc1!Neto, FormatoDecimalPrecio)
                    If IsNull(MiRc1!iva) = False Then Me.Cuota2Txt.Text = Format(MiRc1!iva, FormatoDecimalPrecio)
                Case 3
                    If IsNull(MiRc1!Neto) = False Then Me.Base3Txt.Text = Format(MiRc1!Neto, FormatoDecimalPrecio)
                    If IsNull(MiRc1!iva) = False Then Me.Cuota3Txt.Text = Format(MiRc1!iva, FormatoDecimalPrecio)
                Case 4
                    If IsNull(MiRc1!Neto) = False Then Me.Base4Txt.Text = Format(MiRc1!Neto, FormatoDecimalPrecio)
                    If IsNull(MiRc1!iva) = False Then Me.Cuota4Txt.Text = Format(MiRc1!iva, FormatoDecimalPrecio)
                Case 5
                    If IsNull(MiRc1!Neto) = False Then Me.Base5Txt.Text = Format(MiRc1!Neto, FormatoDecimalPrecio)
                    If IsNull(MiRc1!iva) = False Then Me.Cuota5Txt.Text = Format(MiRc1!iva, FormatoDecimalPrecio)
            End Select

        End If

        MiRc1.Close

        'Siguiente tipo de iva
        MiRc.MoveNext
        Fila = Fila + 1

    Wend

End If

MiRc.Close

Fila = NumSM(Me.Base1Txt.Text) + NumSM(Me.Base2Txt.Text) + NumSM(Me.Base3Txt.Text) + NumSM(Me.Base4Txt.Text) + NumSM(Me.Base5Txt.Text) + NumSM(Me.Cuota1Txt.Text) + NumSM(Me.Cuota2Txt.Text) + NumSM(Me.Cuota3Txt.Text) + NumSM(Me.Cuota4Txt.Text) + NumSM(Me.Cuota5Txt.Text)
Me.Label11(10).Caption = Format(Fila, FormatoDecimalStandar)

Set MiRc = Base.OpenRecordset("Select Tipo_Retencion,IVA_1,IVA_2,IVA_3,IVA_4,IVA_5," _
    & "Base_1,Base_2,Base_3,Base_4,Base_5,Cuota_1,Cuota_2,Cuota_3,Cuota_4,Cuota_5," _
    & "Base_Documento,Cuota_IVA_Documento,Importe_Total_Documento,Porcentaje_Retencion,Importe_Retencion," _
    & "Importe_Retencion_Pendiente_Cobro,Importe_Retencion_Pendiente_Facturar" _
    & " from Cabecera_Documentos_Venta" _
    & " where Id=" & IdDocumento)
If MiRc.EOF = False Then

    MiRc.MoveFirst

    CobradoPrevio = 0
    ImportePendientePrevio = 0
    ImporteTotalPrevio = 0
    If IsNull(MiRc!Importe_Total_Documento) = False Then ImporteTotalPrevio = MiRc!Importe_Total_Documento
    CobradoPrevio = ImporteTotalPrevio - ImportePendientePrevio

    MiRc.Edit

    MiRc!IVA_1 = NumSM(Me.IVA1Txt.Text)
    MiRc!IVA_2 = NumSM(Me.IVA2Txt.Text)
    MiRc!IVA_3 = NumSM(Me.IVA3Txt.Text)
    MiRc!IVA_4 = NumSM(Me.IVA4Txt.Text)
    MiRc!IVA_5 = NumSM(Me.IVA5Txt.Text)
    MiRc!Base_1 = NumSM(Me.Base1Txt.Text)
    MiRc!Base_2 = NumSM(Me.Base2Txt.Text)
    MiRc!Base_3 = NumSM(Me.Base3Txt.Text)
    MiRc!Base_4 = NumSM(Me.Base4Txt.Text)
    MiRc!Base_5 = NumSM(Me.Base5Txt.Text)
    MiRc!Cuota_1 = NumSM(Me.Cuota1Txt.Text)
    MiRc!Cuota_2 = NumSM(Me.Cuota2Txt.Text)
    MiRc!Cuota_3 = NumSM(Me.Cuota3Txt.Text)
    MiRc!Cuota_4 = NumSM(Me.Cuota4Txt.Text)
    MiRc!Cuota_5 = NumSM(Me.Cuota5Txt.Text)
    MiRc!Base_Documento = NumSM(Me.Base1Txt.Text) + NumSM(Me.Base2Txt.Text) + NumSM(Me.Base3Txt.Text) + NumSM(Me.Base4Txt.Text) + NumSM(Me.Base5Txt.Text)
    Me.BaseImponibleTxt.Text = Format(MiRc!Base_Documento, FormatoDecimalStandar)
    MiRc!Cuota_IVA_Documento = NumSM(Me.Cuota1Txt.Text) + NumSM(Me.Cuota2Txt.Text) + NumSM(Me.Cuota3Txt.Text) + NumSM(Me.Cuota4Txt.Text) + NumSM(Me.Cuota5Txt.Text)
    Me.CuotaIvaTxt.Text = Format(MiRc!Cuota_IVA_Documento, FormatoDecimalStandar)
    MiRc!Importe_Total_Documento = NumSM(Me.Label11(10).Caption)
    Me.TotalFacturaTxt.Text = Format(MiRc!Importe_Total_Documento, FormatoDecimalStandar)
    Call CalculaRetenciones
    MiRc!Porcentaje_Retencion = NumSM(Me.PorcentajeRetencionTxt.Text)
    MiRc!Importe_Retencion = NumSM(Me.RetencionesTxt.Text)
    MiRc!Tipo_Retencion = Val(Me.PorcentajeRetencionTxt.Tag)
    MiRc!Importe_Retencion_Pendiente_Cobro = CalculoRetencionPendienteCobro(MiRc!Importe_Retencion, IdDocumento)
    MiRc!Importe_Retencion_Pendiente_Facturar = 0

    MiRc.Update
End If
MiRc.Close

'Calculando importe pendiente cobro
Call ActualizarPendienteCobroFactura(IdDocumento)

'Calculando pie de certificación si procede
If TipoDocumento = 2 Or TipoDocumento = 4 Then
    Me.CertificacionAOrigenTxt.Text = Format(CalculoCertificacionOrigen(Val(Me.ExpedienteTxt.Tag), Val(Me.VersionTxt.Text)), FormatoDecimalStandar)
    Me.CertificacionActualTxt.Text = Format(NumSM(Me.BaseImponibleTxt.Text) + NumSM(Me.ProvisionFondosTxt.Text), FormatoDecimalStandar)
    If Val(Me.CertificacionAOrigenTxt.Text) <> 0 Then Me.CertificacionesAnterioresTxt.Text = Format(NumSM(Me.CertificacionAOrigenTxt.Text) - NumSM(Me.CertificacionActualTxt.Text), FormatoDecimalStandar)
End If

'Actualizando datos de la solicitud si existe
If Val(Me.SolicitudTxt.Tag) <> 0 Then Call ActualizarDatosSolicitud(Val(Me.SolicitudTxt.Tag))

'Actualizando datos cabecera proyecto si es necesarios
If Val(Me.ExpedienteTxt.Tag) <> 0 Then Call CalcularImporteProyecto(Val(Me.ExpedienteTxt.Tag))

Call GrabarNumeroRegistro

'Actualizando inventario si procede
If TipoDocumento = 3 Then
    Call ActualizarInventario(2, IdLinea, IdArticuloTemporal, Val(Me.AlmacenTxt.Tag), Unidades, ImporteTotal, Me.FechaDocumentoTxt.Text)
    Auxiliar = CapturaLote(IdLoteTemporal)
    Call ActualizarLotes(2, 0, Auxiliar, 0, IdLinea, IdArticuloTemporal)
    Me.GridDetalle.TextMatrix(Me.GridDetalle.Row, 13) = Auxiliar
End If

If TipoDocumento = 2 Then Call ActualizarStockDisponible(Val(Me.CodigoArticuloTxt.Tag))

'Saliendo
IdArticuloTemporal = 0
IdLoteTemporal = 0

Debug.Print vbCrLf

End Sub
```
### `PorcentajeIVAArticuloDefecto`
Reemplazar por

```vb
Private Function PorcentajeIVAArticuloDefecto(ByVal IdA As Double) As Double

Dim MiRc As Recordset
Dim RsIVA As DAO.Recordset
Dim IndiceIva As Double
Dim idFamilia As Double

IndiceIva = 0
idFamilia = 0
PorcentajeIVAArticuloDefecto = 0

If IdA = 0 Then Exit Function

'Datos de IVA generales Artículo
'Set MiRc = Base.OpenRecordset("Select Tipo_IVA_Id,Familia_Id" _
'    & " from Articulos" _
'    & " where Id=" & IdA)
'If MiRc.EOF = False Then
'    MiRc.MoveFirst
'    IndiceIva = MiRc!Tipo_Iva_Id
'    idFamilia = MiRc!Familia_Id
'End If
'MiRc.Close

'Si el Artículo no tiene iva se coge el de la familia
'If IndiceIva = 0 Then
'    Set MiRc = Base.OpenRecordset("Select Tipo_IVA" _
'        & " from Familias" _
'        & " where Id=" & idFamilia)
'    If MiRc.EOF = False Then
'        MiRc.MoveFirst
'        IndiceIva = MiRc!Tipo_IVA
'    End If
'    MiRc.Close
'End If

' Combinación del código de arriba pero en una sola consulta
Set MiRc = Base.OpenRecordset( _
    " SELECT A.Tipo_IVA_Id, A.Familia_Id, F.Tipo_IVA AS TipoIvaFamilia" & _
    " FROM Articulos AS A" & _
    " LEFT JOIN Familias AS F ON F.Id = A.Familia_Id" & _
    " WHERE A.Id=" & IdA _
)

If Not MiRc.EOF Then
    Call MiRc.MoveFirst
    
    IndiceIva = MiRc!Tipo_Iva_Id
    
    If IndiceIva = 0 Then
        If Not IsNull(MiRc!TipoIvaFamilia) Then
            IndiceIva = MiRc!TipoIvaFamilia
        End If
    End If
End If

Call MiRc.Close

'Si hay Proyecto asociado se coge el IVA del Proyecto si está seleccionado)
If Val(Me.ExpedienteTxt.Tag) <> 0 Then
    Set MiRc = Base.OpenRecordset("Select Porcentaje_IVA" _
        & " from Servicios" _
        & " Where Id=" & Val(Me.ExpedienteTxt.Tag))
    If MiRc.EOF = False Then
        MiRc.MoveFirst
        If IsNull(MiRc!Porcentaje_IVA) = False Then IndiceIva = MiRc!Porcentaje_IVA
    End If
    MiRc.Close
End If

'Estableciendo valor de iva en tanto por 100
PorcentajeIVAArticuloDefecto = 0

If 1 <= IndiceIva And IndiceIva <= 15 Then
    Set RsIVA = Base.OpenRecordset("SELECT IVA_" & IndiceIva & " AS Iva FROM Tipos_IVA")
    
    If Not RsIVA.EOF Then
        If Not IsNull(RsIVA.Fields(0).value) Then
            PorcentajeIVAArticuloDefecto = Format(RsIVA.Fields(0).value, FormatoDecimalStandar)
        End If
    End If
    
    Call RsIVA.Close
End If
    
End Function
```