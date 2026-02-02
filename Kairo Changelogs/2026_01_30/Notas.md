
El campo **Contrato** corresponde a la tabla `Servicios.Texto_Contrato_Base` y es un **Texto Corto**
Los datos se cargan en `CargaAOrigen`.


```vb
TextoOrden = ""
Select Case Me.OrdenListadoCombo.ListIndex
    Case 0  'Proyecto
        TextoOrden = " order by Codigo_Proyecto asc, Descripcion asc"
    Case 1  'Cliente
        TextoOrden = " order by Nombre_Ficha asc, Descripcion asc"
    Case 2  'Responsable Técnico
        TextoOrden = " order by Responsable_Tecnico_Id asc, Descripcion asc"
    Case 3  'Técnico
        TextoOrden = " order by Tecnico_Id asc, Descripcion asc"
    Case Else
        TextoOrden = " order by Descripcion asc"
End Select
```

## Mapeo General
**0** -> `EProyectoGeneralCol.GeneralId`
**1** -> `EProyectoGeneralCol.GeneralEmpresa`
**2** -> `EProyectoGeneralCol.GeneralEstado`
**3** -> `EProyectoGeneralCol.GeneralFechaInicio`
**4** -> `EProyectoGeneralCol.GeneralFechaFin`
**5** -> `EProyectoGeneralCol.GeneralCodigo`
**(nuevo)** -> `EProyectoGeneralCol.GeneralContrato`
**6** -> `EProyectoGeneralCol.GeneralProyecto`
**7** -> `EProyectoGeneralCol.GeneralCliente`
**8** -> `EProyectoGeneralCol.GeneralPresupuesto`
**9** -> `EProyectoGeneralCol.GeneralCertPrevista`
**10** -> `EProyectoGeneralCol.GeneralCertReal`
**11** -> `EProyectoGeneralCol.GeneralFacturado`
**12** -> `EProyectoGeneralCol.GeneralDesvCertPorcentaje`
**13** -> `EProyectoGeneralCol.GeneralCosteReal`
**14** -> `EProyectoGeneralCol.GeneralResultadoImporte`
**15** -> `EProyectoGeneralCol.GeneralResultadoPorcentaje`
**16** -> `EProyectoGeneralCol.GeneralGastosGenerales`
**17** -> `EProyectoGeneralCol.GeneralGastosGeneralesPorcentaje`
**18** -> `EProyectoGeneralCol.GeneralBeneficioReal`
**19** -> `EProyectoGeneralCol.GeneralBeneficioRealPorcentaje`

## Mapeo Técnico
**0** -> `EProyectoTecnicoCol.TecnicoId`
**1** -> `EProyectoTecnicoCol.TecnicoEmpresa`
**2** -> `EProyectoTecnicoCol.TecnicoCodigo`
**3** -> `EProyectoTecnicoCol.TecnicoPresupuesto`
**4** -> `EProyectoTecnicoCol.TecnicoOT`
**(nuevo)** -> `EProyectoTecnicoCol.Contrato`
**5** -> `EProyectoTecnicoCol.TecnicoProyecto)`
**6** -> `EProyectoTecnicoCol.TecnicoFechaPedido`
**7** -> `EProyectoTecnicoCol.TecnicoPedidoNumero`
**8** -> `EProyectoTecnicoCol.TecnicoEstado`
**9** -> `EProyectoTecnicoCol.TecnicoCliente`
**10** -> `EProyectoTecnicoCol.TecnicoPeticionario`
**11** -> `EProyectoTecnicoCol.TecnicoJefeCompras`
**12** -> `EProyectoTecnicoCol.TecnicoImporte`
**13** -> `EProyectoTecnicoCol.TecnicoFechaEntrega`
**14** -> `EProyectoTecnicoCol.TecnicoFechaEntregaPresupuesto`
**15** -> `EProyectoTecnicoCol.TecnicoTecnico`
**16** -> `EProyectoTecnicoCol.TecnicoNumeroFactura`
**17** -> `EProyectoTecnicoCol.TecnicoFechaCobroFactura`
**18** -> `EProyectoTecnicoCol.TecnicoCobrado`
**19** -> `EProyectoTecnicoCol.TecnicoPedidoCliente`
**20** -> `EProyectoTecnicoCol.TecnicoContacto`
**21** -> `EProyectoTecnicoCol.TecnicoCosteManoObraPresupuestada`
**22** -> `EProyectoTecnicoCol.TecnicoCosteHorasIngenieriaPresupuestada`
**23** -> `EProyectoTecnicoCol.TecnicoCosteMateriales`
**24** -> `EProyectoTecnicoCol.TecnicoHorasManoObra`
**25** -> `EProyectoTecnicoCol.TecnicoHorasConsumidas`
**26** -> `EProyectoTecnicoCol.TecnicoHorasIngenieria`
**27** -> `EProyectoTecnicoCol.TecnicoTotalCostePresupuestado`
**28** -> `EProyectoTecnicoCol.TecnicoFechaCambioEstado`
**29** -> `EProyectoTecnicoCol.TecnicoTipoOT`

## Wtf
**`TecnicoId`** -> t1
**`TecnicoEmpresa`** -> `CodigoEmpresa`
**`TecnicoCodigo`** -> t2
**`TecnicoPresupuesto`** -> T28
**`TecnicoOT`** -> T3
**`TecnicoProyecto`** -> T4
**`TecnicoFechaPedido`** -> T29
**`TecnicoPedidoNumero`** -> T30
**`TecnicoEstado`** -> T6
**`TecnicoCliente`** -> T7
**`TecnicoPeticionario`** -> T8
**`TecnicoJefeCompras`** -> T9
**`TecnicoImporte`** -> T10
**`TecnicoFechaEntrega`** -> T11
**`TecnicoFechaEntregaPresupuesto`** -> T12
**`TecnicoTecnico`** -> T13
**`TecnicoNumeroFactura`** -> T14
**`TecnicoFechaCobroFactura`** -> T15
**`TecnicoCobrado`** -> T16
**`TecnicoPedidoCliente`** -> T17
**`TecnicoContacto`** -> T18
**`TecnicoCosteManoObraPresupuestada`** -> T19
**`TecnicoCosteHorasIngenieriaPresupuestada`** -> T20
**`TecnicoCosteMateriales`** -> T21
**`TecnicoHorasManoObra`** -> T22
**`TecnicoHorasConsumidas`** -> T23
**`TecnicoHorasIngenieria`** -> T24
**`TecnicoTotalCostePresupuestado`** -> T25
**`TecnicoFechaCambioEstado`** -> T26
**`TecnicoTipoOT`** -> T27

