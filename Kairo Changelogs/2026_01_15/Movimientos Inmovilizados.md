Jueves 15/01/2026

## `ConfiguraciónBaseDeDatos.bas`
Añadir
```vb
Call CrearNuevaTabla("Movimientos_Inmovilizado", "Id", True, TipoCampo.TipoAutonumerico, TamanoTipoCampo(TipoCampo.TipoAutonumerico), "0", Base)
Call CrearCampo("Fecha", "Movimientos_Inmovilizado", TipoCampo.TipoFecha, TamanoTipoCampo(TipoCampo.TipoFecha), "", Base, False)
Call CrearCampo("Tipo_Ubicacion", "Movimientos_Inmovilizado", TipoCampo.TipoByte, TamanoTipoCampo(TipoCampo.TipoByte), "0", Base, False)
Call CrearCampo("Ubicacion_Id", "Movimientos_Inmovilizado", TipoCampo.TipoEnteroLargo, TamanoTipoCampo(TipoCampo.TipoEnteroLargo), "", Base, False)
Call CrearCampo("Trabajador_Id", "Movimientos_Inmovilizado", TipoCampo.TipoEnteroLargo, TamanoTipoCampo(TipoCampo.TipoEnteroLargo), "", Base, False)
```

- - -
## `VehiculosFrm.frm`
### Archivo
Abrir con Notepad++, buscar `LockControls` y ponerlo en `0 'False`

- - -
### Controles
- Añadir `TabNuevoMovimiento` y sus hijos.
- Añadir `NuevoMovimientoBtn`
- En `TabDatosInmovilizado`, cambiar el Caption de la segunda tab a `Movimientos`
- En la segunda tab de `TabDatosInmovilizado`, añadir `FrameMovimientos` y sus hijos

- - -
### `Form_Resize`

Añadir
```vb
Dim EspacioFlexibleDisponible As Long
```

- - -

Debajo de
```vb
Me.FrameDatosInmovilizado.Height = Me.TabDatosInmovilizado.Height - 600
```

Añadir
```vb
Me.FrameMovimientos.Top = 480
Me.FrameMovimientos.Left = 150
Me.FrameMovimientos.Width = Me.TabDatosInmovilizado.Width - 300
Me.FrameMovimientos.Height = Me.TabDatosInmovilizado.Height - 600
```

- - -
Debajo de 
```vb
Call ConfigurandoGridTotal
```

Añadir
```vb
' Filtros Movimientos
Me.FrameFiltrosMovimientos.Width = Me.FrameMovimientos.Width
Me.TrabajadorFiltroMovimientosTxt.Width = Me.FrameFiltrosMovimientos.Width - Me.TrabajadorFiltroMovimientosTxt.Left - Me.LabelFiltroUbicacionMovimientos.Width - Me.UbicacionFiltroMovimientosCombo.Width - 150
Me.LabelFiltroUbicacionMovimientos.Left = Me.TrabajadorFiltroMovimientosTxt.Left + Me.TrabajadorFiltroMovimientosTxt.Width
Me.UbicacionFiltroMovimientosCombo.Left = Me.LabelFiltroUbicacionMovimientos.Left + Me.LabelFiltroUbicacionMovimientos.Width

EspacioFlexibleDisponible = Me.FrameFiltrosMovimientos.Width - Me.AlmacenFiltroMovimientosTxt.Left - Me.LabelFiltroClienteMovimientos.Width - Me.LabelFiltroProyectoMovimientos.Width - 150
Me.AlmacenFiltroMovimientosTxt.Width = EspacioFlexibleDisponible / 3
Me.ClienteFiltroMovimientosTxt.Width = EspacioFlexibleDisponible / 3
Me.ProyectoFiltroMovimientosTxt.Width = EspacioFlexibleDisponible / 3

Me.LabelFiltroClienteMovimientos.Left = Me.AlmacenFiltroMovimientosTxt.Left + Me.AlmacenFiltroMovimientosTxt.Width
Me.ClienteFiltroMovimientosTxt.Left = Me.LabelFiltroClienteMovimientos.Left + Me.LabelFiltroClienteMovimientos.Width
Me.LabelFiltroProyectoMovimientos.Left = Me.ClienteFiltroMovimientosTxt.Left + Me.ClienteFiltroMovimientosTxt.Width
Me.ProyectoFiltroMovimientosTxt.Left = Me.LabelFiltroProyectoMovimientos.Left + Me.LabelFiltroProyectoMovimientos.Width
```

- - - 

### `Form_Load`
Debajo de
```vb
Me.ModificarVehiculoBtn.ShowFocusRect = Me.SalirBtn.ShowFocusRect
```

Añadir
```vb
Me.NuevoMovimientoBtn.BackOver = Me.SalirBtn.BackOver
Me.NuevoMovimientoBtn.ButtonType = Me.SalirBtn.ButtonType
Me.NuevoMovimientoBtn.ShowFocusRect = Me.SalirBtn.ShowFocusRect
Me.ActualizarMovimientosBtn.BackOver = Me.SalirBtn.BackOver
Me.ActualizarMovimientosBtn.ButtonType = Me.SalirBtn.ButtonType
Me.ActualizarMovimientosBtn.ShowFocusRect = Me.SalirBtn.ShowFocusRect
```

- - -

Reemplazar
```vb
Me.TabDatosInmovilizado.Tabs = 1
Me.TabDatosInmovilizado.TabsPerRow = 1
```
Por
```vb
Me.TabDatosInmovilizado.Tabs = 2
Me.TabDatosInmovilizado.TabsPerRow = 2
```

- - -

### `ConfigurarResalteTexto`

Nuevo método, debajo de `SalirBtn_Click`

- - -
### `CargaResaltesDeTexto`

Reemplazar por
```vb
Private Sub CargaResaltesDeTexto()
    Call ConfigurarResalteTexto(Me.OperarioTxt, Me.ShapeNuevoTrabajador) 'Edición - Codigo
    Call ConfigurarResalteTexto(Me.TextPorcentaje, Me.ShapeNuevoTrabajador) 'Edición - Porcentaje

    Call ConfigurarResalteTexto(Me.TextCodigo, Me.ShapeResaltesInmovilizado) 'Edición - Codigo
    Call ConfigurarResalteTexto(Me.TextDescripcionVehiculo, Me.ShapeResaltesInmovilizado) 'Edición - Descripcion
    Call ConfigurarResalteTexto(Me.BoxFechaAlta, Me.ShapeResaltesInmovilizado) 'Edición - Fecha Alta
    Call ConfigurarResalteTexto(Me.BoxFechaBaja, Me.ShapeResaltesInmovilizado) 'Edición - Fecha Baja
    Call ConfigurarResalteTexto(Me.TextFijo, Me.ShapeResaltesInmovilizado) 'Edición - Fijo
    
    Call ConfigurarResalteTexto(Me.TextCodigoFiltro, Me.ShapeResalteFiltros) 'Edición - Filtro Codigo
    Call ConfigurarResalteTexto(Me.TextDescripcionFiltro, Me.ShapeResalteFiltros) 'Edición - Filtro Descripcion
    Call ConfigurarResalteTexto(Me.TextTipoFiltro, Me.ShapeResalteFiltros) 'Edición - Filtro Tipo
    Call ConfigurarResalteTexto(Me.BoxFechaAltaDesde, Me.ShapeResalteFiltros) 'Edición - Filtro Fecha Alta
    Call ConfigurarResalteTexto(Me.BoxFechaAltaHasta, Me.ShapeResalteFiltros)

    Call ConfigurarResalteTexto(Me.BoxFechaBajaDesde, Me.ShapeResalteFiltros) 'Edición - Filtro Fecha Baja
    Call ConfigurarResalteTexto(Me.BoxFechaBajaHasta, Me.ShapeResalteFiltros) 'Edición - Filtro Descripcion
    Call ConfigurarResalteTexto(Me.TextTrabajadorFiltro, Me.ShapeResalteFiltros) 'Edición - Filtro Trabajador
    Call ConfigurarResalteTexto(Me.BoxFechaChequeo, Me.ShapeNuevoChequeo) 'Edición - Fecha Chequeo
    Call ConfigurarResalteTexto(Me.TextDescripcionChequeo, Me.ShapeNuevoChequeo) 'Edición - Descripcion Chequeo
    'Call ConfigurarResalteTexto(Me.EdicionMarcatxt, Me.ShapeNuevoChequeo) 'Edición - Marca
    
    Call ConfigurarResalteTexto(Me.OrdenChequeoTxt, Me.Shape42)
    Call ConfigurarResalteTexto(Me.ActividadTxt, Me.Shape42)
    
    Call ConfigurarResalteTexto(Me.BoxFechaChequeo, Me.ShapeNuevoChequeo)
    Call ConfigurarResalteTexto(Me.ComboEstado, Me.ShapeNuevoChequeo)
    Call ConfigurarResalteTexto(Me.TextDescripcionChequeo, Me.ShapeNuevoChequeo)
    
    Call ConfigurarResalteTexto(Me.FechaFiltroMovimientoDesdeTxt, Me.ShapeFiltrosMovimientos)
    Call ConfigurarResalteTexto(Me.FechaFiltroMovimientoHastaTxt, Me.ShapeFiltrosMovimientos)
    Call ConfigurarResalteTexto(Me.TrabajadorFiltroMovimientosTxt, Me.ShapeFiltrosMovimientos)
    Call ConfigurarResalteTexto(Me.AlmacenFiltroMovimientosTxt, Me.ShapeFiltrosMovimientos)
    Call ConfigurarResalteTexto(Me.ClienteFiltroMovimientosTxt, Me.ShapeFiltrosMovimientos)
    Call ConfigurarResalteTexto(Me.ProyectoFiltroMovimientosTxt, Me.ShapeFiltrosMovimientos)
    
    Call ConfigurarResalteTexto(Me.FechaNuevoMovimientoTxt, Me.ShapeNuevoMovimiento)
    Call ConfigurarResalteTexto(Me.ClienteNuevoMovimientoTxt, Me.ShapeNuevoMovimiento)
    Call ConfigurarResalteTexto(Me.ProyectoNuevoMovimientoTxt, Me.ShapeNuevoMovimiento)
    Call ConfigurarResalteTexto(Me.TrabajadorNuevoMovimiento, Me.ShapeNuevoMovimiento)
    Call ConfigurarResalteTexto(Me.ObservacionesNuevoMovimiento, Me.ShapeNuevoMovimiento)
End Sub
```

- - -

Debajo de
```vb
Me.FrameCostes.BackColor = Blanco
```

Añadir
```vb
Me.FrameMovimientos.BackColor = Blanco
Me.FrameFiltrosMovimientos.BackColor = Blanco
```