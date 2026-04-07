# Base de datos
Nueva tabla `Contratos_Proyecto`

```vb
Call CrearNuevaTabla("Contratos_Proyecto", "Id", True, TipoAutonumerico, TamanoTipoCampo(TipoAutonumerico), "0", Base)
Call CrearCampo("Numero", "Contratos_Proyecto", TipoTexto, TamanoTipoCampo(TipoTexto), "", Base, True)
Call CrearCampo("Descripcion", "Contratos_Proyecto", TipoTexto, TamanoTipoCampo(TipoTexto), "", Base, True)
Call CrearCampo("Ficha_Id", "Contratos_Proyecto", TipoEnteroLargo, TamanoTipoCampo(TipoEnteroLargo), "", Base)
Call CrearCampo("FechaInicio", "Contratos_Proyecto", TipoFecha, TamanoTipoCampo(TipoFecha), "", Base)
Call CrearCampo("FechaFin", "Contratos_Proyecto", TipoFecha, TamanoTipoCampo(TipoFecha), "", Base)
Call CrearCampo("FechaFirma", "Contratos_Proyecto", TipoFecha, TamanoTipoCampo(TipoFecha), "", Base)
Call CrearCampo("ImporteTotal", "Contratos_Proyecto", TipoDouble, TamanoTipoCampo(TipoDouble), "", Base)
Call CrearCampo("PorcentajeGastosGenerales", "Contratos_Proyecto", TipoDouble, TamanoTipoCampo(TipoDouble), "", Base)
Call CrearCampo("PorcentajeBeneficiosIndustriales", "Contratos_Proyecto", TipoDouble, TamanoTipoCampo(TipoDouble), "", Base)
Call CrearCampo("PorcentajeDeducciones", "Contratos_Proyecto", TipoDouble, TamanoTipoCampo(TipoDouble), "", Base)
```

# `ResumenEconomicoProyectosObraFrm.frm`
## Controles
Añadir un botón `ContratosBtn` a la izquierda de `VerDetalleBtn` (a lo mejor conviene cambiar el ícono).

## `Form_Load`
Debajo de 
```vb
Me.VerDetalleBtn.ShowFocusRect = Me.SalirBtn.ShowFocusRect
```

Añadir
```vb
Me.ContratosBtn.BackOver = Me.SalirBtn.BackOver
Me.ContratosBtn.ButtonType = Me.SalirBtn.ButtonType
Me.ContratosBtn.ShowFocusRect = Me.SalirBtn.ShowFocusRect
```

## `Form_Resize`
Debajo de 
```vb
Me.VerDetalleBtn.Left = Me.Line2.X1 - Me.VerDetalleBtn.Width - 50
```

Añadir
```vb
Me.ContratosBtn.Top = Me.SalirBtn.Top
Me.ContratosBtn.Left = Me.VerDetalleBtn.Left - Me.ContratosBtn.Width - 50
```

Debajo de
```vb
On Error Resume Next
Me.Shape1(Me.OrdenListadoCombo.index).Visible = Me.OrdenListadoCombo.Visible
On Error GoTo 0
```

Añadir
```vb
Call ContratosModalFrm.Resize
```

## `ContratosBtn_Click`
Nuevo método

## `SalirBtn_Click`

Añadir al final
```vb
Call ContratosModalFrm.Hide
```

# `ContratosModalFrm`
Archivo nuevo

cobro
detalletarjetaregalo