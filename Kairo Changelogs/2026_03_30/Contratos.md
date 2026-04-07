# Base de datos
Nueva tabla `Contratos_Proyecto`

```vb
Call CrearNuevaTabla("Contratos_Proyecto", "Id", True, TipoAutonumerico, TamanoTipoCampo(TipoAutonumerico), "0", Base)
Call CrearCampo("Numero", "Contratos_Proyecto", TipoTexto, TamanoTipoCampo(TipoTexto), "", Base, True)
Call CrearCampo("Descripcion", "Contratos_Proyecto", TipoTexto, TamanoTipoCampo(TipoTexto), "", Base, True)
Call CrearCampo("Estado", "Contratos_Proyecto", TipoByte, TamanoTipoCampo(TipoByte), "", Base, True)
Call CrearCampo("Ficha_Id", "Contratos_Proyecto", TipoEnteroLargo, TamanoTipoCampo(TipoEnteroLargo), "", Base)
Call CrearCampo("Fecha_Inicio", "Contratos_Proyecto", TipoFecha, TamanoTipoCampo(TipoFecha), "", Base)
Call CrearCampo("Fecha_Fin", "Contratos_Proyecto", TipoFecha, TamanoTipoCampo(TipoFecha), "", Base)
Call CrearCampo("Fecha_Firma", "Contratos_Proyecto", TipoFecha, TamanoTipoCampo(TipoFecha), "", Base)
Call CrearCampo("Importe_Total", "Contratos_Proyecto", TipoDouble, TamanoTipoCampo(TipoDouble), "", Base)
Call CrearCampo("Porcentaje_Gastos_Generales", "Contratos_Proyecto", TipoDouble, TamanoTipoCampo(TipoDouble), "", Base)
Call CrearCampo("Porcentaje_Beneficios_Industriales", "Contratos_Proyecto", TipoDouble, TamanoTipoCampo(TipoDouble), "", Base)
Call CrearCampo("Porcentaje_Deducciones", "Contratos_Proyecto", TipoDouble, TamanoTipoCampo(TipoDouble), "", Base)
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

## `ContratosBtn_Click`
Nuevo método
```vb
Private Sub ContratosBtn_Click()
    Call ContratosModalFrm.Show
    ContratosModalFrm.Left = Me.Left + (Me.Width - ContratosModalFrm.Width) / 2
    ContratosModalFrm.Top = Me.Top + (Me.Height - ContratosModalFrm.Height) / 2
    Call ContratosModalFrm.ZOrder(0)
End Sub
```

## `SalirBtn_Click`

Añadir al final
```vb
Call ContratosModalFrm.Hide
```

# `ContratosModalFrm`
Archivo nuevo

cobro
detalletarjetaregalo