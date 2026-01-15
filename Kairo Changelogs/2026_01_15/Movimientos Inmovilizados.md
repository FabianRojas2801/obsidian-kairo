Jueves 15/01/2026

## `VehiculosFrm.frm`
### `Form_Resize`
*Debajo* de 
```vb
Me.ModificarVehiculoBtn.Left = Me.NuevoVehiculoBtn.Left - Me.NuevoVehiculoBtn.Width - 50
```
Añadir
```vb
Me.MovimientosBtn.Top = Me.SalirBtn.Top
Me.MovimientosBtn.Left = Me.ModificarVehiculoBtn.Left - Me.ModificarVehiculoBtn.Width - 50
```

### `Form_Load`
*Debajo* de
```vb
Me.ModificarVehiculoBtn.ShowFocusRect = Me.SalirBtn.ShowFocusRect
```

Añadir
```vb
Me.MovimientosBtn.BackOver = Me.SalirBtn.BackOver
Me.MovimientosBtn.ButtonType = Me.SalirBtn.ButtonType
Me.MovimientosBtn.ShowFocusRect = Me.SalirBtn.ShowFocusRect
```