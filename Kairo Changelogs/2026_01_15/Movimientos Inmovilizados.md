Jueves 15/01/2026

## `VehiculosFrm.frm`
## Archivo
Abrir con Notepad++, buscar `LockControls` y ponerlo en `0 'False`

- - -
## Controles
- Añadir `TabNuevoMovimiento` y sus hijos.
- Añadir `NuevoMovimientoBtn`

- - -
### `Form_Resize`
*Debajo* de 
```vb
Me.ModificarVehiculoBtn.Left = Me.NuevoVehiculoBtn.Left - Me.NuevoVehiculoBtn.Width - 50
```
Añadir
```vb
Me.NuevoMovimientoBtn.Top = Me.SalirBtn.Top
Me.NuevoMovimientoBtn.Left = Me.ModificarVehiculoBtn.Left - Me.NuevoMovimientoBtn.Width - 50
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

## `Form_Resize`
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