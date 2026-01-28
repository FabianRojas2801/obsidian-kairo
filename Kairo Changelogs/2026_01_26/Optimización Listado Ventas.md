Lunes 26/01/2026 - Martes 27/01/2026. Descartado en la rama `optimizar_autocompletado_ventas`.

## `DocumentosVentaFrm`

Añadir
```vb
Private Type TFiltrosArticulos
    Origen As Byte
    Referencia As String
    Descripcion As String
    TotalPaginas As Long
    Pagina As Long
End Type

Private FiltrosArticulos As TFiltrosArticulos
```

**Borrar**
```vb
Dim PageSize As Integer
Dim CurrentPage As Integer
Dim TotalPages As Integer
Dim FilasRegistros As Recordset
Dim CantidadRegistros As Long
```
### `CargarSelectorArticulos`
Nuevo método, bajo `IdCuentaPredeterminada`.
### `SiguientePaginaArticulo`
Reemplazar por
```vb
Private Sub SiguientePaginaArticulo()
    If FiltrosArticulos.Pagina < FiltrosArticulos.TotalPaginas Then
        FiltrosArticulos.Pagina = FiltrosArticulos.Pagina + 1
        Call CargarSelectorArticulos(TipoDetalleDocumento = 1)
    End If
End Sub
```

### `AnteriorPaginaArticulo`
Reemplazar por
```vb
Private Sub AnteriorPaginaArticulo()
    If FiltrosArticulos.Pagina > 1 Then
        FiltrosArticulos.Pagina = FiltrosArticulos.Pagina - 1
        Call CargarSelectorArticulos(TipoDetalleDocumento = 1)
    End If
End Sub
```

### `CargarListadoReferenciasArticulos`

Reemplazar por
```vb
Private Sub CargarListadoReferenciasArticulos(ByVal Referencia As String, ByVal Descripcion As String, ByVal Origen As Byte)
    FiltrosArticulos.Referencia = Referencia
    FiltrosArticulos.Descripcion = Descripcion
    FiltrosArticulos.Origen = Origen
    FiltrosArticulos.Pagina = 1
    
    Call CargarSelectorArticulos(False)
    
    Me.FrameSelectorArticulos.Top = (Me.Height - Me.FrameSelectorArticulos.Height) / 2
    Me.FrameSelectorArticulos.Left = (Me.Width - Me.FrameSelectorArticulos.Width) / 2

    Me.FrameSelectorArticulos.ZOrder 0
    Call CargaConfiguracionUsuario(1)
    Me.FrameSelectorArticulos.Visible = True
    Me.KeyPreview = False

    Me.GridSelectorArticulos.Refresh
    DoEvents
    
    Call PasarFoco(Me.GridSelectorArticulos)
End Sub
```
### `CargarListadoReferenciasArticulosSimplificada`

Reemplazar por
```vb
Private Sub CargarListadoReferenciasArticulosSimplificada(ByVal Referencia As String, ByVal Descripcion As String, ByVal Origen As Byte)
    FiltrosArticulos.Referencia = Referencia
    FiltrosArticulos.Descripcion = Descripcion
    FiltrosArticulos.Origen = Origen
    FiltrosArticulos.Pagina = 1
    
    Call CargarSelectorArticulos(True)
    
    Me.FrameSelectorArticulos.Top = (Me.Height - Me.FrameSelectorArticulos.Height) / 2
    Me.FrameSelectorArticulos.Left = (Me.Width - Me.FrameSelectorArticulos.Width) / 2

    Me.FrameSelectorArticulos.ZOrder 0
    Call CargaConfiguracionUsuario(1)
    Me.FrameSelectorArticulos.Visible = True
    Me.KeyPreview = False

    Me.GridSelectorArticulos.Refresh
    DoEvents
    
    Call PasarFoco(Me.GridSelectorArticulos)
End Sub
```

### `LoadPage`
**Eliminar método.**

### `LoadPageSimplificada`
**Eliminar método.**



