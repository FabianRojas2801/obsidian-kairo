## Base de datos
Añadir
```vb
Call CrearNuevaTabla("Alergenos", "Id", True, TipoAutonumerico, TamanoTipoCampo(TipoAutonumerico), "0", Base)
Call CrearCampo("Orden", "Alergenos", TipoEnteroLargo, TamanoTipoCampo(TipoEnteroLargo), "0", Base)
Call CrearCampo("Descripcion", "Alergenos", TipoTexto, TamanoTipoCampo(TipoTexto), "", Base, False)

Call CrearNuevaTabla("Articulos_Alergenos", "Id", True, TipoAutonumerico, TamanoTipoCampo(TipoAutonumerico), "0", Base)
Call CrearCampo("Articulo_Id", "Articulos_Alergenos", TipoEnteroLargo, TamanoTipoCampo(TipoEnteroLargo), "0", Base)
Call CrearCampo("Alergeno_Id", "Articulos_Alergenos", TipoEnteroLargo, TamanoTipoCampo(TipoEnteroLargo), "0", Base)
Call CrearCampo("Tipo", "Articulos_Alergenos", TipoEnteroLargo, TamanoTipoCampo(TipoEnteroLargo), "0", Base)
```

## `AlergenosModalFrm`
Nuevo formulario.

**Para abrir:**
```vb
Call AlergenosModalFrm.Show
AlergenosModalFrm.Left = (EntornoFrm.ScaleWidth - AlergenosModalFrm.Width) / 2
AlergenosModalFrm.Top = (EntornoFrm.ScaleHeight - AlergenosModalFrm.Height) / 2
Call AlergenosModalFrm.ZOrder(0)
```

## `AlergenosArticulosModalFrm`
Nuevo formulario

**Para abrir:**
> [!info] Info
> Se necesita `IdArticulo` y `MostrarEditable` (Boolean)

```vb
Call AlergenosArticuloModalFrm.Cargar(IdArticulo, MostrarEditable)
    
Call AlergenosArticuloModalFrm.Show
AlergenosArticuloModalFrm.Left = (EntornoFrm.ScaleWidth - AlergenosArticuloModalFrm.Width) / 2
AlergenosArticuloModalFrm.Top = (EntornoFrm.ScaleHeight - AlergenosArticuloModalFrm.Height) / 2
Call AlergenosArticuloModalFrm.ZOrder(0)
```

