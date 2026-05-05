## Cambios en la base de datos
```vb
Call CrearNuevaTabla("Rutas", "Id", True, TipoAutonumerico, TamanoTipoCampo(TipoAutonumerico), "0", Base)
Call CrearCampo("Codigo", "Rutas", TipoTexto, TamanoTipoCampo(TipoTexto), "0", Base)
Call CrearCampo("Descripcion", "Rutas", TipoTexto, TamanoTipoCampo(TipoTexto), "", Base, False)
```

## `ConfiguradorRutasModalFrm`

Formulario nuevo
## `ConfiguracionPedidosCliente`
Formulario nuevo