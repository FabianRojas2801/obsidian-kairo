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

**Demasiados cambios para listar**

