Jueves 15/01/2026

## `ConfiguraciónBaseDeDatos.bas`
Añadir
```vb
Call CrearNuevaTabla("Movimientos_Inmovilizado", "Id", True, TipoCampo.TipoAutonumerico, TamanoTipoCampo(TipoCampo.TipoAutonumerico), "0", Base)

Call CrearCampo("Fecha", "Movimientos_Inmovilizado", TipoCampo.TipoFecha, TamanoTipoCampo(TipoCampo.TipoFecha), "", Base, False)
Call CrearCampo("Tipo_Ubicacion", "Movimientos_Inmovilizado", TipoCampo.TipoByte, TamanoTipoCampo(TipoCampo.TipoByte), "0", Base, False)
Call CrearCampo("Ubicacion_Id", "Movimientos_Inmovilizado", TipoCampo.TipoEnteroLargo, TamanoTipoCampo(TipoCampo.TipoEnteroLargo), "", Base, False)
Call CrearCampo("Trabajador_Id", "Movimientos_Inmovilizado", TipoCampo.TipoEnteroLargo, TamanoTipoCampo(TipoCampo.TipoEnteroLargo), "", Base, False)
Call CrearCampo("Observaciones", "Movimientos_Inmovilizado", TipoCampo.TipoMemo, TamanoTipoCampo(TipoCampo.TipoMemo), "", Base, True)
```

- - -
## `VehiculosFrm.frm`

**Demasiados en código cambios para listar**

- Añadido un nuevo tab para los movimientos
- Añadido una columna a los inmovilizados para mostrar su última ubicación
- Escondido el botón de "Editar" por que no hace literalmente nada
- Cerrar el Tab de abajo cuando se selecciona la fila de abajo de todo. Esto se debe a que existe la posibilidad de borrar los datos de *todos* los inmovilizados si se interactúa con el tab teniendo seleccionado la fila de "nada".

