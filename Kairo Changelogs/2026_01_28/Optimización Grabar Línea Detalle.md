Miércoles 28/01/2026. ==Pendiente pruebas en Kedeke==, rama `optimización_ventas`.
## `QueryPerformanceCounter.bas`
Nuevo archivo, se usa para hacer un benchmark de funciones especificas.

## `DocumentoVentaFrm.frm`

### `GrabarLineaDetalle`


## Comparación
Antes

| Método             | Paso ID | Paso                      | Duración   | Extra                                                                    |
| ------------------ | ------- | ------------------------- | ---------- | ------------------------------------------------------------------------ |
| GrabarLineaDetalle | 00      | START                     | 0,00 ms    | 20260128-114944;IdDocumento=46;Fila=6;TipoDocumento=3;IdLineaModificar=0 |
| GrabarLineaDetalle | 10      | OpenRecordset+AddNew/Edit | 35,18 ms   | 20260128-114944;IdDocumento=46;Fila=6;TipoDocumento=3;IdLineaModificar=0 |
| GrabarLineaDetalle | 20      | AsignacionCampos          | 330,56 ms  | 20260128-114944;IdDocumento=46;Fila=6;TipoDocumento=3;IdLineaModificar=0 |
| GrabarLineaDetalle | 30      | Update+CloseDetalle       | 86,21 ms   | 20260128-114944;IdDocumento=46;Fila=6;TipoDocumento=3;IdLineaModificar=0 |
| GrabarLineaDetalle | 40      | Certificaciones           | 0,00 ms    | 20260128-114944;IdDocumento=46;Fila=6;TipoDocumento=3;IdLineaModificar=0 |
| GrabarLineaDetalle | 50      | PieFacturaIVA             | 130,55 ms  | 20260128-114944;IdDocumento=46;Fila=6;TipoDocumento=3;IdLineaModificar=0 |
| GrabarLineaDetalle | 60      | UpdateCabecera            | 2,96 ms    | 20260128-114944;IdDocumento=46;Fila=6;TipoDocumento=3;IdLineaModificar=0 |
| GrabarLineaDetalle | 99      | TOTAL                     | 1011,99 ms | 20260128-114944;IdDocumento=46;Fila=6;TipoDocumento=3;IdLineaModificar=0 |
Después

| Método             | Paso ID | Paso                      | Duración   | Extra                                                                     |
| ------------------ | ------- | ------------------------- | ---------- | ------------------------------------------------------------------------- |
| GrabarLineaDetalle | 00      | START                     | 0,00 ms    | 20260128-124124;IdDocumento=46;Fila=20;TipoDocumento=3;IdLineaModificar=0 |
| GrabarLineaDetalle | 10      | OpenRecordset+AddNew/Edit | 47,48 ms   | 20260128-124124;IdDocumento=46;Fila=20;TipoDocumento=3;IdLineaModificar=0 |
| GrabarLineaDetalle | 20      | AsignacionCampos          | 314,67 ms  | 20260128-124124;IdDocumento=46;Fila=20;TipoDocumento=3;IdLineaModificar=0 |
| GrabarLineaDetalle | 30      | Update+CloseDetalle       | 70,78 ms   | 20260128-124124;IdDocumento=46;Fila=20;TipoDocumento=3;IdLineaModificar=0 |
| GrabarLineaDetalle | 40      | Certificaciones           | 0,00 ms    | 20260128-124124;IdDocumento=46;Fila=20;TipoDocumento=3;IdLineaModificar=0 |
| GrabarLineaDetalle | 50      | PieFacturaIVA             | 234,14 ms  | 20260128-124124;IdDocumento=46;Fila=20;TipoDocumento=3;IdLineaModificar=0 |
| GrabarLineaDetalle | 60      | UpdateCabecera            | 2,81 ms    | 20260128-124124;IdDocumento=46;Fila=20;TipoDocumento=3;IdLineaModificar=0 |
| GrabarLineaDetalle | 99      | TOTAL                     | 1047,88 ms | 20260128-124124;IdDocumento=46;Fila=20;TipoDocumento=3;IdLineaModificar=0 |


