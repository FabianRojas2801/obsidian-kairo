Martes 03/02/2026

## `ProcesosUtilidades.bas`
Nuevo archivo

## `Procesos.bas`

### `ImprimirDocumentoVentaCompra`
Debajo de
```vb
Dim Copias As Long
```

Añadir
```vb
Dim PidExcel As Long
```

- - -
Reemplazar
```vb
Dim ObjExcel As Variant
Dim xLibro As Variant
```

Por
```vb
Dim ObjExcel As Object
Dim xLibro As Object
```
- - -
Reemplazar
```vb
Dim Hoja As Variant
```

Por
```vb
Dim Hoja As Object
```
- - -
Debajo de
```vb
Set ObjExcel = CreateObject("Excel.application")
```
Añadir
```vb
PidExcel = CapturarPidExcel(ObjExcel)
```
- - -
Debajo de
```vb
'Actualizando biblioteca
Call ControlarNuevosFicherosImpresion(IdExpediente)
```

Reemplazar el siguiente código
```vb
xLibro.Close True
ObjExcel.Quit
```

Por
```vb
Call CerrarExcelSeguro(ObjExcel, xLibro, Hoja, PidExcel)
```
- - -
Arriba de
```vb
MsgBox "Se ha producido un error al imprimir el documento"
```

Añadir
```vb
Call CerrarExcelSeguro(ObjExcel, xLibro, Hoja, PidExcel)
```