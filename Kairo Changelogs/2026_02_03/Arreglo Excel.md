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
Dim DejarExcelAbierto As Boolean

DejarExcelAbierto = False
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

**Reemplazar** el siguiente código
```vb
xLibro.Close True
ObjExcel.Quit
```

Por
```vb
If Not DejarExcelAbierto Then
    Call CerrarExcelSeguro(ObjExcel, xLibro, Hoja, PidExcel)
End If
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

- - -
Debajo de 
```vb
ZExe = ShellExecute(Control.hWnd, "Open", NombreFicheroDestino, "", "", 1)
```

Añade
```vb
DejarExcelAbierto = True
```
****