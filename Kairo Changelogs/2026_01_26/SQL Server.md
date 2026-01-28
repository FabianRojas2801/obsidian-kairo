Lunes 26/01/2026. Cambios descartados a la branch `mssqlsv`.

## `DatabaseWrapper.cls`
Archivo nuevo, envuelve a `DAO.Database` para añadir la opción `dbSeeChanges` a **todos** los `Base.OpenRecordset`.

## `ConfiguracionBaseDeDatos.bas`

> [!danger] TODO
> - Hay que refactorizar el archivo entero, la lógica actual de creación del esquema no es compatible con Microsoft SQL Server.
> - Hay que lanzar un error explicito si el ODBC del `.mdb` no se encuentra en el ordenador.
> 

Reemplazar
```vb
Public Base As Database
```
por 
```vb
Public Base As DatabaseWrapper
```

> [!warning] Aviso
> Este cambio causa *Procedimiento demasiado largo* y hay que mover líneas hasta que compile

### `CrearNuevaTabla`
En el encabezado, reemplazar el tipo de `BaseTrabajo` a
```vb
BaseTrabajo As Object
```

### `CrearCampo`
En el encabezado, reemplazar el tipo de `BaseTrabajo` a
```vb
BaseTrabajo As Object
```

## `EntornoFrm`
### `MDIForm_Unload`
Reemplazar
```vb
Base.Close
```
Por
```vb
Base.CloseDB
```

### `ProcesoLanzamientoAplicacion`
Reemplazar
```vb
Set Base = EspacioDeTrabajo.OpenDatabase(BaseDeDatos, False, False, ";pwd=3n+z}%6mmi2kfh2ft5\:")
```

Por
```vb
Set Base = New DatabaseWrapper
Call Base.Init(EspacioDeTrabajo.OpenDatabase(BaseDeDatos, False, False, ";pwd=3n+z}%6mmi2kfh2ft5\:"))
```
