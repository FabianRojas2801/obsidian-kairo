## `Daemon\Module1.bas` (`MainService`)
### `CreateBases`

Debajo de:
```vb
BaseDeDatos = BaseDeDatos & "\Datos.Mdb"
```

Añadir:
```vb
On Error GoTo ErrCreateBases
```

- - -
Debajo de _(línea incompleta)_:
```vb
Set Base = EspacioDeTrabajo.OpenDatabase(BaseDeDatos, False, False,
```

Añadir
```vb
Exit Sub

ErrCreateBases:
    MsgBox "Error al abrir base de datos." & vbCrLf & vbCrLf & _
           "App.Path: " & App.Path & vbCrLf & _
           "BaseDeDatos: " & BaseDeDatos & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Service - Error CreateBases"
```