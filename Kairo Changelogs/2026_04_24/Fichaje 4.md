Viernes 24/04/2026

## `Daemon/ProcesosPrincipales`

### `CargarPartesTrabajo`
Debajo de
```vb
    If Not MiRc.EOF Then
			'... (código omitido)
			
            MiRc2.Update
            MiRc.Update
            MiRc2.Close
            MiRc.MoveNext
        Wend
```

Añadir
```vb
DBEngine.SetOption dbLockRetry, 20
```

### `CargarPicajeAPP`
Debajo de
```vb
If Not MiRc.EOF Then

        MiRc.MoveFirst
```

Añadir
```vb
DBEngine.SetOption dbLockRetry, 0
```

- - -
Debajo de
```vb
        While Not MiRc.EOF

            IdPicaje = 0

            Set MiRc2 = Base.OpenRecordset("SELECT * FROM Picaje_Usuarios WHERE Id=" & IIf(IsNull(MiRc!id_origen), 0, MiRc!id_origen))

            '... (código omitido)

            ' Grabando ID Origen
            MiRc!id_origen = IdPicaje
            MiRc.UpdateBatch

            ' Siguiente parte de trabajo
            MiRc.MoveNext
        Wend
```

Añadir
```vb
DBEngine.SetOption dbLockRetry, 20
```

## `Procesos_App`
### `CargarPicajeAPP`

Remplazar
```vb
Set MiRc2 = Base.OpenRecordset("Picaje_Usuarios", dbOpenDynaset, dbSeeChanges)
```

Por
```vb
DBEngine.SetOption dbLockRetry, 0
Set MiRc2 = Base.OpenRecordset("Picaje_Usuarios", dbOpenDynaset, dbSeeChanges, dbOptimistic)
```

- - -
Debajo de
```vb
    Do While Not MiRc.EOF
 
        IdPicaje = 0
```

Añadir
```vb
IniciarProceso:
        On Error GoTo ErrorProceso
```

- - -
Debajo de
```vb
        If Not IsNull(MiRc!localizacion_salida) Then
            MiRc2!Ubicacion_Salida = MiRc!localizacion_salida
        End If
 
        MiRc2.Update
```

Añadir
```vb
On Error GoTo 0
```

- - -
Debajo de
```vb
        'Obtener ID si es nuevo
        If IdPicaje = 0 Then
            MiRc2.Bookmark = MiRc2.LastModified
            IdPicaje = MiRc2!Id
        End If
```

Añadir
```vb
        GoTo FinProceso
ErrorProceso:
        On Error Resume Next
        MiRc2.CancelUpdate
        On Error GoTo 0
        Resume SiguienteRegistro
FinProceso:
```

- - -
Debajo de
```vb
        'Actualizar SQL
        MiRc!id_origen = IdPicaje
 
        If Not IsNull(MiRc!Hora_salida) Then
            MiRc!estado_traspaso = 1
        Else
            MiRc!estado_traspaso = 0
        End If
 
        MiRc.Update
```

Añadir
```vb
SiguienteRegistro:
```

- - -
Debajo de 
```vb
        MiRc.MoveNext
 
    Loop
 
    MiRc2.Close
```

Añadir
```vb
DBEngine.SetOption dbLockRetry, 20
```