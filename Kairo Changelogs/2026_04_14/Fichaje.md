Martes 14/04/2026 - Miércoles 15/04/2026

## Resumen
- Modificar fichaje perdía el ID del trabajador
- Fichajes abiertos y modificados perdían sus modificaciones al actualizar datos de la app
- Los ajustes de parada obligatoria no se sincronizaban con SQL Server

## `Daemon/ProcesosPrincipales.bas`

### `CargarPicajeAPP`
Debajo de 
```vb
Dim idsEmpresaSql As String
```

Añadir
```vb
Dim EsNuevo As Boolean
```

- - -
Remplazar
```vb
            If MiRc2.EOF Then
                MiRc2.AddNew
                IdPicaje = MiRc2!id
            Else
                MiRc2.Edit
                IdPicaje = MiRc2!id
            End If
```

Por
```vb
            If MiRc2.EOF Then
                MiRc2.AddNew
                IdPicaje = MiRc2!id
                EsNuevo = True
            Else
                MiRc2.Edit
                IdPicaje = MiRc2!id
                EsNuevo = False
            End If
```
- - -
Remplazar
```vb
            MiRc2!usuario_id = MiRc!UsuarioIdOrigen
            MiRc2!Fecha_Entrada = MiRc!Hora_entrada
            MiRc2!Fecha_Salida = 0

            If IsNull(MiRc!Hora_salida) Then
                fecha = 0
                Horas = 0
            Else
                MiRc2!Fecha_Salida = MiRc!Hora_salida
                Horas = DateDiff("n", MiRc2!Fecha_Entrada, MiRc2!Fecha_Salida) / 60
            End If

            MiRc2!Horas = Horas
            If Not IsNull(MiRc!localizacion_entrada) And Trim(MiRc!localizacion_entrada) <> "" Then MiRc2!Ubicacion_Entrada = MiRc!localizacion_entrada
            If Not IsNull(MiRc!localizacion_salida) And Trim(MiRc!localizacion_salida) <> "" Then MiRc2!Ubicacion_Salida = MiRc!localizacion_salida
```

Por
```vb
            MiRc2!usuario_id = MiRc!UsuarioIdOrigen
            MiRc2!Fecha_Entrada = MiRc!Hora_entrada
            If IsNull(MiRc!Hora_salida) Then
                If EsNuevo Then
                    MiRc2!Fecha_Salida = 0
                    MiRc2!Horas = 0
                End If
                ' Si es edición: preservar Fecha_Salida y Horas introducidos manualmente
            Else
                MiRc2!Fecha_Salida = MiRc!Hora_salida
                Horas = DateDiff("n", MiRc2!Fecha_Entrada, MiRc2!Fecha_Salida) / 60
                MiRc2!Horas = Horas
            End If
            If Not IsNull(MiRc!localizacion_entrada) And Trim(MiRc!localizacion_entrada) <> "" Then MiRc2!Ubicacion_Entrada = MiRc!localizacion_entrada
            If Not IsNull(MiRc!localizacion_salida) And Trim(MiRc!localizacion_salida) <> "" Then MiRc2!Ubicacion_Salida = MiRc!localizacion_salida
```

## `ListadoPicajeFrm.frm`
### `capturaUsuarioApp`
Remplazar
```vb
    ' --- 1. Obtener empresa_id desde Access (Ficha_Personal) ---
    Set rsAccess = Base.OpenRecordset( _
        "SELECT empresa_id " & _
        "FROM Ficha_Personal " & _
        "WHERE usuario_id = " & Usuario_Id)
```

Por
```vb
    ' --- 1. Obtener empresa_id desde Access (Ficha_Personal) ---
    Set rsAccess = Base.OpenRecordset( _
        "SELECT empresa_id " & _
        "FROM Ficha_Personal " & _
        "WHERE Id = " & Usuario_Id)
```

(Se cambia `usuario_id` por `Id`)

### `ComprobandoParadaObligatoria`
Debajo de
```vb
            'Caso 1) En este caso se divide en dos el picaje para respetar la parada obligatoria
            FechaAuxiliar = 0
            FechaAuxiliar = MiRc!Fecha_Salida
            
            'Modificando picaje para ajustar a salida obligatoria
            MiRc.Edit
            MiRc!Fecha_Salida = Format(MiRc!Fecha_Salida, "dd/mm/yyyy") & " " & Format(InicioParada, "hh:mm")
            MiRc!Horas = DateDiff("n", MiRc!fecha_entrada, MiRc!Fecha_Salida) / 60
            MiRc.Update
```
Añadir
```vb
If BaseAPPActiva Then guardarPicajeApp(MiRc!Id)
```
- - -
Debajo de
```vb
            'Caso 3) En este caso se ajusta la hora de entrada para que conincida con la parada obligatoria
            MiRc.Edit
            MiRc!fecha_entrada = Format(MiRc!fecha_entrada, "dd/mm/yyyy") & " " & Format(FinParada, "hh:mm")
            MiRc!Horas = DateDiff("n", MiRc!fecha_entrada, MiRc!Fecha_Salida) / 60
            MiRc.Update
```

Añadir
```vb
If BaseAPPActiva Then guardarPicajeApp(MiRc!Id)
```
- - -
Debajo de
```vb
'Caso 4) En este caso el fichaje se borra directamente (está dentro de la parada obligatoria de la empresa
```
Añadir
```vb
If BaseAPPActiva Then borrarPicajeApp(MiRc!Id)
```