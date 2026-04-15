Martes 14/04/2026 - Miércoles 15/04/2026

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

