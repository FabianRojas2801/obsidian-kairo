## Bug: El Daemon sobreescribe correcciones manuales de fichajes

### Contexto

El Daemon sincroniza los fichajes de la web (SQL Server, tabla `Control_Asistencias`) hacia la base de datos local (Access, tabla `Picaje_Usuarios`). Esta sincronización se ejecuta periódicamente desde `ProcesosPrincipales.bas` → `CargarPicajeAPP`.

---

### Cómo decide el Daemon qué fichajes procesar

La query de selección (línea 516) captura registros SQL que cumplan **cualquiera** de estas condiciones:

```sql
WHERE (id_origen IS NULL OR estado_traspaso = 0)
```

- `id_origen IS NULL` → nunca traspasado a Access
- `estado_traspaso = 0` → ya traspasado, pero marcado como incompleto

---

### El problema: `estado_traspaso` nunca se marca como completo en fichajes sin salida

Al final de cada iteración (líneas 557-558), el Daemon actualiza `estado_traspaso`:

```vb
If Not IsNull(MiRc!id_origen) Then MiRc!estado_traspaso = 1
If IsNull(MiRc!Hora_salida) And IsNull(MiRc!localizacion_salida) Then MiRc!estado_traspaso = 0
```

La segunda línea **sobreescribe** la primera. El resultado según el estado del fichaje en SQL:

| `Hora_salida` | `localizacion_salida` | `estado_traspaso` resultante |
|---|---|---|
| Tiene valor | (cualquiera) | 1 — completo ✓ |
| NULL | Tiene valor | 1 — completo ✓ |
| NULL | NULL | **0 — incompleto ✗** |

El último caso es el más habitual: el empleado **solo ha fichado la entrada** desde la web. En ese momento `Hora_salida` y `localizacion_salida` son NULL, por lo que el fichaje queda permanentemente con `estado_traspaso = 0`.

---

### Qué ocurre en cada ciclo posterior del Daemon

Como el registro sigue cumpliendo `estado_traspaso = 0`, el Daemon vuelve a encontrarlo. Dado que `id_origen` ya tiene valor, entra por el camino de edición:

```vb
Set MiRc2 = Base.OpenRecordset("SELECT * FROM Picaje_Usuarios WHERE Id=" & MiRc!id_origen)
' MiRc2 no es EOF → edita el registro existente
MiRc2.Edit
```

Y sobreescribe los campos del registro Access con los datos actuales del SQL:

```vb
MiRc2!Fecha_Entrada = MiRc!Hora_entrada   ' sobreescribe
MiRc2!Fecha_Salida = 0                    ' siempre se resetea a 0
MiRc2!Horas = 0                           ' porque Hora_salida es NULL en SQL
```

**Esto ocurre en cada ejecución del Daemon mientras el empleado no fiche la salida desde la web.**

---

### Escenario concreto de pérdida de datos

1. El empleado ficha **entrada** desde la web → SQL registra `Hora_entrada`, `Hora_salida = NULL`
2. El Daemon traspasa el fichaje a Access correctamente
3. Alguien en la oficina **añade manualmente la `Fecha_Salida`** en Kairo (el empleado no fichó la salida)
4. El Daemon vuelve a ejecutarse (pocos minutos después)
5. El registro SQL sigue con `estado_traspaso = 0` — no ha cambiado nada en la web
6. **El Daemon sobreescribe el registro Access** → `Fecha_Salida = 0`, `Horas = 0`
7. La corrección manual **desaparece**

---

### Síntomas que produce

- Horas trabajadas que aparecen como 0 sin motivo aparente
- Fechas de salida que desaparecen tras haberse introducido manualmente
- Días que "se borran" si el cálculo de jornada depende de `Fecha_Salida`

---

### Causa raíz

La condición de la línea 558 mezcla dos responsabilidades distintas:

1. **Saber si el fichaje está completo** (tiene salida) — para no volver a procesarlo
2. **Saber si hay datos de localización** — información secundaria

Un fichaje sin salida no necesita re-procesarse cada ciclo. Cuando el empleado fiche la salida desde la web, ese registro entrará de nuevo en la query por `id_origen IS NULL` (nuevo registro de salida) o por otro mecanismo de actualización. El re-procesado continuo del registro de entrada no aporta nada y destruye ediciones manuales.

---

### Archivos afectados

- `Daemon/ProcesosPrincipales.bas` — líneas 557-558
