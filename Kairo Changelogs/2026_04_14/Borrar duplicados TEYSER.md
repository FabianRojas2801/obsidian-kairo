## SQL SERVER (Kong1App)
### Ver duplicados (anotar ID usuario y fecha)
```sql
SELECT CA_dup.id AS id_duplicado, CA_keep.id AS id_conservar, 
       CA_dup.id_usuario,
       U.nomapes, CA_dup.fecha, 
       CA_dup.hora_entrada AS entrada_dup, 
       CA_keep.hora_entrada AS entrada_conservar,
       DATEDIFF(MINUTE, CA_dup.hora_entrada, CA_keep.hora_entrada) AS diff_minutos
FROM CONTROL_ASISTENCIAS CA_dup
INNER JOIN USUARIOS U ON CA_dup.id_usuario = U.id
INNER JOIN (
    SELECT CA2.id_usuario, CA2.fecha, MAX(CA2.id) AS max_id
    FROM CONTROL_ASISTENCIAS CA2
    INNER JOIN USUARIOS U2 ON CA2.id_usuario = U2.id
    WHERE U2.id_empresa = 16
    GROUP BY CA2.id_usuario, CA2.fecha
) grp ON CA_dup.id_usuario = grp.id_usuario AND CA_dup.fecha = grp.fecha
INNER JOIN CONTROL_ASISTENCIAS CA_keep ON CA_keep.id = grp.max_id
WHERE U.id_empresa = 16
AND CA_dup.id <> grp.max_id
AND ABS(DATEDIFF(MINUTE, CA_dup.hora_entrada, CA_keep.hora_entrada)) < 5
ORDER BY CA_dup.id_usuario, CA_dup.fecha;
```

### Confirmar duplicados (anotar IDs)
```sql
SELECT *
FROM CONTROL_ASISTENCIAS
WHERE id_usuario = <ID USUARIO> AND fecha = '<FECHA>';
```

### Borrar duplicados
```sql
DELETE FROM CONTROL_ASISTENCIAS
WHERE id IN (<IDs DUPLICADAS>);
```

## MS Access


