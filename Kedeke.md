Tienen **`Articulos.Stock_Actual`** como **Fecha/Hora**, hay que cambiarlo a Número antes que nada.

Añadir la columna `Stock_Actual_Num` como Número con tamaño Double.

Ejecutar SQL:
```sql
UPDATE Articulos
SET Stock_Actual_Num = CDbl([Stock_Actual])
WHERE Stock_Actual Is Not Null;
```

Borrar Stock_Actual y renombrar Stock_Actual_Num


Hay que añadir un paso de verificación de ConfiguracionBaseDeDatos con la base de datos real.