# Total
En todas las filas, $\text{Total}$ corresponde a $\text{Cantidad} \mul \text{Precio}$ redondeado a **2 decimales**.

# Jgo Soporte
- **Precio:** Corresponde a la tabla "Brazo Invisible Epsylon", fila **Jgo. Sop. Epsylon** columna **C/Dto**

# Lona

- **Cantidad:** Redondea a **2 decimales** la siguiente fórmula: $(\text{Linea} \div 100) \mul [(\text{Salida} + 60) \div 100]$
$$
\begin{flalign}
\text{Cantidad Lona} = \frac{\text{Linea}}{100} \mul \frac{\text{Salida} + {60}}{100} &&
\end{flalign}
$$
- **Precio:** Corresponde a la tabla "Brazo Invisible Epsylon", Fila **Lona** - **Rec-Sat** columna **C/Dto**

# Confección
- **Cantidad:** Busca en la tabla de Confección (de `AC:4` a `AN:18`) el valor que corresponde a la intersección entre la fila **Salida** y la columna **Línea**, con coincidencia exacta en ambos casos.
- **Precio:** Corresponde a la tabla "Brazo Invisible Epsylon", fila **Hora Manual** columna **C/Dto**
# Fabricación
- **Precio:** Corresponde a la tabla "Brazo Invisible Epsylon", fila **Hora Manual** columna **C/Dto** (igual que confección)
# Manivela
- **Precio:** Corresponde a la tabla "Brazo Invisible Epsylon", fila **Manivela** columna **C/Dto**
# Varilla
- **Cantidad:** Realiza la siguiente fórmula:  $\text{Linea} \div 100 \mul 3$
$$
\begin{flalign}
\text{Cantidad Varilla} = \frac{\text{Linea}}{100} \mul 3 &&
\end{flalign}
$$
- **Precio:** Corresponde a la tabla "Brazo Invisible Epsylon", fila **Varilla Vaina** columna **C/Dto**
# Vivo
- **Cantidad:** Realiza la siguiente fórmula:  $\text{Linea} \div 100$
$$
\begin{flalign}
\text{Cantidad Vivo} = \frac{\text{Linea}}{100} &&
\end{flalign}
$$
# Taco+Tornillo
- **Precio:** Corresponde a la tabla "Brazo Invisible Epsylon", fila **Taco + tornillo** columna **C/Dto**
# P. Delfin (m)
- **Cantidad:**  Es un valor "por tramos" según el valor de `Linea`
$$
\begin{flalign}
\text{Cantidad P. Delfin (m)}=
\begin{cases}
2{,}5 & \text{si } \text{Linea}\le 250\\
3{,}  & \text{si } 250<\text{Linea}\le 300\\
3{,}5 & \text{si } 300<\text{Linea}\le 350\\
5     & \text{si } 350<\text{Linea}\le 500\\
6     & \text{si } 500<\text{Linea}\le 600\\
7     & \text{si } \text{Linea}>600
\end{cases} &&
\end{flalign}
$$
- **Precio:** Selecciona entre dos valores de la segunda tabla superior, de `G3` a `K18`
Si $\text{Linea} \leq 500$ y $\text{Salida} \leq 275$ el precio es "**Perfil carga Delfin Ø75**", en caso contrario es "**Perfil carga Delfin Ø 100**", en ambos caso se selecciona la columna **Precio**.
$$
\begin{flalign}
\text{Precio P.Delfin (m) = }
\begin{cases}
\text{Precio Perfil carga } \diameter 75,  & \text{si } (\text{Linea}\le 500)\land(\text{Salida}\le 275)\\
\text{Precio Perfil carga } \diameter 100, & \text{en otro caso}
\end{cases} &&
\end{flalign}
$$
# Tapetas
- **Precio:** Selecciona entre dos valores de la segunda tabla superior, de `G3` a `K18`
Si $\text{Linea} \geq 500$ o $\text{Salida} \gt 250$ el precio es "**Tapetas  Delfin Ø 100**", en caso contrario es "**Tapetas  Delfin Ø 75**", en ambos casos se selecciona la columna **Precio**.
$$
\begin{flalign}
\text{Precio Tapetas} =
\begin{cases}
\text{Precio Tapetas Delfin } \diameter 100 & \text{si } (\text{Linea} \geq 500) \lor (\text{Salida} \gt 250)\\
\text{Precio Tapetas Delfin } \diameter 75 & \text{en otro caso}
\end {cases} &&
\end{flalign}
$$

# T.Enrolle
- **Cantidad:** Es un valor "por tramos" según el valor de `Linea` (exactamente igual que `P. Delfin (m)`)
$$
\begin{flalign}
\text{Cantidad T.Enrolle}=
\begin{cases}
2{,}5 & \text{si } \text{Linea}\le 250\\
3     & \text{si } 250<\text{Linea}\le 300\\
3{,}5 & \text{si } 300<\text{Linea}\le 350\\
5     & \text{si } 350<\text{Linea}\le 500\\
6     & \text{si } 500<\text{Linea}\le 600\\
7     & \text{si } \text{Linea}>600
\end{cases} &&
\end{flalign}
$$

- **Precio:** Selecciona entre dos valores de la segunda tabla superior, de `G3` a `K18`
Si $\text{Linea} \leq 500$ y $\text{Salida} \leq 275$ el precio es "**Tubo enrrolle Ø 70**", en caso contrario es "**Tubo enrrolle Ø 80**", en ambos casos se selecciona la columna **Precio**.

$$
\begin{flalign}
\text{Precio T.Enrolle} =
\begin{cases}
\text{Precio T.Enrolle } \diameter 70 & \text{si } (\text{Linea} \leq 500) \land (\text{Salida} \leq 275)\\
\text{Precio T.Enrolle } \diameter 80 & \text{en otro caso}
\end {cases} &&
\end{flalign}
$$

# Casquillo
- **Precio:** Suma 2 valores de la segunda tabla superior, de `G3` a `K18`.
Si $\text{Linea} \leq 500$ y $\text{Salida} \leq 275$ se suman "**Casquillo Ø 70**" y "**Casq. Maquina Ø 70**", en caso contrario se suman "**Casquillo Ø 80**" y "**Casq. Maquina Ø 80 nylon**"
$$
\begin{flalign}
\text{Precio Casquillo}=
\begin{cases}
\text{Precio Casquillo Ø 70} + \text{Precio Casq. Maquina Ø 70} &
\text{si } (\text{Linea} \leq 500) \land (\text{Salida} \leq 275) \\
\text{Precio Casquillo Ø 80} + \text{Precio Casq. Maquina Ø 80 nylon} &
\text{en otro caso}
\end{cases} &&
\end{flalign}
$$

# Brazos
- **Precio:** Corresponde a la tercera tabla superior, fila corresponde a **Salida** y la columna **Precio**.
Si **Salida** no coincide con ninguna fila, se usa la fila inmediatamente inferior a Salida.

# Cruce
Solo presente en las hojas `BI_Cruce` y `BI_Cruce_EpsyDlux`.
- **Precio:** Corresponde a la segunda tabla superior, fila "**Cruce Vinci**" y la columna **Precio**

# Tranformadoor, Transfo
Solo presente en las hojas `BI_EpsyDLUX` (como `Transformadoor`) y `BI_Cruce_EpsyDlux` (como `Transfo`).
- **Potencia:** *Es un valor calculado internamente, no se muestra en ninguna celda.* Corresponde a la tabla "**Brazos B. I.**", fila **Salida** y columna **Potencia**, si **Salida** no coincide con ninguna fila, se usa la fila inmediatamente inferior a Salida.
- **Precio:** Corresponde a la tabla "**Brazo Invisible Epsylon Dlux Cruce**", si $\text{Potencia} \leq 140$ se usa el precio de "**Transformador 24V IP67 150W**", si $\text{Potencia} \leq 190$ se usa el precio de "**Transformador 24V IP67 240W**", si no, se usa el precio de "**Transformador 24V IP67 320W**".
$$
\begin{flalign}
\text{Precio Transformador} = 
\begin{cases}
\text{Precio Transformador 24V IP67 150W} & \text{si } \text{Potencia} \leq 140 \\
\text{Precio Transformador 24V IP67 240W} & \text{si } \text{Potencia} \leq 190 \\
\text{Precio Transformador 24V IP67 320W} & \text{en otro caso}
\end{cases} &&
\end{flalign}
$$

# Costos GVM
La suma de todos los totales anteriores.

# Tarifa GVM
Se calcula a partir de "**Costos GVM**", aplicando primero los recargos de "**Gastos Gen**." y "**Ganancias**", y después convirtiendo ese precio "neto" en un precio de tarifa antes de aplicar el “Descuento”. Al final redondea a 1 decimal.

$$
\begin{flalign}
\text{Tarifa GVM} = 
\text{Costos GVM} \mul \left( 1 + \frac{\text{Gastos Gen. GVM}}{100}+\frac{\text{Ganancias GVM}}{100} \right) \div \left( 1 - \frac{\text{Descuento GVM}}{100} \right) &&
\end{flalign}
$$
**Gastos Gen.**, **Ganancias** y **Descuento** son valores fijos.
# Coste TVM
No se presenta en la hoja `BI_Epsylon`.
Se calcula aplicando un descuento porcentual a la **Tarifa GVM** y redondeando el resultado a **2 decimales**.
El descuento se encuentra en la celda **K46** en la hoja `BI_Cruce`,  **O44** en la hoja `BI_EpsyDLUX` y **K44** en `BI_Cruce_EpsyDlux`.
$$
\begin{flalign}
\text{Coste TVM} = \text{Tarifa GVM} \mul \left(1 - \frac{\text{Descuento Coste TVM}}
{100}\right) &&
\end{flalign}
$$
# Tarifa TVM
No se presenta en la hoja `BI_Epsylon`.
Se calcula a partir del "**Coste TVM**" aplicando recargos por **Gastos Generales** y **Ganancias**, y después "deshaciendo" (dividiendo) el efecto de un descuento comercial, redondeando a 1 decimal.
**Ganancias TVM** se encuentra en **K58** en la hoja `BI_Cruce`, **O55** en la hoja `BI_EpsyDLUX` y **K55** en la hoja `BI_Cruce_EpsyDlux`.
**Gastos Gen. TVM** y **Descuento TVM** se encuentran debajo de **Ganancias TVM**.

$$
\begin{flalign}
\text{Tarifa TVM} =
\frac{
	\text{Coste TVM} + \text{Coste TVM} \mul \frac{\text{Gastos Gen. TVM}}{100} + \text{Coste TVM} \mul \frac{\text{Ganancias TVM}}{100}
} {
	1 - \frac{\text{Descuento TVM}}{100}
} &&
\end{flalign}
$$
Se puede factorizar:
$$
\begin{flalign}
\text{Tarifa TVM} =
\text{Coste TVM} \mul \left(1 + \frac{\text{Gastos Gen. TVM}}{100} + \frac{\text{Ganancias TVM}}{100}\right)
\div
\left(1 - \frac{\text{Descuento TVM}}{100}\right) &&
\end{flalign}
$$
# DTO
No se presenta en la hoja `BI_Epsylon`.
Se calcula como un porcentaje de la **Tarifa TVM**, en este caso el porcentaje queda fijo en **10%**.
El **porcentaje** se define en la celda amarilla adyacente a la celda **DTO**.
$$
\begin{flalign}
\text{DTO} =
\text{Tarifa TVM} \mul \frac{\text{Porcentaje}}{100} &&
\end{flalign}
$$

# Total TVM
No se presenta en la hoja `BI_Epsylon`.
Se calcula a partir de **Tarifa TVM** y el **porcentaje** de **DTO**.
$$
\begin{flalign}
\text{Total TVM}= \text{Tarifa TVM} \mul \left( 1 - \frac{\text{Porcentaje}}{100} \right) &&
\end{flalign}
$$

