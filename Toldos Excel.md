# Total
En todas las filas, $\text{Precio}$ corresponde a $\text{Cantidad} \mul \text{Precio}$ redondeado a **2 decimales**.

# Jgo Soporte
- - **Precio:** Corresponde a la tabla "Brazo Invisible Epsylon", fila **Jgo. Sop. Epsylon** columna **C/Dto**

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
\text{Cantidad Vivo} = \frac{\text{Salida}}{100} &&
\end{flalign}
$$
# Taco+Tornillo
- **Precio:** Corresponde a la tabla "Brazo Invisible Epsylon", fila **Taco + tornillo** columna **C/Dto**
# P. Delfin (m)
- **Cantidad:**  Es un valor "por tramos" según el valor de `Linea`
$$
\begin{flalign}
\text{Cantidad P. Delfin (m)}=
f(\text{Linea})=
\begin{cases}
2{,}5 & \si \text{Linea}\le 250\\
3{,}  & \si 250<\text{Linea}\le 300\\
3{,}5 & \si 300<\text{Linea}\le 350\\
5     & \si 350<\text{Linea}\le 500\\
6     & \si 500<\text{Linea}\le 600\\
7     & \si \text{Linea}>600
\end{cases} &&
\end{flalign}
$$
- **Precio:** 
# T.Enrolle
- **Cantidad:** Es un valor "por tramos" según el valor de `Linea` (exactamente igual que `P. Delfin (m)`)
$$
\begin{flalign}
\text{Cantidad T.Enrolle}=
f(\text{Linea})=
\begin{cases}
2{,}5 & \si \text{Linea}\le 250\\
3     & \si 250<\text{Linea}\le 300\\
3{,}5 & \si 300<\text{Linea}\le 350\\
5     & \si 350<\text{Linea}\le 500\\
6     & \si 500<\text{Linea}\le 600\\
7     & \si \text{Linea}>600
\end{cases} &&
\end{flalign}
$$

