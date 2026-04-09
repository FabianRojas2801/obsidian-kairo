Lunes 30/03/2026 - Jueves 09/04/2026

# Base de datos (`!!ActualizadorBD/MainBas.bas`)

## Tabla `Servicios`
Añadir
```vb
Call CrearCampo("Contrato_Id", "Servicios", TipoEnteroLargo, TamanoTipoCampo(TipoEnteroLargo), "", Base)
```

## Nueva tabla `Contratos_Proyecto`
```vb
Call CrearNuevaTabla("Contratos_Proyecto", "Id", True, TipoAutonumerico, TamanoTipoCampo(TipoAutonumerico), "0", Base)
Call CrearCampo("Numero", "Contratos_Proyecto", TipoTexto, TamanoTipoCampo(TipoTexto), "", Base, True)
Call CrearCampo("Descripcion", "Contratos_Proyecto", TipoTexto, TamanoTipoCampo(TipoTexto), "", Base, True)
Call CrearCampo("Estado", "Contratos_Proyecto", TipoByte, TamanoTipoCampo(TipoByte), "", Base, True)
Call CrearCampo("Ficha_Id", "Contratos_Proyecto", TipoEnteroLargo, TamanoTipoCampo(TipoEnteroLargo), "", Base)
Call CrearCampo("Fecha_Inicio", "Contratos_Proyecto", TipoFecha, TamanoTipoCampo(TipoFecha), "", Base)
Call CrearCampo("Fecha_Fin", "Contratos_Proyecto", TipoFecha, TamanoTipoCampo(TipoFecha), "", Base)
Call CrearCampo("Fecha_Firma", "Contratos_Proyecto", TipoFecha, TamanoTipoCampo(TipoFecha), "", Base)
Call CrearCampo("Importe_Total", "Contratos_Proyecto", TipoDouble, TamanoTipoCampo(TipoDouble), "", Base)
Call CrearCampo("Porcentaje_Gastos_Generales", "Contratos_Proyecto", TipoDouble, TamanoTipoCampo(TipoDouble), "", Base)
Call CrearCampo("Porcentaje_Beneficios_Industriales", "Contratos_Proyecto", TipoDouble, TamanoTipoCampo(TipoDouble), "", Base)
Call CrearCampo("Porcentaje_Deducciones", "Contratos_Proyecto", TipoDouble, TamanoTipoCampo(TipoDouble), "", Base)
```

---

# `ContratosModalFrm`
Archivo nuevo

---

# `ResumenEconomicoProyectosObraFrm.frm`

## Controles
Añadir un botón `ContratosBtn` a la izquierda de `VerDetalleBtn`.

## `Form_Load`
Debajo de
```vb
Me.VerDetalleBtn.ShowFocusRect = Me.SalirBtn.ShowFocusRect
```

Añadir
```vb
Me.ContratosBtn.BackOver = Me.SalirBtn.BackOver
Me.ContratosBtn.ButtonType = Me.SalirBtn.ButtonType
Me.ContratosBtn.ShowFocusRect = Me.SalirBtn.ShowFocusRect
```

## `Form_Resize`
Debajo de
```vb
Me.VerDetalleBtn.Left = Me.Line2.X1 - Me.VerDetalleBtn.Width - 50
```

Añadir
```vb
Me.ContratosBtn.Top = Me.SalirBtn.Top
Me.ContratosBtn.Left = Me.VerDetalleBtn.Left - Me.ContratosBtn.Width - 50
```

## `ContratosBtn_Click`
Nuevo método
```vb
Private Sub ContratosBtn_Click()
    Call ContratosModalFrm.Show
    ContratosModalFrm.Left = Me.Left + (Me.Width - ContratosModalFrm.Width) / 2
    ContratosModalFrm.Top = Me.Top + (Me.Height - ContratosModalFrm.Height) / 2
    Call ContratosModalFrm.ZOrder(0)
End Sub
```

## `SalirBtn_Click`
Añadir al final
```vb
Call ContratosModalFrm.Hide
```

## Cambios en `CargaGeneral`, `CargaAOrigen`, `CargaMes`, `CargaPorFechas`

Los 4 métodos llevan los mismos cambios.

### `filtroDescripcion`
Buscar
```vb
filtroDescripcion = " and Descripcion like '*" & Replace(Me.DescripcionTxt.Text, " ", "*' and Descripcion like '*") & "*'"
```

Cambiar por
```vb
filtroDescripcion = " and Servicios.Descripcion like '*" & Replace(Me.DescripcionTxt.Text, " ", "*' and Servicios.Descripcion like '*") & "*'"
```

### `TextoOrden`
Buscar y reemplazar `Descripcion asc` por `Servicios.Descripcion asc` en las 7 líneas de `TextoOrden`. Son estas:
```vb
TextoOrden = " order by Codigo_Proyecto asc, Descripcion asc"
TextoOrden = " order by Nombre_Ficha asc, Descripcion asc"
TextoOrden = " order by Responsable_Tecnico_Id asc, Descripcion asc"
TextoOrden = " order by Tecnico_Id asc, Descripcion asc"
TextoOrden = " ORDER BY IIF(TRIM(Texto_Contrato_Base) <> '', 0, 1) ASC, TRIM(Texto_Contrato_Base) ASC, Descripcion ASC"
TextoOrden = " order by Descripcion asc"
```

### Consulta SQL
Buscar
```vb
Set MiRc = Base.OpenRecordset("select Servicios.*,Fichas.Nombre_Ficha" _
    & " from Servicios left join Fichas on Servicios.Ficha_Id=Fichas.Id" _
```

Cambiar por
```vb
Set MiRc = Base.OpenRecordset("select Servicios.*,Fichas.Nombre_Ficha," _
    & " Contratos_Proyecto.Numero AS Contrato_Numero, Contratos_Proyecto.Descripcion AS Contrato_Desc" _
    & " from (Servicios left join Fichas on Servicios.Ficha_Id=Fichas.Id)" _
    & " left join Contratos_Proyecto on Servicios.Contrato_Id=Contratos_Proyecto.Id" _
```

### Subtotales
Hay 2 bloques idénticos en cada método. Buscar (aparece 2 veces)
```vb
                    If Trim(Nz(MiRc!Texto_Contrato_Base, "")) <> "" Then
```

Cambiar cada uno por
```vb
                    If Not IsNull(MiRc!Contrato_Id) Then
                        TextoSubtotalActivo = MiRc!Contrato_Numero & " - " & MiRc!Contrato_Desc
                    ElseIf Trim(Nz(MiRc!Texto_Contrato_Base, "")) <> "" Then
```

> En el primer bloque, `TextoSubtotalActivo` se asigna en la línea añadida. En el segundo bloque, la variable es `TextoAuxiliar`:
```vb
                    If Not IsNull(MiRc!Contrato_Id) Then
                        TextoAuxiliar = MiRc!Contrato_Numero & " - " & MiRc!Contrato_Desc
                    ElseIf Trim(Nz(MiRc!Texto_Contrato_Base, "")) <> "" Then
```

### Celda del grid
Buscar
```vb
        If Not IsNull(MiRc!Texto_Contrato_Base) Then Me.GridProyectos.TextMatrix(Posicion, EProyectoGeneralCol.GeneralContrato) = Trim(MiRc!Texto_Contrato_Base)
```

Cambiar por
```vb
        If Not IsNull(MiRc!Contrato_Id) Then
            Me.GridProyectos.TextMatrix(Posicion, EProyectoGeneralCol.GeneralContrato) = MiRc!Contrato_Numero & " - " & MiRc!Contrato_Desc
        ElseIf Not IsNull(MiRc!Texto_Contrato_Base) Then
            Me.GridProyectos.TextMatrix(Posicion, EProyectoGeneralCol.GeneralContrato) = Trim(MiRc!Texto_Contrato_Base)
        End If
```

> En `CargaGeneral` la línea original es multilínea en lugar de una sola:
> ```vb
>         If Not IsNull(MiRc!Texto_Contrato_Base) Then
>             Me.GridProyectos.TextMatrix(Posicion, EProyectoGeneralCol.GeneralContrato) = Trim(MiRc!Texto_Contrato_Base)
>         End If
> ```
> El reemplazo es el mismo.

---

## Cambios en `ActualizarTecnico`

### `filtroDescripcion`
Buscar
```vb
    If Trim(Me.DescripcionTxt.Text) <> "" Then filtroDescripcion = " and Descripcion like '*" & Replace(Me.DescripcionTxt.Text, " ", "*' and Descripcion like '*") & "*'"
```

Cambiar por
```vb
    If Trim(Me.DescripcionTxt.Text) <> "" Then filtroDescripcion = " and Servicios.Descripcion like '*" & Replace(Me.DescripcionTxt.Text, " ", "*' and Servicios.Descripcion like '*") & "*'"
```

### `TextoOrden`
Buscar y reemplazar `Descripcion asc` por `Servicios.Descripcion asc` en las 7 líneas. Son estas:
```vb
TextoOrden = " order by Servicios.Fecha_Alta desc, Codigo_Proyecto desc, Descripcion asc"
TextoOrden = " order by Nombre_Ficha asc, Servicios.Fecha_Alta desc, Codigo_Proyecto desc, Descripcion asc"
TextoOrden = " order by Responsable_Tecnico_Id asc, Servicios.Fecha_Alta desc, Codigo_Proyecto desc, Descripcion asc"
TextoOrden = " order by Tecnico_Id asc, Servicios.Fecha_Alta desc, Codigo_Proyecto desc, Descripcion asc"
TextoOrden = " ORDER BY IIF(TRIM(Texto_Contrato_Base) <> '', 0, 1) ASC, TRIM(Texto_Contrato_Base) ASC, Descripcion ASC"
TextoOrden = " order by Servicios.Fecha_Alta desc, Codigo_Proyecto desc, Descripcion asc"
```

### Consulta SQL
Buscar
```vb
               "Estados_Proyectos.Estado AS EP_Estado, Estados_Proyectos.Color_Fondo AS EP_ColorF, Estados_Proyectos.Color_Letra AS EP_ColorL " & _
               "FROM (((((Servicios " & _
```

Cambiar por
```vb
               "Estados_Proyectos.Estado AS EP_Estado, Estados_Proyectos.Color_Fondo AS EP_ColorF, Estados_Proyectos.Color_Letra AS EP_ColorL , " & _
               "Contratos_Proyecto.Numero AS Contrato_Numero, Contratos_Proyecto.Descripcion AS Contrato_Desc " & _
               "FROM ((((((Servicios " & _
```

Y buscar
```vb
               "LEFT JOIN Estados_Proyectos ON Servicios.Estado = Estados_Proyectos.ID) " & _
```

Cambiar por
```vb
               "LEFT JOIN Estados_Proyectos ON Servicios.Estado = Estados_Proyectos.ID) " & _
               "LEFT JOIN Contratos_Proyecto ON Servicios.Contrato_Id = Contratos_Proyecto.Id) " & _
```

### Subtotales
Igual que en las otras 4 vistas: buscar (2 veces)
```vb
                    If Trim(Nz(MiRc!Texto_Contrato_Base, "")) <> "" Then
```

Primer bloque, cambiar por
```vb
                    If Not IsNull(MiRc!Contrato_Id) Then
                        TextoSubtotalActivo = MiRc!Contrato_Numero & " - " & MiRc!Contrato_Desc
                    ElseIf Trim(Nz(MiRc!Texto_Contrato_Base, "")) <> "" Then
```

Segundo bloque, cambiar por
```vb
                    If Not IsNull(MiRc!Contrato_Id) Then
                        TextoAuxiliar = MiRc!Contrato_Numero & " - " & MiRc!Contrato_Desc
                    ElseIf Trim(Nz(MiRc!Texto_Contrato_Base, "")) <> "" Then
```

### Celda del grid
Buscar
```vb
            If Not IsNull(MiRc!Texto_Contrato_Base) Then
```

Cambiar por
```vb
            If Not IsNull(MiRc!Contrato_Id) Then
                Me.GridProyectos.TextMatrix(Posicion, EProyectoTecnicoCol.TecnicoContrato) = MiRc!Contrato_Numero & " - " & MiRc!Contrato_Desc
            ElseIf Not IsNull(MiRc!Texto_Contrato_Base) Then
```

---

# `DetalleEconomicoProyectoObraFrm.frm`

## `TextoContratoTxt_KeyPress`
Reemplazar el método completo. Buscar
```vb
Private Sub TextoContratoTxt_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Call PasarFoco(Me.ImporteContratoTxt)

If KeyAscii = 27 Then
    Me.TextoContratoTxt.Text = ""
    Call PasarFoco(Me.TextoContratoTxt)
    Exit Sub
End If

End Sub
```

Cambiar por
```vb
Private Sub TextoContratoTxt_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Call PasarFoco(Me.ImporteContratoTxt)
    
    If KeyAscii = 27 Then
        Me.TextoContratoTxt.Text = ""
        Me.TextoContratoTxt.Tag = ""
        Call PasarFoco(Me.TextoContratoTxt)
        Exit Sub
    End If

    Me.TextoContratoTxt.Tag = ""

End Sub
```

## `TextoContratoTxt_KeyDown`
Nuevo método, añadir después de `TextoContratoTxt_KeyPress`
```vb
Private Sub TextoContratoTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
        KeyCode = 0
        
        Call SelectorFrm.Cargar( _
                cGridManager.Of( _
                    SelectorFrm.Grid, _
                    CollectionOf( _
                        cColumna.Of(0, "Id", " ", False), _
                        cColumna.Of(1, "Numero", "Nro Contrato", False, 1500), _
                        cColumna.Of(2, "Descripcion", "Descripcion", True, 3000))), _
                "Contratos", _
                "SELECT Id, Numero, Descripcion FROM Contratos_Proyecto WHERE UCase(Numero) LIKE '*" & UCase(Me.TextoContratoTxt.Text) & "*' OR UCase(Descripcion) LIKE '*" & UCase(Me.TextoContratoTxt.Text) & "*'", _
                cFunction.Of(Me, "callbackSelectorContrato"), _
                cIntegerSupplier.OfFunction(cFunction.Of(Me, "widthIntegerSupplierSelectorContrato")), _
                cIntegerSupplier.OfVal(5000), _
                Me.TextoContratoTxt)
    End If
End Sub
```

## `callbackSelectorContrato`
Nuevo método
```vb
Public Sub callbackSelectorContrato()
    Me.TextoContratoTxt.Text = SelectorFrm.Grid.TextMatrix(SelectorFrm.Grid.Row, SelectorFrm.gm.p("Numero")) & " - " & SelectorFrm.Grid.TextMatrix(SelectorFrm.Grid.Row, SelectorFrm.gm.p("Descripcion"))
    Me.TextoContratoTxt.Tag = SelectorFrm.Grid.TextMatrix(SelectorFrm.Grid.Row, SelectorFrm.gm.p("Id"))
    Unload SelectorFrm
End Sub
```

## `widthIntegerSupplierSelectorContrato`
Nuevo método
```vb
Public Function widthIntegerSupplierSelectorContrato() As Integer
    widthIntegerSupplierSelectorContrato = Max(Me.TextoContratoTxt.Width + 275, 5000 + 275)
End Function
```

## Guardado del servicio
Buscar
```vb
If Trim(Me.TextoContratoTxt.Text) = "" Then
    MiRc!Texto_Contrato_Base = " "
Else
    MiRc!Texto_Contrato_Base = Trim(Me.TextoContratoTxt.Text)
End If
```

Cambiar por
```vb
If Trim(Me.TextoContratoTxt.Text) = "" Then
    MiRc!Texto_Contrato_Base = " "
    MiRc!Contrato_Id = Null
Else
    MiRc!Texto_Contrato_Base = Trim(Me.TextoContratoTxt.Text)
    If Trim(Me.TextoContratoTxt.Tag) <> "" Then
        MiRc!Contrato_Id = CLng(Me.TextoContratoTxt.Tag)
    Else
        MiRc!Contrato_Id = Null
    End If
End If
```

## Carga del servicio
Debajo de
```vb
    If IsNull(MiRc!Texto_Contrato_Base) = False Then Me.TextoContratoTxt.Text = Trim(MiRc!Texto_Contrato_Base)
```

Añadir
```vb
    If IsNull(MiRc!Contrato_Id) = False Then Me.TextoContratoTxt.Tag = MiRc!Contrato_Id
```
