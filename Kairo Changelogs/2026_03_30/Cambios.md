# Cambios ContratosModalFrm - Lógica Completa

## Resumen
Implementación completa de la lógica del formulario modal de Contratos. Incluye carga de datos, CRUD completo, filtros y validaciones.

---

## 1. Cambios en Declaraciones de Variables

**Línea ~993, cambiar:**
```vb
Private IdContrato As Long
```

**Por:**
```vb
Private IdContratoActual As Long
```

---

## 2. Cambios en Form_Load

**Línea ~1127, cambiar:**
```vb
IdContrato = -1
```

**Por:**
```vb
IdContratoActual = -1
```

**Al final de Form_Load, después de `HaInicializado = True` (línea ~1169), agregar:**
```vb
Call CargarContratos
```

---

## 3. Event Handlers de Botones

Agregar después de la función `widthIntegerSupplierSelectorCliente()`:

```vb
Private Sub ActualizarBtn_Click()
    Call CargarContratos
End Sub

Private Sub AñadirContratoBtn_Click()
    IdContratoActual = -1
    Call LimpiarFrameContrato
    Call AbrirFrameContrato
    Call PasarFoco(Me.NroContratoTxt)
End Sub

Private Sub ModificarContratoBtn_Click()
    If ContratosGm.SelectedRow < 0 Then
        MsgBox "Debe seleccionar un contrato", vbInformation
        Exit Sub
    End If
    
    IdContratoActual = CLng(Me.GridContratos.TextMatrix(ContratosGm.SelectedRow, 0))
    Call CargarDatosContrato(IdContratoActual)
    Call AbrirFrameContrato
End Sub

Private Sub BorrarContratoBtn_Click()
    If ContratosGm.SelectedRow < 0 Then
        MsgBox "Debe seleccionar un contrato", vbInformation
        Exit Sub
    End If
    
    If MsgBox("¿Desea eliminar este contrato?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    IdContratoActual = CLng(Me.GridContratos.TextMatrix(ContratosGm.SelectedRow, 0))
    
    Base.Execute "DELETE FROM Contratos_Proyecto WHERE id = " & IdContratoActual
    
    MsgBox "Contrato eliminado", vbInformation
    Call CerrarFrameContrato
    IdContratoActual = -1
    Call CargarContratos
End Sub

Private Sub GrabarContratoBtn_Click()
    Dim RsContrato As DAO.Recordset
    
    If Not ValidarDatosContrato() Then
        Exit Sub
    End If
    
    Set RsContrato = Base.OpenRecordset( _
        " SELECT *" & _
        " FROM Contratos_Proyecto" & _
        " WHERE Id = " & IdContratoActual)
    
    If RsContrato.EOF Then
        Call RsContrato.AddNew
    Else
        Call RsContrato.MoveFirst
        Call RsContrato.Edit
    End If
    
    ' Asignar valores al recordset
    RsContrato!Numero = Me.NroContratoTxt.Text
    RsContrato!Descripcion = Me.DescripcionContratoTxt.Text
    RsContrato!Ficha_Id = CLng(Me.ClienteContratoTxt.Tag)
    RsContrato!FechaInicio = CDate(Me.FechaInicioContratoTxt.Text)
    RsContrato!FechaFin = CDate(Me.FechaFinContratoTxt.Text)
    RsContrato!FechaFirma = CDate(Me.FechaFirmaContratoTxt.Text)
    RsContrato!ImporteTotal = CDbl(Me.ImporteTotalContratoTxt.Text)
    RsContrato!PorcentajeGastosGenerales = CDbl(Me.GastosGeneralesContratoTxt.Text)
    RsContrato!PorcentajeBeneficiosIndustriales = CDbl(Me.BeneficioIndustrialesTxt.Text)
    RsContrato!PorcentajeDeducciones = CDbl(Me.DeduccionesContratoTxt.Text)
    RsContrato!Estado = Me.EstadoContratoCombo.ListIndex
    
    Call RsContrato.Update
    Call RsContrato.Close
    Set RsContrato = Nothing
    
    MsgBox "Contrato grabado exitosamente", vbInformation
    Call CerrarFrameContrato
    IdContratoActual = -1
    Call CargarContratos
End Sub

Private Sub GridContratos_DblClick()
    Call ModificarContratoBtn_Click
End Sub

Private Sub ContratoFiltrosTxt_Change()
    Call CargarContratos
End Sub

Private Sub ClienteFiltrosTxt_Change()
    Call CargarContratos
End Sub

Private Sub EstadoFiltrosCombo_Change()
    Call CargarContratos
End Sub
```

---

## 4. Funciones Auxiliares

Agregar al final del archivo:

```vb
Private Sub LimpiarFrameContrato()
    Me.NroContratoTxt.Text = ""
    Me.DescripcionContratoTxt.Text = ""
    Me.ClienteContratoTxt.Text = ""
    Me.ClienteContratoTxt.Tag = ""
    Me.FechaInicioContratoTxt.Text = FECHA_VACIA
    Me.FechaFinContratoTxt.Text = FECHA_VACIA
    Me.FechaFirmaContratoTxt.Text = FECHA_VACIA
    Me.ImporteTotalContratoTxt.Text = ""
    Me.GastosGeneralesContratoTxt.Text = ""
    Me.BeneficioIndustrialesTxt.Text = ""
    Me.DeduccionesContratoTxt.Text = ""
    Me.EstadoContratoCombo.ListIndex = 0
End Sub

Private Sub CargarContratos()
    Dim RsContratos As DAO.Recordset
    Dim Sql As String
    Dim Posicion As Long
    
    Sql = _
        " SELECT Contratos_Proyecto.id, Contratos_Proyecto.Numero, Contratos_Proyecto.Descripcion," & _
        " Contratos_Proyecto.Importe_Total, Fichas.Nombre_Ficha FROM Contratos_Proyecto" & _
        " LEFT JOIN Fichas ON Contratos_Proyecto.Ficha_Id = Fichas.Id WHERE 1 = 1"
    
    ' Aplicar filtros
    If Me.ContratoFiltrosTxt.Text <> "" Then
        Sql = Sql & " AND Numero LIKE '*" & Me.ContratoFiltrosTxt.Text & "*'"
    End If
    
    If Me.ClienteFiltrosTxt.Text <> "" Then
        Sql = Sql & " AND Nombre_Ficha LIKE '*" & Me.ClienteFiltrosTxt.Text & "*'"
    End If
    
    If Me.EstadoFiltrosCombo.ListIndex > 0 Then
        Sql = Sql & " AND Estado = " & (Me.EstadoFiltrosCombo.ListIndex - 1)
    End If
    
    Set RsContratos = Base.OpenRecordset(Sql)
    
    ' Limpiar grid
    Me.GridContratos.Rows = 1
    Me.GridContratos.Redraw = False
    
    Posicion = 1
    Do While Not RsContratos.EOF
        ' Agregar fila
        Me.GridContratos.Rows = Me.GridContratos.Rows + 1
        
        ' Llenar datos usando identificadores de cGridManager
        Me.GridContratos.TextMatrix(Posicion, ContratosGm.p("NroContrato")) = RsContratos!Numero
        Me.GridContratos.TextMatrix(Posicion, ContratosGm.p("Descripcion")) = RsContratos!Descripcion
        Me.GridContratos.TextMatrix(Posicion, ContratosGm.p("Importe")) = RsContratos!Importe_Total
        Me.GridContratos.TextMatrix(Posicion, ContratosGm.p("Certificado")) = ""
        Me.GridContratos.TextMatrix(Posicion, ContratosGm.p("Pendiente")) = ""
        
        ' Guardar ID en Tag de la fila para referencia
        Me.GridContratos.Row = Posicion
        Me.GridContratos.Tag = RsContratos!id
        
        Posicion = Posicion + 1
        RsContratos.MoveNext
    Loop
    
    Me.GridContratos.Redraw = True
    RsContratos.Close
    Set RsContratos = Nothing
End Sub

Private Sub CargarDatosContrato(Id As Long)
    Dim rs As DAO.Recordset
    
    Set rs = Base.OpenRecordset("SELECT Contratos_Proyecto.*, Fichas.Nombre_Ficha FROM Contratos_Proyecto " & _
        "LEFT JOIN Fichas ON Contratos_Proyecto.Ficha_Id = Fichas.Id WHERE Contratos_Proyecto.id=" & Id)
    
    If Not rs.EOF Then
        Me.NroContratoTxt.Text = rs("Numero")
        Me.DescripcionContratoTxt.Text = rs("Descripcion")
        Me.ClienteContratoTxt.Text = rs("Nombre_Ficha")
        Me.ClienteContratoTxt.Tag = rs("Ficha_Id")
        Me.FechaInicioContratoTxt.Text = Format(rs("FechaInicio"), "dd/MM/yyyy")
        Me.FechaFinContratoTxt.Text = Format(rs("FechaFin"), "dd/MM/yyyy")
        Me.FechaFirmaContratoTxt.Text = Format(rs("FechaFirma"), "dd/MM/yyyy")
        Me.ImporteTotalContratoTxt.Text = rs("ImporteTotal")
        Me.GastosGeneralesContratoTxt.Text = rs("PorcentajeGastosGenerales")
        Me.BeneficioIndustrialesTxt.Text = rs("PorcentajeBeneficiosIndustriales")
        Me.DeduccionesContratoTxt.Text = rs("PorcentajeDeducciones")
        Me.EstadoContratoCombo.ListIndex = rs("Estado")
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Function ValidarDatosContrato() As Boolean
    ValidarDatosContrato = True
    
    If Me.NroContratoTxt.Text = "" Then
        MsgBox "Ingrese número de contrato", vbExclamation
        Call PasarFoco(Me.NroContratoTxt)
        ValidarDatosContrato = False
        Exit Function
    End If
    
    If Me.DescripcionContratoTxt.Text = "" Then
        MsgBox "Ingrese descripción", vbExclamation
        Call PasarFoco(Me.DescripcionContratoTxt)
        ValidarDatosContrato = False
        Exit Function
    End If
    
    If Me.ClienteContratoTxt.Tag = "" Then
        MsgBox "Seleccione un cliente", vbExclamation
        Call PasarFoco(Me.ClienteContratoTxt)
        ValidarDatosContrato = False
        Exit Function
    End If
End Function
```

---

## Resumen de Funcionalidad

✅ **Cargar contratos** al abrir el formulario  
✅ **Agregar** contrato nuevo  
✅ **Modificar** contrato seleccionado  
✅ **Borrar** contrato con confirmación  
✅ **Grabar/Actualizar** con validación  
✅ **Filtrar** por número, cliente y estado  
✅ **Doble-click** en grid para editar  
✅ **Gestión del estado** del contrato  

---

## Notas Importantes

- **IdContratoActual**: Variable de la instancia actual para evitar conflicto con enum EContrato
- **CargarContratos()**: Sin parámetros, aplica filtros desde los controles del formulario
- **Validación**: Solo verifica campos obligatorios (número, descripción, cliente)
- **Fechas**: Formato dd/MM/yyyy en la interfaz, se ajustan en Form_Load si falta el año
- **Certificado y Pendiente**: Quedan vacíos, implementar cuando existan las tablas relacionadas
