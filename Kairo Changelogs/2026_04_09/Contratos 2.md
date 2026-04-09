Jueves 09/04/2026

## `ContratosModalFrm`
Archivo modificado

## `DetalleEconomicoProyectoObraFrm`
### `TextoContratoTxt_KeyPress`
Método eliminado

## `TextoContratoTxt_Change`
Nuevo método
```vb
Private Sub TextoContratoTxt_Change()
    If Me.TextoContratoTxt.Tag <> "" Then
        Me.TextoContratoTxt.Tag = ""
        Me.TextoContratoTxt.Text = ""
    End If
End Sub
```

## `TextoContratoTxt_KeyDown`
Remplazar método
```vb
Private Sub TextoContratoTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Me.TextoContratoTxt.Text = ""
        Me.TextoContratoTxt.Tag = ""
        Call PasarFoco(Me.TextoContratoTxt)
        Exit Sub
    End If
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
        KeyCode = 0
        
        Call AbrirSelectorContrato
    End If
End Sub
```

## `TextoContratoTxt_LostFocus`
Remplazar método
```vb
Private Sub TextoContratoTxt_LostFocus()
    Me.Shape7(Me.TextoContratoTxt.TabIndex).BorderColor = ColorBordeFrame

    If Trim(Me.TextoContratoTxt.Text) <> "" And Trim(Me.TextoContratoTxt.Tag) = "" Then
        Call AbrirSelectorContrato
    End If
End Sub
```

## `AbrirSelectorContrato`
Nuevo método
```vb
Private Sub AbrirSelectorContrato()
    Call SelectorFrm.Cargar( _
        cGridManager.Of( _
            SelectorFrm.Grid, _
            cGridConfig.Default, _
            CollectionOf( _
                cColumna.Of(0, "Id", " ", False), _
                cColumna.Of(1, "Numero", "Nro Contrato", False, 1500), _
                cColumna.Of(2, "Descripcion", "Descripción", True, 3000))), _
        "Contratos", _
        " SELECT Id, Numero, Descripcion " & _
            " FROM Contratos_Proyecto " & _
            " WHERE UCase(Numero) LIKE '*" & UCase(Me.TextoContratoTxt.Text) & "*'" & _
            "   OR UCase(Descripcion) LIKE '*" & UCase(Me.TextoContratoTxt.Text) & "*'" & _
            "   OR UCase(Numero & ' - ' & Descripcion) LIKE '*" & UCase(Me.TextoContratoTxt.Text) & "*'", _
        cFunction.Of(Me, "callbackSelectorContrato"), _
        cIntegerSupplier.OfFunction(cFunction.Of(Me, "widthIntegerSupplierSelectorContrato")), _
        cIntegerSupplier.OfVal(5000), _
        Me.TextoContratoTxt)
    SelectorFrm.Top = SelectorFrm.Top + Me.TextoContratoTxt.Height
End Sub
```

## `callbackSelectorContrato`
Remplazar método
```vb
Public Sub callbackSelectorContrato()
    Dim Id As String
    Id = SelectorFrm.Grid.TextMatrix(SelectorFrm.Grid.Row, SelectorFrm.gm.p("Id"))
    
    If NumSM(Id) > 0 Then
        Me.TextoContratoTxt.Text = SelectorFrm.Grid.TextMatrix(SelectorFrm.Grid.Row, SelectorFrm.gm.p("Numero")) & " - " & SelectorFrm.Grid.TextMatrix(SelectorFrm.Grid.Row, SelectorFrm.gm.p("Descripcion"))
        Me.TextoContratoTxt.Tag = Id
    Else
        Call PasarFoco(Me.TextoContratoTxt)
    End If
    
    Unload SelectorFrm
End Sub
```