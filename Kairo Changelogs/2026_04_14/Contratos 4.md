Miércoles 15/04/2026 - Jueves 16/04/2026
## Base de datos
Remplazar
```vb
Call CrearCampo("Estado", "Contratos_Proyecto", TipoByte, TamanoTipoCampo(TipoByte), "", Base, True)
```

Por
```vb
Call CrearCampo("Estado", "Contratos_Proyecto", TipoByte, TamanoTipoCampo(TipoByte), "", Base)
```

## `DetalleEconomicoProyectoObraFrm`
Debajo de
```vb
    If KeyCode = vbKeyEscape Then
        Me.TextoContratoTxt.Text = ""
        Me.TextoContratoTxt.Tag = ""
        Call PasarFoco(Me.TextoContratoTxt)
        Exit Sub
    End If
```

Añadir
```vb
    If KeyCode = vbKeyReturn And (Me.TextoContratoTxt.Tag <> "" Or Trim(Me.TextoContratoTxt.Text) = "") Then
        Call PasarFoco(Me.ImporteContratoTxt)
        Exit Sub
    End If
```

## `ContratosModalFrm`
Archivo modificado
