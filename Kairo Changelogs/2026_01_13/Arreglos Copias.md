Martes 13/01/2026, **añadido a Kairo el Miércoles 14/01/2026**.
## `DocumentosCompraFrm.frm`
### `Form_Load`
Arreglar carga de copias (al final del método)
```vb
' Copias
For Posicion = 0 To 9
    Me.CopiasCombo.AddItem CStr(Posicion + 1), Posicion
Next Posicion
Me.CopiasCombo.ListIndex = 0

Set MiRc = Base.OpenRecordset( _
    " SELECT Copias_Compra" & _
    " FROM Configuracion_Empresa" _
)

If Not MiRc.EOF Then
    Call MiRc.MoveFirst
    
    If Not IsNull(MiRc!Copias_Compra) Then
        Me.CopiasCombo.ListIndex = CLng(MiRc!Copias_Compra) - 1
    End If
End If
```

## `DocumentosVentaFrm.frm`
### `Form_Load`
Arreglar carga de copias (al final del método)
```vb
' Copias
For Posicion = 0 To 9
    Me.CopiasCombo.AddItem CStr(Posicion + 1), Posicion
Next Posicion
Me.CopiasCombo.ListIndex = 0

Set MiRc = Base.OpenRecordset( _
    " SELECT Copias_Venta" & _
    " FROM Configuracion_Empresa" _
)

If Not MiRc.EOF Then
    Call MiRc.MoveFirst
    
    If Not IsNull(MiRc!Copias_Venta) Then
        Me.CopiasCombo.ListIndex = CLng(MiRc!Copias_Venta) - 1
    End If
End If
```

## `ConfiguracionEmpresaDatosAdicionalesFrm.frm`

### Controles
- Setear `CopiasVentasCombo.TabStop` a true, asegurarse de que su `TabIndex` sea 86.
- Setear `TabStop` a true para todos los `NDV_Combo`, los `NCD_Combo`, `CopiasVentasCombo` y `CopiasComprasCombo`.

### `CopiasVentasCombo_KeyPress`
Nuevo método, debajo de `ContraseñaFTPTxt_LostFocus`
```vb
Private Sub CopiasVentasCombo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call PasarFoco(Me.NDC1Combo)
    End If
End Sub
```

### `CopiasComprasCombo_KeyPress`
Nuevo método, debajo del anterior
```vb
Private Sub CopiasComprasCombo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call PasarFoco(Me.ContenedorTab)
    End If
End Sub
```

### `Vencimiento2Txt_KeyPress`
Reemplazar el método con:
```vb
Private Sub Vencimiento2Txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call PasarFoco(Me.CopiasVentasCombo)
    End If
    
    If KeyAscii = vbKeyEscape Then
        Me.Vencimiento2Txt.Text = ""
    End If

    Call SoloNumeros(KeyAscii, False)
End Sub
```

### `TC4Chk_KeyPress`
Reemplazar el método con:
```vb
Private Sub TC4Chk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call PasarFoco(Me.CopiasComprasCombo)
    End If
End Sub
```

