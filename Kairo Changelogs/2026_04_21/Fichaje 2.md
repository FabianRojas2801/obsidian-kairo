Martes 21/04/2026
## `FichaPersonalFrm`
Remplazar
```vb
"WHERE user_name = '" & Replace(Trim(Me.UsuarioAPPTxt.Text), "'", "''") & "'", _
```
Por
```vb
"WHERE user_name = '" & Replace(Trim(Me.UsuarioAPPTxt.Text), "'", "''") & "' " & _
"AND id_empresa = " & obtenerIdsEmpresaApp(CStr(IdEmpresaCombo(Me.EmpresaCombo.ListIndex))), _
```

## `Procesos_App.bas`
Remplazar
```vb
"WHERE user_name = '" & Replace(UsuarioAPPTxt, "'", "''") & "'"
```
Por
```vb
"WHERE user_name = '" & Replace(UsuarioAPPTxt, "'", "''") & "' " & _
"AND id_empresa = " & obtenerIdsEmpresaApp(IdEmpresaGeneral)
```

