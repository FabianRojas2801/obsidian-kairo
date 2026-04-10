Viernes 10/04/2026
## `ContratosModalFrm`
### `BorrarContratoBtn_Click`
Borrar la siguiente línea
```vb
MsgBox "Contrato eliminado", vbInformation
```


### `Form_Load`
Debajo de
```vb
Call ContratosGm.ApplyCols(0)
```

Añadir
```
Call StartBorderedGridSubClass(Me.GridContratos)
```
- - -
Debajo de
```vb
Call ProyectosGm.ApplyCols(0)
```

Añadir
```
Call StartBorderedGridSubClass(Me.GridContratos)
```
- - -
Debajo de
```vb
Call TotalesGm.ApplyCols(0)
```

Añadir
```
Call StartBorderedGridSubClass(Me.GridContratos)
```
