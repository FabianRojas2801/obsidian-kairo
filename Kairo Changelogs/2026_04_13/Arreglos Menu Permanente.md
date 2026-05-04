
## `MenuSuperior.frm`
Debajo de
```vb
MenuPermanenteFrm.GridMenu.Redraw = True
```

Añadir
```vb
Call GridArrows.RefreshGridArrows(MenuPermanenteFrm.GridMenu)
```

(dos veces)
## `MenuPermanenteFrm.frm`
> [!warning] Demasiados cambios

> [!info]
> La altura máxima de la tabla de Accesos directo se puede cambiar alterando `AlturaPersonalizadaMaximaPorcentaje` que es un porcentaje (de 0.0 a 1.0).


## `GridArrow.bas`
Archivo nuevo.

## `MouseWheelHoverScroll.bas`
Archivo nuevo.

