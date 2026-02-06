Miércoles 04/02/2026

## `FormHost.bas`
Nuevo archivo

## `VehiculosCuentaResultadoFrm.frm`
Formulario nuevo
## `VehiculosFrm.frm`

Añadir
```vb
Private CuentaResultadoFrm As VehiculosCuentaResultadoFrm
```
- - -
Borrar
```vb
Private Enum ColGRC
    id_ColGRC = 0
    tipo_ColGRC = 1
    total_ColGRC = 2
    primeraColDinamica_ColGRC = 3
End Enum
```
- - -
Borrar
```vb
Private Enum Mes3L
    Ene = 1
    Feb = 2
    Mar = 3
    Abr = 4
    May = 5
    Jun = 6
    Jul = 7
    Ago = 8
    Sep = 9
    Oct = 10
    Nov = 11
    dic = 12
End Enum
```
### Controles
- [ ] Cambiar el caption del primer tab de `FrameCostes` a **"Cuenta Resultado"**
- [ ] Borrar `FrameResumenCostes`
- [ ] Insertar un  **`PictureBox`** con las siguientes propiedades:
	- [ ] Nombre: `PictureCuentaResultado`
	- [ ] Left: 75
	- [ ] Top: 465
	- [ ] BorderStyle: 0 - None

### `FrameCostes_Click`
Reemplazar
```vb
Call RellenarGridResumenCostes
Call RellenarFechaResumenCostes
```
Por
```vb
Call CuentaResultadoFrm.Mostrar
```

### `FrameCostes_Click`
Reemplazar
```vb
Call RellenarGridResumenCostes
Call RellenarFechaResumenCostes
```

Por
```vb
Call CuentaResultadoFrm.Mostrar
```

### `TabDatosInmovilizado_Click`
Reemplazar
```vb
Call RellenarGridResumenCostes
Call RellenarFechaResumenCostes
```

Por
```vb
Call CuentaResultadoFrm.Mostrar
```

### `GridDatosInmovilizado_Click`
Reemplazar
```vb
Call RellenarGridResumenCostes
Call RellenarFechaResumenCostes
```

Por
```vb
Call CuentaResultadoFrm.Mostrar
```

### `rellenarFrameDatosInmovilizados`
Debajo de
```vb
Me.LabelIdInmovilizado.Caption = Id
```

Añadir
```vb
CuentaResultadoFrm.IdInmovilizado = Id
```

### `Form_Load`
Debajo de
```vb
'Configurando formulario
If Me.WindowState <> 1 Then
    Me.Icon = EntornoFrm.ListaImagenesEntornoAgora.ListImages(3).Picture
    Call PosicionAbsolutaFormulario(Me)
End If
```

Añadir
```vb
Set CuentaResultadoFrm = New VehiculosCuentaResultadoFrm
Set CuentaResultadoFrm.Padre = Me
Call HostFormInContainer(CuentaResultadoFrm, Me.PictureCuentaResultado)
```

- - -
Borrar
```vb
Call StartBorderedGridSubClass(Me.GridResumenCostes)
```

- - -
Borrar
```vb
Me.FrameResumenCostes.BackColor = Me.FrameDatosInmovilizado.BackColor
```

- - -
Borrar
```vb
Me.FiltrosRCBtn.BackOver = Me.SalirBtn.BackOver
```

- - -
Borrar
```vb
Me.FiltrosRCBtn.ButtonType = Me.SalirBtn.ButtonType
```
- - -
```vb
Me.FiltrosRCBtn.ShowFocusRect = Me.SalirBtn.ShowFocusRect
```
- - -
Borrar todo el bloque separado por comentarios
```vb
Me.GridResumenCostes.Font = FontNameGrid
Me.GridResumenCostes.Font.Size = SizeFontGrid
Me.GridResumenCostes.Rows = 2
Me.GridResumenCostes.Cols = ColGRC.primeraColDinamica_ColGRC
Me.GridResumenCostes.FixedRows = 1
Me.GridResumenCostes.GridColorFixed = ColorBordeFrame
Me.GridResumenCostes.SelectionMode = flexSelectionByRow
Me.GridResumenCostes.BackColorFixed = ColorCabeceraGrid
Me.GridResumenCostes.RowHeightMin = TamañoMinimoFila
Me.GridResumenCostes.BackColorSel = ColorSeleccionGrid
Me.GridResumenCostes.ForeColorSel = ColorFuenteSeleccionGrid

Me.GridResumenCostes.ColWidth(ColGRC.id_ColGRC) = 0
Me.GridResumenCostes.ColWidth(ColGRC.tipo_ColGRC) = 3000
Me.GridResumenCostes.ColWidth(ColGRC.total_ColGRC) = 1200


Me.GridResumenCostes.ColAlignment(ColGRC.tipo_ColGRC) = 1
Me.GridResumenCostes.ColAlignment(ColGRC.total_ColGRC) = 7

Me.GridResumenCostes.FixedAlignment(ColGRC.tipo_ColGRC) = 4
Me.GridResumenCostes.FixedAlignment(ColGRC.total_ColGRC) = 4

Me.GridResumenCostes.Row = 0
Me.GridResumenCostes.Col = ColGRC.tipo_ColGRC
Me.GridResumenCostes.Text = "Mes"
Me.GridResumenCostes.CellFontBold = True
Me.GridResumenCostes.CellForeColor = Blanco
Me.GridResumenCostes.Col = ColGRC.total_ColGRC
Me.GridResumenCostes.Text = "Año"
Me.GridResumenCostes.CellFontBold = True
Me.GridResumenCostes.CellForeColor = Blanco
Me.GridResumenCostes.Row = 1
Me.GridResumenCostes.Col = 0
Me.GridResumenCostes.ColSel = Me.GridResumenCostes.Cols - 1
```

### `Form_Resize`
Añadir al final

Reemplazar
```vb
Me.FrameResumenCostes.Top = Me.FrameCostes.Top
Me.FrameResumenCostes.Left = Me.FrameCostes.Left
Me.FrameResumenCostes.Width = Me.FrameCostes.Width - 20
Me.FrameResumenCostes.Height = Me.FrameCostes.Height - Me.FrameFactura.Top
```

Por
```vb
Me.PictureCuentaResultado.Top = Me.FrameCostes.Top
Me.PictureCuentaResultado.Left = Me.FrameCostes.Left
Me.PictureCuentaResultado.Width = Me.FrameCostes.Width - 20
Me.PictureCuentaResultado.Height = Me.FrameCostes.Height - Me.FrameFactura.Top

If Not CuentaResultadoFrm Is Nothing Then
    Call ResizeHostedForm(CuentaResultadoFrm, Me.PictureCuentaResultado)
End If
```

- - -
Borrar
```vb
If Me.GridResumenCostes.Cols > 1 Then
    Me.FDesdeRCLabel.Top = 100
    Me.FDesdeRCLabel.Left = 125
    Me.FDesdeRCTxt.Top = Me.FDesdeRCLabel.Top
    Me.FDesdeRCTxt.Left = Me.FDesdeRCLabel.Left + Me.FDesdeRCLabel.Width + separacion1
    Me.FDesdeRCImg.Top = Me.FDesdeRCLabel.Top
    Me.FDesdeRCImg.Left = Me.FDesdeRCTxt.Left + Me.FDesdeRCTxt.Width + separacion1
    Me.FHastaRCLabel.Top = Me.FDesdeRCLabel.Top
    Me.FHastaRCLabel.Left = Me.FDesdeRCImg.Left + Me.FDesdeRCImg.Width + separacion2
    Me.FHastaRCTxt.Top = Me.FDesdeRCLabel.Top
    Me.FHastaRCTxt.Left = Me.FHastaRCLabel.Left + Me.FHastaRCLabel.Width + separacion1
    Me.FHastaRCImg.Top = Me.FDesdeRCLabel.Top
    Me.FHastaRCImg.Left = Me.FHastaRCTxt.Left + Me.FHastaRCTxt.Width + separacion1
    Me.FiltrosRCBtn.Top = Me.FDesdeRCLabel.Top
    Me.FiltrosRCBtn.Left = Me.FHastaRCImg.Left + Me.FHastaRCImg.Width + separacion2


    Me.GridResumenCostes.Top = Me.FDesdeRCLabel.Top + Me.FDesdeRCLabel.Height + 300
    Me.GridResumenCostes.Left = 125
    Me.GridResumenCostes.Width = Me.FramePersonal.Width - 275
    Me.GridResumenCostes.Height = Me.FramePersonal.Height - 325 - Me.FDesdeRCLabel.Top - Me.FDesdeRCLabel.Height
End If
```


### `CargaResaltesDeTextos`
Borrar
```vb
Call ConfigurarResalteTexto(Me.FDesdeRCTxt, Me.ShapeResumenCostes)
Call ConfigurarResalteTexto(Me.FHastaRCTxt, Me.ShapeResumenCostes)
```

- - -

## Borrar métodos
- [ ] FDesdeRCTxt_GotFocus
- [ ] FDesdeRCImg_Click
- [ ] FDesdeRCTxt_KeyPress
- [ ] FDesdeRCTxt_LostFocus
- [ ] RellenarGridResumenCostes
- [ ] rellenarfecharesumencostes
- [ ] FHastaRCImg_Click
- [ ] FHastaRCTxt_GotFocus
- [ ] FHastaRCTxt_KeyPress
- [ ] FHastaRCTxt_LostFocus
- [ ] FiltrosRCBtn_Click
- [ ] NzDbl
