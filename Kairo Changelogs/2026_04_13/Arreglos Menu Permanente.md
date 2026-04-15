Lunes 13/04/2026

## `MenuPermanenteFrm`
### (General)
Añadir
```vb
Private RelacionAspectoLogo As Double
Private AnchoLogo As Double
Private HayAccesoRapido As Boolean
```

## `Form_Load`
Debajo de
```vb
PosicionSuperiorInicial = 100
```

Añadir
```vb
RelacionAspectoLogo = 1
AnchoLogo = 100
HayAccesoRapido = False
```

## `CargarImagenCliente`

Remplazar
```vb
If ArchivoTxt = "" Then Exit Sub
```

Por
```vb
If ArchivoTxt = "" Then
    Call Form_Resize
    Exit Sub
End If
```

- - -
Remplazar
```vb
    On Error Resume Next
    Set Me.LogoCliente.Picture = LoadPicture(ArchivoTxt)
    On Error GoTo 0
```

Por
```vb
    On Error Resume Next
    Me.LogoCliente.Stretch = False
    Set Me.LogoCliente.Picture = LoadPicture(ArchivoTxt)
    RelacionAspectoLogo = Me.LogoCliente.Picture.Width / Me.LogoCliente.Picture.Height
    AnchoLogo = Me.LogoCliente.Width
    Me.LogoCliente.Stretch = True
    On Error GoTo 0
```


## `Form_Resize`

Remplazar el método por:
```vb
Public Sub Form_Resize()
    Const AlturaPersonalizadaMaximaPorcentaje As Double = 0.3
    Dim EspacioFlexibleDisponible As Double
    
    Call PosicionarRecuadro(30, 0, Me.Width - 30, Me.Height)
    
    If Me.LogoCliente.Visible Then
        Me.LogoCliente.Width = Min(AnchoLogo, Me.Width - 60)
        Me.LogoCliente.Left = (Me.ScaleWidth - Me.LogoCliente.Width) / 2
        
        Me.LogoCliente.Height = Me.LogoCliente.Width / RelacionAspectoLogo
        Me.LogoCliente.Top = Me.ScaleHeight - Me.LogoCliente.Height - 300
    End If
    
    Me.Label1.Top = Me.Height - Me.Label1.Height - 50
    
    ' Posicion de los tabs
    EspacioFlexibleDisponible = Me.Height
    
    If Me.LogoCliente.Visible Then
        EspacioFlexibleDisponible = EspacioFlexibleDisponible - (Me.Height - Me.LogoCliente.Top) - 100
    Else
        EspacioFlexibleDisponible = EspacioFlexibleDisponible - (Me.Height - Me.Label1.Top) - 100
    End If
    
    ' Menu personalizado
    If Me.MenuPersonalizadoTab.Visible Then
        Me.MenuPersonalizadoTab.Left = (Me.ScaleWidth - Me.MenuPersonalizadoTab.Width) / 2
    
        Me.MenuPersonalizadoTab.Top = Me.Recuadro(0).Top + 50
        EspacioFlexibleDisponible = EspacioFlexibleDisponible - Me.MenuPersonalizadoTab.Top
        
        If HayAccesoRapido Then
            Me.GriAccesosDirectos.Height = TamañoMinimoFila * Me.GriAccesosDirectos.Rows
            Me.GriAccesosDirectos.Visible = True
            Me.MenuPersonalizadoTab.Height = Me.GriAccesosDirectos.Height + 510
        Else
            Me.GriAccesosDirectos.Height = 0
            Me.GriAccesosDirectos.Visible = False
            Me.MenuPersonalizadoTab.Height = TamañoMinimoFila * 1.65
        End If
        
        EspacioFlexibleDisponible = EspacioFlexibleDisponible - Me.MenuPersonalizadoTab.Height
        
        ' Espaciado de 100 del menú normal
        Me.MenuTab.Top = Me.MenuPersonalizadoTab.Top + Me.MenuPersonalizadoTab.Height + 100
        EspacioFlexibleDisponible = EspacioFlexibleDisponible - 100
    Else
        Me.MenuTab.Top = Me.Recuadro(0).Top + 50
        EspacioFlexibleDisponible = EspacioFlexibleDisponible - 50
    End If
    
    ' Menú normal
    Me.MenuTab.Left = (Me.ScaleWidth - Me.MenuTab.Width) / 2
    Me.MenuTab.Height = EspacioFlexibleDisponible
End Sub
```

### `CargarMenuPersonalizado`
Debajo de 
```vb
Me.GriAccesosDirectos.Redraw = True
```

Añadir
```vb
HayAccesoRapido = Existe
```
- - -
Eliminar
```vb
If Existe = False Then
    'Estableciendo tamaño final del contenedor
    Posicion = TamañoMinimoFila * 1.65
    Me.GriAccesosDirectos.Height = 0
    Me.GriAccesosDirectos.Visible = False
    Me.MenuPersonalizadoTab.Height = Posicion
Else
    'Estableciendo tamaño final del contenedor
    Posicion = Posicion * TamañoMinimoFila
    Me.GriAccesosDirectos.Height = Posicion
    Me.MenuPersonalizadoTab.Height = Me.GriAccesosDirectos.Height + 510
    Me.GriAccesosDirectos.Visible = True
End If
```
- - -
Debajo de 
```vb
Me.MenuPersonalizadoTab.Visible = True
```

Añadir
```vb
Call Form_Resize
```
### `PosicionarRecuadro`
Nuevo método
```vb
Private Sub PosicionarRecuadro(ByVal Left As Double, ByVal Top As Double, ByVal Width As Double, ByVal Height As Double)
    ' Shape superior
    Me.Recuadro(0).Top = Top
    Me.Recuadro(0).Left = Left
    Me.Recuadro(0).Width = Width
    
    ' Shape inferior
    Me.Recuadro(2).Top = Height - Me.Recuadro(2).Height
    Me.Recuadro(2).Left = Left
    Me.Recuadro(2).Width = Width
    
    ' Shape del medio
    Me.Recuadro(1).Top = Me.Recuadro(0).Top + Me.Recuadro(0).Height / 2
    Me.Recuadro(1).Left = Left
    Me.Recuadro(1).Width = Width
    Me.Recuadro(1).Height = Max(Me.Recuadro(2).Top - Me.Recuadro(1).Top + Me.Recuadro(2).Height / 2, 1)
End Sub
```

## `cGridManager`

### `ResizeCols`

Remplazar encabezado por
```vb
Public Sub ResizeCols(Optional ByVal MargenScrollbar As Long = 275)
```
- - -
Remplazar
```vb
AnchoVariable = (Disponible \ cantidadVariables) - (275 \ cantidadVariables)
```

Por
```vb
AnchoVariable = (Disponible \ cantidadVariables) - (MargenScrollbar \ cantidadVariables)
```

