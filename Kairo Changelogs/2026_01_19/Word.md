Lunes 19/01/2026 - Jueves 22/01/2026

## `ConfiguracionImpresionFrm.frm`

Volver a una versión anterior modificada, añadir multiempresa a esa versión.
- - -
## `Procesos.bas`
Añadir los siguientes enums
```vb
Public Enum EModoSalidaImpresion
    ModoImpresora = 0
    ModoPdf = 1
    ModoDocumento = 2
End Enum
```
- - -
### `ImprimirDocumentoVentaCompraWord`
Método nuevo, debajo de `ImprimirDocumentoVentaCompra`
- - -
### `RellenarDocumentoVentaCompra`
Método nuevo, debajo del anterior
- - -
### `ObtenerConceptoDetalle`
Método nuevo, debajo del anterior
- - -
### `CalculaPosicionImpresion`
Reemplazar por
```vb
Public Sub CalculaPosicionImpresion(ByRef ColumnaImpresion, ByRef FilaImpresion, ByRef CampoVisible, ByVal TipoDocumento As Byte, ByVal TipoLinea As Byte, Optional IdPlantilla As Long = 0, Optional ByRef Marcador As String = "")

Dim MiRc As Recordset

'Reajustando tipo de documento si es necesarios
Select Case TipoDocumento
    Case 5  'Pedido compra
        TipoDocumento = 6
    Case 6  'Albarán compra
        TipoDocumento = 7
    Case 7  'Factura compra
        TipoDocumento = 8
End Select

'Estableciendo posicion
ColumnaImpresion = 0
FilaImpresion = 0
CampoVisible = False

If IdPlantilla = 0 Then
    Set MiRc = Base.OpenRecordset("Select Visible,Columna,Fila,Marcador" _
        & " from Configuracion_Impresion" _
        & " where Tipo_Documento=" & TipoDocumento & " and Tipo_Linea=" & TipoLinea)
Else
    Set MiRc = Base.OpenRecordset("Select Visible,Columna,Fila,Marcador" _
        & " from Configuracion_Impresion" _
        & " where Tipo_Linea=" & TipoLinea _
        & " and Id_Plantilla_Impresion=" & IdPlantilla)
End If

If MiRc.EOF = False Then
    
    MiRc.MoveFirst
    If IsNull(MiRc!Visible) = False Then CampoVisible = MiRc!Visible
    If IsNull(MiRc!Columna) = False Then ColumnaImpresion = MiRc!Columna
    If IsNull(MiRc!Fila) = False Then FilaImpresion = MiRc!Fila
    If IsNull(MiRc!Marcador) = False Then Marcador = MiRc!Marcador
End If

MiRc.Close

End Sub
```
- - -
## `DocumentoCompraVenta.cls`
Archivo nuevo

- - -
## `DocumentoCompraVentaWord.bas`
Archivo nuevo

- - -
## `DocumentosVentaFrm.frm`

### `ImprimirDocumentoBtn_Click`
**==Renombrar==** a **`ImprimirExcel`**, luego poner el siguiente método arriba 
```vb
Private Sub ImprimirDocumentoBtn_Click()
    Dim IdPlantilla As Double
    Dim EsPorDefecto As Boolean
    Dim RsPlantilla As Recordset
    Dim Extension As String
    
    ' Hay que determinar el tipo de plantilla seleccionada (Excel o Word)
    IdPlantilla = 0
    If Me.PlantillaImpresionComho.ListIndex <> -1 Then
        IdPlantilla = IdPlantillaImpresion(Me.PlantillaImpresionComho.ListIndex)
    End If
    
    ' Todas las por defecto son Excel
    If IdPlantilla = 0 Then
        Call ImprimirExcel
        Exit Sub
    End If
    
    Set RsPlantilla = Base.OpenRecordset( _
        " SELECT Id, Archivo, Por_Defecto" & _
        " FROM Plantilla_Impresion" & _
        " WHERE Id = " & IdPlantilla _
    )
    
    If Not RsPlantilla.EOF Then
        Call RsPlantilla.MoveFirst
        
        Extension = LCase$(Mid$(RsPlantilla!Archivo, InStrRev(RsPlantilla!Archivo, ".") + 1))
        
        Select Case Extension
            Case "xls", "xlsx", "xlt", "xltx", "xlsm", "xltm"
                Call ImprimirExcel
            Case "doc", "docx", "dot", "dotx", "docm", "dotm"
                Call ImprimirWord
            Case Else
                Call Err.Raise(5, "ImprimirDocumentoBtn_Click", "Extension de plantilla no admitida: " & Extension)
        End Select
    Else
        Call ImprimirExcel
    End If
    
    Call RsPlantilla.Close
    Set RsPlantilla = Nothing
End Sub
```

- - -
### `ImprimirWord`
Función nueva, debajo de la anterior
```vb
Private Sub ImprimirWord()

Dim MiRc As Recordset
Dim carpeta As String
Dim CarpetaSecundaria As String
Dim fso As FileSystemObject
Dim Archivo As String
Dim Objeto As String
Dim Mensaje As String
Dim IdPlantilla As Double
Dim Copias As Long

Me.MousePointer = 11

Me.Label7.Visible = True

'Estableciendo carpeta
carpeta = ""
CarpetaSecundaria = ""
Archivo = ""
Objeto = ""
If Trim(Me.CarpetaDestinoTxt.Text) = "" Then
    Select Case TipoDocumento
        Case 1  'Presupuesto
            carpeta = App.Path & "\Presupuestos_Cliente\"
            Archivo = App.Path & "\Oferta.pdf"
            Objeto = "Información Oferta Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
            Mensaje = "Adjuntamos Oferta Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
        Case 2  'Pedido
            carpeta = App.Path & "\Pedidos_Cliente\"
            Archivo = App.Path & "\Pedido.pdf"
            Objeto = "Información Pedido Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
            Mensaje = "Adjuntamos Pedido Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
        Case 3  'Albarán
            carpeta = App.Path & "\Albaranes_Cliente\"
            Archivo = App.Path & "\Albarán.pdf"
            Objeto = "Información Albarán Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
            Mensaje = "Adjuntamos Albarán Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
        Case 4  'Factura
            carpeta = App.Path & "\Facturas_Cliente\"
            Archivo = App.Path & "\Factura.pdf"
            Objeto = "Información Factura Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
            Mensaje = "Adjuntamos Factura Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
    End Select
Else
    carpeta = Trim(Me.CarpetaDestinoTxt.Text) & "\"
    CarpetaSecundaria = carpeta
End If

'Estableciendo plantilla
IdPlantilla = 0
If Me.PlantillaImpresionComho.ListIndex <> -1 Then IdPlantilla = IdPlantillaImpresion(Me.PlantillaImpresionComho.ListIndex)

'Informando a Hacienda (Verifactu)
If VerifactuActivado = True And TipoDocumento = 4 And IdDocumento <> 0 Then
    If Me.ImgVerifactu(0).Visible = False Then     'No está registrada ya sea porque no se ha hecho, ya sea por error
        'Informando
        Me.Label7.Caption = "Comunicando Hacienda"
        Me.Label7.Refresh
        'Comunicando
        Call EnvioFacturaVerifactuBase64(IdDocumento)
        DoEvents
        'Actualizando
        Me.ImgVerifactu(0).Visible = False
        Me.ImgVerifactu(1).Visible = False
        Set MiRc = Base.OpenRecordset("Select Subido_Hacienda_Verifactu,Mensaje_Error_Verifactu" _
            & " from Cabecera_Documentos_Venta" _
            & " where Id=" & IdDocumento)
        If MiRc.EOF = False Then
            If IsNull(MiRc!Subido_Hacienda_Verifactu) = False And MiRc!Subido_Hacienda_Verifactu = True Then
                Me.ImgVerifactu(0).Visible = True
            Else
                If IsNull(MiRc!Subido_Hacienda_Verifactu) = False And MiRc!Subido_Hacienda_Verifactu = False Then
                    If IsNull(MiRc!Mensaje_Error_Verifactu) = False Then
                        Me.ImgVerifactu(1).Visible = True
                    End If
                End If
            End If
        End If
        MiRc.Close
    End If
End If

'Imprimiendo
Me.Label7.Caption = "Imprimiendo"
Me.Label7.Refresh

Copias = Me.CopiasCombo.ListIndex + 1

If Me.ImpresoraChk.value = vbChecked Then
    Call ImprimirDocumentoVentaCompraWord(TipoDocumento, IdDocumento, ModoImpresora, Me, Archivo, IdPlantillaUsuario:=IdPlantilla, nombreImpresora:=Me.ImpresorasCombo.Text, CantidadCopias:=Copias)
ElseIf Me.PdfChk.value = vbChecked Then
    Call ImprimirDocumentoVentaCompraWord(TipoDocumento, IdDocumento, ModoPdf, Me, Archivo, IdPlantillaUsuario:=IdPlantilla, CantidadCopias:=Copias)
ElseIf Me.ExcelChk.value = vbChecked Then
    Call ImprimirDocumentoVentaCompraWord(TipoDocumento, IdDocumento, ModoDocumento, Me, Archivo, IdPlantillaUsuario:=IdPlantilla)
ElseIf Me.MailChk.value = vbChecked Then
    If Trim(Me.DireccionCorreoTxt.Text) <> "" Then
        If MsgBox("¿Desea enviar el documento por correo electrónico?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirme Envío") = vbYes Then
            Call ImprimirDocumentoVentaCompraWord(TipoDocumento, IdDocumento, ModoPdf, Me, Archivo, IdPlantillaUsuario:=IdPlantilla)
            Call EnviarCorreo(Trim(Me.DireccionCorreoTxt.Text), "", "", Objeto, Mensaje, Archivo)
        End If
    End If
End If

'Saliendo
Me.MousePointer = 0

Call SalirImpresionBtn_Click

If CarpetaSecundaria = "" Then MsgBox "Impresión Finalizada", vbInformation, "Impresión Documentos"

'Configuración TPV
If FormatoTPV = True Then Call NuevoDocumentoBtn_Click

End Sub
```

- - -
## `DocumentosCompraFrm.frm`
### `ImprimirDocumentoBtn_Click`
**==Renombrar==** a **`ImprimirExcel`**, luego poner el siguiente método arriba

```vb
Private Sub ImprimirDocumentoBtn_Click()
    Dim IdPlantilla As Double
    Dim EsPorDefecto As Boolean
    Dim RsPlantilla As Recordset
    Dim Extension As String
    
    ' Hay que determinar el tipo de plantilla seleccionada (Excel o Word)
    IdPlantilla = 0
    If Me.PlantillaImpresionComho.ListIndex <> -1 Then
        IdPlantilla = IdPlantillaImpresion(Me.PlantillaImpresionComho.ListIndex)
    End If
    
    ' Todas las por defecto son Excel
    If IdPlantilla = 0 Then
        Call ImprimirExcel
        Exit Sub
    End If
    
    Set RsPlantilla = Base.OpenRecordset( _
        " SELECT Id, Archivo, Por_Defecto" & _
        " FROM Plantilla_Impresion" & _
        " WHERE Id = " & IdPlantilla _
    )
    
    If Not RsPlantilla.EOF Then
        Call RsPlantilla.MoveFirst
        
        Extension = LCase$(Mid$(RsPlantilla!Archivo, InStrRev(RsPlantilla!Archivo, ".") + 1))
        
        Select Case Extension
            Case "xls", "xlsx", "xlt", "xltx", "xlsm", "xltm"
                Call ImprimirExcel
            Case "doc", "docx", "dot", "dotx", "docm", "dotm"
                Call ImprimirWord
            Case Else
                Call Err.Raise(5, "ImprimirDocumentoBtn_Click", "Extension de plantilla no admitida: " & Extension)
        End Select
    Else
        Call ImprimirExcel
    End If
    
    Call RsPlantilla.Close
    Set RsPlantilla = Nothing
End Sub
```

- - -

### `ImprimirWord`
Función nueva, debajo de la anterior
```vb
Private Sub ImprimirWord()

Dim carpeta As String
Dim CarpetaSecundaria As String
Dim fso As FileSystemObject
Dim Archivo As String
Dim Objeto As String
Dim Mensaje As String
Dim TipoDocumentoImpresion As Byte
Dim IdPlantilla As Double
Dim Copias As Long

Me.MousePointer = 11

Me.Label7.Visible = True

'Estableciendo tipo de documento de impresión
Select Case TipoDocumento
    Case 2  'Pedido
        TipoDocumentoImpresion = 5
    Case 3  'Albaran
        TipoDocumentoImpresion = 6
    Case 4  'Factura
        TipoDocumentoImpresion = 7
End Select

'Estableciendo carpeta
carpeta = ""
CarpetaSecundaria = ""
Archivo = ""
Objeto = ""
If Trim(Me.CarpetaDestinoTxt.Text) = "" Then
    Select Case TipoDocumento
        Case 1  'Presupuesto
            carpeta = App.Path & "\Presupuestos_Proveedor\"
            Archivo = App.Path & "\Oferta_Proveedor.pdf"
            Objeto = "Información Presupuesto Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
            Mensaje = "Adjuntamos Presupuesto Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
        Case 2  'Pedido
            carpeta = App.Path & "\Pedidos_Proveedor\"
            Archivo = App.Path & "\Pedido_Proveedor.pdf"
            Objeto = "Información Pedido Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
            Mensaje = "Adjuntamos Pedido Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
        Case 3  'Albarán
            carpeta = App.Path & "\Albaranes_Proveedor\"
            Archivo = App.Path & "\Albarán_Proveedor.pdf"
            Objeto = "Información Albarán Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
            Mensaje = "Adjuntamos Albarán Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
        Case 4  'Factura
            carpeta = App.Path & "\Facturas_Proveedor\"
            Archivo = App.Path & "\Factura_Proveedor.pdf"
            Objeto = "Información Factura Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
            Mensaje = "Adjuntamos Factura Nº" & Me.SeriesCombo.Text & "/" & Trim(str(Me.NumeroDocumentoTxt.Text))
    End Select
Else
    carpeta = Trim(Me.CarpetaDestinoTxt.Text) & "\"
    CarpetaSecundaria = carpeta
End If

'Estableciendo plantilla
IdPlantilla = 0
If Me.PlantillaImpresionComho.ListIndex <> -1 Then IdPlantilla = IdPlantillaImpresion(Me.PlantillaImpresionComho.ListIndex)

'Imprimiendo
Me.Label7.Caption = "Imprimiendo"
Me.Label7.Refresh

Copias = Me.CopiasCombo.ListIndex + 1

If Me.ImpresoraChk.value = vbChecked Then
    Call ImprimirDocumentoVentaCompraWord(TipoDocumento, IdDocumento, ModoImpresora, Me, Archivo, IdPlantillaUsuario:=IdPlantilla, nombreImpresora:=Me.ImpresorasCombo.Text, CantidadCopias:=Copias)
ElseIf Me.PdfChk.value = vbChecked Then
    Call ImprimirDocumentoVentaCompraWord(TipoDocumento, IdDocumento, ModoPdf, Me, Archivo, IdPlantillaUsuario:=IdPlantilla, CantidadCopias:=Copias)
ElseIf Me.ExcelChk.value = vbChecked Then
    Call ImprimirDocumentoVentaCompraWord(TipoDocumento, IdDocumento, ModoDocumento, Me, Archivo, IdPlantillaUsuario:=IdPlantilla)
ElseIf Me.MailChk.value = vbChecked Then
    If Trim(Me.DireccionCorreoTxt.Text) <> "" Then
        If MsgBox("¿Desea enviar el documento por correo electrónico?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirme Envío") = vbYes Then
            Call ImprimirDocumentoVentaCompraWord(TipoDocumento, IdDocumento, ModoPdf, Me, Archivo, IdPlantillaUsuario:=IdPlantilla)
            Call EnviarCorreo(Trim(Me.DireccionCorreoTxt.Text), "", "", Objeto, Mensaje, Archivo)
        End If
    End If
End If

'Abriendo carpeta si es necesario
If CarpetaSecundaria <> "" Then ShellExecute 0, "Open", CarpetaSecundaria, "", "", 1

'Saliendo
Me.MousePointer = 0

Call SalirImpresionBtn_Click

If CarpetaSecundaria = "" Then MsgBox "Impresión Finalizada", vbInformation, "Impresión Documentos"

End Sub
```
