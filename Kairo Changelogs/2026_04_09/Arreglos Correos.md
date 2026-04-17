 (MViernes 10/04/2026
## `DocumentosVentaFrm`

Remplazar
```vb
If IsNull(MiRc!Mail_Contacto) = False Then Me.MailContactoTxt.Text = Trim(MiRc!Mail_Contacto)
```

Por
```vb
Me.MailContactoTxt.Text = IIF(IsNull(MiRc!Mail_Contacto), "", Trim(MiRc!Mail_Contacto))
```

- - -

Remplazar
```vb
If IsNull(MiRc!EMail) = False Then Me.MailContactoTxt.Text = Trim(MiRc!EMail)
```

Por
```vb
Me.MailContactoTxt.Text = IIf(IsNull(MiRc!EMail), "", Trim(MiRc!EMail))
```
