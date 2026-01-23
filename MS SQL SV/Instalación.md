## Instalar Microsoft SQL Express

## Habilitar conexiones
Entrar a **SQL Server Configuration Manager**
- [ ] Bajo **Servicios de SQL Server** habilitar **SQL Server Browser**

Entrar a **Configuración de red de SQL Server/Protocolos de SQLEXPRESS** y abrir **TCP/IP**
- [ ] Bajo **Protocolo** habilitar **Escuchar todo** y **Habilitado**

Abrir el CMD como administrador y ejecutar
```
netsh advfirewall firewall add rule name="SQL Server TCP 1433" dir=in action=allow protocol=TCP localport=1433
```

