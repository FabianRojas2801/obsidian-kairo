## Requisitos locales
- [ ] **Microsoft SQL Server Express** instalado y conectado mediante **SQL Server Management Studio**
- [ ] **[Migrador Kairo](https://github.com/Devs-Kong-Software/MigradorKairo)** con el worker registrado
## Requisitos cliente
- [ ] **Microsoft SQL Server Express** instalado y conectado mediante **SQL Server Management Studio**

> [!info]
> La conexión a SQL Server se usará para ejecutar un script de SQL, no hace falta conectarse directamente desde la oficina.
> Con tener acceso remoto y poder pasar el script SQL basta.

## Instalación y configuración de Microsoft SQL Server Express

> [!info]
> Guía tanto para local como para el servidor destino

### Descarga e instalación
Dirígete a https://www.microsoft.com/es-es/sql-server/sql-server-downloads, descarga e instala **SQL Server 2025 Express** o **Desarrollador de SQL Server 2025**.
![[brave_djfm0AjCY2.png]]

### Habilitar conexión
Entra a **SQL Server Configuration Manager**
- [ ] Bajo **Servicios de SQL Server** habilita **SQL Server Browser**

Entra a **Configuración de red de SQL Server/Protocolos de SQLEXPRESS** y abre **TCP/IP** (doble click)
- [ ] Bajo **Protocolo** habilita **Escuchar todo** y **Habilitado**

Abre CMD como administrador
- [ ]  Ejecuta
```
netsh advfirewall firewall add rule name="SQL Server TCP 1433" dir=in action=allow protocol=TCP localport=1433
```

Conéctate a SQL mediante el **SQL Server Management Studio** y entra a las propiedades del servidor.
![[SSMS_lz5mZvV3EM.png]]
Ve a **Seguridad** y selecciona **Modo de autenticación de Windows y SQL Server**
![[SSMS_pEmkiSm57m.png]]

Ahora reinicia el ordenador.
## Crear usuario y base de datos "Kairo" en SQL Local
Ejecutar el siguiente script en Microsoft SQL Server Express
![[kwqFVih6aC.png]]

> [!caution]
> Tienes que reemplazar `CONTRASEÑA DE DATOS.MDB` por la contraseña de **Datos.mdb**

> [!warning]
> Si ya existe una base de datos llamada Kairo, hay que borrarla manualmente (se perderán datos)
> ```sql
> USE [master];
> DROP DATABASE Kairo;
> ```


```sql
-- CREAR BASE
USE [master];
IF DB_ID('Kairo') IS NULL
BEGIN
    CREATE DATABASE [Kairo];
END;
GO

USE [Kairo];
GO

-- CREAR LOGIN
IF EXISTS (SELECT 1 FROM sys.server_principals WHERE name = 'Kairo')
BEGIN
    DROP LOGIN [Kairo];
END;
CREATE LOGIN [Kairo] WITH PASSWORD = N'CONTRASEÑA DE DATOS.MDB';
GO

-- CREAR USUARIO
IF EXISTS (SELECT 1 FROM sys.database_principals WHERE name = 'Kairo')
BEGIN
    DROP USER [Kairo];
END;
CREATE USER [Kairo] FOR LOGIN [Kairo];
GO

-- PERMISOS
USE [Kairo];
EXEC sp_addrolemember N'db_owner', N'Kairo'
```

Ahora será posible iniciar sesión a la tabla `Kairo` con el usuario `Kairo` y la contraseña especificada.

## Configuración del ODBC

> [!info]
> Guía tanto para local como para el servidor destino

Abre el **Administrador de origen de datos ODBC (32 bits)** (`odbcad32.exe`), el migrador facilita un botón para abrirlo directamente.
![[Obsidian_zCd1PFul8Z.png]]![[1wTm6gaNlb.png]]

Añadir un nuevo ODBC con el controlador **`ODBC Driver 18 for SQL Server`**.
El servidor para SQL local es `localhost\sqlexpress`, para un SQL en otro ordenador hay que poner la IP, por ejemplo `192.168.0.22`

![[WBwPTEHbIc.png]]
![[odbcad32_njH7wtj2TD.png]]
![[odbcad32_TMreE54eln.png]]

Presionando Siguiente podremos especificar el método de autenticación, en nuestro caso será **Con la autenticación de SQL Server**![[odbcad32_Y6Whj1VCBn.png]]

En la siguiente pantalla dejamos todo por defecto. En la pantalla final marcamos **Confiar en el certificado del servidor.**