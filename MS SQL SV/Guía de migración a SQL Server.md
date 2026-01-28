Se puede migrar de distintas maneras
1. Copiando `Datos.mdb` al ordenador local, ejecutando el migrador localmente y exportar la base de datos creada para importarla en el servidor del cliente
2. Ejecutando el migrador directamente en el servidor del cliente
3. Copiando `Datos.mdb` al ordenador local, ejecutando el migrador localmente apuntando al servidor del cliente (esto requiere acceso directo al servidor del cliente desde la oficina, lo que es muy poco probable, sin embargo la opción existe)

> [!info]
> La migración puede ejecutarse con éxito, pero fallar al exportarse o importarse. En tal caso el migrador deberá escribir al SQL Server destino directamente
## Requisitos
- **Microsoft SQL Server Express** instalado y conectado mediante **SQL Server Management Studio**
- **[.NET 10.0 Desktop Runtime](https://dotnet.microsoft.com/es-es/download/dotnet/thank-you/runtime-desktop-10.0.2-windows-x86-installer?cid=getdotnetcore)** instalado (requisito del migrador)
- **[Migrador Kairo](https://github.com/Devs-Kong-Software/MigradorKairo)** con el worker registrado
- Un **ODBC** apuntando al SQL Server (hay que instalar **ODBC Driver for SQL Server 18**)

> [!warning] Aviso
> El ODBC debe estar configurado en cada ordenador que corra Kairo, no solo en ordenador que ejecute el migrador

## Instalación y configuración de Microsoft SQL Server Express
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
## Crear usuario y base de datos "Kairo"
Ejecutar el siguiente script en Microsoft SQL Server Express
![[kwqFVih6aC.png]]

> [!caution] Atención
> Tienes que reemplazar `CONTRASEÑA DE DATOS.MDB` por la contraseña de **Datos.mdb**

> [!warning] Advertencia
> Si ya existe una base de datos llamada **Kairo**, para que el script pueda borrar al usuario Kairo hay que borrar la base de datos Kairo.
> 
> ```sql
> USE [master];
> DROP DATABASE Kairo;
> ```

```sql
USE [master];
GO

IF DB_ID('Kairo') IS NULL
BEGIN
    CREATE DATABASE [Kairo];
END
GO

IF EXISTS (SELECT 1 FROM sys.server_principals WHERE name = 'Kairo')
BEGIN
    DROP LOGIN [Kairo];
END
GO

CREATE LOGIN [Kairo]
WITH PASSWORD = 'CONTRASEÑA DE DATOS.MDB',
     CHECK_POLICY = OFF,
     CHECK_EXPIRATION = OFF,
     DEFAULT_DATABASE = [Kairo];
GO

USE [Kairo];
GO

IF EXISTS (SELECT 1 FROM sys.database_principals WHERE name = 'Kairo')
BEGIN
    DROP USER [Kairo];
END
GO

CREATE USER [Kairo] FOR LOGIN [Kairo];
EXEC sp_addrolemember N'db_owner', N'Kairo';
GO

ALTER LOGIN [Kairo] ENABLE;
GO

```

Ahora será posible iniciar sesión a la tabla `Kairo` con el usuario `Kairo` y la contraseña especificada.

## Configuración del ODBC

> [!info]
> Se recomienda realizar la migración a un servidor local y luego exportar la base de datos en forma de script desde el Management Studio.
> Sin embargo, el migrador tiene la capacidad de realizar la migración a una base de datos remota si hiciera falta.

Abre el **Administrador de origen de datos ODBC (32 bits)** (`odbcad32.exe`), el migrador facilita un botón para abrirlo directamente.
![[Obsidian_zCd1PFul8Z.png]]![[1wTm6gaNlb.png]]

Añadir un nuevo ODBC con el controlador **`ODBC Driver 18 for SQL Server`**.

> [!info]
> Si no aparece la opción, hace falta instalar **[ODBC Driver for SQL Server](https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server)**.
> **En un Windows de 64 bits hay que instalar la versión de 64 bits ya que incluye el controlador de 32 bits también.**

![[WBwPTEHbIc.png]]![[odbcad32_njH7wtj2TD.png]]

El servidor para SQL local es `localhost\sqlexpress`, para un SQL en otro ordenador hay que poner la IP, por ejemplo `192.168.0.22`
![[odbcad32_TMreE54eln.png]]

Presionando Siguiente podremos especificar el método de autenticación.
En el caso de una conexión remota, el método será **Con la autenticación de SQL Server**.
Si el servidor está en localhost, la **Autenticación integrada de Windows** debería funcionar.
![[odbcad32_Y6Whj1VCBn.png]]

En la siguiente pantalla dejamos todo por defecto. En la pantalla final marcamos **Confiar en el certificado del servidor.**
![[odbcad32_xsAFFjCySB.png]]

Al presionar Finalizar podremos comprobar si la conexión tiene éxito
![[odbcad32_YiE1mLUdDs.png]] ![[odbcad32_cusa3aZeAU.png]]

## Migración al SQL Express Local
Abrir el **[Asistente de Migración Kairo](https://github.com/Devs-Kong-Software/MigradorKairo)**,  rellenar los datos y presionar el botón de iniciar.
![[MigradorKairo_jsP9w9SEHk.png]]

> [!info]
> Se pueden omitir los campos de **Login** si el inicio de sesión se configuró como inicio de sesión de Windows

El programa te pedirá confirmación antes de comenzar la migración.
![[MigradorKairo_q0OJlOwXKB.png]] ![[MigradorKairo_hqfAMZJrtd.png]]

> [!caution] Precaución
> Luego de la migración, comprobar que el ODBC apunte al servidor destino.
> **Kairo utilizará el ODBC del ordenador donde se ejecute**, no el del servidor (en el caso de carpetas compartidas).

