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
Dirigirse a https://www.microsoft.com/es-es/sql-server/sql-server-downloads, descargar e instalar **SQL Server 2025 Express** o **Desarrollador de SQL Server 2025**.
![[brave_djfm0AjCY2.png]]

### Habilitar conexión
Entrar a **SQL Server Configuration Manager**
- [ ] Bajo **Servicios de SQL Server** habilitar **SQL Server Browser**

Entrar a **Configuración de red de SQL Server/Protocolos de SQLEXPRESS** y abrir **TCP/IP**
- [ ] Bajo **Protocolo** habilitar **Escuchar todo** y **Habilitado**

Abrir el CMD como administrador
- [ ]  Ejecutar
```
netsh advfirewall firewall add rule name="SQL Server TCP 1433" dir=in action=allow protocol=TCP localport=1433
```
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