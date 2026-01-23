## Requisitos locales
- [ ] **Microsoft SQL Server Express** instalado y conectado mediante **SQL Server Management Studio**
- [ ] **[Migrador Kairo](https://github.com/Devs-Kong-Software/MigradorKairo)** con el worker registrado
- [ ] 
## Requisitos cliente
- [ ] **Microsoft SQL Server Express** instalado y conectado mediante **SQL Server Management Studio**

> [!info]
> La conexión a SQL Server se usará para ejecutar un script de SQL, no hace falta conectarse directamente desde la oficina.
> Con tener acceso remoto y poder pasar el script SQL basta.
## Crear usuario Kairo en SQL Local
Ejecutar el siguiente script en Microsoft SQL Server Express
![[kwqFVih6aC.png]]

> [!caution]
> Tienes que reemplazar `CONTRASEÑA DE DATOS.MDB` por la contraseña de **Datos.mdb**

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

