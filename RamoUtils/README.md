# RamoUtils - Sistema de Consulta de Stock SAP

## Descripción
Sistema Windows Forms para validar stock disponible en SAP HANA antes de migrar facturas desde una base de datos intermedia SQL Server.

## Requisitos Previos

### 1. Instalar el Driver de SAP HANA

**Opción A: Desde Visual Studio (Recomendado)**
1. Clic derecho en el proyecto `RamoUtils` en el Solution Explorer
2. Seleccionar **Manage NuGet Packages**
3. Ir a la pestaña **Browse**
4. Buscar: `Sap.Data.Hana`
5. Instalar el paquete `Sap.Data.Hana.Core.v2.1` o `Sap.Data.Hana` (según tu versión de SAP HANA)

**Opción B: Package Manager Console**
```powershell
Install-Package Sap.Data.Hana.Core.v2.1
```

**Opción C: .NET CLI**
```bash
cd RamoUtils
dotnet add package Sap.Data.Hana.Core.v2.1
```

### 2. Configurar las Cadenas de Conexión

Editar el archivo `FrmConsultaStock.cs` en el método `InicializarFormulario()` (líneas 27-28):

```csharp
// CAMBIAR ESTAS CADENAS POR LAS TUYAS
string sqlConn = "Server=TU_SERVIDOR_SQL;Database=TU_BD_INTERMEDIA;User Id=usuario;Password=pass;";
string hanaConn = "Server=TU_SERVIDOR_HANA:30015;Database=TU_BD_SAP;UserID=usuario;Password=pass;";
```

**Ejemplo de cadena de conexión SQL Server:**
```
Server=192.168.1.100;Database=FacturasIntermedia;User Id=sa;Password=MiPassword123;
```

**Ejemplo de cadena de conexión SAP HANA:**
```
Server=192.168.1.200:30015;Database=SBODEMO;UserID=MANAGER;Password=MiPassword123;
```

### 3. Ajustar las Consultas SQL

Editar el archivo `StockService.cs` en el método `ObtenerArticulosPendientesSQL()`:

**Cambiar los nombres de las tablas según tu estructura:**

```csharp
// LÍNEA 126 - Ajustar nombres de tablas y columnas
string query = @"
    SELECT 
        h.DocNum,           -- Número de factura
        h.DocEntry,         -- ID interno
        d.ItemCode,         -- Código de artículo
        d.Quantity,         -- Cantidad requerida
        d.WhsCode,          -- Código de almacén
        h.DocDate           -- Fecha del documento
    FROM DetalleFacturas d      -- CAMBIAR POR TU TABLA DE DETALLE
    INNER JOIN EncabezadoFacturas h ON d.DocEntry = h.DocEntry  -- CAMBIAR POR TU TABLA DE ENCABEZADO
    WHERE (h.Migrado = 0 OR h.ErrorMigracion = 1)  -- AJUSTAR CONDICIONES
    AND CAST(h.DocDate AS DATE) = @Fecha
    ORDER BY h.DocNum, d.ItemCode";
```

## Estructura del Proyecto

```
RamoUtils/
?
??? Program.cs                      # Punto de entrada de la aplicación
??? FrmPrincipal.cs                 # Formulario principal con menú
??? FrmPrincipal.Designer.cs        # Diseño del formulario principal
??? FrmConsultaStock.cs             # Formulario de consulta de stock
??? FrmConsultaStock.Designer.cs    # Diseño del formulario de consulta
??? StockService.cs                 # Lógica de negocio y consultas
??? App.config                      # Configuración de la aplicación
??? UserControl1.cs                 # (Archivo original - puede eliminarse)
```

## Cómo Usar

1. **Compilar el Proyecto**
   - Presiona `F5` o `Ctrl+F5` en Visual Studio
   - Se abrirá el formulario principal

2. **Consultar Stock**
   - Ir al menú: **Consultas > Consultar Stock**
   - Seleccionar una fecha en el DateTimePicker
   - Hacer clic en el botón **Buscar**
   - El sistema mostrará los artículos con stock insuficiente

3. **Interpretar Resultados**
   - **Nº Factura**: Número de la factura pendiente
   - **Código Artículo**: SKU del producto
   - **Descripción**: Nombre del artículo desde SAP
   - **Almacén**: Código del warehouse
   - **Cant. Requerida**: Cantidad que necesita la factura
   - **Stock Disponible**: Stock actual en SAP (OnHand - IsCommited)
   - **Faltante**: Diferencia (en rojo)

## Funcionalidades

? **Menú Principal MDI** con submenú de consultas
? **Filtro por Fecha** para buscar facturas específicas
? **Consulta Dual** SQL Server + SAP HANA
? **Cálculo de Stock Real** (OnHand - IsCommited)
? **Interfaz Intuitiva** con DataGridView formateado
? **Búsqueda Asíncrona** que no bloquea la UI
? **Mensajes de Estado** en barra inferior

## Solución de Problemas

### Error: "Could not find Sap.Data.Hana"
- Instalar el paquete NuGet `Sap.Data.Hana.Core.v2.1` (ver paso 1)

### Error: "A network-related or instance-specific error"
- Verificar que el servidor SQL/HANA esté accesible
- Revisar firewall y puertos (SQL: 1433, HANA: 30015)
- Verificar credenciales de conexión

### Error: "Invalid object name 'DetalleFacturas'"
- Ajustar los nombres de tablas en `StockService.cs` (ver paso 3)

### No aparecen resultados
- Verificar que existan facturas con `Migrado = 0` o `ErrorMigracion = 1`
- Revisar que la fecha seleccionada tenga registros
- Verificar los nombres de columnas en la consulta SQL

## Modificar el Proyecto

### Cambiar el Output Type a Windows Application

Si el proyecto se ejecuta como consola, cambiar en `RamoUtils.csproj`:

```xml
<PropertyGroup>
  <OutputType>WinExe</OutputType>  <!-- Cambiar de Exe a WinExe -->
  <StartupObject>RamoUtils.Program</StartupObject>
</PropertyGroup>
```

### Agregar más consultas al menú

Editar `FrmPrincipal.Designer.cs` y agregar más ToolStripMenuItem al menú "Consultas".

## Contacto y Soporte

Para dudas o problemas, revisar los logs de error en la aplicación o contactar al equipo de desarrollo.

---

**Versión:** 1.0.0  
**Última Actualización:** Enero 2025  
**.NET Framework:** 4.7.2  
**Lenguaje:** C# 7.3
