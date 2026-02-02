# GUÍA DE USO - RAMOUTILS v2.0

## ?? CAMBIOS PRINCIPALES

### 1?? Nuevo Login SAP
Ya no se conecta automáticamente desde App.config. Ahora debe ingresar credenciales manualmente al inicio.

### 2?? Rango de Fechas
El formulario de consulta de stock ahora permite seleccionar un rango de fechas (inicio - fin).

---

## ?? INICIO DE LA APLICACIÓN

### Paso 1: Ejecutar RamoUtils.exe

Al iniciar, verá el formulario **"Login SAP Business One"**:

![Login](https://via.placeholder.com/400x300?text=Formulario+de+Login)

### Paso 2: Completar Credenciales

#### **Configuración del Servidor**
- **Servidor**: `NDB@sles.navarrete.local:30013`
- **Base Datos**: `SBO_DISTRIBUIDORA`
- **Tipo de Base Datos**: `SAP HANA`
- **Usuario Base de Datos**: `B1SYSTEM`
- **Contraseña Base Datos**: `Fq6p9FZtz4yCs`

#### **Credenciales SAP**
- **Usuario**: `integraRML` (o su usuario SAP)
- **Contraseña**: Su contraseña SAP

### Paso 3: Conectar

Haga clic en **"Conectar"**. Si los datos son correctos:
- ? Verá mensaje: "Conexión exitosa a SAP Business One"
- ? Se abrirá el formulario principal

Si hay error:
- ? Verifique las credenciales
- ? Asegúrese de tener acceso al servidor SAP
- ? Consulte con el administrador si el problema persiste

---

## ?? CONSULTAR STOCK

### Paso 1: Abrir Consulta de Stock

En el menú principal: **Consultas ? Consultar Stock**

### Paso 2: Seleccionar Rango de Fechas

**Nuevo diseño con DOS fechas:**
```
[Fecha Inicio: 01/12/2024] [Fecha Fin: 15/12/2024] [Buscar]
```

**Ejemplos de uso:**

| Caso | Fecha Inicio | Fecha Fin | Resultado |
|------|--------------|-----------|-----------|
| Un solo día | 10/12/2024 | 10/12/2024 | Facturas del 10 de diciembre |
| Una semana | 01/12/2024 | 07/12/2024 | Facturas del 1 al 7 de diciembre |
| Un mes | 01/12/2024 | 31/12/2024 | Facturas de todo diciembre |

### Paso 3: Buscar

Haga clic en **"Buscar"**. El sistema:
1. Consulta SQL Server (facturas pendientes)
2. Consulta SAP HANA (stock disponible)
3. Compara y muestra artículos con stock insuficiente

### Paso 4: Interpretar Resultados

**Columnas importantes:**
- **Código Artículo**: SKU del producto
- **Descripción**: Nombre del artículo
- **Cant. Requerida**: Cantidad que necesita la(s) factura(s)
- **Unidad**: Unidad de medida (ej: CAJA, UND)
- **Factor**: Factor de conversión a unidad base
- **Cant. Base (UND)**: Cantidad convertida a unidad base
- **Stock Disponible (UND)**: Stock actual en SAP
- **Faltante (UND)**: ?? Cantidad faltante (en ROJO)

### Paso 5: Exportar (Opcional)

Haga clic en **"Exportar Excel"** para generar:
- ?? CSV: Archivo compatible con Excel
- ?? HTML: Archivo con formato visual

---

## ?? CONFIGURACIÓN (App.config)

El archivo App.config **solo se usa para**:

### ? Conexión SQL Server (OBLIGATORIO)
```xml
<connectionStrings>
  <add name="SQLIntermedia" 
       connectionString="Server=SRVDISNAV06;Database=BD_INT_NAVARRETE_PROD;User Id=sa;Password=Business456;" 
       providerName="System.Data.SqlClient"/>
</connectionStrings>
```

### ? Stored Procedure SQL (OPCIONAL)
```xml
<add key="SQL_SP_FacturasPendientes" value="RML_GET_ARTICULOS_PORFECHA"/>
```

### ? Sugerencias para Login SAP (OPCIONAL)
```xml
<add key="SAP_Server" value="NDB@sles.navarrete.local:30013"/>
<add key="SAP_CompanyDB" value="SBO_DISTRIBUIDORA"/>
<add key="SAP_DbUserName" value="B1SYSTEM"/>
```

**Nota:** Las credenciales SAP ya NO se toman de App.config, debe ingresarlas manualmente.

---

## ?? SOLUCIÓN DE PROBLEMAS

### ? "No se pudo establecer conexión con SAP"

**Posibles causas:**
1. Credenciales incorrectas
2. Usuario sin permisos en SAP
3. Servidor SAP inaccesible
4. Puerto bloqueado (30013, 30015)

**Solución:**
- Verifique usuario y contraseña SAP
- Pruebe conectarse con SAP Business One Client
- Consulte con el administrador SAP

---

### ? "Error de configuración SQL"

**Posibles causas:**
1. Cadena de conexión SQL incorrecta en App.config
2. Servidor SQL no accesible
3. Base de datos no existe

**Solución:**
- Verifique la conexión en App.config
- Pruebe conectarse con SQL Server Management Studio
- Asegúrese de tener acceso a `BD_INT_NAVARRETE_PROD`

---

### ? "Advertencia de conexión SQL"

**Causa:**
- La aplicación pudo conectarse a SAP pero no a SQL Server

**Solución:**
- Puede continuar (solo si NO va a consultar stock)
- Verifique la conexión SQL en App.config
- Consulte con el administrador SQL Server

---

### ? "La fecha de inicio no puede ser mayor que la fecha fin"

**Causa:**
- Fecha inicio > Fecha fin

**Solución:**
- Asegúrese de que: `Fecha Inicio ? Fecha Fin`

---

## ?? DIFERENCIAS CON LA VERSIÓN ANTERIOR

| Característica | Versión 1.0 (Anterior) | Versión 2.0 (Actual) |
|----------------|------------------------|----------------------|
| **Login SAP** | Automático desde App.config | Manual con formulario |
| **Conexiones SAP** | Una por formulario | Una global compartida |
| **Filtro de fechas** | Solo una fecha | Rango de fechas |
| **Validación DIAPI** | En cada formulario | Solo al inicio |
| **Seguridad** | Contraseña en App.config | Contraseña ingresada manualmente |

---

## ?? PREGUNTAS FRECUENTES

### ? ¿Dónde están las credenciales SAP?
**R:** Ya no se almacenan en App.config. Debe ingresarlas cada vez que inicie la aplicación.

### ? ¿Puedo usar una sola fecha como antes?
**R:** Sí, simplemente ponga la misma fecha en "Fecha Inicio" y "Fecha Fin".

### ? ¿La aplicación guarda las credenciales?
**R:** No, por seguridad, debe ingresar las credenciales cada vez que inicie la aplicación.

### ? ¿Qué pasa si cierro el formulario de login sin conectar?
**R:** La aplicación se cerrará y deberá ejecutarla nuevamente.

### ? ¿Puedo cambiar de usuario SAP sin cerrar la aplicación?
**R:** No, debe cerrar y volver a abrir la aplicación para cambiar de usuario.

### ? ¿Cuántas conexiones SAP simultáneas usa?
**R:** Solo UNA conexión SAP compartida por toda la aplicación.

---

## ?? SOPORTE TÉCNICO

**Desarrollador:** Sistema RamoUtils  
**Versión:** 2.0  
**Última actualización:** 2024

Para soporte adicional:
1. Revise los logs en Visual Studio (Output Window)
2. Consulte el archivo `CAMBIOS_REALIZADOS.md` para detalles técnicos
3. Contacte al administrador del sistema

---

## ? CARACTERÍSTICAS NUEVAS

### ? Login Manual
- Mayor seguridad
- No almacena contraseñas en archivos
- Validación en tiempo real

### ? Rango de Fechas
- Consultar múltiples días a la vez
- Análisis de períodos completos
- Exportación con rango incluido

### ? Conexión Única
- Mejor rendimiento
- Menos consumo de recursos
- Más estable

---

**¡Gracias por usar RamoUtils v2.0!** ??
