using System;
using System.Configuration;

namespace RamoUtils
{
    /// <summary>
    /// Helper para leer configuraciones desde App.config de forma segura
    /// </summary>
    public static class ConfigHelper
    {
        #region SQL Server Configuration

        /// <summary>
        /// Obtiene la cadena de conexión de SQL Server desde App.config
        /// </summary>
        public static string GetSQLConnectionString()
        {
            return GetConnectionString("SQLIntermedia");
        }

        /// <summary>
        /// Obtiene el nombre de la tabla de encabezado
        /// </summary>
        public static string GetSQLTablaEncabezado()
        {
            return GetAppSetting("SQL_TablaEncabezado", "EncabezadoFacturas");
        }

        /// <summary>
        /// Obtiene el nombre de la tabla de detalle
        /// </summary>
        public static string GetSQLTablaDetalle()
        {
            return GetAppSetting("SQL_TablaDetalle", "DetalleFacturas");
        }

        /// <summary>
        /// Obtiene el nombre del SP para facturas pendientes
        /// </summary>
        public static string GetSQLStoredProcFacturasPendientes()
        {
            return GetAppSetting("SQL_SP_FacturasPendientes", "");
        }

        #endregion

        #region SAP DI API Configuration

        /// <summary>
        /// Obtiene el servidor SAP HANA
        /// </summary>
        public static string GetSAPServer()
        {
            return GetAppSetting("SAP_Server", "");
        }

        /// <summary>
        /// Obtiene el tipo de servidor de base de datos SAP (dst_HANADB = 9)
        /// </summary>
        public static string GetSAPDbServerType()
        {
            return GetAppSetting("SAP_DbServerType", "dst_HANADB");
        }

        /// <summary>
        /// Obtiene el puerto del servidor HANA
        /// </summary>
        public static int GetSAPDbPort()
        {
            string port = GetAppSetting("SAP_DbPort", "30015");
            int result;
            return int.TryParse(port, out result) ? result : 30015;
        }

        /// <summary>
        /// Obtiene el nombre de la base de datos SAP
        /// </summary>
        public static string GetSAPCompanyDB()
        {
            return GetAppSetting("SAP_CompanyDB", "");
        }

        /// <summary>
        /// Obtiene el usuario SAP
        /// </summary>
        public static string GetSAPUserName()
        {
            return GetAppSetting("SAP_UserName", "");
        }

        /// <summary>
        /// Obtiene la contraseña SAP
        /// </summary>
        public static string GetSAPPassword()
        {
            return GetAppSetting("SAP_Password", "");
        }

        /// <summary>
        /// Obtiene el usuario de la base de datos HANA
        /// </summary>
        public static string GetSAPDbUserName()
        {
            return GetAppSetting("SAP_DbUserName", "");
        }

        /// <summary>
        /// Obtiene la contraseña de la base de datos HANA
        /// </summary>
        public static string GetSAPDbPassword()
        {
            return GetAppSetting("SAP_DbPassword", "");
        }

        /// <summary>
        /// Obtiene el servidor de licencias (opcional)
        /// </summary>
        public static string GetSAPLicenseServer()
        {
            return GetAppSetting("SAP_LicenseServer", "");
        }

        /// <summary>
        /// Obtiene el nombre del SP HANA para consultar stock
        /// </summary>
        public static string GetHANAStoredProcConsultarStock()
        {
            return GetAppSetting("HANA_SP_ConsultarStock", "");
        }

        #endregion

        #region General Configuration

        /// <summary>
        /// Obtiene el timeout para queries en segundos
        /// </summary>
        public static int GetQueryTimeout()
        {
            string timeout = GetAppSetting("QueryTimeout", "60");
            int result;
            return int.TryParse(timeout, out result) ? result : 60;
        }

        #endregion

        #region Helper Methods

        /// <summary>
        /// Lee una cadena de conexión desde connectionStrings
        /// </summary>
        private static string GetConnectionString(string name)
        {
            try
            {
                var connString = ConfigurationManager.ConnectionStrings[name];
                if (connString != null)
                {
                    return connString.ConnectionString;
                }
                throw new ConfigurationErrorsException($"No se encontró la cadena de conexión '{name}' en App.config");
            }
            catch (Exception ex)
            {
                throw new ConfigurationErrorsException($"Error al leer la cadena de conexión '{name}': {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Lee un valor de appSettings con valor por defecto
        /// </summary>
        private static string GetAppSetting(string key, string defaultValue = "")
        {
            try
            {
                string value = ConfigurationManager.AppSettings[key];
                return string.IsNullOrEmpty(value) ? defaultValue : value;
            }
            catch
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// Valida que todas las configuraciones SAP estén completas
        /// </summary>
        public static void ValidarConfiguracionSAP()
        {
            if (string.IsNullOrEmpty(GetSAPServer()))
                throw new ConfigurationErrorsException("Falta configurar SAP_Server en App.config");

            if (string.IsNullOrEmpty(GetSAPCompanyDB()))
                throw new ConfigurationErrorsException("Falta configurar SAP_CompanyDB en App.config");

            if (string.IsNullOrEmpty(GetSAPUserName()))
                throw new ConfigurationErrorsException("Falta configurar SAP_UserName en App.config");

            if (string.IsNullOrEmpty(GetSAPPassword()))
                throw new ConfigurationErrorsException("Falta configurar SAP_Password en App.config");

            if (string.IsNullOrEmpty(GetSAPDbUserName()))
                throw new ConfigurationErrorsException("Falta configurar SAP_DbUserName en App.config");

            if (string.IsNullOrEmpty(GetSAPDbPassword()))
                throw new ConfigurationErrorsException("Falta configurar SAP_DbPassword en App.config");
        }

        /// <summary>
        /// Valida que la configuración SQL esté completa
        /// </summary>
        public static void ValidarConfiguracionSQL()
        {
            string connString = GetSQLConnectionString();
            if (string.IsNullOrEmpty(connString) || connString.Contains("TU_SERVIDOR_SQL"))
                throw new ConfigurationErrorsException("Falta configurar la cadena de conexión SQLIntermedia en App.config");
        }

        #endregion
    }
}
