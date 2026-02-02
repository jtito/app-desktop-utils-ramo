using System;
using System.Data.SqlClient;
using System.Threading;
using SAPbobsCOM;

namespace RamoUtils
{
    /// <summary>
    /// Gestor centralizado de conexiones a SQL Server y SAP DI API (Thread-Safe)
    /// </summary>
    public class ConnectionManager : IDisposable
    {
        private Company sapCompany;
        private bool disposed = false;
        private readonly object lockObject = new object();  // Para thread-safety
        private int connectionRefCount = 0;  // Contador de referencias
        private bool useGlobalConnection = true; // Usar conexión global del login

        #region SQL Server Connection

        /// <summary>
        /// Crea una nueva conexión a SQL Server desde la configuración
        /// </summary>
        public SqlConnection GetSQLConnection()
        {
            try
            {
                string connectionString = ConfigHelper.GetSQLConnectionString();
                SqlConnection conn = new SqlConnection(connectionString);
                return conn;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al crear conexión SQL: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Valida la conexión SQL Server
        /// </summary>
        public bool TestSQLConnection(out string errorMessage)
        {
            errorMessage = string.Empty;
            SqlConnection conn = null;

            try
            {
                conn = GetSQLConnection();
                conn.Open();
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
            finally
            {
                if (conn != null)
                {
                    if (conn.State == System.Data.ConnectionState.Open)
                    {
                        conn.Close();
                    }
                    conn.Dispose();
                }
            }
        }

        #endregion

        #region SAP DI API Connection (Thread-Safe)

        /// <summary>
        /// Obtiene la conexión SAP Company de forma thread-safe
        /// Usa la conexión global establecida en el login
        /// </summary>
        public Company GetSAPCompany()
        {
            lock (lockObject)
            {
                // Si está habilitado el uso de conexión global y existe
                if (useGlobalConnection && Program.SAPCompanyGlobal != null && Program.SAPCompanyGlobal.Connected)
                {
                    sapCompany = Program.SAPCompanyGlobal;
                    Interlocked.Increment(ref connectionRefCount);
                    return sapCompany;
                }

                // Si no hay conexión global, intentar crear una nueva (fallback)
                if (sapCompany == null || !sapCompany.Connected)
                {
                    // Limpiar conexión anterior si existe
                    if (sapCompany != null && sapCompany != Program.SAPCompanyGlobal)
                    {
                        try
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(sapCompany);
                        }
                        catch { }
                        sapCompany = null;
                    }

                    // Crear nueva conexión (solo si hay configuración en App.config)
                    try
                    {
                        sapCompany = ConnectToSAP();
                    }
                    catch
                    {
                        throw new Exception("No hay conexión SAP disponible. Por favor, reinicie la aplicación e inicie sesión.");
                    }
                }

                // Incrementar contador de referencias
                Interlocked.Increment(ref connectionRefCount);

                return sapCompany;
            }
        }

        /// <summary>
        /// Conecta a SAP usando DI API con la configuración del App.config
        /// </summary>
        private Company ConnectToSAP()
        {
            Company company = null;

            try
            {
                // Validar configuración
                ConfigHelper.ValidarConfiguracionSAP();

                // Crear instancia de Company
                company = new Company();

                // Configurar conexión desde App.config
                company.Server = ConfigHelper.GetSAPServer();
                company.CompanyDB = ConfigHelper.GetSAPCompanyDB();
                company.UserName = ConfigHelper.GetSAPUserName();
                company.Password = ConfigHelper.GetSAPPassword();

                // Configurar tipo de servidor (HANA)
                string dbServerType = ConfigHelper.GetSAPDbServerType();
                if (dbServerType == "dst_HANADB")
                {
                    company.DbServerType = BoDataServerTypes.dst_HANADB;
                }
                else
                {
                    company.DbServerType = BoDataServerTypes.dst_HANADB;
                }

                // Configurar UseTrusted (importante para multiconexión)
                company.UseTrusted = false;

                // Conectar
                int result = company.Connect();

                if (result != 0)
                {
                    string errorMsg = company.GetLastErrorDescription();
                    int errorCode = company.GetLastErrorCode();

                    // Liberar recursos
                    if (company != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(company);
                        company = null;
                    }

                    throw new Exception($"Error al conectar a SAP. Código: {errorCode}, Mensaje: {errorMsg}");
                }

                return company;
            }
            catch (Exception ex)
            {
                if (company != null)
                {
                    try
                    {
                        if (company.Connected)
                        {
                            company.Disconnect();
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(company);
                    }
                    catch { }
                    company = null;
                }

                throw new Exception($"Error al conectar con SAP DI API: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Valida la conexión SAP (usa la conexión global)
        /// </summary>
        public bool TestSAPConnection(out string errorMessage)
        {
            errorMessage = string.Empty;

            try
            {
                // Verificar conexión global
                if (Program.SAPCompanyGlobal != null && Program.SAPCompanyGlobal.Connected)
                {
                    return true;
                }
                else
                {
                    errorMessage = "No hay conexión SAP activa";
                    return false;
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Crea un RecordSet para ejecutar queries en SAP HANA (Thread-Safe)
        /// </summary>
        public Recordset CreateRecordset()
        {
            try
            {
                Company company = GetSAPCompany();

                lock (lockObject)
                {
                    Recordset recordset = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    return recordset;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al crear RecordSet: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Libera un RecordSet de forma segura
        /// </summary>
        public void ReleaseRecordset(Recordset recordset)
        {
            if (recordset != null)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset);

                    // Decrementar contador de referencias
                    Interlocked.Decrement(ref connectionRefCount);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error al liberar RecordSet: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Forzar reconexión (útil cuando la conexión se pierde)
        /// </summary>
        public void ForceReconnect()
        {
            lock (lockObject)
            {
                // No desconectar la conexión global
                if (sapCompany != null && sapCompany != Program.SAPCompanyGlobal)
                {
                    try
                    {
                        if (sapCompany.Connected)
                        {
                            sapCompany.Disconnect();
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sapCompany);
                    }
                    catch { }
                    sapCompany = null;
                }

                connectionRefCount = 0;
            }
        }

        #endregion

        #region Dispose Pattern

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // Liberar recursos administrados
                    System.Diagnostics.Debug.WriteLine($"Disposing ConnectionManager. Referencias activas: {connectionRefCount}");
                }

                // NO desconectar SAP si es la conexión global (se maneja en Program.cs)
                lock (lockObject)
                {
                    if (sapCompany != null && sapCompany != Program.SAPCompanyGlobal)
                    {
                        try
                        {
                            if (sapCompany.Connected)
                            {
                                System.Diagnostics.Debug.WriteLine("Desconectando de SAP...");
                                sapCompany.Disconnect();
                                System.Diagnostics.Debug.WriteLine("SAP desconectado exitosamente");
                            }

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(sapCompany);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error al desconectar SAP: {ex.Message}");
                        }
                        finally
                        {
                            sapCompany = null;
                        }
                    }
                    
                    connectionRefCount = 0;
                }

                disposed = true;
            }
        }

        ~ConnectionManager()
        {
            Dispose(false);
        }

        #endregion
    }
}
