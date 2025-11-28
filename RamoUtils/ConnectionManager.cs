using System;
using System.Data.SqlClient;
using SAPbobsCOM;

namespace RamoUtils
{
    /// <summary>
    /// Gestor centralizado de conexiones a SQL Server y SAP DI API
    /// </summary>
    public class ConnectionManager : IDisposable
    {
        private Company sapCompany;
        private bool disposed = false;

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
                if (conn != null && conn.State == System.Data.ConnectionState.Open)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }

        #endregion

        #region SAP DI API Connection

        /// <summary>
        /// Obtiene la conexión SAP Company (singleton)
        /// </summary>
        public Company GetSAPCompany()
        {
            if (sapCompany == null || !sapCompany.Connected)
            {
                sapCompany = ConnectToSAP();
            }

            return sapCompany;
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
                //company.DbUserName = ConfigHelper.GetSAPDbUserName();
                //company.DbPassword = ConfigHelper.GetSAPDbPassword();

                // Configurar tipo de servidor (HANA)
                string dbServerType = ConfigHelper.GetSAPDbServerType();
                if (dbServerType == "dst_HANADB")
                {
                    company.DbServerType = BoDataServerTypes.dst_HANADB;
                }
                else
                {
                    // Soportar otros tipos si es necesario
                    company.DbServerType = BoDataServerTypes.dst_HANADB;
                }

                
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
        /// Valida la conexión SAP
        /// </summary>
        public bool TestSAPConnection(out string errorMessage)
        {
            errorMessage = string.Empty;
            Company testCompany = null;

            try
            {
                testCompany = ConnectToSAP();
                return testCompany != null && testCompany.Connected;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
            finally
            {
                if (testCompany != null)
                {
                    try
                    {
                        if (testCompany.Connected)
                        {
                            testCompany.Disconnect();
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(testCompany);
                    }
                    catch { }
                }
            }
        }

        /// <summary>
        /// Crea un RecordSet para ejecutar queries en SAP HANA
        /// </summary>
        public Recordset CreateRecordset()
        {
            try
            {
                Company company = GetSAPCompany();
                Recordset recordset = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                return recordset;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al crear RecordSet: {ex.Message}", ex);
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
                }

                // Desconectar SAP
                if (sapCompany != null)
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
