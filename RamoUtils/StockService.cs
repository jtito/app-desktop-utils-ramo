using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace RamoUtils
{
    /// <summary>
    /// Servicio de negocio para consultar stock y validar disponibilidad
    /// </summary>
    public class StockService : IDisposable
    {
        private ConnectionManager connectionManager;
        private DIAPIService diapiService;
        private bool disposed = false;

        /// <summary>
        /// Constructor que inicializa los servicios de conexión
        /// </summary>
        public StockService(ConnectionManager connManager)
        {
            connectionManager = connManager ?? throw new ArgumentNullException(nameof(connManager));
            diapiService = new DIAPIService(connectionManager);
        }

        /// <summary>
        /// Clase para representar artículos sin stock suficiente
        /// </summary>
        public class ArticuloSinStock
        {
            public string NumeroFactura { get; set; }
            public string ItemCode { get; set; }
            public string ItemName { get; set; }
            public string WhsCode { get; set; }
            public decimal CantidadRequerida { get; set; }
            public decimal CantidadRequeridaBase { get; set; }  // Nueva: cantidad en unidad base
            public decimal StockDisponible { get; set; }
            public decimal Faltante { get; set; }
            public int UomEntry { get; set; }
            public decimal FactorConversion { get; set; }
            public string MensajeError { get; set; }
            public string UomCode { get; set; }
            public string UomName { get; set; }
        }

        /// <summary>
        /// Consulta artículos sin stock suficiente por fecha
        /// </summary>
        public List<ArticuloSinStock> ConsultarStockPorFecha(DateTime fecha)
        {
            return ConsultarStockPorRangoFechas(fecha, fecha);
        }

        /// <summary>
        /// Consulta artículos sin stock suficiente por rango de fechas
        /// </summary>
        public List<ArticuloSinStock> ConsultarStockPorRangoFechas(DateTime fechaInicio, DateTime fechaFin)
        {
            List<ArticuloSinStock> resultado = new List<ArticuloSinStock>();

            try
            {
                // 1. Obtener artículos pendientes de SQL por rango de fechas
                DataTable pendientes = ObtenerArticulosPendientesSQLRango(fechaInicio, fechaFin);

                if (pendientes.Rows.Count == 0)
                    return resultado;

                // 2. Extraer códigos únicos
                List<string> itemCodes = pendientes.AsEnumerable()
                    .Select(r => r.Field<string>("ItemCode"))
                    .Distinct()
                    .ToList();

                List<string> whsCodes = pendientes.AsEnumerable()
                    .Select(r => r.Field<string>("WhsCode"))
                    .Where(w => !string.IsNullOrEmpty(w))
                    .Distinct()
                    .ToList();

                if (itemCodes.Count == 0)
                    return resultado;

                // 3. Obtener stock de HANA usando DI API
                DataTable stockHana = ObtenerStockHANA(itemCodes, whsCodes);

                List<int> uomEntries = pendientes.AsEnumerable()
                .Where(r => r["UomEntry"] != DBNull.Value)
                .Select(r => Convert.ToInt32(r["UomEntry"]))
                .Distinct()
                .ToList();

                Dictionary<int, Tuple<string, string>> uomInfo = diapiService.ObtenerInformacionUoM(uomEntries);

                // 4. Comparar y generar reporte
                foreach (DataRow row in pendientes.Rows)
                {
                    string itemCode = row["ItemCode"].ToString();
                    string whsCode = row["WhsCode"]?.ToString() ?? "";
                    decimal cantidadRequerida = Convert.ToDecimal(row["Quantity"]);
                    string mensajeError = row["Mensaje"]?.ToString() ?? "";
                    int uomEntry = row["UomEntry"] != DBNull.Value ? Convert.ToInt32(row["UomEntry"]) : 0;
                    decimal factorConversion = diapiService.ObtenerFactorConversionUoM(itemCode, uomEntry);
                    decimal cantidadRequeridaBase = cantidadRequerida * factorConversion;

                    if (uomEntry > 0)
                    {
                        factorConversion = diapiService.ObtenerFactorConversionUoM(itemCode, uomEntry);
                    }
                    
                    // Buscar stock en HANA
                    decimal stockDisponible = 0;
                    string itemName = "";

                    if (!string.IsNullOrEmpty(whsCode))
                    {
                        DataRow[] stockRows = stockHana.Select(
                            $"ItemCode = '{itemCode}' AND WhsCode = '{whsCode}'");

                        if (stockRows.Length > 0)
                        {
                            stockDisponible = Convert.ToDecimal(stockRows[0]["StockReal"]);
                            itemName = stockRows[0]["ItemName"].ToString();
                        }
                    }
                    else
                    {
                        // Si no hay almacén específico, sumar todos los almacenes
                        DataRow[] stockRows = stockHana.Select($"ItemCode = '{itemCode}'");
                        if (stockRows.Length > 0)
                        {
                            stockDisponible = stockRows.Sum(r => Convert.ToDecimal(r["StockReal"]));
                            itemName = stockRows[0]["ItemName"].ToString();
                        }
                    }

                    // Obtener información de UoM
                    string uomCode = "";
                    string uomName = "";
                    if (uomEntry > 0 && uomInfo.ContainsKey(uomEntry))
                    {
                        uomCode = uomInfo[uomEntry].Item1;
                        uomName = uomInfo[uomEntry].Item2;
                    }

                    // Si no hay stock suficiente, agregar al reporte
                    if (stockDisponible < cantidadRequeridaBase)
                    {
                        resultado.Add(new ArticuloSinStock
                        {
                            ItemCode = itemCode,
                            ItemName = itemName,
                            WhsCode = whsCode,
                            CantidadRequerida = cantidadRequerida,  // Cantidad original (ej: 2 CAJAS)
                            CantidadRequeridaBase = cantidadRequeridaBase,  // Cantidad convertida (ej: 48 UND)
                            StockDisponible = Math.Max(0, stockDisponible),  // Stock en unidad base
                            Faltante = cantidadRequeridaBase - Math.Max(0, stockDisponible),  // Diferencia en unidad base
                            MensajeError = mensajeError,
                            UomEntry = uomEntry,
                            UomCode = uomCode,
                            UomName = uomName,
                            FactorConversion = factorConversion
                        });
                    }
                }

                return resultado.OrderBy(x => x.ItemCode).ToList();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al consultar stock por rango de fechas: {ex.Message}", ex);
            }
        }

        #region Métodos SQL Server

        /// <summary>
        /// Obtiene artículos pendientes desde SQL Server por rango de fechas
        /// </summary>
        private DataTable ObtenerArticulosPendientesSQLRango(DateTime fechaInicio, DateTime fechaFin)
        {
            DataTable dt = new DataTable();

            try
            {
                // Verificar si hay un SP configurado
                string spName = ConfigHelper.GetSQLStoredProcFacturasPendientes();

                using (SqlConnection conn = connectionManager.GetSQLConnection())
                {
                    conn.Open();

                    if (!string.IsNullOrEmpty(spName))
                    {
                        // Usar Stored Procedure
                        using (SqlCommand cmd = new SqlCommand(spName, conn))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.CommandTimeout = ConfigHelper.GetQueryTimeout();
                            cmd.Parameters.AddWithValue("@FechaInicio", fechaInicio.Date);
                            cmd.Parameters.AddWithValue("@FechaFin", fechaFin.Date);

                            using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                            {
                                adapter.Fill(dt);
                            }
                        }
                    }
                    else
                    {
                        // Usar consulta SQL directa con nombres de tablas configurables
                        string tablaEncabezado = ConfigHelper.GetSQLTablaEncabezado();
                        string tablaDetalle = ConfigHelper.GetSQLTablaDetalle();

                        string query = $@"
                        SELECT 
                            C.NroOperacion AS DocNum,
                            C.IdVenta AS DocEntry,
                            D.SAP_ItemCode AS ItemCode,
                            D.Cantidad AS Quantity,
                            OC.U_SEIN_ALMA AS WhsCode,
                            C.FechaEmision AS DocDate,
                            D.SAP_UoMEntry AS UomEntry,
                            C._ISMensajeError AS Mensaje
                        FROM {tablaDetalle} D
                        INNER JOIN {tablaEncabezado} C ON D.IdVenta = C.IdVenta
                        INNER JOIN OCRD OC ON OC.CardCode = C.SAP_CodigoCaja
                        WHERE C.SAP_Estado IN ('A', 'E')
                        AND CAST(C.FechaEmision AS DATE) BETWEEN @FechaInicio AND @FechaFin
                        AND C.SAP_ObjType <> 14
                        AND D.SAP_ItemCode IS NOT NULL
                        ORDER BY C.NroOperacion, D.SAP_ItemCode";

                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.CommandTimeout = ConfigHelper.GetQueryTimeout();
                            cmd.Parameters.AddWithValue("@FechaInicio", fechaInicio.Date);
                            cmd.Parameters.AddWithValue("@FechaFin", fechaFin.Date);

                            using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                            {
                                adapter.Fill(dt);
                            }
                        }
                    }
                }

                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al obtener artículos pendientes de SQL: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Obtiene artículos pendientes desde SQL Server
        /// Puede usar Stored Procedure o consulta directa
        /// </summary>
        private DataTable ObtenerArticulosPendientesSQL(DateTime fecha)
        {
            return ObtenerArticulosPendientesSQLRango(fecha, fecha);
        }

        #endregion

        #region Métodos SAP HANA (DI API)

        /// <summary>
        /// Obtiene stock disponible desde SAP HANA usando DI API
        /// </summary>
        private DataTable ObtenerStockHANA(List<string> itemCodes, List<string> whsCodes)
        {
            try
            {
                return diapiService.ConsultarStockHANA(itemCodes, whsCodes);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al obtener stock de HANA: {ex.Message}", ex);
            }
        }

        #endregion

        #region Métodos Públicos Adicionales

        /// <summary>
        /// Consulta stock de un artículo específico
        /// </summary>
        public DataTable ConsultarStockPorArticulo(string itemCode, string whsCode = null)
        {
            List<string> items = new List<string> { itemCode };
            List<string> warehouses = string.IsNullOrEmpty(whsCode) ? null : new List<string> { whsCode };
            
            return diapiService.ConsultarStockHANA(items, warehouses);
        }

        /// <summary>
        /// Valida las conexiones SQL y SAP
        /// </summary>
        public void ValidarConexiones(out bool sqlOk, out bool sapOk, out string sqlError, out string sapError)
        {
            sqlOk = connectionManager.TestSQLConnection(out sqlError);
            sapOk = connectionManager.TestSAPConnection(out sapError);
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
                    if (connectionManager != null)
                    {
                        connectionManager.Dispose();
                        connectionManager = null;
                    }
                }

                disposed = true;
            }
        }

        ~StockService()
        {
            Dispose(false);
        }

        #endregion
    }
}
