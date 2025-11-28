using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using SAPbobsCOM;

namespace RamoUtils
{
    /// <summary>
    /// Servicio para ejecutar consultas en SAP HANA usando DI API RecordSet
    /// </summary>
    public class DIAPIService
    {
        private ConnectionManager connectionManager;
        private const int BATCH_SIZE = 500; // Aumentado para reducir round-trips
        private Dictionary<string, decimal> cacheFactoresUoM; // Caché de factores de conversión

        public DIAPIService(ConnectionManager connManager)
        {
            connectionManager = connManager;
            cacheFactoresUoM = new Dictionary<string, decimal>();
        }

        /// <summary>
        /// Ejecuta un Stored Procedure en SAP HANA y retorna un DataTable
        /// </summary>
        public DataTable ExecuteStoredProcedure(string procedureName, Dictionary<string, object> parameters = null)
        {
            Recordset rs = null;
            DataTable dt = new DataTable();

            try
            {
                rs = connectionManager.CreateRecordset();

                // Construir llamada al SP
                string query = BuildStoredProcedureCall(procedureName, parameters);

                // Ejecutar
                rs.DoQuery(query);

                // Convertir RecordSet a DataTable
                dt = RecordsetToDataTable(rs);

                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al ejecutar SP '{procedureName}': {ex.Message}", ex);
            }
            finally
            {
                if (rs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                }
            }
        }

        /// <summary>
        /// Ejecuta una consulta SQL directa en SAP HANA
        /// </summary>
        public DataTable ExecuteQuery(string query)
        {
            Recordset rs = null;
            DataTable dt = new DataTable();

            try
            {
                rs = connectionManager.CreateRecordset();
                rs.DoQuery(query);
                dt = RecordsetToDataTable(rs);
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al ejecutar query: {ex.Message}", ex);
            }
            finally
            {
                if (rs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                }
            }
        }

        /// <summary>
        /// Consulta el stock disponible en SAP HANA por artículo y almacén (método legacy)
        /// </summary>
        public DataTable ConsultarStockHANA(List<string> itemCodes, List<string> whsCodes = null)
        {
            if (itemCodes == null || itemCodes.Count == 0)
            {
                return new DataTable();
            }

            try
            {
                // Verificar si hay un SP configurado
                string spName = ConfigHelper.GetHANAStoredProcConsultarStock();

                if (!string.IsNullOrEmpty(spName))
                {
                    // Usar Stored Procedure
                    var parameters = new Dictionary<string, object>
                    {
                        { "ItemCodes", string.Join(",", itemCodes) },
                        { "WhsCodes", whsCodes != null ? string.Join(",", whsCodes) : "" }
                    };

                    return ExecuteStoredProcedure(spName, parameters);
                }
                else
                {
                    // Usar consulta SQL directa
                    return ConsultarStockDirecto(itemCodes, whsCodes);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al consultar stock en HANA: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Consulta stock directamente sin SP
        /// </summary>
        private DataTable ConsultarStockDirecto(List<string> itemCodes, List<string> whsCodes)
        {
            // Construir lista de items con comillas simples
            List<string> quotedItems = new List<string>();
            foreach (string item in itemCodes)
            {
                quotedItems.Add($"'{EscapeSQLString(item)}'");
            }

            string query = $@"
                SELECT 
                    T0.""ItemCode"",
                    T0.""ItemName"",
                    T1.""WhsCode"",
                    T1.""OnHand"" AS ""StockDisponible"",
                    T1.""IsCommited"" AS ""Comprometido"",
                    (T1.""OnHand"" - T1.""IsCommited"") AS ""StockReal""
                FROM ""OITM"" T0
                INNER JOIN ""OITW"" T1 ON T0.""ItemCode"" = T1.""ItemCode""
                WHERE T0.""ItemCode"" IN ({string.Join(",", quotedItems)})";

            // Agregar filtro de almacenes si existen
            if (whsCodes != null && whsCodes.Count > 0)
            {
                List<string> quotedWhs = new List<string>();
                foreach (string whs in whsCodes)
                {
                    quotedWhs.Add($"'{EscapeSQLString(whs)}'");
                }
                query += $" AND T1.\"WhsCode\" IN ({string.Join(",", quotedWhs)})";
            }

            return ExecuteQuery(query);
        }

        /// <summary>
        /// Consulta stock de HANA por lotes (OPTIMIZADO para grandes volúmenes)
        /// </summary>
        public DataTable ConsultarStockHANAPorLotes(DataTable facturasPendientes)
        {
            DataTable resultadoCompleto = CrearEstructuraStockTable();

            try
            {
                if (facturasPendientes == null || facturasPendientes.Rows.Count == 0)
                    return resultadoCompleto;

                // PRE-CARGAR TODOS LOS FACTORES UoM EN UNA SOLA CONSULTA
                PreCargarFactoresUoM(facturasPendientes);

                // Extraer solo ItemCodes únicos (consultar todos los almacenes)
                var itemCodesUnicos = facturasPendientes.AsEnumerable()
                    .Select(row => row.Field<string>("ItemCode"))
                    .Where(x => !string.IsNullOrEmpty(x))
                    .Distinct()
                    .ToList();

                if (itemCodesUnicos.Count == 0)
                    return resultadoCompleto;

                // Procesar en lotes MÁS GRANDES
                int totalLotes = (int)Math.Ceiling((double)itemCodesUnicos.Count / BATCH_SIZE);

                for (int lote = 0; lote < totalLotes; lote++)
                {
                    var itemsLote = itemCodesUnicos
                        .Skip(lote * BATCH_SIZE)
                        .Take(BATCH_SIZE)
                        .ToList();

                    DataTable stockLote = ConsultarStockLoteOptimizado(itemsLote);
                    
                    // Combinar resultados
                    foreach (DataRow row in stockLote.Rows)
                    {
                        resultadoCompleto.ImportRow(row);
                    }
                }

                return resultadoCompleto;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al consultar stock por lotes: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// PRE-CARGA todos los factores UoM en UNA SOLA consulta batch
        /// </summary>
        private void PreCargarFactoresUoM(DataTable facturasPendientes)
        {
            try
            {
                // Extraer combinaciones únicas de ItemCode + UomEntry
                var combinacionesUom = facturasPendientes.AsEnumerable()
                    .Where(r => r["UomEntry"] != DBNull.Value && Convert.ToInt32(r["UomEntry"]) > 0)
                    .Select(r => new
                    {
                        ItemCode = r.Field<string>("ItemCode"),
                        UomEntry = Convert.ToInt32(r["UomEntry"])
                    })
                    .Where(x => !string.IsNullOrEmpty(x.ItemCode))
                    .Distinct()
                    .ToList();

                if (combinacionesUom.Count == 0)
                    return;

                // Construir IN clauses
                var itemsList = string.Join(",", combinacionesUom
                    .Select(x => $"'{EscapeSQLString(x.ItemCode)}'")
                    .Distinct());

                var uomsList = string.Join(",", combinacionesUom
                    .Select(x => x.UomEntry)
                    .Distinct());

                // UNA SOLA consulta para TODOS los factores
                string query = $@"
                    SELECT 
                        T0.""ItemCode"",
                        U.""UomEntry"",
                        U.""BaseQty"" AS ""Factor""
                    FROM ""OITM"" T0
                    INNER JOIN ""UGP1"" U ON T0.""UgpEntry"" = U.""UgpEntry""
                    WHERE T0.""ItemCode"" IN ({itemsList})
                      AND U.""UomEntry"" IN ({uomsList})";

                DataTable factores = ExecuteQuery(query);

                // Poblar caché
                foreach (DataRow row in factores.Rows)
                {
                    string itemCode = row["ItemCode"].ToString();
                    int uomEntry = Convert.ToInt32(row["UomEntry"]);
                    decimal factor = row["Factor"] != DBNull.Value 
                        ? Convert.ToDecimal(row["Factor"]) 
                        : 1m;

                    string cacheKey = $"{itemCode}_{uomEntry}";
                    cacheFactoresUoM[cacheKey] = factor;
                }
            }
            catch (Exception ex)
            {
                // Si falla la pre-carga, continuamos con factores = 1
                System.Diagnostics.Debug.WriteLine($"Advertencia pre-carga UoM: {ex.Message}");
            }
        }

        /// <summary>
        /// Consulta un lote de artículos OPTIMIZADO (sin filtro de WhsCode)
        /// </summary>
        private DataTable ConsultarStockLoteOptimizado(List<string> itemCodes)
        {
            Recordset rs = null;
            DataTable dt = CrearEstructuraStockTable();

            try
            {
                rs = connectionManager.CreateRecordset();

                // Construir lista con comillas
                string itemsList = string.Join(",", itemCodes.Select(i => $"'{EscapeSQLString(i)}'"));

                // Query SIN filtro de WhsCode (consultar todos los almacenes)
                string query = $@"
                    SELECT 
                        T0.""ItemCode"",
                        T0.""ItemName"",
                        T1.""WhsCode"",
                        T1.""OnHand"" AS ""StockDisponible"",
                        T1.""IsCommited"" AS ""Comprometido"",
                        (T1.""OnHand"" - T1.""IsCommited"") AS ""StockReal"",
                        T0.""InvntryUom"" AS ""BaseUoM"",
                        T0.""UgpEntry""
                    FROM ""OITM"" T0
                    INNER JOIN ""OITW"" T1 ON T0.""ItemCode"" = T1.""ItemCode""
                    WHERE T0.""ItemCode"" IN ({itemsList})
                      AND T0.""InvntItem"" = 'Y'
                    ORDER BY T0.""ItemCode"", T1.""WhsCode""";

                rs.DoQuery(query);
                dt = RecordsetToDataTable(rs);

                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en lote de consulta: {ex.Message}", ex);
            }
            finally
            {
                if (rs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                }
            }
        }

        /// <summary>
        /// Clase auxiliar para agrupar ItemCode y WhsCode
        /// </summary>
        private class ItemWhsKey
        {
            public string ItemCode { get; set; }
            public string WhsCode { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is ItemWhsKey other)
                {
                    return ItemCode == other.ItemCode && WhsCode == other.WhsCode;
                }
                return false;
            }

            public override int GetHashCode()
            {
                return (ItemCode ?? "").GetHashCode() ^ (WhsCode ?? "").GetHashCode();
            }
        }

        /// <summary>
        /// Consulta un lote específico de artículos
        /// </summary>
        private DataTable ConsultarStockLote(List<ItemWhsKey> itemsLote)
        {
            Recordset rs = null;
            DataTable dt = CrearEstructuraStockTable();

            try
            {
                rs = connectionManager.CreateRecordset();

                // Construir query optimizado
                string query = ConstruirQueryLote(itemsLote);

                rs.DoQuery(query);
                dt = RecordsetToDataTable(rs);

                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en lote de consulta: {ex.Message}", ex);
            }
            finally
            {
                if (rs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                }
            }
        }

        /// <summary>
        /// Construye query optimizado para un lote
        /// </summary>
        private string ConstruirQueryLote(List<ItemWhsKey> items)
        {
            // Extraer ItemCodes y WhsCodes únicos
            var itemCodes = items.Select(x => x.ItemCode).Distinct().ToList();
            var whsCodes = items.Select(x => x.WhsCode)
                .Where(w => !string.IsNullOrEmpty(w))
                .Distinct()
                .ToList();

            // Construir listas con comillas
            string itemsList = string.Join(",", itemCodes.Select(i => $"'{EscapeSQLString(i)}'"));
            string whsList = string.Join(",", whsCodes.Select(w => $"'{EscapeSQLString(w)}'"));

            string query = $@"
                SELECT 
                    T0.""ItemCode"",
                    T0.""ItemName"",
                    T1.""WhsCode"",
                    T1.""OnHand"" AS ""StockDisponible"",
                    T1.""IsCommited"" AS ""Comprometido"",
                    (T1.""OnHand"" - T1.""IsCommited"") AS ""StockReal"",
                    T0.""InvntryUom"" AS ""BaseUoM"",
                    T0.""UgpEntry""
                FROM ""OITM"" T0
                INNER JOIN ""OITW"" T1 ON T0.""ItemCode"" = T1.""ItemCode""
                WHERE T0.""ItemCode"" IN ({itemsList})
                  AND T0.""InvntItem"" = 'Y'";

            if (whsCodes.Count > 0)
            {
                query += $@" AND T1.""WhsCode"" IN ({whsList})";
            }

            query += @" ORDER BY T0.""ItemCode"", T1.""WhsCode""";

            return query;
        }

        /// <summary>
        /// Obtiene el factor de conversión desde caché (ya no consulta HANA)
        /// </summary>
        public decimal ObtenerFactorConversionUoM(string itemCode, int uomEntry)
        {
            if (uomEntry <= 0)
                return 1m;

            string cacheKey = $"{itemCode}_{uomEntry}";
            
            if (cacheFactoresUoM.ContainsKey(cacheKey))
            {
                return cacheFactoresUoM[cacheKey];
            }

            // Si no está en caché, retornar 1 (ya debería estar pre-cargado)
            return 1m;
        }

        /// <summary>
        /// Crea la estructura de DataTable para stock
        /// </summary>
        private DataTable CrearEstructuraStockTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ItemCode", typeof(string));
            dt.Columns.Add("ItemName", typeof(string));
            dt.Columns.Add("WhsCode", typeof(string));
            dt.Columns.Add("StockDisponible", typeof(decimal));
            dt.Columns.Add("Comprometido", typeof(decimal));
            dt.Columns.Add("StockReal", typeof(decimal));
            dt.Columns.Add("BaseUoM", typeof(string));
            dt.Columns.Add("UgpEntry", typeof(int));
            return dt;
        }

        #region Helper Methods

        /// <summary>
        /// Construye la llamada a un Stored Procedure
        /// </summary>
        private string BuildStoredProcedureCall(string procedureName, Dictionary<string, object> parameters)
        {
            if (parameters == null || parameters.Count == 0)
            {
                return $"CALL {procedureName}()";
            }

            List<string> paramList = new List<string>();
            foreach (var param in parameters)
            {
                string value = FormatParameterValue(param.Value);
                paramList.Add(value);
            }

            return $"CALL {procedureName}({string.Join(", ", paramList)})";
        }

        /// <summary>
        /// Formatea un valor de parámetro para SQL
        /// </summary>
        private string FormatParameterValue(object value)
        {
            if (value == null)
            {
                return "NULL";
            }

            if (value is string)
            {
                return $"'{EscapeSQLString(value.ToString())}'";
            }

            if (value is DateTime)
            {
                DateTime dt = (DateTime)value;
                return $"'{dt:yyyy-MM-dd}'";
            }

            if (value is bool)
            {
                return (bool)value ? "1" : "0";
            }

            return value.ToString();
        }

        /// <summary>
        /// Escapa caracteres especiales en strings SQL
        /// </summary>
        private string EscapeSQLString(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return value;
            }

            return value.Replace("'", "''");
        }

        /// <summary>
        /// Convierte un RecordSet de SAP a DataTable
        /// </summary>
        private DataTable RecordsetToDataTable(Recordset rs)
        {
            DataTable dt = new DataTable();

            try
            {
                if (rs.RecordCount == 0)
                {
                    return dt;
                }

                // Crear columnas
                for (int i = 0; i < rs.Fields.Count; i++)
                {
                    Field field = rs.Fields.Item(i);
                    DataColumn column = new DataColumn(field.Name);

                    // Mapear tipos de datos
                    switch (field.Type)
                    {
                        case BoFieldTypes.db_Alpha:
                        case BoFieldTypes.db_Memo:
                            column.DataType = typeof(string);
                            break;
                        case BoFieldTypes.db_Numeric:
                            column.DataType = typeof(decimal);
                            break;
                        case BoFieldTypes.db_Date:
                            column.DataType = typeof(DateTime);
                            break;
                        case BoFieldTypes.db_Float:
                            column.DataType = typeof(double);
                            break;
                        default:
                            column.DataType = typeof(object);
                            break;
                    }

                    dt.Columns.Add(column);
                }

                // Cargar datos
                rs.MoveFirst();
                while (!rs.EoF)
                {
                    DataRow row = dt.NewRow();

                    for (int i = 0; i < rs.Fields.Count; i++)
                    {
                        Field field = rs.Fields.Item(i);
                        object value = field.Value;

                        if (value != null && value != DBNull.Value)
                        {
                            row[i] = value;
                        }
                        else
                        {
                            row[i] = DBNull.Value;
                        }
                    }

                    dt.Rows.Add(row);
                    rs.MoveNext();
                }

                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al convertir RecordSet a DataTable: {ex.Message}", ex);
            }
        }

        #endregion
    }
}
