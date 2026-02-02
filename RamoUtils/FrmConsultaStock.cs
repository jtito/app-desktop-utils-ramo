using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RamoUtils
{
    public partial class FrmConsultaStock : Form
    {
        private StockService stockService;
        private ConnectionManager connectionManager;

        public FrmConsultaStock()
        {
            InitializeComponent();
            InicializarFormulario();
        }

        private void InicializarFormulario()
        {
            // Configurar DateTimePickers para rango de fechas
            dtpFechaInicio.Format = DateTimePickerFormat.Short;
            dtpFechaFin.Format = DateTimePickerFormat.Short;
            dtpFechaInicio.Value = DateTime.Now;
            dtpFechaFin.Value = DateTime.Now;

            // Configurar DataGridView
            ConfigurarGrilla();

            // Inicializar servicios de conexión
            try
            {
                // Validar que existe conexión SAP global
                if (Program.SAPCompanyGlobal == null || !Program.SAPCompanyGlobal.Connected)
                {
                    throw new Exception("No hay conexión SAP activa. Por favor, reinicie la aplicación.");
                }

                // Validar configuraciones SQL
                ValidarConfiguraciones();

                // Crear gestor de conexiones (usa la conexión SAP global)
                connectionManager = new ConnectionManager();

                // Crear servicio de stock
                stockService = new StockService(connectionManager);

                // Validar solo conexión SQL (SAP ya está validado por el login)
                ValidarConexionesInicio();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error al inicializar el formulario:\n\n{ex.Message}\n\n" +
                    "Por favor, revise la configuración.",
                    "Error de Configuración",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                lblEstado.Text = "Error de configuración";
                lblEstado.ForeColor = Color.Red;
                lblEstado.Visible = true;
                btnBuscar.Enabled = false;
            }
        }

        private void ValidarConfiguraciones()
        {
            try
            {
                // Validar solo configuración SQL
                ConfigHelper.ValidarConfiguracionSQL();
            }
            catch (Exception ex)
            {
                throw new Exception($"Configuración SQL incompleta: {ex.Message}", ex);
            }
        }

        private void ValidarConexionesInicio()
        {
            try
            {
                lblEstado.Text = "Validando conexión SQL...";
                lblEstado.ForeColor = Color.Blue;
                lblEstado.Visible = true;
                Application.DoEvents();

                bool sqlOk, sapOk;
                string sqlError, sapError;

                stockService.ValidarConexiones(out sqlOk, out sapOk, out sqlError, out sapError);

                if (!sqlOk)
                {
                    string mensaje = $"Advertencia de conexión SQL:\n\n{sqlError}\n\n¿Desea continuar de todos modos?";

                    var result = MessageBox.Show(
                        mensaje,
                        "Advertencia de Conexión",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

                    if (result == DialogResult.No)
                    {
                        btnBuscar.Enabled = false;
                        lblEstado.Text = "Conexión SQL no validada";
                        lblEstado.ForeColor = Color.Red;
                        return;
                    }
                }

                lblEstado.Text = "Listo para consultar";
                lblEstado.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error al validar conexiones:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                lblEstado.Text = "Advertencia: No se pudo validar la conexión";
                lblEstado.ForeColor = Color.Orange;
            }
        }

        private void ConfigurarGrilla()
        {
            dgvResultados.AutoGenerateColumns = false;
            dgvResultados.AllowUserToAddRows = false;
            dgvResultados.ReadOnly = true;
            dgvResultados.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dgvResultados.MultiSelect = true;
            dgvResultados.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText;
            dgvResultados.MultiSelect = false;

            // Configurar columnas
            dgvResultados.Columns.Clear();

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "ItemCode",
                HeaderText = "Código Artículo",
                DataPropertyName = "ItemCode",
                Width = 120
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "ItemName",
                HeaderText = "Descripción",
                DataPropertyName = "ItemName",
                Width = 250
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "WhsCode",
                HeaderText = "Almacén",
                DataPropertyName = "WhsCode",
                Width = 80
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "WhsCode",
                HeaderText = "Unidad Medida",
                DataPropertyName = "WhsCode",
                Width = 80
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "CantidadRequerida",
                HeaderText = "Cant. Requerida",
                DataPropertyName = "CantidadRequerida",
                Width = 100,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight }
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "UomCode",
                HeaderText = "Unidad",
                DataPropertyName = "UomCode",
                Width = 60
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "FactorConversion",
                HeaderText = "Factor",
                DataPropertyName = "FactorConversion",
                Width = 60,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", Alignment = DataGridViewContentAlignment.MiddleCenter }
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "CantidadRequeridaBase",
                HeaderText = "Cant. Base (UND)",
                DataPropertyName = "CantidadRequeridaBase",
                Width = 100,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Format = "N2",
                    Alignment = DataGridViewContentAlignment.MiddleRight,
                    BackColor = Color.LightYellow
                }
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "StockDisponible",
                HeaderText = "Stock Disponible (UND)",
                DataPropertyName = "StockDisponible",
                Width = 120,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight }
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Faltante",
                HeaderText = "Faltante (UND)",
                DataPropertyName = "Faltante",
                Width = 100,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Format = "N2",
                    Alignment = DataGridViewContentAlignment.MiddleRight,
                    ForeColor = Color.Red,
                    Font = new Font(dgvResultados.Font, FontStyle.Bold)
                }
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "UomName",
                HeaderText = "Nombre Unidad",
                DataPropertyName = "UomName",
                Width = 120
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "MensajeError",
                HeaderText = "Mensaje de Error",
                DataPropertyName = "MensajeError",
                Width = 300,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    ForeColor = Color.DarkRed,
                    Font = new Font(dgvResultados.Font, FontStyle.Italic)
                }
            });
        }

        private async void btnBuscar_Click(object sender, EventArgs e)
        {
            try
            {
                // Validar rango de fechas
                if (dtpFechaInicio.Value.Date > dtpFechaFin.Value.Date)
                {
                    MessageBox.Show(
                        "La fecha de inicio no puede ser mayor que la fecha fin.",
                        "Validación",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Deshabilitar controles durante la búsqueda
                btnBuscar.Enabled = false;
                dtpFechaInicio.Enabled = false;
                dtpFechaFin.Enabled = false;
                Cursor = Cursors.WaitCursor;

                lblEstado.Text = "Consultando base de datos...";
                lblEstado.ForeColor = Color.Blue;
                lblEstado.Visible = true;

                // Realizar la consulta con rango de fechas
                DateTime fechaInicio = dtpFechaInicio.Value.Date;
                DateTime fechaFin = dtpFechaFin.Value.Date;
                
                var resultados = await System.Threading.Tasks.Task.Run(() =>
                    stockService.ConsultarStockPorRangoFechas(fechaInicio, fechaFin));

                // Mostrar resultados
                dgvResultados.DataSource = null;
                dgvResultados.DataSource = resultados;

                lblEstado.Text = $"Se encontraron {resultados.Count} artículos con stock insuficiente";
                lblEstado.ForeColor = resultados.Count > 0 ? Color.Red : Color.Green;

                if (resultados.Count == 0)
                {
                    lblEstado.Text = "? No se encontraron artículos con problemas de stock en el rango de fechas seleccionado";
                    lblEstado.ForeColor = Color.Green;
                }
            }
            catch (Exception ex)
            {
                lblEstado.Text = "Error al consultar";
                lblEstado.ForeColor = Color.Red;
                MessageBox.Show($"Error al consultar stock:\n\n{ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Rehabilitar controles
                btnBuscar.Enabled = true;
                dtpFechaInicio.Enabled = true;
                dtpFechaFin.Enabled = true;
                Cursor = Cursors.Default;
            }
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            if (dgvResultados.Rows.Count == 0)
            {
                MessageBox.Show("No hay datos para exportar", "Información",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Configurar SaveFileDialog
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Archivo CSV (*.csv)|*.csv|Archivo Excel (*.xls)|*.xls|Todos los archivos (*.*)|*.*";
                saveDialog.FilterIndex = 1;
                saveDialog.Title = "Exportar Reporte de Stock";
                saveDialog.FileName = $"ReporteStock_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                saveDialog.DefaultExt = "csv";
                saveDialog.AddExtension = true;

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    Cursor = Cursors.WaitCursor;
                    lblEstado.Text = "Exportando datos...";
                    lblEstado.ForeColor = Color.Blue;
                    Application.DoEvents();

                    // Exportar según la extensión seleccionada
                    string extension = Path.GetExtension(saveDialog.FileName).ToLower();

                    if (extension == ".csv")
                    {
                        ExportarCSV(saveDialog.FileName);
                    }
                    else
                    {
                        ExportarHTML(saveDialog.FileName);
                    }

                    lblEstado.Text = $"Archivo exportado exitosamente: {Path.GetFileName(saveDialog.FileName)}";
                    lblEstado.ForeColor = Color.Green;

                    // Preguntar si desea abrir el archivo
                    var result = MessageBox.Show(
                        $"Archivo exportado exitosamente:\n{saveDialog.FileName}\n\n¿Desea abrir el archivo ahora?",
                        "Exportación Exitosa",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information);

                    if (result == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(saveDialog.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                lblEstado.Text = "Error al exportar";
                lblEstado.ForeColor = Color.Red;
                MessageBox.Show($"Error al exportar datos:\n\n{ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void ExportarCSV(string filePath)
        {
            using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                // Escribir encabezado
                StringBuilder headerLine = new StringBuilder();
                foreach (DataGridViewColumn column in dgvResultados.Columns)
                {
                    if (column.Visible)
                    {
                        headerLine.Append(EscapeCSV(column.HeaderText));
                        headerLine.Append(",");
                    }
                }
                if (headerLine.Length > 0)
                {
                    headerLine.Length--; // Remover última coma
                }
                sw.WriteLine(headerLine.ToString());

                // Escribir datos
                foreach (DataGridViewRow row in dgvResultados.Rows)
                {
                    if (row.IsNewRow) continue;

                    StringBuilder dataLine = new StringBuilder();
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.OwningColumn.Visible)
                        {
                            string value = cell.Value?.ToString() ?? "";
                            dataLine.Append(EscapeCSV(value));
                            dataLine.Append(",");
                        }
                    }
                    if (dataLine.Length > 0)
                    {
                        dataLine.Length--; // Remover última coma
                    }
                    sw.WriteLine(dataLine.ToString());
                }
            }
        }

        private void ExportarHTML(string filePath)
        {
            using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                // Escribir encabezado HTML con estilos
                sw.WriteLine("<html>");
                sw.WriteLine("<head>");
                sw.WriteLine("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />");
                sw.WriteLine($"<title>Reporte de Stock - {DateTime.Now:dd/MM/yyyy}</title>");
                sw.WriteLine("<style>");
                sw.WriteLine("body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; }");
                sw.WriteLine("table { border-collapse: collapse; width: 100%; }");
                sw.WriteLine("th { background-color: #4472C4; color: white; font-weight: bold; padding: 8px; text-align: left; border: 1px solid #ddd; }");
                sw.WriteLine("td { padding: 6px; border: 1px solid #ddd; }");
                sw.WriteLine("tr:nth-child(even) { background-color: #f2f2f2; }");
                sw.WriteLine(".number { text-align: right; }");
                sw.WriteLine(".faltante { color: red; font-weight: bold; }");
                sw.WriteLine("h1 { color: #4472C4; }");
                sw.WriteLine(".info { margin: 10px 0; color: #666; }");
                sw.WriteLine("</style>");
                sw.WriteLine("</head>");
                sw.WriteLine("<body>");

                // Título y metadatos
                sw.WriteLine($"<h1>Reporte de Artículos con Stock Insuficiente</h1>");
                sw.WriteLine($"<div class=\"info\">");
                sw.WriteLine($"<p><strong>Rango de Fechas:</strong> {dtpFechaInicio.Value:dd/MM/yyyy} - {dtpFechaFin.Value:dd/MM/yyyy}</p>");
                sw.WriteLine($"<p><strong>Fecha de Generación:</strong> {DateTime.Now:dd/MM/yyyy HH:mm:ss}</p>");
                sw.WriteLine($"<p><strong>Total de Registros:</strong> {dgvResultados.Rows.Count}</p>");
                sw.WriteLine($"</div>");

                // Tabla
                sw.WriteLine("<table>");

                // Encabezado
                sw.WriteLine("<thead><tr>");
                foreach (DataGridViewColumn column in dgvResultados.Columns)
                {
                    if (column.Visible)
                    {
                        sw.WriteLine($"<th>{EscapeHTML(column.HeaderText)}</th>");
                    }
                }
                sw.WriteLine("</tr></thead>");

                // Datos
                sw.WriteLine("<tbody>");
                foreach (DataGridViewRow row in dgvResultados.Rows)
                {
                    if (row.IsNewRow) continue;

                    sw.WriteLine("<tr>");
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.OwningColumn.Visible)
                        {
                            string value = cell.Value?.ToString() ?? "";
                            string cssClass = "";

                            // Aplicar estilos según el tipo de columna
                            if (cell.OwningColumn.Name == "Faltante")
                            {
                                cssClass = " class=\"number faltante\"";
                            }
                            else if (cell.OwningColumn.Name == "CantidadRequerida" ||
                                     cell.OwningColumn.Name == "StockDisponible")
                            {
                                cssClass = " class=\"number\"";
                            }

                            sw.WriteLine($"<td{cssClass}>{EscapeHTML(value)}</td>");
                        }
                    }
                    sw.WriteLine("</tr>");
                }
                sw.WriteLine("</tbody>");
                sw.WriteLine("</table>");

                // Pie de página
                sw.WriteLine("<br/>");
                sw.WriteLine($"<div class=\"info\">");
                sw.WriteLine($"<p><em>Generado por RamoUtils - Sistema de Consulta de Stock SAP</em></p>");
                sw.WriteLine($"</div>");

                sw.WriteLine("</body>");
                sw.WriteLine("</html>");
            }
        }

        private string EscapeCSV(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";

            // Si contiene coma, comillas o salto de línea, encerrar entre comillas
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n") || value.Contains("\r"))
            {
                // Duplicar comillas internas
                value = value.Replace("\"", "\"\"");
                return $"\"{value}\"";
            }

            return value;
        }

        private string EscapeHTML(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";

            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&#39;");
        }

        private void dgvResultados_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Evento para futuras funcionalidades
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("Cerrando formulario FrmConsultaStock...");

                // Limpiar DataGridView
                if (dgvResultados != null)
                {
                    dgvResultados.DataSource = null;
                }

                // Liberar StockService
                if (stockService != null)
                {
                    System.Diagnostics.Debug.WriteLine("Liberando StockService...");
                    try
                    {
                        stockService.Dispose();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error al liberar StockService: {ex.Message}");
                    }
                    stockService = null;
                }

                // Liberar ConnectionManager
                if (connectionManager != null)
                {
                    System.Diagnostics.Debug.WriteLine("Liberando ConnectionManager...");
                    try
                    {
                        connectionManager.Dispose();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error al liberar ConnectionManager: {ex.Message}");
                    }
                    connectionManager = null;
                }

                // Forzar recolección de basura para liberar objetos COM
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                System.Diagnostics.Debug.WriteLine("Formulario FrmConsultaStock cerrado correctamente");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en OnFormClosing: {ex.Message}");
            }
            finally
            {
                base.OnFormClosing(e);
            }
        }

        private void FrmConsultaStock_Load(object sender, EventArgs e)
        {

        }
    }
}
