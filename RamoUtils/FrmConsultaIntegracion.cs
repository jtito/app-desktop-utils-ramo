using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RamoUtils
{
    public partial class FrmConsultaIntegracion : Form
    {
        private ConnectionManager connectionManager;

        public FrmConsultaIntegracion()
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

            // Configurar ComboBox de Estados
            cboEstado.Items.Clear();
            cboEstado.Items.Add(new ComboBoxItem("TODOS", null));
            cboEstado.Items.Add(new ComboBoxItem("PENDIENTE", "PEN"));
            cboEstado.Items.Add(new ComboBoxItem("ERROR", "ERR"));
            cboEstado.Items.Add(new ComboBoxItem("FINALIZADO", "FIN"));
            cboEstado.DisplayMember = "Text";
            cboEstado.ValueMember = "Value";
            cboEstado.SelectedIndex = 0; // TODOS por defecto

            // Configurar DataGridView
            ConfigurarGrilla();

            // Inicializar conexión
            try
            {
                ConfigHelper.ValidarConfiguracionSQL();
                connectionManager = new ConnectionManager();
                lblEstado.Text = "Listo para consultar";
                lblEstado.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error al inicializar el formulario:\n\n{ex.Message}",
                    "Error de Configuración",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                lblEstado.Text = "Error de configuración";
                lblEstado.ForeColor = Color.Red;
                btnBuscar.Enabled = false;
            }
        }

        private void ConfigurarGrilla()
        {
            dgvResultados.AutoGenerateColumns = false;
            dgvResultados.AllowUserToAddRows = false;
            dgvResultados.ReadOnly = true;
            dgvResultados.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvResultados.MultiSelect = false;

            // Configurar columnas
            dgvResultados.Columns.Clear();

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "ID",
                HeaderText = "ID",
                DataPropertyName = "_key",
                Width = 80
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "MENSAJE",
                HeaderText = "Mensaje",
                DataPropertyName = "_ISMensajeError",
                Width = 300
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "FECHA_ING",
                HeaderText = "Fecha Ingreso",
                DataPropertyName = "FechaIngreso",
                Width = 120,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy HH:mm" }
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "ID_VENTA",
                HeaderText = "ID Venta",
                DataPropertyName = "U_RML_ID",
                Width = 100
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "TIPO",
                HeaderText = "Tipo",
                DataPropertyName = "U_BPP_MDTD",
                Width = 60
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "SERIE",
                HeaderText = "Serie",
                DataPropertyName = "U_BPP_MDSD",
                Width = 80
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "CORRELATIVO",
                HeaderText = "Correlativo",
                DataPropertyName = "U_BPP_MDCD",
                Width = 100
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "FECHA_DOC",
                HeaderText = "Fecha Doc",
                DataPropertyName = "Fecha",
                Width = 100,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy" }
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "TIPO_DOC",
                HeaderText = "Tipo Documento",
                DataPropertyName = "TipoDocumento",
                Width = 120
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
                cboEstado.Enabled = false;
                Cursor = Cursors.WaitCursor;

                lblEstado.Text = "Consultando base de datos...";
                lblEstado.ForeColor = Color.Blue;
                lblEstado.Visible = true;

                // Obtener parámetros
                string fechaInicio = dtpFechaInicio.Value.ToString("yyyyMMdd");
                string fechaFin = dtpFechaFin.Value.ToString("yyyyMMdd");
                ComboBoxItem estadoItem = (ComboBoxItem)cboEstado.SelectedItem;
                string estado = estadoItem.Value;

                // Realizar la consulta
                var resultados = await System.Threading.Tasks.Task.Run(() =>
                    ConsultarIntegracion(fechaInicio, fechaFin, estado));

                // Mostrar resultados
                dgvResultados.DataSource = null;
                dgvResultados.DataSource = resultados;

                lblEstado.Text = $"Se encontraron {resultados.Rows.Count} registros";
                lblEstado.ForeColor = resultados.Rows.Count > 0 ? Color.Green : Color.Orange;
            }
            catch (Exception ex)
            {
                lblEstado.Text = "Error al consultar";
                lblEstado.ForeColor = Color.Red;
                MessageBox.Show($"Error al consultar integración:\n\n{ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Rehabilitar controles
                btnBuscar.Enabled = true;
                dtpFechaInicio.Enabled = true;
                dtpFechaFin.Enabled = true;
                cboEstado.Enabled = true;
                Cursor = Cursors.Default;
            }
        }

        private DataTable ConsultarIntegracion(string fechaInicio, string fechaFin, string estado)
        {
            DataTable dt = new DataTable();

            try
            {
                using (SqlConnection conn = connectionManager.GetSQLConnection())
                {
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_RML_CONSULTAR_INTEGRACION", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = ConfigHelper.GetQueryTimeout();

                        cmd.Parameters.AddWithValue("@FechaInicio", fechaInicio);
                        cmd.Parameters.AddWithValue("@FechaFin", fechaFin);
                        cmd.Parameters.AddWithValue("@Estado", (object)estado ?? DBNull.Value);

                        using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                        {
                            adapter.Fill(dt);
                        }
                    }
                }

                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al ejecutar SP_RML_CONSULTAR_INTEGRACION: {ex.Message}", ex);
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
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Archivo CSV (*.csv)|*.csv|Archivo Excel (*.xls)|*.xls";
                saveDialog.FilterIndex = 1;
                saveDialog.Title = "Exportar Reporte de Integración";
                saveDialog.FileName = $"ReporteIntegracion_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                saveDialog.DefaultExt = "csv";
                saveDialog.AddExtension = true;

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    Cursor = Cursors.WaitCursor;
                    lblEstado.Text = "Exportando datos...";
                    lblEstado.ForeColor = Color.Blue;
                    Application.DoEvents();

                    string extension = Path.GetExtension(saveDialog.FileName).ToLower();

                    if (extension == ".csv")
                    {
                        ExportarCSV(saveDialog.FileName);
                    }
                    else
                    {
                        ExportarHTML(saveDialog.FileName);
                    }

                    lblEstado.Text = $"Archivo exportado: {Path.GetFileName(saveDialog.FileName)}";
                    lblEstado.ForeColor = Color.Green;

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
                    headerLine.Length--;
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
                        dataLine.Length--;
                    sw.WriteLine(dataLine.ToString());
                }
            }
        }

        private void ExportarHTML(string filePath)
        {
            using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                sw.WriteLine("<html>");
                sw.WriteLine("<head>");
                sw.WriteLine("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />");
                sw.WriteLine($"<title>Reporte de Integración - {DateTime.Now:dd/MM/yyyy}</title>");
                sw.WriteLine("<style>");
                sw.WriteLine("body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; }");
                sw.WriteLine("table { border-collapse: collapse; width: 100%; }");
                sw.WriteLine("th { background-color: #4472C4; color: white; font-weight: bold; padding: 8px; text-align: left; border: 1px solid #ddd; }");
                sw.WriteLine("td { padding: 6px; border: 1px solid #ddd; }");
                sw.WriteLine("tr:nth-child(even) { background-color: #f2f2f2; }");
                sw.WriteLine("h1 { color: #4472C4; }");
                sw.WriteLine(".info { margin: 10px 0; color: #666; }");
                sw.WriteLine("</style>");
                sw.WriteLine("</head>");
                sw.WriteLine("<body>");

                sw.WriteLine($"<h1>Reporte de Integración</h1>");
                sw.WriteLine($"<div class=\"info\">");
                sw.WriteLine($"<p><strong>Rango de Fechas:</strong> {dtpFechaInicio.Value:dd/MM/yyyy} - {dtpFechaFin.Value:dd/MM/yyyy}</p>");
                sw.WriteLine($"<p><strong>Estado:</strong> {cboEstado.Text}</p>");
                sw.WriteLine($"<p><strong>Fecha de Generación:</strong> {DateTime.Now:dd/MM/yyyy HH:mm:ss}</p>");
                sw.WriteLine($"<p><strong>Total de Registros:</strong> {dgvResultados.Rows.Count}</p>");
                sw.WriteLine($"</div>");

                sw.WriteLine("<table>");
                sw.WriteLine("<thead><tr>");
                foreach (DataGridViewColumn column in dgvResultados.Columns)
                {
                    if (column.Visible)
                        sw.WriteLine($"<th>{EscapeHTML(column.HeaderText)}</th>");
                }
                sw.WriteLine("</tr></thead>");

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
                            sw.WriteLine($"<td>{EscapeHTML(value)}</td>");
                        }
                    }
                    sw.WriteLine("</tr>");
                }
                sw.WriteLine("</tbody>");
                sw.WriteLine("</table>");

                sw.WriteLine("<br/>");
                sw.WriteLine($"<div class=\"info\">");
                sw.WriteLine($"<p><em>Generado por RamoUtils - Sistema de Consulta SAP</em></p>");
                sw.WriteLine($"</div>");

                sw.WriteLine("</body>");
                sw.WriteLine("</html>");
            }
        }

        private string EscapeCSV(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";

            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n") || value.Contains("\r"))
            {
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

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            try
            {
                if (dgvResultados != null)
                    dgvResultados.DataSource = null;

                if (connectionManager != null)
                {
                    connectionManager.Dispose();
                    connectionManager = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch { }
            finally
            {
                base.OnFormClosing(e);
            }
        }

        private void FrmConsultaIntegracion_Load(object sender, EventArgs e)
        {
        }

        // Clase auxiliar para ComboBox
        private class ComboBoxItem
        {
            public string Text { get; set; }
            public string Value { get; set; }

            public ComboBoxItem(string text, string value)
            {
                Text = text;
                Value = value;
            }
        }
    }
}
