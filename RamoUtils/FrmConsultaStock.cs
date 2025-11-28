using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
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
            // Configurar DateTimePicker
            dtpFecha.Format = DateTimePickerFormat.Short;
            dtpFecha.Value = DateTime.Now;

            // Configurar DataGridView
            ConfigurarGrilla();

            // Inicializar servicios de conexión
            try
            {
                // Validar configuraciones
                ValidarConfiguraciones();

                // Crear gestor de conexiones
                connectionManager = new ConnectionManager();

                // Crear servicio de stock
                stockService = new StockService(connectionManager);

                // Validar conexiones
                ValidarConexionesInicio();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error al inicializar el formulario:\n\n{ex.Message}\n\n" +
                    "Por favor, revise el archivo App.config y configure correctamente " +
                    "las cadenas de conexión y parámetros SAP.",
                    "Error de Configuración",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                lblEstado.Text = "Error de configuración - Revise App.config";
                lblEstado.ForeColor = Color.Red;
                lblEstado.Visible = true;
                btnBuscar.Enabled = false;
            }
        }

        private void ValidarConfiguraciones()
        {
            try
            {
                // Validar configuración SQL
                ConfigHelper.ValidarConfiguracionSQL();

                // Validar configuración SAP
                ConfigHelper.ValidarConfiguracionSAP();
            }
            catch (Exception ex)
            {
                throw new Exception($"Configuración incompleta: {ex.Message}", ex);
            }
        }

        private void ValidarConexionesInicio()
        {
            try
            {
                lblEstado.Text = "Validando conexiones...";
                lblEstado.ForeColor = Color.Blue;
                lblEstado.Visible = true;
                Application.DoEvents();

                bool sqlOk, sapOk;
                string sqlError, sapError;

                stockService.ValidarConexiones(out sqlOk, out sapOk, out sqlError, out sapError);

                if (!sqlOk || !sapOk)
                {
                    string mensaje = "Advertencia de conexión:\n\n";

                    if (!sqlOk)
                    {
                        mensaje += $"? SQL Server: {sqlError}\n\n";
                    }
                    else
                    {
                        mensaje += "? SQL Server: Conectado\n\n";
                    }

                    if (!sapOk)
                    {
                        mensaje += $"? SAP HANA: {sapError}\n\n";
                    }
                    else
                    {
                        mensaje += "? SAP HANA: Conectado\n\n";
                    }

                    mensaje += "¿Desea continuar de todos modos?";

                    var result = MessageBox.Show(
                        mensaje,
                        "Advertencia de Conexión",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

                    if (result == DialogResult.No)
                    {
                        btnBuscar.Enabled = false;
                        lblEstado.Text = "Conexiones no validadas";
                        lblEstado.ForeColor = Color.Red;
                        return;
                    }
                }

                lblEstado.Text = "Conexiones validadas correctamente - Listo para consultar";
                lblEstado.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error al validar conexiones:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                lblEstado.Text = "Advertencia: No se pudieron validar las conexiones";
                lblEstado.ForeColor = Color.Orange;
            }
        }

        private void ConfigurarGrilla()
        {
            dgvResultados.AutoGenerateColumns = false;
            dgvResultados.AllowUserToAddRows = false;
            dgvResultados.ReadOnly = true;
            dgvResultados.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dgvResultados.MultiSelect = true; // Permitir selección múltiple
            dgvResultados.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText;
            dgvResultados.MultiSelect = false;

            // Configurar columnas
            dgvResultados.Columns.Clear();

            //dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            //{
            //    Name = "NumeroFactura",
            //    HeaderText = "Nº Factura",
            //    DataPropertyName = "NumeroFactura",
            //    Width = 100
            //});

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
                Name = "CantidadRequerida",
                HeaderText = "Cant. Requerida",
                DataPropertyName = "CantidadRequerida",
                Width = 100,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight }
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "StockDisponible",
                HeaderText = "Stock Disponible",
                DataPropertyName = "StockDisponible",
                Width = 100,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight }
            });

            dgvResultados.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Faltante",
                HeaderText = "Faltante",
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
        }

        private async void btnBuscar_Click(object sender, EventArgs e)
        {
            try
            {
                // Deshabilitar controles durante la búsqueda
                btnBuscar.Enabled = false;
                dtpFecha.Enabled = false;
                Cursor = Cursors.WaitCursor;

                lblEstado.Text = "Consultando base de datos...";
                lblEstado.ForeColor = Color.Blue;
                lblEstado.Visible = true;

                // Realizar la consulta
                DateTime fechaSeleccionada = dtpFecha.Value.Date;
                var resultados = await System.Threading.Tasks.Task.Run(() =>
                    stockService.ConsultarStockPorFecha(fechaSeleccionada));

                // Mostrar resultados
                dgvResultados.DataSource = null;
                dgvResultados.DataSource = resultados;

                lblEstado.Text = $"Se encontraron {resultados.Count} artículos con stock insuficiente";
                lblEstado.ForeColor = resultados.Count > 0 ? Color.Red : Color.Green;

                if (resultados.Count == 0)
                {
                    lblEstado.Text = "? No se encontraron artículos con problemas de stock para la fecha seleccionada";
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
                dtpFecha.Enabled = true;
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

            // Implementar exportación a Excel aquí si lo necesitas
            MessageBox.Show("Funcionalidad de exportación pendiente", "Información",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void dgvResultados_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Evento para futuras funcionalidades
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Liberar recursos
            if (stockService != null)
            {
                stockService.Dispose();
                stockService = null;
            }

            if (connectionManager != null)
            {
                connectionManager.Dispose();
                connectionManager = null;
            }

            base.OnFormClosing(e);
        }
    }
}
