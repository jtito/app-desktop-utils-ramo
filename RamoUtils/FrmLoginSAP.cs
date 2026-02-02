using System;
using System.Windows.Forms;
using SAPbobsCOM;

namespace RamoUtils
{
    public partial class FrmLoginSAP : Form
    {
        public Company SAPCompany { get; private set; }

        public FrmLoginSAP()
        {
            InitializeComponent();
        }

        private void btnConectar_Click(object sender, EventArgs e)
        {
            if (!ValidarCampos())
                return;

            // Deshabilitar controles durante la conexión
            HabilitarControles(false);
            lblEstado.Text = "Conectando a SAP Business One...";
            lblEstado.Visible = true;
            Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            try
            {
                // Validar que exista configuración SAP en App.config
                ConfigHelper.ValidarConfiguracionSAP();

                // Crear instancia de Company
                SAPCompany = new Company();

                // Configurar conexión desde App.config
                SAPCompany.Server = ConfigHelper.GetSAPServer();
                SAPCompany.CompanyDB = ConfigHelper.GetSAPCompanyDB();
                SAPCompany.DbUserName = ConfigHelper.GetSAPDbUserName();
                SAPCompany.DbPassword = ConfigHelper.GetSAPDbPassword();

                // Credenciales SAP desde el formulario
                SAPCompany.UserName = txtUsuario.Text.Trim();
                SAPCompany.Password = txtPassword.Text;

                // Tipo de servidor desde App.config
                string dbServerType = ConfigHelper.GetSAPDbServerType();
                if (dbServerType == "dst_HANADB")
                    SAPCompany.DbServerType = BoDataServerTypes.dst_HANADB;
                else
                    SAPCompany.DbServerType = BoDataServerTypes.dst_MSSQL2019;

                SAPCompany.UseTrusted = false;

                // Intentar conectar
                int resultado = SAPCompany.Connect();

                if (resultado != 0)
                {
                    string errorMsg = SAPCompany.GetLastErrorDescription();
                    int errorCode = SAPCompany.GetLastErrorCode();

                    // Liberar recursos
                    if (SAPCompany != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(SAPCompany);
                        SAPCompany = null;
                    }

                    MessageBox.Show(
                        $"Error al conectar a SAP Business One:\n\n" +
                        $"Código: {errorCode}\n" +
                        $"Mensaje: {errorMsg}",
                        "Error de Conexión",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    lblEstado.Text = "Error de conexión";
                    HabilitarControles(true);
                    Cursor = Cursors.Default;
                    return;
                }

                // Conexión exitosa
                lblEstado.Text = "Conexión exitosa";
                MessageBox.Show(
                    "Conexión exitosa a SAP Business One",
                    "Éxito",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                // Limpiar en caso de error
                if (SAPCompany != null)
                {
                    try
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(SAPCompany);
                    }
                    catch { }
                    SAPCompany = null;
                }

                MessageBox.Show(
                    $"Error al conectar a SAP:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                lblEstado.Text = "Error de conexión";
                HabilitarControles(true);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private bool ValidarCampos()
        {
            if (string.IsNullOrWhiteSpace(txtUsuario.Text))
            {
                MessageBox.Show("Debe ingresar el usuario SAP", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtUsuario.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                MessageBox.Show("Debe ingresar la contraseña SAP", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtPassword.Focus();
                return false;
            }

            return true;
        }

        private void HabilitarControles(bool habilitar)
        {
            txtUsuario.Enabled = habilitar;
            txtPassword.Enabled = habilitar;
            btnConectar.Enabled = habilitar;
            btnCancelar.Enabled = habilitar;
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void FrmLoginSAP_Load(object sender, EventArgs e)
        {
            lblEstado.Visible = false;
            
            // Mostrar información de configuración desde App.config
            try
            {
                lblServidor.Text = $"Servidor: {ConfigHelper.GetSAPServer()}";
                lblCompanyDB.Text = $"CompanyDB: {ConfigHelper.GetSAPCompanyDB()}";
            }
            catch
            {
                lblServidor.Text = "Servidor: (No configurado)";
                lblCompanyDB.Text = "CompanyDB: (No configurado)";
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void lblServidor_Click(object sender, EventArgs e)
        {

        }
    }
}
