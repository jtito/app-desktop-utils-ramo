using System;
using System.Windows.Forms;

namespace RamoUtils
{
    public partial class FrmPrincipal : Form
    {
        public FrmPrincipal()
        {
            InitializeComponent();
            ConfigurarMenu();
        }

        private void ConfigurarMenu()
        {
            this.Text = "Sistema de Migración SAP - Utilidades";
            this.WindowState = FormWindowState.Maximized;
            this.IsMdiContainer = true;
        }

        private void consultarStockToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Verificar si ya existe el formulario abierto
            foreach (Form frm in this.MdiChildren)
            {
                if (frm is FrmConsultaStock)
                {
                    frm.Activate();
                    return;
                }
            }

            // Crear nueva instancia del formulario
            FrmConsultaStock frmStock = new FrmConsultaStock();
            frmStock.MdiParent = this;
            frmStock.Show();
        }

        private void integracionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Verificar si ya existe el formulario abierto
            foreach (Form frm in this.MdiChildren)
            {
                if (frm is FrmConsultaIntegracion)
                {
                    frm.Activate();
                    return;
                }
            }

            // Crear nueva instancia del formulario
            FrmConsultaIntegracion frmIntegracion = new FrmConsultaIntegracion();
            frmIntegracion.MdiParent = this;
            frmIntegracion.Show();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("¿Está seguro que desea salir?", "Confirmar",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }
    }
}
