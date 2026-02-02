using System;
using System.Windows.Forms;

namespace RamoUtils
{
    static class Program
    {
        // Variable global para mantener la conexión SAP
        public static SAPbobsCOM.Company SAPCompanyGlobal { get; set; }

        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Mostrar formulario de login SAP primero
            using (FrmLoginSAP frmLogin = new FrmLoginSAP())
            {
                DialogResult resultado = frmLogin.ShowDialog();

                if (resultado == DialogResult.OK && frmLogin.SAPCompany != null && frmLogin.SAPCompany.Connected)
                {
                    // Guardar la conexión SAP globalmente
                    SAPCompanyGlobal = frmLogin.SAPCompany;

                    try
                    {
                        // Iniciar el formulario principal
                        Application.Run(new FrmPrincipal());
                    }
                    finally
                    {
                        // Limpiar conexión SAP al cerrar la aplicación
                        if (SAPCompanyGlobal != null)
                        {
                            try
                            {
                                if (SAPCompanyGlobal.Connected)
                                {
                                    SAPCompanyGlobal.Disconnect();
                                }
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(SAPCompanyGlobal);
                            }
                            catch { }
                            finally
                            {
                                SAPCompanyGlobal = null;
                            }
                        }
                    }
                }
                else
                {
                    // Usuario canceló el login o falló la conexión
                    MessageBox.Show(
                        "No se pudo establecer conexión con SAP Business One.\nLa aplicación se cerrará.",
                        "Conexión Requerida",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
        }
    }
}
