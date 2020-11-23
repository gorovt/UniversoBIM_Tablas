using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UniBimFamilyTools
{
    public partial class frmWarning : Form
    {
        public frmWarning(string titulo, string descripcion, string mensaje)
        {
            InitializeComponent();
            this.Text = titulo;
            this.lblDescripcion.Text = descripcion;
            this.txtMensaje.Text = mensaje;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
